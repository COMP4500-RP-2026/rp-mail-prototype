"""Read public RP+ project participants and their published profile emails."""
import json
import re
import time
from collections import OrderedDict
from datetime import datetime, timezone
from urllib.parse import urljoin, urlsplit, urlunsplit, unquote
from bs4 import BeautifulSoup

HOST = 'researchportalplus.anu.edu.au'
FIELDS = ['project_url', 'email_source_url', 'email_lookup_status', 'email_lookup_time',
          'email_candidates', 'contact_basis', 'email_lookup_warnings']
EMAIL = re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

def canonical(url, base='https://' + HOST):
    parsed = urlsplit(urljoin(base, url))
    if parsed.scheme != 'https' or parsed.hostname != HOST or parsed.username or parsed.password or parsed.port not in (None,443):
        raise ValueError('只读取 RP+ 官方 HTTPS 网页，不访问表格中的其他网站。')
    if not any(parsed.path.startswith('/en/'+p+'/') for p in ('projects','persons','publications')):
        raise ValueError('网页地址不是 RP+ 项目、人员或成果页面。')
    return urlunsplit(('https', HOST, parsed.path.rstrip('/')+'/', '', ''))

def soup_of(html):
    return BeautifulSoup(html, 'html.parser')

def people(html, url):
    soup=soup_of(html)
    intro=soup.select_one('.introduction')
    if not intro:return []
    result=[];seen=set()
    # Only the primary project's member list / publication's author list,
    # never related outputs, page footer or global navigation.
    for anchor in intro.select('a[href]'):
        href=urljoin(url,anchor.get('href',''))
        if '/en/persons/' not in href:continue
        try:profile=canonical(href)
        except ValueError:continue
        if profile in seen:continue
        name=anchor.get_text(' ',strip=True)
        if not name:continue
        context=anchor.find_parent('li')
        role=context.get_text(' ',strip=True) if context else ''
        seen.add(profile)
        result.append({'name':name,'profile_url':profile,
                       'principal':bool(re.search(r'\((?:PI|CI)\)',role,re.I))})
    return result

def project_info(html):
    soup=soup_of(html)
    heading=soup.select_one('.introduction h1')
    return {'project_name':heading.get_text(' ',strip=True)} if heading else {}

def profile_emails(html):
    soup=soup_of(html)
    section=soup.select_one('section.profile')
    if section is None:return []
    emails=[]
    # Read contact controls only, not email-looking text in publications/bios.
    for anchor in section.select('a.email, a[href^="mailto:"]'):
        for script in anchor.select('script,style'):script.decompose()
        href=anchor.get('href','')
        texts=[anchor.get_text('',strip=True)]
        if href.lower().startswith('mailto:'):texts.append(unquote(href[7:].split('?')[0]))
        for text in texts:
            for email in EMAIL.findall(text):
                if email.lower() not in {e.lower() for e in emails}:emails.append(email)
    return emails

def display_name(name):
    if name.count(',')==1:
        last,first=name.split(',',1)
        return first.strip()+' '+last.strip()
    return name

class BrowserFetcher:
    """Use an ordinary visible browser; no saved credentials or bypasses."""
    def __init__(self,stop=None,report=None):
        self.stop=stop;self.report=report or (lambda _:None)

    def __enter__(self):
        from playwright.sync_api import sync_playwright
        self.runtime=sync_playwright().start()
        try:
            errors=[]
            for channel in ('msedge','chrome',None):
                try:
                    options={'headless':False}
                    if channel:options['channel']=channel
                    self.browser=self.runtime.chromium.launch(**options)
                    break
                except Exception as exc:
                    errors.append(f'{channel or "Playwright Chromium"}: {exc}')
            else:
                raise LookupFatal('无法打开浏览器。请安装 Microsoft Edge 或 Google Chrome，或运行 python -m playwright install chromium。\n具体原因：'+errors[-1])
            self.context=self.browser.new_context()
            self.page=self.context.new_page()
        except Exception as exc:
            self.runtime.stop()
            if isinstance(exc,LookupFatal):raise
            raise LookupFatal(f'无法打开浏览器。请检查 Edge、Chrome 或 Playwright Chromium。\n具体原因：{exc}') from exc
        self.cache=OrderedDict();self.last_request=0
        return self

    def __call__(self,url):
        if self.stop and self.stop.is_set():raise LookupStopped()
        url=canonical(url)
        if url in self.cache:return self.cache[url]
        time.sleep(max(0,1.2-(time.monotonic()-self.last_request)))
        self.last_request=time.monotonic()
        self.report('正在读取：'+url)
        from playwright.sync_api import TimeoutError as BrowserTimeout
        for attempt in range(2):
            try:
                response=self.page.goto(url,wait_until='domcontentloaded',timeout=45000)
                if response and response.status in (401,403,429):
                    raise LookupFatal(f'网站拒绝访问或限流（HTTP {response.status}）。\n页面：{url}\n请在浏览器中确认页面能正常访问；软件不会绕过网站验证。')
                if response and response.status>=400:
                    raise PageUnavailable(f'HTTP {response.status}：{url}')
                # Wait for the actual RP+ content, not just a heading in a cookie/challenge page.
                selector='section.profile h1' if '/en/persons/' in url else '.introduction h1'
                self.page.locator(selector).first.wait_for(state='visible',timeout=25000)
                final=canonical(self.page.url)
                html=self.page.content()
                break
            except BrowserTimeout as exc:
                title=self.page.title().lower() if not self.page.is_closed() else ''
                if any(x in title for x in ('just a moment','access denied','verify','captcha')):
                    raise LookupFatal(f'网站显示浏览器验证页面，未自动处理验证。\n页面：{url}') from exc
                if self.stop and self.stop.is_set():raise LookupStopped()
                if attempt==0:
                    self.report('页面超时，重试一次：'+url)
                    continue
                raise PageUnavailable(f'网页读取超时：{url}') from exc
            except (LookupFatal,PageUnavailable,LookupStopped):raise
            except Exception as exc:
                if self.page.is_closed() or not self.browser.is_connected():
                    raise LookupFatal('查找用的浏览器窗口已关闭。请重新点击“自动查找邮箱”，查找期间保持窗口打开。') from exc
                raise PageUnavailable(f'网页读取失败：{url}\n{exc}') from exc
        self.cache[url]=html;self.cache[final]=html
        while len(self.cache)>256:self.cache.popitem(last=False)
        return html

    def __exit__(self,*args):
        try:self.browser.close()
        finally:self.runtime.stop()

class LookupStopped(Exception):
    pass

class LookupFatal(RuntimeError):
    pass

class PageUnavailable(RuntimeError):
    pass

def lookup(row,fetch):
    project=canonical(row.get('project_url','')) if row.get('project_url') else None
    if not project:raise ValueError('该行缺少 project_url 项目网站地址。')
    html=fetch(project)
    members=people(html,project)
    result=project_info(html)
    result['email_lookup_time']=datetime.now(timezone.utc).isoformat()
    candidates=[]
    if not members:
        result.update(email_lookup_status='项目成员没有公开主页链接',email_candidates='[]')
        return result
    authors=set();warnings=[]
    output=row.get('output_url')
    if output:
        try:authors={p['profile_url'] for p in people(fetch(canonical(output)),output)}
        except (PageUnavailable,ValueError) as exc:
            warnings.append('成果作者页未读取；仅匹配项目成员：'+str(exc))
    for member in members:
        try:member_html=fetch(member['profile_url'])
        except PageUnavailable as exc:
            warnings.append(member['name']+' 主页未读取：'+str(exc));continue
        for email in profile_emails(member_html):
            shared=member['profile_url'] in authors
            candidates.append(dict(name=display_name(member['name']),email=email,
                source_url=member['profile_url'],project_url=project,
                basis='项目成员且为该成果作者' if shared else '项目页列出的成员（尚未确认是成果作者）',
                rank=(2 if shared else 0)+(1 if member['principal'] else 0)))
    candidates.sort(key=lambda c:-c['rank'])
    result['email_candidates']=json.dumps(candidates,ensure_ascii=False)
    result['email_lookup_status']='已找到公开邮箱' if candidates else '成员主页未公开邮箱'
    result['email_lookup_warnings']='\n'.join(warnings)
    if warnings:result['email_lookup_status']='部分网页未读取；已保留可用结果' if candidates else '网页读取不完整，需重试'
    if candidates:
        chosen=candidates[0]
        result.update(researcher_name=chosen['name'],researcher_email=chosen['email'],
                      email_source_url=chosen['source_url'],contact_basis=chosen['basis'])
        if chosen['rank']>=2 and not row.get('suggestion_reason'):
            result['suggestion_reason']=f"{chosen['name']} is listed both as a participant on the RP+ project page and as an author on the RP+ research output page. This is a possible connection, not confirmation of project funding."
    return result

def merge_result(row,result):
    """Fill blanks; preserve manually supplied identity and verified data."""
    row=dict(row)
    identity_present=bool(row.get('researcher_email') or row.get('researcher_name'))
    for field in ['email_lookup_status','email_lookup_time','email_candidates','email_lookup_warnings']:
        if field in result:row[field]=result[field]
    if identity_present:
        if result.get('researcher_email'):row['email_lookup_status']='已找到候选；保留现有收件人'
    else:
        for field in ['researcher_email','researcher_name','email_source_url','contact_basis']:
            if result.get(field):row[field]=result[field]
        if result.get('researcher_email'):row['reviewed']=''
    for field in ['project_name','suggestion_reason']:
        if not row.get(field) and result.get(field):
            row[field]=result[field];row['reviewed']=''
    return row
