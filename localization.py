"""Presentation-only translations. Source records and email bodies stay unchanged."""
import tkinter as tk

language = 'en'

EN = {
    '个人研究确认':'Personal research confirmation', '尚未设置个人邮箱':'No sending account configured',
    '导入表格':'Import', '自动查找邮箱':'Find emails', '导出清单':'Export',
    '个人邮箱设置':'Account settings', '检查 Outlook':'Check Outlook', '检查发件连接':'Check sending connection', '停止查找':'Stop lookup',
    '搜索':'Search', '邮件工作区':'Mail workspace', '全部记录':'All records',
    '待补充':'Needs information', '可发送':'Ready to send', '已存草稿':'Draft saved',
    '已提交':'Submitted', '结果待核查':'Check result', '已有关系':'Already linked',
    '服务器已接受':'Accepted by server', '已交给 Outlook':'Submitted to Outlook',
    '以个人身份联系\n使用自己的邮箱\n回复在 Outlook 查看':'Personal correspondence\nYour own email account\nRead replies in Outlook',
    '以个人身份联系\n使用自己的邮箱\n回复在邮箱中查看':'Personal correspondence\nYour own email account\nRead replies in your mailbox',
    '研究关系':'Research relationships', '研究人员':'Researcher', '研究成果':'Research output', '状态':'Status',
    'Ctrl / Shift 多选记录；发送仅处理选中项。':'Use Ctrl / Shift to select multiple records. Only selected records are processed.',
    '⌘ / Shift 多选记录；发送仅处理选中项。':'Use Command / Shift to select multiple records. Only selected records are processed.',
    '试发给自己':'Test to myself', '生成草稿':'Create drafts', '发送选中项':'Send selected',
    '收件人和研究信息':'Record details', '邮件预览':'Email preview', '邮箱来源与候选':'Contact sources',
    '项目网站上的公开联系人':'Public contacts from the project website',
    '优先同一论文作者，再考虑项目负责人和其他项目成员。\n仅表示公开联系信息；研究关系仍需核实。':'Publication authors are preferred, followed by project leads and other members.\nA public contact does not confirm the research relationship.',
    '姓名':'Name', '公开邮箱':'Public email', '来源依据':'Match basis',
    '使用选中的联系人':'Use selected contact', '选择一条记录，点击“自动查找邮箱”。':'Select a record and click “Find emails”.',
    '打开项目网站':'Open project page', '打开邮箱来源':'Open email source',
    '研究人员姓名':'Researcher name', '收件邮箱':'Recipient email', '研究成果标题':'Output title',
    '项目名称':'Project name', '资助编号':'Grant ID', '发表日期':'Publication date', '资助机构':'Funder',
    '项目开始日期':'Project start date', '项目结束日期':'Project end date', 'RP+ 成果链接':'RP+ output URL',
    '推荐理由（英文，需核实）':'Suggested reason (English)',
    '我已核对收件人、成果和推荐理由':'I have checked the recipient, output and supporting evidence',
    '保存修改并更新预览':'Save and update preview', '选择左侧记录开始编辑':'Select a record to edit',
    '就绪 · 尚未发送任何邮件':'Ready · No emails sent in this session', '未填写':'Not provided',
    '此条目已具备发送信息':'This record has the required information', '尚未查找':'Not looked up yet',
    '暂无来源':'No source available', '该条记录尚无此来源链接。':'This record does not have a source link yet.',
    '链接无效':'Invalid link', '无需查找':'Nothing to look up',
    '没有选中记录，也没有缺少邮箱的记录。':'No records are selected and no email addresses are missing.',
    '批量读取公开网页':'Look up public contacts',
    '将查找 ':'Records to look up: ',
    ' 条记录。软件会打开浏览器读取项目和个人主页，并缓存重复页面。':'\nA browser will open project and profile pages and cache repeated pages.',
    '正在打开浏览器查找公开邮箱… 无需邮箱登录，不会发送邮件。':'Opening a browser to find public emails… No mailbox login or sending is involved.',
    ' 条记录。软件会打开 Edge 读取项目和个人主页，并缓存重复页面。':'\nEdge will open project and profile pages and cache repeated pages.',
    '可随时点击“停止查找”，已完成的结果会保留。继续？':'You can stop the lookup; completed results will be retained. Continue?',
    '正在打开 Edge 查找公开邮箱… 无需邮箱登录，不会发送邮件。':'Opening Edge to find public emails… No mailbox login or sending is involved.',
    '正在查找 ':'Looking up ', '此条查找失败，可重试':'Lookup failed for this record; retry is available',
    '邮箱查找已停止：已完成 ':'Lookup stopped. Completed records: ', ' 条，结果已保留。':'. Results have been saved.',
    '邮箱查找完成：处理 ':'Lookup complete. Records processed: ', '邮箱查找已停止：处理 ':'Lookup stopped. Records processed: ',
    ' 条，其中 ':'. Failed records: ', ' 条读取失败。请查看“邮箱来源与候选”，核对后再发送。':'. Review “Contact sources” before sending.',
    '邮箱查找中断，已完成 ':'Lookup interrupted. Completed records: ', ' 条并保留结果。':'. Completed results have been saved.',
    '详细记录：':'Detailed log: ', '导入新的清单':'Import a new list',
    '新清单会替换当前工作清单，发送历史会保留。是否继续？\n需要保留旧清单时请先导出。':'This will replace the working list and retain sending history. Continue?\nExport the current list first if you want to keep a separate copy.',
    '已读取工作表：':'Worksheet loaded: ', '已按 UTF-8 读取':'Read as UTF-8',
    '源文件不是 UTF-8，已按 Windows-1252 读取；请核对标题乱码':'Read as Windows-1252 because the source is not UTF-8; check titles for encoding errors',
    '缺少原始字段：':'Missing source columns: ', ' 条记录 · ':' records · ', '原文件未修改':'Source file unchanged',
    '无法导入':'Import failed', '已导出清单（CSV UTF-8）':'List exported as CSV UTF-8',
    '使用 Outlook 中已登录的个人邮箱':'Use your personal account in Outlook',
    '支持 Windows 经典版 Outlook。不会索取或保存邮箱密码。':'Requires classic Outlook for Windows. No email password is requested or stored.',
    '选择发信方式并配置已获准使用的个人邮箱':'Choose a sending method and configure your authorized account',
    'Windows 可使用经典版 Outlook；macOS 和 Windows 均可使用 SMTP。软件不保存邮箱密码。':'Classic Outlook is available on Windows. SMTP works on macOS and Windows. No password is stored.',
    '发信方式':'Sending method', 'SMTP 服务器':'SMTP host', 'SMTP 端口':'SMTP port',
    'SMTP 登录邮箱':'SMTP login address', '连接安全':'Connection security', '登录方式':'Sign-in method',
    'Microsoft 租户 ID':'Microsoft tenant ID', 'Microsoft 应用 ID':'Microsoft application ID',
    '密码环境变量名称':'Password environment variable name',
    '你的姓名（英文署名）':'Your name (English signature)', '个人发件邮箱':'Personal sending email', '回复邮箱（可留空）':'Reply-to email (optional)',
    '保存个人邮箱':'Save account', '请检查填写内容':'Check your settings', '先设置个人邮箱':'Configure your sending account first',
    '选择记录':'Select records', '请先选择左侧需要处理的记录。':'Select the records you want to process from the list.',
    '没有可处理的记录':'No eligible records',
    '请补齐信息并勾选已核对。已生成草稿或已提交的记录不会重复发送。':'Complete the required fields and mark the records as reviewed. Existing drafts and submissions will not be processed again.',
    '发送个人确认邮件':'Send confirmation emails', '将从 ':'From: ', ' 发送 ':'\nEmails to send: ', ' 封邮件。':'',
    '请先在右侧预览内容。现在发送？':'Review the email preview before proceeding. Send now?', '……共 ':'… Total: ', ' 封':' emails',
    '保存到 Outlook 草稿箱':'Save Outlook drafts', 'SMTP 不支持 Outlook 草稿；可以先试发给自己。':'SMTP cannot create Outlook drafts. Send a test to yourself first.', '将生成 ':'Drafts to create: ',
    ' 封草稿，不会发送。':'\nNo emails will be sent.',
    '草稿随后需在 Outlook 中手动发送；软件不会再自动发送这些记录。':'Send these drafts manually in Outlook. The application will skip these records in subsequent automatic sends.',
    '是否继续？':'Continue?', '正在处理… 请保持 Outlook 打开，等待结果。':'Processing… Keep Outlook open and wait for the result.',
    '正在处理… 请保持应用打开，等待结果。':'Processing… Keep the application open until the result appears.',
    '已连接经典版 Outlook，并找到指定个人邮箱。未创建或发送邮件。':'Connected to classic Outlook and matched the sending account. No emails were created or sent.',
    'SMTP 登录与连接成功。未发送邮件。':'SMTP sign-in and connection succeeded. No email was sent.',
    '已将 1 封测试邮件交给 Outlook，收件人为你自己的邮箱。':'One test email was submitted to Outlook, addressed to your own account.',
    '已将 1 封测试邮件交给发件服务，收件人为你自己的邮箱。':'One test email was submitted to the sending service, addressed to your own account.',
    '草稿处理完成，请在 Outlook 草稿箱中检查并手动发送。':'Draft processing finished. Review and send the drafts manually in Outlook.',
    '选中邮件处理完成，请在 Outlook 中检查发件箱、已发送和退信。':'Processing finished. Check the Outbox, Sent Items and bounce messages in Outlook.',
    '选中邮件处理完成，请在邮箱中检查已发送、退信和回复。':'Processing finished. Check Sent Items, bounce messages and replies in your mailbox.',
    'Microsoft 登录':'Microsoft sign-in', '取消登录':'Cancel sign-in',
    '等待 Microsoft 登录；可点击“取消登录”。':'Waiting for Microsoft sign-in. You can select “Cancel sign-in”.',
    '正在取消 Microsoft 登录…':'Cancelling Microsoft sign-in…',
    'Microsoft 登录已取消，未发送邮件。':'Microsoft sign-in cancelled. No email was sent.',
    '请在打开的网页输入代码 ':'Enter code ',
    ' 完成登录。\n\n登录网址：':' in the opened page to sign in.\n\nSign-in URL: ',
    '操作未完成':'Operation incomplete', '若已开始处理邮件，请先检查 Outlook，再决定是否重试。':'If email processing started, check Outlook before retrying.',
    '若已开始处理邮件，请先检查邮箱的已发送、发件箱和退信，再决定是否重试。':'If email processing started, check Sent Items, the Outbox and bounce messages before retrying.',
    '处理结果':'Result', '正在处理':'Processing', '请等待本批操作结束后关闭，避免发送结果不确定。':'Wait for the current operation to finish before closing the application.',
    '正在读取：':'Reading: ', '页面超时，重试一次：':'Page timed out; retrying once: ',
    '只读取 RP+ 官方 HTTPS 网页，不访问表格中的其他网站。':'Only official RP+ HTTPS pages are supported.',
    '网页地址不是 RP+ 项目、人员或成果页面。':'The URL must be an RP+ project, person or publication page.',
    '无法打开 Microsoft Edge。请确认本机已安装 Edge。':'Unable to open Microsoft Edge. Check that Edge is installed.',
    '具体原因：':'Details: ', '网站拒绝访问或限流（HTTP ':'Access denied or rate limited (HTTP ',
    '）。':').', '页面：':'Page: ', '请在浏览器中确认页面能正常访问；软件不会绕过网站验证。':'Check whether the page opens normally in your browser. The application does not bypass verification.',
    '网站显示浏览器验证页面，未自动处理验证。':'The website is requesting browser verification. No verification was completed automatically.',
    '网页读取超时：':'Page timed out: ', '网页读取失败：':'Page could not be read: ',
    '查找用的 Edge 窗口已关闭。请重新点击“自动查找邮箱”，查找期间保持窗口打开。':'The lookup Edge window was closed. Click “Find emails” again and keep that window open during lookup.',
    '该行缺少 project_url 项目网站地址。':'This record is missing its project_url.',
    '项目成员没有公开主页链接':'Project members have no linked public profiles',
    '成果作者页未读取；仅匹配项目成员：':'Publication authors could not be checked; only project membership was used: ',
    ' 主页未读取：':' profile could not be read: ', '项目成员且为该成果作者':'Project member and publication author',
    '项目页列出的成员（尚未确认是成果作者）':'Project member; publication authorship unconfirmed',
    '已找到公开邮箱':'Public email found', '成员主页未公开邮箱':'No public email on member profiles',
    '部分网页未读取；已保留可用结果':'Some pages failed; available results retained',
    '网页读取不完整，需重试':'Incomplete lookup; retry needed', '已找到候选；保留现有收件人':'Candidates found; existing recipient preserved',
    '语言':'Language', '操作结束后才能切换语言。':'Wait for the operation to finish before switching languages.',
}

ZH_ERRORS = {
    'reviewed must be yes after verifying recipient and evidence':'请核对收件人和证据，并勾选已核对',
    'Relationship already exists or status is unknown':'关系已存在或状态不明',
    'One valid recipient email is required':'请填写一个有效的收件邮箱',
    'Unexpected RP+ output URL':'RP+ 成果链接无效',
    'Missing ':'缺少字段：', 'Configure ':'请配置：', ' in config.json':'',
    'Invalid sender_email or reply_to':'发件邮箱或回复邮箱无效',
    'sender_email does not match an Outlook sending account. No default account will be substituted.':'发件邮箱与 Outlook 中的账号不匹配，未改用默认账号。',
    'Cannot access classic Outlook. Open classic Outlook with a configured profile; new Outlook does not support this connection.':'无法连接经典版 Outlook。请打开已配置账号的经典版 Outlook；新版不支持此连接。',
    'Outlook could not resolve the recipient or reply address':'Outlook 无法解析收件人或回复地址',
    'transport must be outlook or smtp':'发信方式必须是 Outlook 或 SMTP',
    'smtp_port must be between 1 and 65535':'SMTP 端口必须介于 1 至 65535',
    'Set the configured password environment variable':'请设置已配置的密码环境变量',
    'Operation failed or result is uncertain. Stopped; inspect Outlook/mail server before any manual retry.':'操作失败或结果不确定，已停止。重试前请检查 Outlook 或邮件服务器。',
}

def set_language(value):
    global language
    language = value if value in ('en','zh') else 'en'

def tr(text):
    text=str(text)
    mapping=EN if language=='en' else ZH_ERRORS
    # Replace in one pass so a translated value cannot be translated again.
    import re
    pattern='|'.join(re.escape(k) for k in sorted(mapping,key=len,reverse=True))
    return re.sub(pattern,lambda m:mapping[m.group()],text)

class DisplayVar(tk.StringVar):
    def __init__(self,master=None,value='',**kwargs):
        self.raw=value
        super().__init__(master=master,value=tr(value),**kwargs)
    def set(self,value):
        self.raw=value
        super().set(tr(value))
