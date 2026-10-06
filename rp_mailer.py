"""RP+ confirmation mailer. Python 3.10+; Outlook or cross-platform SMTP."""
import argparse
from contextlib import closing
import csv
import hashlib
import json
import os
from pathlib import Path
import re
import smtplib
import sqlite3
import ssl
import time
from datetime import datetime, timezone
from email.message import EmailMessage
from email.policy import SMTP
from email.utils import formataddr, make_msgid, formatdate

BASE = Path(__file__).resolve().parent
EXTRA = ['researcher_name', 'researcher_email', 'publication_date', 'project_name',
         'funder', 'project_start_date', 'project_end_date', 'suggestion_reason', 'reviewed']
SUBJECT = 'Potential RP+ project–output relationship'
SMTP_SCOPE = ['https://outlook.office.com/SMTP.Send']
_MSAL_APPS = {}

class SignInCancelled(RuntimeError):
    pass

def read_rows(path, encoding='utf-8-sig'):
    with open(path, encoding=encoding, newline='') as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or not {'output_uuid', 'project_uuid', 'title',
                'output_url', 'project_code', 'already_in_relations'} <= set(reader.fieldnames):
            raise ValueError('Input is missing required source columns.')
        fields = [k for k in reader.fieldnames if k]
        return fields, [{k: (r.get(k) or '').strip() for k in fields} for r in reader]

def prepare(source, target, encoding='utf-8-sig'):
    fields, rows = read_rows(source, encoding)
    fields += [k for k in EXTRA if k not in fields]
    with open(target, 'x', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f'Created {target}: {len(rows)} rows. Fill contact details and verified evidence.')

def address(value):
    return bool(re.fullmatch(r'[^\s<>@,;]+@[^\s<>@,;]+\.[^\s<>@,;]+', value))

def key_for(row):
    value = '\0'.join([row['output_uuid'], row['project_uuid'], row['researcher_email'].lower()])
    return hashlib.sha256(value.encode()).hexdigest()

def validate(row):
    if row.get('already_in_relations', '').lower() not in ('false', '0', 'no'):
        return 'Relationship already exists or status is unknown'
    if row.get('reviewed', '').lower() != 'yes':
        return 'reviewed must be yes after verifying recipient and evidence'
    for field in ['researcher_name', 'researcher_email', 'output_uuid', 'project_uuid',
                  'title', 'output_url', 'project_code', 'project_name', 'suggestion_reason']:
        if not row.get(field):
            return f'Missing {field}'
    if not address(row['researcher_email']):
        return 'One valid recipient email is required'
    if not row['output_url'].startswith('https://researchportalplus.anu.edu.au/'):
        return 'Unexpected RP+ output URL'
    return ''

def validate_smtp_port(config):
    try:
        port = int(config.get('smtp_port', ''))
    except (TypeError, ValueError):
        raise ValueError('smtp_port must be between 1 and 65535')
    if not 1 <= port <= 65535:
        raise ValueError('smtp_port must be between 1 and 65535')

def validate_config(config, sending=False):
    for field in ['sender_email', 'sender_name']:
        if not config.get(field) or '\n' in config[field] or '\r' in config[field]:
            raise ValueError(f'Configure {field} in config.json')
    if not address(config['sender_email']) or (config.get('reply_to') and not address(config['reply_to'])):
        raise ValueError('Invalid sender_email or reply_to')
    if config.get('transport', 'smtp') not in ('outlook', 'smtp'):
        raise ValueError('transport must be outlook or smtp')
    if sending:
        if float(config.get('delay_seconds', 2)) < 0:
            raise ValueError('delay_seconds cannot be negative')
        if config.get('transport') == 'outlook':
            return
        if not config.get('smtp_host') or not config.get('smtp_username'):
            raise ValueError('Configure smtp_host and smtp_username')
        validate_smtp_port(config)
        if config.get('smtp_security') not in ('starttls', 'ssl'):
            raise ValueError('smtp_security must be starttls or ssl')
        if config.get('auth_mode') == 'microsoft':
            for field in ['tenant_id', 'client_id']:
                if not config.get(field):
                    raise ValueError(f'Ask institution IT to configure {field}')
        elif config.get('auth_mode') != 'password':
            raise ValueError('auth_mode must be microsoft or password')
        elif not config.get('password_env') or not os.environ.get(config['password_env']):
            raise ValueError('Set the configured password environment variable')

def message(row, config):
    ref = key_for(row)[:16]
    def val(k):
        return row.get(k) or 'Not available in the current dataset'
    body = f'''Dear {row['researcher_name']},

I am reviewing possible relationships between research projects and research outputs listed in Research Portal+ (RP+). I am contacting you in a personal capacity.

Based on the information currently available, I identified the following possible relationship that may be associated with your research. Please first confirm whether this is your research output.

Research output
- Title: {row['title']}
- Publication date: {val('publication_date')}
- RP+ record: {row['output_url']}

Potentially related project
- Project: {row['project_name']}
- Funder and grant ID: {val('funder')} — {row['project_code']}
- Project period: {val('project_start_date')} – {val('project_end_date')}

Why this relationship was suggested
{row['suggestion_reason']}

Could you please reply with one of the following responses?

1. Confirmed — this is my research output and it is related to this project.
2. Not related — this is my research output, but the suggested project relationship is incorrect.
3. Unsure — more information is required.
4. Correction — the output is related to a different project. Please provide the project name or any details you remember.
5. Not my research output — this output is not associated with me.

You are also welcome to include any additional context in your reply.

This email does not make any changes to RP+ records. I will review your response to check my data; any update to RP+ would need to follow the appropriate process.

Thank you for your time and assistance.

Kind regards,
{config['sender_name']}

Reference: {ref}
'''
    msg = EmailMessage(policy=SMTP)
    msg['Subject'] = f'{SUBJECT} [RP-{ref}]'
    msg['From'] = formataddr((config['sender_name'], config['sender_email']))
    msg['To'] = row['researcher_email']
    msg['Reply-To'] = config.get('reply_to') or config['sender_email']
    msg['Date'] = formatdate(localtime=True)
    msg['Message-ID'] = make_msgid(domain=config['sender_email'].split('@')[1])
    msg.set_content(body)
    return msg

class OutlookTransport:
    def __init__(self, config):
        try:
            import win32com.client
        except ImportError:
            raise ValueError('Install Outlook dependency: python -m pip install pywin32')
        try:
            self.app = win32com.client.gencache.EnsureDispatch('Outlook.Application')
            self.session = self.app.Session
            self.account = next((a for a in self.session.Accounts
                                 if str(a.SmtpAddress).lower() == config['sender_email'].lower()), None)
        except Exception as exc:
            raise RuntimeError('Cannot access classic Outlook. Open classic Outlook with a configured profile; new Outlook does not support this connection.') from exc
        if self.account is None:
            raise ValueError('sender_email does not match an Outlook sending account. No default account will be substituted.')
        self.sender = config['sender_email']

    def send_message(self, msg, from_addr, to_addrs, draft=False):
        if from_addr.lower() != self.sender.lower() or len(to_addrs) != 1:
            raise ValueError('Unexpected sender or recipient count')
        # Create in the selected account store, including for secondary accounts.
        folder = self.account.DeliveryStore.GetDefaultFolder(16)  # olFolderDrafts
        mail = folder.Items.Add('IPM.Note')
        mail.SendUsingAccount = self.account
        mail.To = to_addrs[0]
        mail.Subject = str(msg['Subject'])
        mail.Body = msg.get_content()
        mail.ReplyRecipients.Add(str(msg['Reply-To']))
        if not mail.Recipients.ResolveAll() or not mail.ReplyRecipients.ResolveAll():
            raise ValueError('Outlook could not resolve the recipient or reply address')
        if draft:
            mail.Save()
        else:
            mail.Send()

    def close(self):
        # Do not quit the user's Outlook application.
        pass

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

def connect(config, on_auth=None, stop_event=None, on_auth_done=None):
    if config.get('transport') == 'outlook':
        return OutlookTransport(config)
    def stopped():
        return stop_event is not None and stop_event.is_set()
    if stopped():
        raise SignInCancelled('Microsoft sign-in cancelled. No email was sent.')
    token = None
    if config.get('auth_mode') == 'microsoft':
        try:
            import msal
        except ImportError:
            raise ValueError('Install Microsoft login dependency: python -m pip install msal')
        key = (config['tenant_id'], config['client_id'])
        app = _MSAL_APPS.get(key)
        if app is None:
            app = msal.PublicClientApplication(config['client_id'],
                authority='https://login.microsoftonline.com/' + config['tenant_id'],
                token_cache=msal.SerializableTokenCache())
            _MSAL_APPS[key] = app
        result = None
        for account in app.get_accounts(username=config['smtp_username']):
            result = app.acquire_token_silent(SMTP_SCOPE, account=account)
            if result and result.get('access_token'):
                break
        if not result or not result.get('access_token'):
            if stopped():
                raise SignInCancelled('Microsoft sign-in cancelled. No email was sent.')
            flow = app.initiate_device_flow(scopes=SMTP_SCOPE)
            if 'user_code' not in flow:
                raise ValueError('Microsoft sign-in could not start. Check application configuration with IT.')
            if on_auth:
                on_auth(flow['verification_uri'], flow['user_code'])
            else:
                print(flow['message'], flush=True)
            result = app.acquire_token_by_device_flow(flow,
                exit_condition=lambda current: stopped() or current.get('expires_at', 0) < time.time())
        if stopped():
            raise SignInCancelled('Microsoft sign-in cancelled. No email was sent.')
        token = result.get('access_token')
        if not token:
            raise ValueError('Microsoft sign-in failed. Check consent and school access policy.')
        if on_auth_done:
            on_auth_done()
    context = ssl.create_default_context()
    if config['smtp_security'] == 'ssl':
        smtp = smtplib.SMTP_SSL(config['smtp_host'], int(config['smtp_port']), timeout=30, context=context)
    else:
        smtp = smtplib.SMTP(config['smtp_host'], int(config['smtp_port']), timeout=30)
        smtp.ehlo()
        smtp.starttls(context=context)
        smtp.ehlo()
    try:
        if config.get('auth_mode') == 'microsoft':
            auth = f'user={config["smtp_username"]}\x01auth=Bearer {token}\x01\x01'
            smtp.auth('XOAUTH2', lambda challenge=None: auth if challenge is None else '')
        else:
            smtp.login(config['smtp_username'], os.environ[config['password_env']])
    except Exception:
        smtp.close()
        raise
    return smtp

def run(args, on_auth=None, stop_event=None, on_auth_done=None):
    _, rows = read_rows(args.input)
    config = json.loads(Path(args.config).read_text(encoding='utf-8-sig'))
    validate_config(config, args.command in ('send', 'test', 'draft'))
    if args.command == 'draft' and config.get('transport') != 'outlook':
        raise ValueError('draft requires transport=outlook')
    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    eligible, report, seen = [], [], set()
    for number, row in enumerate(rows, 2):
        reason = validate(row)
        key = key_for(row) if not reason else ''
        if key and key in seen:
            reason = 'Duplicate output/project/recipient'
        if key:
            seen.add(key)
        report.append({'csv_row': number, 'status': reason or 'Ready', 'reference': key[:16]})
        if not reason:
            eligible.append(row)
    with (out / 'validation.csv').open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['csv_row', 'status', 'reference'])
        writer.writeheader()
        writer.writerows(report)
    print(f'{len(eligible)} eligible; {len(rows)-len(eligible)} skipped. See {out / "validation.csv"}')
    if args.command == 'preview':
        for row in eligible:
            msg = message(row, config)
            (out / f'{key_for(row)[:16]}.eml').write_bytes(bytes(msg))
        print('Preview only. No emails sent. Old preview files are not automatically removed.')
        return
    if not eligible:
        return
    if args.command == 'test':
        if not address(args.to):
            raise ValueError('Provide a valid test recipient with --to')
        msg = message(eligible[0], config)
        msg.replace_header('To', args.to)
        msg.replace_header('Subject', '[TEST] ' + str(msg['Subject']))
        with connect(config, on_auth=on_auth, stop_event=stop_event, on_auth_done=on_auth_done) as smtp:
            if stop_event is not None and stop_event.is_set():
                raise SignInCancelled('Microsoft sign-in cancelled. No email was sent.')
            smtp.send_message(msg, from_addr=config['sender_email'], to_addrs=[args.to])
        print('One test message sent. Production history unchanged.')
        return
    if args.limit < 1:
        raise ValueError('--limit must be positive')
    # The fixed ledger is deliberately independent of the preview/output directory.
    with closing(sqlite3.connect(BASE / 'send_history.sqlite3', timeout=30)) as db:
        db.execute('CREATE TABLE IF NOT EXISTS history (id TEXT PRIMARY KEY, recipient TEXT, status TEXT, timestamp TEXT, message_id TEXT)')
        db.commit()
        smtp = None
        count = 0
        try:
            for row in eligible:
                if count >= args.limit:
                    break
                key = key_for(row)
                if db.execute('SELECT 1 FROM history WHERE id=?', (key,)).fetchone():
                    continue
                msg = message(row, config)
                if smtp is None:
                    smtp = connect(config, on_auth=on_auth, stop_event=stop_event, on_auth_done=on_auth_done)
                    if stop_event is not None and stop_event.is_set():
                        raise SignInCancelled('Microsoft sign-in cancelled. No email was sent.')
                try:
                    db.execute('INSERT INTO history VALUES (?, ?, ?, ?, ?)',
                               (key, row['researcher_email'], 'pending', datetime.now(timezone.utc).isoformat(), str(msg['Message-ID'])))
                    db.commit()
                except sqlite3.IntegrityError:
                    continue
                try:
                    kwargs = {'draft': True} if args.command == 'draft' else {}
                    smtp.send_message(msg, from_addr=config['sender_email'], to_addrs=[row['researcher_email']], **kwargs)
                except Exception:
                    db.execute('UPDATE history SET status=? WHERE id=?', ('uncertain', key))
                    db.commit()
                    raise RuntimeError('Operation failed or result is uncertain. Stopped; inspect Outlook/mail server before any manual retry.')
                status = 'drafted' if args.command == 'draft' else ('submitted_outlook' if config.get('transport') == 'outlook' else 'accepted')
                db.execute('UPDATE history SET status=? WHERE id=?', (status, key))
                db.commit()
                count += 1
                print(f'{status}: RP-{key[:16]}')
                time.sleep(float(config.get('delay_seconds', 2)))
        finally:
            if smtp is not None:
                smtp.close()
        if args.command == 'draft':
            print(f'{count} Outlook drafts created. Send them manually; saving a draft does not deliver it.')
        else:
            print(f'{count} messages submitted. Submission does not guarantee delivery.')

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    prep = sub.add_parser('prepare')
    prep.add_argument('source')
    prep.add_argument('target')
    prep.add_argument('--encoding', default='utf-8-sig')
    check = sub.add_parser('check-outlook')
    check.add_argument('--config', default=str(BASE / 'config.json'))
    for name in ('preview', 'test', 'send', 'draft'):
        p = sub.add_parser(name)
        p.add_argument('input')
        p.add_argument('--config', default=str(BASE / 'config.json'))
        p.add_argument('--output', default=str(BASE / 'preview'))
        if name == 'test':
            p.add_argument('--to', required=True)
        if name in ('send', 'draft'):
            p.add_argument('--limit', type=int, default=10)
    args = parser.parse_args()
    try:
        if args.command == 'check-outlook':
            config = json.loads(Path(args.config).read_text(encoding='utf-8-sig'))
            validate_config(config)
            with OutlookTransport(config):
                print('Classic Outlook connected; sender account matched. No messages created or sent.')
        elif args.command == 'prepare':
            prepare(args.source, args.target, args.encoding)
        else:
            run(args)
    except Exception as exc:
        parser.exit(1, f'Error: {exc}\n')

if __name__ == '__main__':
    main()
