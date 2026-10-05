"""Cross-platform paths use fakes; no browser, mailbox, or live website is accessed."""
import queue
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import desktop_app as desktop
import email_lookup
import rp_mailer as mailer

with patch.object(desktop.sys, 'frozen', True, create=True):
    with patch.object(desktop.sys, 'platform', 'darwin'):
        assert desktop.app_home() == Path.home() / 'Library' / 'Application Support' / 'RPMail'

config = dict(transport='smtp', sender_name='Example', sender_email='person@example.org',
              smtp_host='smtp.example.org', smtp_port='587', smtp_security='starttls',
              smtp_username='person@example.org', auth_mode='microsoft',
              tenant_id='tenant', client_id='client')
mailer.validate_config(config, True)
for invalid in ('0', 'bad', '65536'):
    try:
        mailer.validate_config(dict(config, smtp_port=invalid), True)
        raise AssertionError('Invalid SMTP port was accepted')
    except ValueError as exc:
        assert 'smtp_port' in str(exc)

events = []
oauth = MagicMock()
oauth.initiate_device_flow.return_value = {
    'verification_uri': 'https://microsoft.com/devicelogin', 'user_code': 'ABC123'}
oauth.acquire_token_by_device_flow.return_value = {'access_token': 'fake-token'}
smtp = MagicMock()
with patch.dict(sys.modules, {'msal': SimpleNamespace(PublicClientApplication=lambda *a, **k: oauth)}):
    with patch.object(mailer.smtplib, 'SMTP', side_effect=lambda *a, **k: events.append('smtp') or smtp):
        connection = mailer.connect(config, on_auth=lambda uri, code: events.append((uri, code)))
assert connection is smtp
assert events[0] == ('https://microsoft.com/devicelogin', 'ABC123')
assert events[1] == 'smtp'
assert 'Bearer fake-token' in smtp.auth.call_args.args[1](None)

browser = MagicMock()
runtime = MagicMock()
runtime.chromium.launch.side_effect = [RuntimeError('Edge unavailable'), browser]
with patch('playwright.sync_api.sync_playwright', return_value=SimpleNamespace(start=lambda: runtime)):
    with email_lookup.BrowserFetcher() as fetcher:
        assert fetcher.browser is browser
assert runtime.chromium.launch.call_args_list[0].kwargs['channel'] == 'msedge'
assert runtime.chromium.launch.call_args_list[1].kwargs['channel'] == 'chrome'

with tempfile.TemporaryDirectory() as directory:
    home = Path(directory)
    with patch.object(desktop, 'HOME', home), patch.object(desktop, 'DATA', home / 'personal_data'):
        app = desktop.App()
        expected_transport = 'outlook' if sys.platform == 'win32' else 'smtp'
        assert app.settings['transport'] == expected_transport
        assert str(app.draft_button['state']) == ('normal' if expected_transport == 'outlook' else 'disabled')
        app.account_settings()
        assert len(app.winfo_children()) > 0
        for child in app.winfo_children():
            if isinstance(child, desktop.tk.Toplevel):
                child.destroy()
        app.destroy()

fake = SimpleNamespace(events=queue.Queue(), auth_prompt=lambda *args: None)
with patch.object(desktop.core, 'connect', return_value=MagicMock()):
    desktop.App.worker(fake, 'check', [], config)
assert fake.events.get()[0] == 'ok'

print('Cross-platform configuration, OAuth prompt, browser fallback and macOS UI checks passed. No emails sent.')
