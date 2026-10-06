"""Cross-platform paths use fakes; no browser, mailbox, or live website is accessed."""
import queue
import sys
import tempfile
import threading
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
        mailer.validate_smtp_port(dict(config, smtp_port=invalid))
        raise AssertionError('Invalid SMTP port was accepted')
    except ValueError as exc:
        assert 'smtp_port' in str(exc)

events = []
oauth = MagicMock()
oauth.initiate_device_flow.return_value = {
    'verification_uri': 'https://microsoft.com/devicelogin', 'user_code': 'ABC123'}
oauth.acquire_token_by_device_flow.return_value = {'access_token': 'fake-token'}
oauth.get_accounts.side_effect = [[], [{'username': config['smtp_username']}]]
oauth.acquire_token_silent.return_value = {'access_token': 'silent-token'}
smtp = MagicMock()
msal = SimpleNamespace(SerializableTokenCache=MagicMock(), PublicClientApplication=MagicMock(return_value=oauth))
with patch.dict(sys.modules, {'msal': msal}), patch.dict(mailer._MSAL_APPS, {}, clear=True):
    with patch.object(mailer.smtplib, 'SMTP', side_effect=lambda *a, **k: events.append('smtp') or smtp):
        connection = mailer.connect(config, on_auth=lambda uri, code: events.append((uri, code)))
        second = mailer.connect(config, on_auth=lambda uri, code: events.append((uri, code)))
assert connection is smtp
assert second is smtp
assert events[0] == ('https://microsoft.com/devicelogin', 'ABC123')
assert events[1] == 'smtp'
assert events[2] == 'smtp'
assert oauth.initiate_device_flow.call_count == 1
assert msal.PublicClientApplication.call_count == 1
assert msal.PublicClientApplication.call_args.kwargs['token_cache'] is msal.SerializableTokenCache.return_value
oauth.acquire_token_silent.assert_called_once_with(mailer.SMTP_SCOPE, account={'username': config['smtp_username']})
assert 'Bearer silent-token' in smtp.auth.call_args.args[1](None)

stop = threading.Event()
oauth = MagicMock()
oauth.get_accounts.return_value = []
oauth.initiate_device_flow.return_value = {
    'verification_uri': 'https://microsoft.com/devicelogin', 'user_code': 'ABC123',
    'expires_at': 9999999999, 'interval': 5}
def cancelled_result(flow, **kwargs):
    assert kwargs['exit_condition'](flow)
    return {'error': 'authorization_pending'}
oauth.acquire_token_by_device_flow.side_effect = cancelled_result
msal = SimpleNamespace(SerializableTokenCache=MagicMock(), PublicClientApplication=MagicMock(return_value=oauth))
with patch.dict(sys.modules, {'msal': msal}), patch.dict(mailer._MSAL_APPS, {}, clear=True):
    with patch.object(mailer.smtplib, 'SMTP') as smtp_factory:
        try:
            mailer.connect(config, on_auth=lambda *_: stop.set(), stop_event=stop)
            raise AssertionError('Cancelled sign-in continued')
        except mailer.SignInCancelled:
            pass
        smtp_factory.assert_not_called()

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
        app.account_settings()
        settings_window = next(child for child in app.winfo_children() if isinstance(child, desktop.tk.Toplevel))
        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)
        if sys.platform == 'darwin':
            assert desktop.UI_FONT == 'Helvetica Neue'
            assert any('Command / Shift' in str(w.cget('text')) for w in descendants(app)
                       if 'text' in w.keys())
        widgets = list(descendants(settings_window))
        method = next(w for w in widgets if isinstance(w, desktop.ttk.Combobox)
                      and 'SMTP (macOS / Windows)' in w.cget('values'))
        method.set('SMTP (macOS / Windows)')
        method.event_generate('<<ComboboxSelected>>')
        entries = [w for w in widgets if isinstance(w, desktop.ttk.Entry)
                   and not isinstance(w, desktop.ttk.Combobox)]
        for index, value in ((0, 'Example'), (1, 'person@example.org'), (4, 'bad')):
            entries[index].delete(0, 'end')
            entries[index].insert(0, value)
        save = next(w for w in widgets if isinstance(w, desktop.ttk.Button) and w.cget('text') == 'Save account')
        with patch.object(desktop.messagebox, 'showerror') as show_error:
            save.invoke()
        assert 'smtp_port' in show_error.call_args.args[1], show_error.call_args
        assert not (home / 'personal_data' / 'settings.json').exists()
        settings_window.destroy()
        app.auth_pending.set()
        with patch.object(desktop.webbrowser, 'open'):
            app.show_auth_prompt('https://microsoft.com/devicelogin', 'ABC123')
        assert str(app.stop_button['state']) == 'normal'
        assert app.auth_window is not None
        app.stop_button.invoke()
        assert app.auth_stop.is_set()
        assert app.auth_window is None
        app.auth_stop.clear()
        app.auth_pending.set()
        with patch.object(desktop.webbrowser, 'open'):
            app.show_auth_prompt('https://microsoft.com/devicelogin', 'ABC123')
        app.busy = True
        app.close_app()
        assert app.close_when_idle and app.auth_stop.is_set()
        app.events.put(('cancelled', 'Microsoft sign-in cancelled. No email was sent.'))
        app.poll()
        assert not app.close_when_idle

fake = SimpleNamespace(events=queue.Queue(), auth_prompt=lambda *args: None,
                       auth_done=lambda: None, auth_stop=threading.Event(),
                       auth_pending=threading.Event())
with patch.object(desktop.core, 'connect', return_value=MagicMock()):
    desktop.App.worker(fake, 'check', [], config)
assert fake.events.get()[0] == 'ok'

print('Cross-platform configuration, OAuth prompt, browser fallback and macOS UI checks passed. No emails sent.')
