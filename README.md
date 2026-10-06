# RP Mail - Experimental Prototype

An experimental Windows and macOS desktop prototype exploring how to review potential relationships between research projects and research outputs in Research Portal+ (RP+), and contacting researchers from a personal email account.

**Status: Experimental prototype - work in progress.**

This repository is an early attempt to explore the workflow. It is not a completed project deliverable or a production-ready system. The existing functionality and limited tests support evaluation only; broader validation and further development are still needed.

**Prototype version: v0.1.0**

## Implemented Prototype Features

- Import CSV or Excel (.xlsx) files, edit records, and preview confirmation emails.
- Find publicly listed email addresses through researcher profiles linked from RP+ project pages.
- Prefer contacts whose profile links appear on both the project and publication pages. Retain alternative contacts, source URLs, and lookup timestamps.
- Send a test email or selected records through SMTP on Windows or macOS. Classic Outlook for Windows remains available for drafts and sending.
- Track drafts and submissions locally to prevent duplicate processing.
- Retry page timeouts once and continue after individual page failures. Stop on access restrictions, rate limits, or verification pages.

Finding an email address does not confirm a research relationship. Records must be reviewed before sending. RP Mail does not update RP+ records.

## Requirements

- Windows 64-bit or macOS
- Python 3.10+ when running from source
- Microsoft Edge, Google Chrome, or Playwright Chromium for public email lookup
- A permitted SMTP account for cross-platform sending, or **Classic Outlook for Windows** with the personal sending account configured

New Outlook is not supported by the Windows COM integration. The SMTP path works independently of Outlook. The application does not store an email password: it uses Microsoft device sign-in with an approved application registration, or a password supplied through an environment variable when the mail provider permits it.

## Getting Started

Platform guides: [Windows desktop guide](docs/桌面软件使用说明.md) · [macOS guide](docs/macOS使用说明.md).

Windows:

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python desktop_app.py
```

macOS:

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python desktop_app.py
```

If Edge or Chrome is installed, no separate Playwright browser download is required. Otherwise run `.venv/bin/python -m playwright install chromium` on macOS, or the equivalent `.venv\Scripts\python -m playwright install chromium` on Windows.

1. Enter your name and personal sending address in the account settings. Choose **SMTP** on macOS; Windows can also choose **Classic Outlook**.
2. Import your research data and select records.
3. Run the email lookup and review the suggested contacts and sources.
4. Check the relationship evidence and email preview.
5. Check the sending connection and send a test email to yourself before sending reviewed records. Outlook drafts are available only with the Windows Outlook method.

For Microsoft 365 SMTP, confirm with Eric and ANU IT that SMTP AUTH, device-code sign-in, and the `SMTP.Send` permission are allowed for the chosen mailbox and application. Ask IT for an approved public-client application ID and tenant ID. Device sign-in displays a code and opens the Microsoft sign-in page; **Cancel sign-in** stops waiting without sending. A successful sign-in is reused for later checks and sends while the app remains open. Tokens are kept in memory only and are not saved across restarts. The default server settings are `smtp.office365.com`, port `587`, and STARTTLS; use your provider's settings for another service. A successful SMTP submission does not prove delivery.

For password-based SMTP, set the named environment variable before starting the app. A Finder-launched macOS app does not inherit Terminal-only environment variables, so Microsoft sign-in is the practical choice for that launch method. Do not put passwords in `settings.json`.

The desktop interface defaults to English. Use the **English / 中文** selector in the top bar to switch languages. Your preference is saved; records and edits are retained when switching. Researcher-facing emails remain in English and use a personal sender's voice.

## Input Data

Required columns: `output_uuid`, `project_uuid`, `project_code`, `title`, `output_url`, and `already_in_relations`.

Automatic email lookup also requires `project_url`. Excel imports use the active worksheet, with column names in the first row.

[examples/sample.csv](examples/sample.csv) contains fictional data for exploring the interface. It is not intended for sending or live lookup.

## Local Data

Working records and settings are stored in `personal_data/`. Processing history is stored in `send_history.sqlite3`. Preserve both when upgrading.

The repository excludes real research datasets, personal account settings, collected contacts, logs, and sending history.

## Testing

Run `.venv\Scripts\python tests/run_tests.py` on Windows or `.venv/bin/python tests/run_tests.py` on macOS.

Automated checks use temporary data and simulated sending. They do not contact researchers or access RP+. The desktop test briefly opens a Tk window.

A live lookup was verified in the packaged Windows application: one project record returned three publicly listed contact emails and saved the results. This does not establish full-dataset coverage or verify email delivery.

## Building the Desktop Application

Windows:

```powershell
.venv\Scripts\python -m pip install -r requirements-dev.txt
.venv\Scripts\python -m PyInstaller --noconfirm RPMail.spec
```

macOS:

```sh
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m PyInstaller --noconfirm RPMail.spec
```

Run `dist/RPMail/RPMail.exe` on Windows or `dist/RPMail.app` on macOS. Build separately on each operating system. Keep the Windows `_internal` folder beside the executable. Windows data remains beside the executable; the macOS app stores settings, working records and send history in `~/Library/Application Support/RPMail/` so an installed app bundle stays read-only.

## Current Limitations

- No inbox synchronization, automatic reply classification, or RP+ updates.
- Researchers without linked public profiles or published email addresses need manual contact details.
- Multiple relationships produce separate emails; grouping by researcher is not implemented.
- Generated Outlook drafts must be sent manually. Later automatic sends skip those records. SMTP has no Outlook-draft feature.
- Submission to Outlook or SMTP does not guarantee delivery. Check Sent Items and bounce messages.
- Microsoft 365 SMTP requires tenant/app consent and mailbox policy to permit SMTP AUTH; some institutional accounts will not allow this route.
- Website changes, network conditions, and Outlook account policies can affect lookup and sending.
- The application does not bypass website access restrictions or verification challenges.
