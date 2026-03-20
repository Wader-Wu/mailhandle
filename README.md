# Mailhandle

## 1. What This Tool Does

`mailhandle` is a local Windows Outlook workspace for reviewing and handling your recent mail.

Major features:

- syncs `Inbox` and `Sent Items` from classic Outlook into a local SQLite database
- shows grouped mail threads in a local browser UI
- lets you filter by time range, priority, and status
- lets you open the original Outlook item directly
- lets you mark status and keep that review state locally
- supports rule-based priority tuning from `scripts\priority_rules.json`
- can generate abstracts, replies, and new emails when Codex CLI is installed

Safety:

- your mail history is stored locally in `data\mailhandle.sqlite`
- the database keeps a local backup copy at `data\mailhandle.backup.sqlite`
- stored abstracts in SQLite are protected with Windows DPAPI
- the tool runs against your local Outlook profile; it is not a hosted mail service
- if Codex CLI is not installed, the core sync and review workflow still works

## 2. How To Install

### Recommended For New Windows Users

Run the bootstrap installer from the project root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\install_mailhandle.ps1"
```

This installer:

- installs Python 3.11
- installs Node.js LTS
- installs `pywin32`
- installs Codex CLI
- runs `codex login`
- copies the app to `%USERPROFILE%\mailhandle`
- launches the workspace

Requirements:

- Windows 10 or Windows 11
- classic Outlook installed and signed in
- `winget` available
- permission to install desktop software

### Manual Install

If you do not want to use the installer:

1. Install Python on Windows.
2. Install `pywin32`:

```powershell
py -m pip install pywin32
```

3. Optional: install Codex CLI if you want LLM-powered abstracts and drafting:

```powershell
npm i -g @openai/codex
codex login
```

4. Open classic Outlook once and wait for mailbox sync to finish.

## 3. How To Use

### Start

From the project root in PowerShell:

```powershell
& ".\scripts\launch_mailhandle.ps1"
```

If you used the bootstrap installer, the default install location is:

```powershell
& "$env:USERPROFILE\mailhandle\scripts\launch_mailhandle.ps1"
```

From `cmd.exe`:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\launch_mailhandle.ps1"
```

### Use The Workspace

After launch:

- the browser should open automatically
- if it does not, open the URL shown in `tmp\mailhandle-last-start.txt`
- review threads in the `Mailbox` tab
- update `status`, filter by `Time range`, `Priority`, and `Status`
- use `Open` to jump to the original Outlook item
- use `Response` to draft a reply
- use `New Email` to draft a new message
- use the `Priority rules` tab to edit `scripts\priority_rules.json`

Notes:

- if you select a second language in the draft dialog, the tool can automatically generate a localized version of the email body
- generated replies and new emails automatically append `Powered by Codex`

Advanced docs:

- [runbook.md](references/runbook.md)
- [database-design.md](references/database-design.md)
- [SKILL.md](SKILL.md)
