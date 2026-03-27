# Mailhandle

## 1. What This Tool Does

`mailhandle` is a local Windows Outlook workspace for reviewing and handling your recent mail.

Major features:

- syncs `Inbox` and `Sent Items` from classic Outlook into a local SQLite database
- shows grouped mail threads in a local browser UI
- supports a command-based CLI mode that does not start a webpage
- lets you filter by time range, priority, and status
- lets you open the original Outlook item directly
- lets you mark status and keep that review state locally
- supports rule-based priority tuning and LLM model selection from `scripts\priority_rules.json`
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

### Start GUI Mode

For end users on Windows Explorer, just double-click:

```text
start_mailhandle_gui.bat
```

That launches the existing PowerShell GUI launcher and opens the browser workspace.

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

From any folder, replace `.\` with the absolute path to your install.

### Use The GUI Workspace

After launch:

- the browser should open automatically
- if it does not, open the URL shown in `tmp\mailhandle-last-start.txt`
- review threads in the `Mailbox` tab
- update `status`, filter by `Time range`, `Priority`, and `Status`
- use `Open` to jump to the original Outlook item
- use `Response` to draft a reply
- use `New Email` to draft a new message
- use the `Priority rules` tab to edit `scripts\priority_rules.json`

### Start CLI Mode

Run the command-based CLI without starting the browser workspace:

```powershell
& ".\scripts\launch_mailhandle.ps1" -Mode cli
```

From `cmd.exe`:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\launch_mailhandle.ps1" -Mode cli
```

Direct Python entrypoint:

```powershell
python .\scripts\mailhandle_cli.py
```

`CLI` mode defaults to `overview`, which syncs recent mail and prints the current open groups in the terminal.

Examples:

```powershell
& ".\scripts\launch_mailhandle.ps1" -Mode cli sync --json
& ".\scripts\launch_mailhandle.ps1" -Mode cli list --status todo
& ".\scripts\launch_mailhandle.ps1" -Mode cli show "<group-key-or-email-id>"
& ".\scripts\launch_mailhandle.ps1" -Mode cli status "<email-id>" done
& ".\scripts\launch_mailhandle.ps1" -Mode cli open "<email-id>"
python .\scripts\mailhandle_cli.py --help
```

### Use CLI Mode

Useful commands:

- `overview`: sync then print the current grouped mailbox view
- `sync`: refresh the local SQLite database without printing the full overview
- `list`: show grouped threads already stored in SQLite
- `show <group-key-or-email-id>`: print one thread in detail
- `status <email-id> <todo|doing|done>`: update local review state
- `open <email-id>`: open the original Outlook item
- `reply-draft <group-key-or-email-id>`: generate a reply draft JSON
- `reply-open <group-key-or-email-id>`: open `Reply All` in Outlook, optionally with a prepared draft
- `new-email-draft`: generate a new email draft JSON
- `new-email-open`: open a new Outlook email, optionally with a prepared draft

In CLI mode, edit `scripts\priority_rules.json` manually when you want to change rule behavior or the shared `llm_model` used for abstracts and drafts.

Drafting notes:

- `reply-draft` and `new-email-draft` require Codex CLI to be installed and logged in
- the LLM model for abstracts and drafts comes from `scripts\priority_rules.json` under `llm_model` and defaults to `codex-mini-latest`
- in GUI mode, the draft dialog supports a `second language` selector for `Thailand` and `Chinese`
- in CLI mode, request any second-language output in your notes text or notes file
- generated replies and new emails automatically append `Powered by Codex`

Advanced docs:

- [runbook.md](references/runbook.md)
- [database-design.md](references/database-design.md)
- [SKILL.md](SKILL.md)
