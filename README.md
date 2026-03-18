# Mailhandle

`mailhandle` is a Codex CLI skill for native Windows that reads Outlook mail, stores review history in SQLite, and opens a local browser workspace for triage.

Default usage is through Codex CLI. If the environment is not clearly defined, assume Codex CLI and the local browser workspace.

Recommended Codex prompt:

```text
$mailhandle start
```

Accepted aliases:

```text
$mailhandle run
$mailhandle launch
```

## What It Does

- Reads Outlook mail from classic Outlook on Windows
- Syncs mail into a local SQLite database
- Preserves review state such as `status`
- Groups related mail by thread or normalized subject
- Uses Codex CLI to generate concise abstracts from mail content
- Opens a local browser workspace with:
  - a mailbox tab for review, filters, response drafting, and Outlook opening
  - an embedded priority-rules editor

## Requirements

Required:

- Windows 10 or Windows 11
- Classic Outlook desktop app installed and signed in
- Node.js with `npm`, to install Codex CLI
- Python 3.10 or newer on Windows
- `pywin32` installed in that Python environment

Local storage:

- `data\mailhandle.sqlite` for history and review state
- `.cache\mailhandle_abstracts.json` for abstract caching

## Outlook Compatibility

This skill uses Outlook COM through `win32com`.

Supported:

- Classic Outlook for Windows

Not supported:

- New Outlook when classic COM automation is unavailable
- Outlook Web
- macOS or Linux native Outlook access

## Clean-Machine Setup

### 1. Install Python

Install Python 3.10+ for Windows and confirm `python` or `py` works:

```powershell
python --version
```

or

```powershell
py --version
```

### 2. Install Python dependency

Install `pywin32`:

```powershell
python -m pip install pywin32
```

If needed:

```powershell
py -m pip install pywin32
```

### 3. Ensure Outlook is ready

- Install classic Outlook
- Sign in to the mailbox
- Let `Inbox` and `Sent Items` finish syncing
- Open Outlook once before running the skill

### 4. Copy the skill folder

Copy this folder to:

```text
C:\Users\<username>\.codex\skills\mailhandle
```

## Configuration

Create `scripts\.env` from `scripts\.env.example`.

Recommended fields:

```env
MAIL_OWNER_EMAIL=your.name@company.com
MAIL_OWNER_NAME=Your Name
```

Optional:

```env
WINDOWS_PYTHON_EXE=C:\Path\To\Python\python.exe
MAILHANDLE_ABSTRACT_MODEL=codex-mini-latest
```

## Install Node.js and Codex CLI

Use the official Node.js Windows installer:

- https://nodejs.org/en/download

Recommended:

- install the current LTS release
- keep `npm` enabled

Verify:

```powershell
node --version
npm --version
```

Install Codex CLI:

```powershell
npm i -g @openai/codex
```

Start Codex once:

```powershell
codex
```

Upgrade later:

```powershell
npm i -g @openai/codex@latest
```

Notes:

- OpenAI's current docs say Windows support is still experimental
- This skill is intended for native Windows because it depends on classic Outlook COM automation

## Primary Workflow

Start the local browser workspace:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

From `cmd.exe`:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%USERPROFILE%\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

In Codex CLI, the intended skill prompt is `start`. `run` and `launch` are equivalent aliases.

The Codex launcher starts the workspace server in the background, requests browser auto-open, waits for the startup summary, and prints the workspace URL or current status.

These commands do not depend on the current working directory and are safe to run from any folder.

`scripts\start_mailhandle.cmd` is still available as a thin wrapper, but the direct `cmd.exe` command above is the recommended manual CMD entrypoint.

Read the current startup result from:

```powershell
Get-Content "$env:USERPROFILE\.codex\skills\mailhandle\tmp\mailhandle-last-start.txt"
```

or

```bat
type "%USERPROFILE%\.codex\skills\mailhandle\tmp\mailhandle-last-start.txt"
```

That file is the authoritative source for:

- workspace URL
- startup confirmation
- stdout/stderr log paths

The launcher also tries to open that local page in your default browser. The page provides:

- mailbox history from SQLite
- time-range filters such as `today`, `last_2days`, and `last_7_days`
- priority and status filters
- inline `status` editing
- an `Open` action for a specific Outlook item
- thread-level `Response` drafting that opens `Reply All` in Outlook
- an embedded priority-rules editor

If you need to adjust priority matching behavior, use the embedded editor in the workspace.

## Priority Tuning

Edit:

```text
scripts\priority_rules.json
```

Useful fields:

- `owner_aliases`
- `manager_senders`
- `greeting_terms`

## Troubleshooting

### `No module named win32com`

Install `pywin32` in the Windows Python environment used by the skill.

### Outlook access fails or is slow

- Open classic Outlook first
- Wait for sync to finish
- Try again if Outlook is busy

### Only new Outlook is installed

This skill is not expected to work reliably there. Use classic Outlook.

### Abstracts are slow on first run

The first run for unseen emails is slower because Codex CLI generates and caches abstracts.

### The launcher does not show the URL immediately

From `cmd.exe`, prefer the direct launcher command:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%USERPROFILE%\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

If you are using the wrapper, read the saved startup summary:

```bat
type "%USERPROFILE%\.codex\skills\mailhandle\tmp\mailhandle-last-start.txt"
```

If the browser still does not appear, open the reported workspace URL manually.

## Files

- `SKILL.md`: skill usage instructions for Codex
- `scripts\run_mail_database.py`: local browser workspace and sync entrypoint
- `scripts\start_mailhandle.cmd`: optional thin CMD wrapper
- `scripts\launch_mailhandle.ps1`: detached launcher helper
- `scripts\start_mailhandle.ps1`: PowerShell implementation behind the launcher
- `scripts\edit_priority_rules.py`: browser editor that saves rules back to disk
- `scripts\mailhandle_db.py`: SQLite storage, DPAPI helpers, and Outlook open helper
- `scripts\read_outlook.py`: Python-to-Windows reader bridge
- `scripts\read_outlook_win.py`: Outlook COM reader
- `scripts\priority_rules.json`: priority tuning rules
- `references\runbook.md`: developer notes
- `references\database-design.md`: SQLite-backed history design and rollout plan
- `references\example-prompts.md`: example prompts

## Notes for Sharing

When you share this skill, tell users:

- It is Windows-only
- It requires classic Outlook
- It requires Python plus `pywin32`
- Codex CLI is the default intended usage

The default installation target is:

```text
C:\Users\<username>\.codex\skills\mailhandle
```

The normal workflow is:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

From `cmd.exe`, use:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%USERPROFILE%\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

Those commands request browser auto-open and print the workspace URL.
