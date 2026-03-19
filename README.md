# Mailhandle

`mailhandle` is a native Windows Outlook workspace that syncs mail into a local SQLite database and serves a local browser UI for review.

It can run in three modes:

- standalone Windows tool launched from PowerShell or `cmd.exe`
- standalone Windows tool with optional Codex CLI features enabled
- Codex CLI skill installed under `~/.codex/skills/mailhandle`

The runtime itself is not tied to the Codex skill environment. The code resolves paths from the script location and can run from any copied install folder that keeps the `scripts\` contents together.

## What It Does

- reads Outlook mail from classic Outlook on Windows
- syncs mail into `data\mailhandle.sqlite`
- keeps one same-folder backup copy at `data\mailhandle.backup.sqlite`, refreshed once per local day
- preserves review state such as `status`
- keeps stored history rows stable unless the same email is synced again
- groups related mail by thread or normalized subject
- opens a local browser workspace for triage
- provides Outlook `Open` actions for specific items
- provides thread-level response drafting and `Reply All` opening
- embeds a priority-rules editor for `scripts\priority_rules.json`

## Works Without Codex CLI

These features work with only Windows, Outlook, Python, and `pywin32`:

- mailbox sync into SQLite
- browser workspace startup
- priority filtering and a default `All_Open` status view that hides `Done` items
- inline `status` editing
- Outlook `Open` actions
- priority-rules editing
- fallback non-LLM abstracts based on message content

## Optional Codex CLI Features

If the `codex` executable is installed, `mailhandle` can also use it for:

- cleaner LLM-generated mail abstracts
- thread reply drafting in the workspace

Without Codex CLI:

- the workspace still runs
- mail abstracts fall back to local heuristics
- reply-draft generation is unavailable

## Reply Drafting

In the response modal:

- `Additional notes` remains the free-text guidance box
- `second language` is a dropdown with `None`, `Thailand`, and `Chinese`
- the dropdown default is `None`
- selecting a second language automatically instructs the LLM, so you do not need to repeat that request in notes

Generated Outlook replies are pasted as normal mail content:

- `greeting`
- `body_en`
- optional `body_local`

When `body_local` is present:

- it is appended after the English body
- Outlook HTML paste shows it with a light darker background
- it begins with a compact marker such as `[Language_TH]`

The pasted reply does not show JSON keys, language section titles like `EN` or `TH`, or the structured `closing` field.

## Required Runtime

Required:

- Windows 10 or Windows 11
- classic Outlook desktop app installed and signed in
- Python 3.10 or newer on Windows
- `pywin32` installed in that Python environment

Local runtime files:

- `scripts\priority_rules.json`
- `data\mailhandle.sqlite`
- `data\mailhandle.backup.sqlite`
- `.cache\mailhandle_abstracts.json`
- `tmp\mailhandle-last-start.txt`

Install `pywin32`:

```powershell
python -m pip install pywin32
```

If needed:

```powershell
py -m pip install pywin32
```

Before first run:

- open classic Outlook once
- wait for `Inbox` and `Sent Items` to finish syncing

## Optional Codex CLI Setup

Install this only if you want LLM-generated abstracts and reply drafts.

Requirements:

- Node.js with `npm`
- Codex CLI installed on `PATH`

Install:

```powershell
npm i -g @openai/codex
```

Verify:

```powershell
codex --version
```

Optional model settings in `scripts\.env`:

```env
MAILHANDLE_ABSTRACT_MODEL=codex-mini-latest
MAILHANDLE_RESPONSE_MODEL=codex-mini-latest
```

## Install Location

Standalone install:

- copy the project folder anywhere on disk
- keep the `scripts\` directory contents together

Codex skill install:

```text
C:\Users\<username>\.codex\skills\mailhandle
```

## Configuration

Create `scripts\.env` from `scripts\.env.example`.

The launcher and priority logic now auto-detect the mailbox owner from Outlook.

Use these only as overrides when:

- Outlook autodetect is wrong
- you want to force a specific mailbox identity
- you are debugging an unusual environment

Optional owner overrides:

```env
MAIL_OWNER_EMAIL=your.name@company.com
MAIL_OWNER_NAME=Your Name
```

Optional:

```env
WINDOWS_PYTHON_EXE=C:\Path\To\Python\python.exe
MAILHANDLE_ABSTRACT_MODEL=codex-mini-latest
MAILHANDLE_RESPONSE_MODEL=codex-mini-latest
```

## Start The Workspace

From the `mailhandle` root folder in PowerShell:

```powershell
& ".\scripts\launch_mailhandle.ps1"
```

From the `mailhandle` root folder in `cmd.exe`:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\launch_mailhandle.ps1"
```

From any folder, use the absolute path to your install:

```powershell
& "C:\path\to\mailhandle\scripts\launch_mailhandle.ps1"
```

or:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\path\to\mailhandle\scripts\launch_mailhandle.ps1"
```

If installed as a Codex skill, the same launcher is typically here:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

or:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%USERPROFILE%\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

The launcher:

- starts the Python workspace server in the background
- requests browser auto-open
- prints the workspace URL or current startup status
- writes the startup summary to `<mailhandle-root>\tmp\mailhandle-last-start.txt`

Read the latest startup summary:

```powershell
Get-Content "C:\path\to\mailhandle\tmp\mailhandle-last-start.txt"
```

or:

```bat
type "C:\path\to\mailhandle\tmp\mailhandle-last-start.txt"
```

If the browser does not appear, open the reported workspace URL manually.

`scripts\start_mailhandle.cmd` remains available as a thin wrapper, but the direct launcher commands above are the preferred manual entrypoints.

## Codex Skill Usage

If you install this under `~/.codex/skills/mailhandle`, the intended prompts are:

```text
$mailhandle start
$mailhandle run
$mailhandle launch
```

Those prompts ultimately call the same `scripts\launch_mailhandle.ps1` launcher.

## Priority Tuning

Use the embedded browser editor in the workspace when possible.

Useful fields in `scripts\priority_rules.json`:

- `owner_aliases`
- `manager_senders`
- `greeting_terms`

Notes:

- `manager_senders` can match a display name, an SMTP address, or a combined form like `Bin Tan <Bin.Tan@lumentum.com>`
- changing `priority_rules.json` affects future synced items; existing rows already stored in SQLite keep their saved priority/history

## Outlook Compatibility

This project uses Outlook COM through `win32com`.

Supported:

- classic Outlook for Windows

Not supported:

- new Outlook when classic COM automation is unavailable
- Outlook Web
- macOS or Linux native Outlook access

## Troubleshooting

### `No module named win32com`

Install `pywin32` in the Windows Python environment used by the launcher.

### `codex` is not installed

The workspace can still start and sync mail.

Effects:

- abstracts fall back to local heuristics
- reply-draft generation is unavailable

### Outlook access fails or is slow

- open classic Outlook first
- wait for mailbox sync to finish
- rerun after Outlook becomes responsive

### Only new Outlook is installed

This project is not expected to work reliably there. Use classic Outlook.

### Browser did not open automatically

Open the workspace URL from the launcher output or from:

```text
<mailhandle-root>\tmp\mailhandle-last-start.txt
```

### Abstracts are slow on first run

If Codex CLI is installed, first-run abstracts for unseen mail are slower because they are generated once and cached under `.cache\`.

### I changed priority rules but old items did not move

That is the intended behavior for the current runtime.

- new rule changes apply to future synced items
- historical rows already stored in `data\mailhandle.sqlite` keep their existing saved priority unless you explicitly rebuild or replace the database

## Files

- `scripts\launch_mailhandle.ps1`: detached launcher entrypoint
- `scripts\start_mailhandle.ps1`: PowerShell launcher implementation
- `scripts\start_mailhandle.cmd`: optional thin CMD wrapper
- `scripts\run_mail_database.py`: browser workspace and sync entrypoint
- `scripts\mailhandle_db.py`: SQLite storage and Outlook helpers
- `scripts\read_outlook.py`: Python-to-Windows reader bridge
- `scripts\read_outlook_win.py`: Outlook COM reader
- `scripts\summarize_mail.py`: summary and abstract generation logic
- `scripts\priority_rules.json`: priority tuning rules
- `SKILL.md`: Codex skill instructions
- `references\runbook.md`: developer and packaging notes
- `references\database-design.md`: SQLite-backed history design
- `references\release-notes-next.md`: draft notes and checklist for the next release
- `references\example-prompts.md`: Codex prompt examples

## Sharing

If you share this project with another user, tell them:

- it runs as a standalone Windows tool
- Codex CLI is optional, not mandatory for startup
- classic Outlook is required
- Python plus `pywin32` is required
- the Codex skill install is just one packaging mode
