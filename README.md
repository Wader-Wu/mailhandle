# Mailhandle

`mailhandle` reads a Windows Outlook mailbox and turns matching messages into a prioritized todo list.

Default usage is through Codex CLI.

Direct Python execution exists only as a debug path. If usage is not clearly defined, assume Codex CLI.

## What It Does

- Reads Outlook inbox mail with filters like `today`, `unread`, sender, subject, and project keyword
- Checks `Sent Items` for later replies in the same thread
- Marks threads you already replied to
- Produces a concise todo-oriented JSON or plain-text summary
- Uses Codex CLI to turn mail body content into cleaner per-email abstracts
- Always writes a self-contained HTML review page with filters, checkboxes, and local export
- Automatically tries to open the generated HTML report on Windows

## Requirements

This skill is designed for native Windows.

Required:

- Windows 10 or Windows 11
- Classic Outlook desktop app installed and signed in
- Node.js with `npm`, to install Codex CLI
- Python 3.10 or newer installed on Windows
- `pywin32` installed in that Python environment

Optional:

- Direct Python execution for debugging

## Outlook Compatibility

This skill uses the Outlook COM API through `win32com`.

Supported:

- Classic Outlook for Windows

Not supported:

- New Outlook when classic COM automation is unavailable
- Outlook Web
- macOS or Linux native Outlook access

If a user only has the new Outlook experience, they should switch to classic Outlook before using this skill.

## Clean-Machine Setup

### 1. Install Python

Install Python 3.10+ for Windows and make sure `python` or `py` works in PowerShell.

Check:

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

If `python` is not on `PATH`, use:

```powershell
py -m pip install pywin32
```

### 3. Ensure Outlook is ready

- Install classic Outlook
- Sign in to the mailbox
- Let `Inbox` and `Sent Items` finish syncing
- If Outlook was just installed, open it once before running the script

### 4. Copy the skill folder

Copy this folder to one of these locations:

For Codex CLI:

```text
C:\Users\<username>\.codex\skills\mailhandle
```

For debug-only direct script use:

```text
Any folder you want
```

If the target user is not explicitly using debug-only direct script execution, install it as a Codex CLI skill.

## Configuration

Create a local config file by copying:

```text
scripts\.env.example
```

to:

```text
scripts\.env
```

Recommended fields:

```env
MAIL_OWNER_EMAIL=your.name@company.com
MAIL_OWNER_NAME=Your Name
```

Optional field:

```env
WINDOWS_PYTHON_EXE=C:\Path\To\Python\python.exe
```

Optional abstract-model override:

```env
MAILHANDLE_ABSTRACT_MODEL=codex-mini-latest
```

Notes:

- On normal native Windows installs, `WINDOWS_PYTHON_EXE` is usually not needed
- Set it only if Outlook must be accessed through a specific Windows Python interpreter
- If `MAILHANDLE_ABSTRACT_MODEL` is not set, Codex CLI uses its normal default model/profile

## Priority Tuning

Edit:

```text
scripts\priority_rules.json
```

The default ranking now follows this order:

1. High priority:
   - you are in `To` and the mail opens with `Hi` or `Hello` plus your name or configured nickname
   - you are `@` mentioned in the mail body
   - the mail is from a configured manager
   - the mail is from an Agile-style sender and asks for your approval
2. Medium priority:
   - you are in `To` but the high-priority cases above do not apply
3. Low priority:
   - you are only in `Cc`

Useful fields in `priority_rules.json`:

- `owner_aliases`: add nicknames or short-name matches such as `["Luke"]`
- `manager_senders`: add manager names or email fragments
- `greeting_terms`: customize greeting detection at the start of the email
- rule keys:
  - `attention_flags_any`
  - `body_contains_any`
  - `sender_matches_manager`

If you want a browser editor instead of editing JSON by hand:

```powershell
python scripts\run_priority_rules_editor.py
```

This opens a local HTML page with labeled form fields for the top-level settings and each rule card. It validates the generated JSON and saves it directly back to the file.

## Default Usage: Codex CLI

### Install Node.js

Use the official Node.js Windows installer from:

- https://nodejs.org/en/download

Recommended:

- install the current `LTS` version
- use the standard Windows installer
- keep `npm` enabled during setup

After installation, open a new PowerShell window and verify:

```powershell
node --version
npm --version
```

If either command is missing, close and reopen PowerShell and try again.

### Install Codex CLI

OpenAI's current CLI setup uses `npm`.

1. Install Codex CLI:

```powershell
npm i -g @openai/codex
```

2. Start Codex once:

```powershell
codex
```

3. On first run, sign in with your ChatGPT account or an API key.
4. To upgrade later:

```powershell
npm i -g @openai/codex@latest
```

Notes:

- OpenAI's docs say Codex CLI is available on macOS and Linux, and Windows support is still experimental.
- OpenAI recommends WSL for the best Windows experience.
- This `mailhandle` skill is intended for native Windows because it depends on classic Outlook COM automation.

### Install This Skill

1. Copy this skill to:

```text
C:\Users\<username>\.codex\skills\mailhandle
```

2. Restart Codex CLI if it was already running.
3. Ask for mail summaries using prompts such as:

- `summary today's mail`
- `show unread mail today`
- `review yesterday's emails and summarize action items`

The skill entrypoint is:

```text
scripts\run_summary.py
```

Every normal summary run now also writes and opens a static review page.

By default, the summary pipeline also asks Codex CLI to generate abstract text from each matched email body. The first run for new emails is slower because those abstracts are generated once and then cached locally; repeat runs reuse the cached abstracts and stay much faster.

Example:

```powershell
python scripts\run_summary.py --date-preset today --json
```

This writes both:

- `records\<YYYY-MM-DD>\summary-<timestamp>.json`
- `records\<YYYY-MM-DD>\report-<timestamp>.html`

## Debug Only

Use this only when:

- debugging the skill locally
- validating Outlook access

If the target environment is not clearly defined, do not default to this mode.

From the skill folder:

```powershell
python scripts\run_summary.py --date-preset today --json
```

Examples:

```powershell
python scripts\run_summary.py --unread-only --date-preset today --json
python scripts\run_summary.py --date-preset yesterday --json
python scripts\run_summary.py --project zhuque --date-preset this_month --json
python scripts\run_summary.py --from-contains Jeffrey --date-preset last_7_days --json
```

If `python` is unavailable but `py` works:

```powershell
py scripts\run_summary.py --date-preset today --json
```

## Common Commands

Default JSON summary:

```powershell
python scripts\run_summary.py --json
```

Unread mail today:

```powershell
python scripts\run_summary.py --unread-only --date-preset today --json
```

Yesterday's mail:

```powershell
python scripts\run_summary.py --date-preset yesterday --json
```

Project mail this month:

```powershell
python scripts\run_summary.py --project zhuque --date-preset this_month --json
```

Include notification-style mail:

```powershell
python scripts\run_summary.py --date-preset today --include-notifications --json
```

Keep every matching email separate:

```powershell
python scripts\run_summary.py --date-preset today --no-collapse --json
```

Generate a static HTML review report directly:

```powershell
python scripts\run_review_report.py --date-preset today
```

Generate a report from an existing summary JSON:

```powershell
python scripts\build_review_report.py --input-json path\to\summary.json
```

Open the priority-rules editor:

```powershell
python scripts\run_priority_rules_editor.py
```

## Output

The JSON output includes:

- summary text
- filters used
- counts by priority
- todo entries
- whether you already replied
- next action suggestion
- `report_meta` with saved JSON/HTML paths
- `report_opened` to show whether the HTML report was opened automatically

The HTML review report includes:

- grouped sections for `Needs action`, `Already replied`, and `Done`
- local done checkboxes per item
- search and filter controls
- `Export review JSON`
- `Export reviewed HTML`
- browser-local persistence while you review

## Troubleshooting

### `No module named win32com`

Install `pywin32`:

```powershell
python -m pip install pywin32
```

### `Server execution failed`

Usually Outlook is not ready for COM access.

Try:

- Open classic Outlook manually
- Wait for mailbox sync to finish
- Close and reopen Outlook
- Run the command again

### No messages returned

Check:

- Outlook is signed into the expected mailbox
- The mailbox has mail in `Inbox`
- The date filter is correct
- You are not filtering too aggressively with `--unread-only`, sender, or subject filters

### Sent-item reply detection is missing

Check:

- `Sent Items` is synced
- You are using classic Outlook
- The sent reply is in the same thread

### Abstracts are slow or look stale

Check:

- The first run for new emails is expected to be slower because Codex CLI generates and caches abstracts
- Codex CLI is installed and logged in
- `MAILHANDLE_ABSTRACT_MODEL` is set only if you really want to override the normal Codex model choice
- If you want to force regeneration, delete `.cache\mailhandle_abstracts.json` and run again

### Only new Outlook is installed

This skill is not expected to work reliably there. Use classic Outlook.

## Files

- `SKILL.md`: skill usage instructions for Codex
- `scripts\run_review_report.py`: wrapper to build a static HTML review report
- `scripts\run_priority_rules_editor.py`: wrapper to open the browser editor for `priority_rules.json`
- `scripts\build_review_report.py`: HTML report generator
- `scripts\edit_priority_rules.py`: local HTTP editor that saves rule changes back to disk
- `scripts\run_summary.py`: main wrapper entrypoint
- `scripts\summarize_mail.py`: summary logic
- `scripts\read_outlook.py`: Python-to-Windows reader bridge
- `scripts\read_outlook_win.py`: Outlook COM reader
- `scripts\priority_rules.json`: priority tuning rules
- `references\runbook.md`: developer notes
- `references\example-prompts.md`: example prompts

## Notes for Sharing

If you share this skill with another user, tell them up front:

- It is Windows-only
- It requires classic Outlook, not just the new Outlook app
- It requires Python plus `pywin32`
- Codex CLI is the default intended usage
- Direct Python execution is for debug only

The default target setup on a new machine is:

```text
C:\Users\<username>\.codex\skills\mailhandle
```

If you are explicitly debugging without Codex CLI, the fastest check is:

```powershell
python scripts\run_summary.py --date-preset today --json
```

If you want a readable review artifact for personal use and record-keeping:

```powershell
python scripts\run_review_report.py --date-preset today
```

## GitHub Release Checklist

When you publish `mailhandle` as a GitHub release, ship only the skill source and docs that are needed to run it.

Include:

- `README.md`
- `SKILL.md`
- `references\`
- `scripts\`

Exclude:

- `.cache\`
- `records\`
- `sessions\`
- `tmp\`
- `log\`
- `.sandbox\`
- `.sandbox-bin\`
- `__pycache__\`
- `*.pyc`
- local auth or state files such as `auth.json`, `history.jsonl`, `models_cache.json`, `state_*.sqlite*`
- any machine-specific `.env` file

If you package a release archive manually, make sure those ignore rules are applied before uploading the zip or attaching the GitHub release asset.
