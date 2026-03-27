# Runbook

This project is a standalone Windows runtime that can also be packaged as a Codex skill.

## Runtime Modes

- Standalone Windows app in GUI mode: launch `scripts\launch_mailhandle.ps1` from any copied install folder.
- Standalone Windows app in CLI mode: launch `scripts\launch_mailhandle.ps1 -Mode cli` to use terminal commands only.
- End-user double-click entrypoint for GUI mode: `start_mailhandle_gui.bat` from the install root.
- Standalone Windows app with Codex CLI: same runtime, plus LLM-generated abstracts and reply drafts.
- Codex skill install: same runtime copied under `~/.codex/skills/mailhandle` and launched through prompt aliases.

## Source Layout

- Install root: any copied `mailhandle` folder that preserves the `scripts\` contents
- Runtime code: `scripts/`
- Runtime rule config: `scripts/priority_rules.json`
- Runtime env template: `scripts/.env.example`
- SQLite history design: `references/database-design.md`
- Local DB path: `data/mailhandle.sqlite`
- Daily backup path: `data/mailhandle.backup.sqlite`

## GitHub Release Packaging

Include:

- `README.md`
- `SKILL.md`
- `references/`
- `references/release-notes-next.md`
- `scripts/`

Exclude:

- `.cache/`
- `records/`
- `sessions/`
- `tmp/`
- `log/`
- `.sandbox/`
- `.sandbox-bin/`
- `__pycache__/`
- `*.pyc`
- `data/`
- `auth.json`
- `history.jsonl`
- `models_cache.json`
- `state_*.sqlite*`
- any local `.env` file

Recommended release outputs:

- `mailhandle-<version>-source.zip`
- `mailhandle-<version>-windows-portable.zip`

The portable zip should include a bundled Windows Python runtime at `runtime\python\python.exe`. The launcher now auto-detects that path before falling back to `.env` or `python` on `PATH`.

## Wrapper Usage

Recommended first-time Windows setup for a new user:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\install_mailhandle.ps1"
```

That bootstrap flow:

- installs Python 3.11 with `winget`
- installs Node.js LTS with `winget`
- installs `pywin32`
- installs Codex CLI with `npm`
- runs `codex login`
- copies the project to `%USERPROFILE%\mailhandle`
- launches the workspace

Limits:

- `winget` must already be available on the machine
- `codex login` still requires user interaction
- classic Outlook still must already be installed and signed in

Primary standalone GUI workflow from the install root:

```powershell
& ".\scripts\launch_mailhandle.ps1"
```

Manual `cmd.exe` launch:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\launch_mailhandle.ps1"
```

From any folder, replace `.\` with the absolute path to your install.

Command-based CLI mode from the install root:

```powershell
& ".\scripts\launch_mailhandle.ps1" -Mode cli
```

Examples:

```powershell
& ".\scripts\launch_mailhandle.ps1" -Mode cli list --status todo
& ".\scripts\launch_mailhandle.ps1" -Mode cli show "introduction of ofs 980-26 based wdm"
& ".\scripts\launch_mailhandle.ps1" -Mode cli status "<email-id>" done
python .\scripts\mailhandle_cli.py --help
```

Codex skill install example:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

Canonical Codex prompt after skill installation:

```text
$mailhandle start
```

Accepted aliases:

```text
$mailhandle run
$mailhandle launch
```

This:

- starts the Python workspace server in the background
- requests browser auto-open for the workspace URL
- returns control to Codex after printing the startup summary
- writes the workspace URL and startup status to `tmp\mailhandle-last-start.txt`
- syncs the local SQLite database
- preserves `status`
- lets you open specific Outlook items
- lets you draft thread responses and open `Reply All` in Outlook
- includes a `second language` dropdown in the response modal with `None`, `Thailand`, and `Chinese`
- embeds the priority-rules editor in the same page

After launch, read:

```powershell
Get-Content ".\tmp\mailhandle-last-start.txt"
```

or

```bat
type ".\tmp\mailhandle-last-start.txt"
```

If the browser does not appear, open the reported workspace URL manually.

Direct `python scripts/run_mail_database.py` is still valid for manual debugging, but it should not be the default Codex skill entrypoint because it keeps the CLI attached to the long-lived server process.

Direct `python scripts/mailhandle_cli.py` is the command-based terminal entrypoint. With no subcommand it defaults to `overview`.

`scripts\start_mailhandle.cmd` remains available as a thin wrapper, but the direct `cmd.exe` command above is the recommended manual CMD entrypoint.

## Optional Codex CLI Features

If the `codex` executable is on `PATH`, the runtime can use it for:

- LLM-generated mail abstracts
- reply drafting for a thread

Reply drafting details:

- the modal uses free-text `Additional notes` plus a separate `second language` dropdown
- the dropdown automatically requests the localized version, so users do not need to type that into notes
- the LLM returns one shared structured contract for reply and new-email drafting: `subject`, `greeting`, `body_en`, optional `body_local`, `local_language`, and `closing`
- pasted Outlook replies keep a normal mail layout: greeting, English body, then optional local-language body, then `[ Powered by Codex ]`
- new Outlook emails use `subject` for the mail subject and paste greeting, English body, optional local-language body, `[ Powered by Codex ]`, then `closing`
- the optional local-language body is shown with a light darker background in Outlook HTML and a compact marker like `[Language_TH]`

Without Codex CLI:

- startup still works
- sync still works
- the command-based CLI still works
- the browser workspace still works
- reply drafting is unavailable

To edit `priority_rules.json`:

- GUI mode: use the embedded browser editor in the workspace
- CLI mode: edit `scripts\priority_rules.json` manually

## Priority Tuning

Useful fields in `scripts/priority_rules.json`:

- `default_sync_period`
- `llm_model`
- `owner_aliases`
- `manager_senders`
- `greeting_terms`

Notes:

- `default_sync_period` controls the startup sync window and default mailbox UI range filter; supported values are `today`, `last_1day`, `last_2days`, `last_7_days`, `this_month`, and `last_month`
- `llm_model` controls the shared Codex model for mail abstracts, reply drafts, and new-email drafts; it defaults to `codex-mini-latest` if omitted
- `manager_senders` can match `Bin Tan`, `Bin.Tan@lumentum.com`, or `Bin Tan <Bin.Tan@lumentum.com>`
- owner-directed priority is now controlled only by explicit rules and flags; the old `boost_owner_attention` top-level switch has been removed
- changing `priority_rules.json` is forward-looking; it affects future synced items and does not retroactively rescore historical rows already stored in SQLite

## Release Prep

Before cutting the next release:

- update `references/release-notes-next.md`
- verify Python syntax:

```powershell
Get-ChildItem scripts -Filter *.py | ForEach-Object { python -m py_compile $_.FullName }
```

- verify launcher startup:

```powershell
& ".\scripts\launch_mailhandle.ps1"
```

- confirm the startup summary exists:

```powershell
Get-Content ".\tmp\mailhandle-last-start.txt"
```

- confirm the release bundle excludes local runtime state such as `data\`, `tmp\`, `.cache\`, and local `.env` files

To build release zips:

```powershell
& ".\scripts\build_release.ps1" -Version "0.1.0"
```

To build both source and Windows-portable zips:

```powershell
& ".\scripts\build_release.ps1" -Version "0.1.0" -PortablePythonDir "C:\path\to\python-embed"
```

Packaging recommendation for new Windows users:

- make `windows-portable` the primary download
- include `runtime\python` in that artifact so users do not need to install Python
- keep Codex CLI documented as an optional add-on instead of a setup prerequisite

## Troubleshooting

- If Outlook access fails, verify classic Outlook is installed and signed in on Windows.
- If Outlook is busy, open it first and rerun.
- If the browser workspace cannot open Outlook items, verify `EntryID` and `StoreID` are present in the stored rows.
- If abstracts are slow on first run, that is expected; new emails are summarized through Codex CLI and then cached under `.cache/`.
- If cached abstracts look stale after prompt changes, delete `.cache/mailhandle_abstracts.json` and rerun.
