# Runbook

This project is a standalone Windows runtime that can also be packaged as a Codex skill.

## Runtime Modes

- Standalone Windows app: launch `scripts\launch_mailhandle.ps1` from any copied install folder.
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

## Wrapper Usage

Primary standalone workflow from the install root:

```powershell
& ".\scripts\launch_mailhandle.ps1"
```

Manual `cmd.exe` launch:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\launch_mailhandle.ps1"
```

From any folder, replace `.\` with the absolute path to your install.

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

`scripts\start_mailhandle.cmd` remains available as a thin wrapper, but the direct `cmd.exe` command above is the recommended manual CMD entrypoint.

## Optional Codex CLI Features

If the `codex` executable is on `PATH`, the runtime can use it for:

- LLM-generated mail abstracts
- reply drafting for a thread

Without Codex CLI:

- startup still works
- sync still works
- the browser workspace still works
- reply drafting is unavailable

To edit `priority_rules.json`, use the embedded browser editor in the workspace.

## Priority Tuning

Useful fields in `scripts/priority_rules.json`:

- `owner_aliases`
- `manager_senders`
- `greeting_terms`

Notes:

- `manager_senders` can match `Bin Tan`, `Bin.Tan@lumentum.com`, or `Bin Tan <Bin.Tan@lumentum.com>`
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

## Troubleshooting

- If Outlook access fails, verify classic Outlook is installed and signed in on Windows.
- If Outlook is busy, open it first and rerun.
- If the browser workspace cannot open Outlook items, verify `EntryID` and `StoreID` are present in the stored rows.
- If abstracts are slow on first run, that is expected; new emails are summarized through Codex CLI and then cached under `.cache/`.
- If cached abstracts look stale after prompt changes, delete `.cache/mailhandle_abstracts.json` and rerun.
