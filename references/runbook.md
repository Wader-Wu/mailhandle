# Runbook

## Source Layout

- Skill root: the installed `mailhandle` folder
- Runtime code: `scripts/`
- Runtime rule config: `scripts/priority_rules.json`
- Runtime env template: `scripts/.env.example`
- SQLite history design: `references/database-design.md`
- Local DB path: `data/mailhandle.sqlite`

## GitHub Release Packaging

Include:

- `README.md`
- `SKILL.md`
- `references/`
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

Primary workflow:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

Manual `cmd.exe` launch:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%USERPROFILE%\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

Canonical Codex prompt:

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
Get-Content "$env:USERPROFILE\.codex\skills\mailhandle\tmp\mailhandle-last-start.txt"
```

or

```bat
type "%USERPROFILE%\.codex\skills\mailhandle\tmp\mailhandle-last-start.txt"
```

If the browser does not appear, open the reported workspace URL manually.

Direct `python scripts/run_mail_database.py` is still valid for manual debugging, but it should not be the default Codex skill entrypoint because it keeps the CLI attached to the long-lived server process.

`scripts\start_mailhandle.cmd` remains available as a thin wrapper, but the direct `cmd.exe` command above is the recommended manual CMD entrypoint.

To edit `priority_rules.json`, use the embedded browser editor in the workspace.

## Priority Tuning

Useful fields in `scripts/priority_rules.json`:

- `owner_aliases`
- `manager_senders`
- `greeting_terms`

## Troubleshooting

- If Outlook access fails, verify classic Outlook is installed and signed in on Windows.
- If Outlook is busy, open it first and rerun.
- If the browser workspace cannot open Outlook items, verify `EntryID` and `StoreID` are present in the stored rows.
- If abstracts are slow on first run, that is expected; new emails are summarized through Codex CLI and then cached under `.cache/`.
- If cached abstracts look stale after prompt changes, delete `.cache/mailhandle_abstracts.json` and rerun.
