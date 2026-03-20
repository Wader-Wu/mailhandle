# Next Release Notes

Draft notes for the next `mailhandle` release.

## Highlights

- clarified that `mailhandle` is a standalone Windows runtime with optional Codex CLI features
- launcher now requests browser auto-open and documents the preferred PowerShell and `cmd.exe` entrypoints
- startup sync status in the workspace now refreshes correctly instead of staying on `Starting sync...`
- `Sent Items` now syncs after `Inbox`, merges into the same thread groups, and supports LLM-backed abstracts when Codex CLI is available
- the default status filter is `All_Open`, which hides `Done` items
- marking an item `Done` automatically forces its effective priority to `Low` without letting later sync overwrite that state
- timestamps in the UI display in the local computer timezone without fractional seconds
- Outlook sender resolution now uses SMTP-aware matching so manager rules can match values like `Bin Tan <Bin.Tan@lumentum.com>`
- database backup uses a single rolling same-folder copy at `data\mailhandle.backup.sqlite`
- incremental sync now uses the stored watermark window instead of the old end-time-only behavior
- reply and new-email drafting now share one structured LLM schema, with `New Email` setting the Outlook subject and appending the generated closing while `Response` keeps the simpler reply-body paste flow
- launcher now prefers a bundled `runtime\python\python.exe`, which enables a portable Windows release without requiring Python on the target machine
- added `scripts\install_mailhandle.ps1` as a one-command Windows bootstrap that installs Python, Node.js, `pywin32`, and Codex CLI, then launches the app
- the default sync period is now configurable through `scripts\priority_rules.json` and now defaults to `last_1day`
- removed the old `boost_owner_attention` toggle so owner-directed priority is controlled only through explicit rules

## Behavioral Notes

- editing `scripts\priority_rules.json` is forward-looking and affects future synced items
- historical rows already stored in SQLite keep their saved priority/history unless the database is explicitly rebuilt or replaced
- `manager_senders` can be entered as a display name, SMTP address, or `Name <email>`
- the easiest release path for new Windows users is now a `windows-portable` zip built with `scripts\build_release.ps1` and a bundled Python runtime

## Pre-Release Validation

Run:

```powershell
Get-ChildItem scripts -Filter *.py | ForEach-Object { python -m py_compile $_.FullName }
```

Launch:

```powershell
& ".\scripts\launch_mailhandle.ps1"
```

Check:

```powershell
Get-Content ".\tmp\mailhandle-last-start.txt"
```

Release bundle exclusions:

- `data/`
- `tmp/`
- `.cache/`
- local `.env` files
- `__pycache__/`
- `*.pyc`
