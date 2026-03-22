---
name: mailhandle
description: Sync recent Outlook mail into the local `mailhandle` SQLite database and answer follow-up summary requests from local data.
---

# Mailhandle

## Use When

- The user wants recent Outlook mail from their own Windows desktop mailbox.
- The user wants to initialize or refresh the local mail database.
- The user wants command-based mailbox review without starting a webpage.
- The user wants thread summaries, priority summaries, follow-up summaries, or quick mailbox reporting from local data.
- The user wants to open a specific Outlook item or draft a reply/new mail from the terminal flow.
- The user wants to tune matching behavior in `scripts/priority_rules.json`.

## Workflow

1. Treat `$mailhandle start` as the canonical skill action for command-based CLI review.
2. Accept `$mailhandle run` and `$mailhandle launch` as equivalent CLI `overview` requests.
3. On those requests, run the launcher in CLI mode:
   ```powershell
   & "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1" -Mode cli overview --json
   ```
4. If `overview --json` returns `used_cached_data: true` or a non-empty `sync_error`, report that the snapshot is from cached SQLite data and tell the user a fresh sync still requires classic Outlook to be open and ready.
5. Treat `$mailhandle sync` as the explicit refresh command:
   ```powershell
   & "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1" -Mode cli sync --json
   ```
6. If `sync --json` returns `ok: false`, report the `sync_error`, mention `last_sync_end` if present, and only describe mailbox state as cached local data.
7. For mailbox review actions, use the CLI subcommands directly instead of opening the browser workspace.
8. Report the command output directly and answer follow-up questions from the local scripts and SQLite-backed runtime state.
9. Do not involve browser UI tasks, GUI walkthroughs, or embedded editor flows in the default Codex skill behavior.
10. Only use GUI mode if the user explicitly asks for the webpage/workspace/browser UI.

## Commands

CLI overview from a Codex skill install:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1" -Mode cli overview --json
```

Explicit sync:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1" -Mode cli sync --json
```

List current open groups:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1" -Mode cli list --json
```

Show one thread:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1" -Mode cli show "<group-key-or-email-id>" --json
```

Update status:

```powershell
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1" -Mode cli status "<email-id>" done --json
```

Manual `cmd.exe` launch for CLI overview:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%USERPROFILE%\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1" -Mode cli overview --json
```

Useful prompt forms:

```text
$mailhandle start
$mailhandle run
$mailhandle launch
$mailhandle sync
```

## Required Config

Supported skill-mode configuration lives in:

- `scripts/priority_rules.json`
  - required
  - the only user-facing configuration file for this skill
  - controls priority logic and the default sync period through `default_sync_period`

Runtime prerequisites outside config files:

- classic Outlook installed, signed in, and already open
- Windows PowerShell and Python available to run the local launcher and CLI scripts

## Rules

- This skill depends on local classic Outlook on Windows.
- Read access is limited to an already opened classic Outlook session; the skill should not try to start Outlook on the user's behalf.
- The underlying runtime is a standalone Windows tool; the Codex skill is only one launch mode.
- The SQLite database is the source of truth for review state once initialized.
- Default to CLI mode via `scripts/launch_mailhandle.ps1 -Mode cli`.
- Use the launcher or `scripts/mailhandle_cli.py` for Codex-driven review actions.
- Use `sync` when the user explicitly asks to refresh/init the database.
- If `sync --json` returns `ok: false`, treat that as a live-sync failure, not a successful refresh.
- `scripts/priority_rules.json` must exist for sync because the runtime reads it during initialization.
- Treat `scripts/priority_rules.json` as the only supported configuration update path in skill mode.
- Do not make GUI interaction part of the default skill flow.
- Use local scripts and stored data for summary answers after sync.
- Use `overview` for a fast mailbox snapshot, `list` for grouped results, and `show` for one detailed thread.
- If `overview` falls back to cached data, say that explicitly instead of treating the run as a launcher failure.
- Use `status`, `open`, `reply-draft`, `reply-open`, `new-email-draft`, and `new-email-open` when the user asks for those concrete actions.
- If the user explicitly asks for the webpage/browser workspace, use GUI mode with `scripts/launch_mailhandle.ps1` and say that this is the non-default path.
- Codex CLI is optional for the runtime itself; without the `codex` executable, sync and local review still work, but LLM-backed abstracts and drafting do not.
- Tune personal priority behavior in `scripts/priority_rules.json`, especially `default_sync_period`, `owner_aliases`, and `manager_senders`.
- `manager_senders` can be configured as a display name, SMTP address, or `Name <email>`.
- In skill mode, all supported configuration updates should be made by editing `scripts/priority_rules.json` manually.
- Priority rule edits are forward-looking. They affect future synced items and should not be described as retroactively rescoring historical SQLite rows unless the user explicitly asks for a rebuild/reset.
- Keep summaries grounded in the returned fields. Do not invent senders, actions, or deadlines.

## References

Read `README.md` for:

- what the tool does
- installation
- GUI and CLI usage

Read `references/runbook.md` for:

- CLI commands
- runtime file locations
- packaging and validation steps
- troubleshooting notes

Read `references/database-design.md` for:

- the SQLite-backed history model
- grouping and sync rules
- DPAPI protection notes
