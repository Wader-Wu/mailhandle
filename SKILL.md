---
name: mailhandle
description: Package the standalone Windows `mailhandle` Outlook workspace as a Codex CLI skill, launch its local browser UI, and tune priority behavior with editable JSON rules.
---

# Mailhandle

## Use When

- The user wants recent Outlook mail from their own desktop mailbox.
- The user wants a local browser workspace for reviewing mail history and status.
- The user wants priority filtering, grouped threads, Outlook open actions, or reply drafting.
- The user wants to tune matching behavior in `priority_rules.json`.

## Workflow

1. Default to Codex CLI usage and treat `$mailhandle start` as the canonical prompt.
2. Accept `$mailhandle run`, `$mailhandle launch`, `mailhandle start`, `mailhandle run`, and `mailhandle launch` as equivalent requests.
3. For those prompts, do not inspect unrelated project files first.
4. Start the workspace immediately with the PowerShell-native command `& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"`.
5. Report the launcher output directly. `launch_mailhandle.ps1` waits for the startup summary and prints the workspace URL or current status itself.
6. If needed, the same status is also written to `%USERPROFILE%\.codex\skills\mailhandle\tmp\mailhandle-last-start.txt`.
7. Do not run `python scripts/run_mail_database.py` directly from Codex for normal startup.
8. The launcher requests browser auto-open and also prints the local URL. If the browser does not appear, tell the user to open the reported URL manually.
9. The page provides:
   - mailbox history from SQLite
   - time-range filters such as `today`, `last_2days`, and `last_7_days`
   - priority filters and a default `All_Open` status filter that hides `Done` items
   - inline `status` editing
   - an `Open` action for a specific Outlook item
   - thread-level response drafting that opens `Reply All` in Outlook
   - an embedded priority-rules editor
10. If the user asks to update priority rules, use the embedded editor in the workspace.

## Commands

Database workspace:
```bash
& "$env:USERPROFILE\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

Manual `cmd.exe` launch:
```bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%USERPROFILE%\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"
```

Prompt forms:
```text
$mailhandle start
$mailhandle run
$mailhandle launch
```

## Rules

- This skill depends on local classic Outlook on Windows and Windows Python configured in `scripts/.env` for installed/runtime use.
- The underlying runtime is a standalone Windows tool; the Codex skill is one distribution and launch mode, not a special host environment.
- Treat Codex CLI as the release execution mode for this skill.
- Prefer the prompt verb `start`; treat `run` and `launch` as equivalent aliases for opening the workspace.
- On `start`/`run`/`launch`, execute the launcher command directly instead of reading README or exploring the workspace first.
- Use the PowerShell-native `launch_mailhandle.ps1` entrypoint from Codex sessions.
- If the user asks for a `cmd.exe` command, provide `powershell -NoProfile -ExecutionPolicy Bypass -File "%USERPROFILE%\.codex\skills\mailhandle\scripts\launch_mailhandle.ps1"` instead of defaulting to `start_mailhandle.cmd`.
- Report the launcher output directly; it comes from `%USERPROFILE%\.codex\skills\mailhandle\tmp\mailhandle-last-start.txt`.
- Use the launcher instead of running `run_mail_database.py` directly from Codex, because the workspace server is a long-lived local process and the launcher handles startup plus browser auto-open.
- Codex CLI is optional for the runtime itself; without the `codex` executable, startup and core workspace features still work, but LLM-backed abstracts and reply drafting do not.
- Keep the documented workflow centered on `scripts/run_mail_database.py`; do not reintroduce removed summary/report entrypoints.
- The SQLite database is the source of truth for review state once initialized.
- Mail abstracts default to Codex-generated body summaries with a local cache; first runs for unseen emails are slower than repeat runs.
- Tune personal priority behavior in `scripts/priority_rules.json`, especially `owner_aliases` and `manager_senders`.
- `manager_senders` can be configured as a display name, SMTP address, or `Name <email>`.
- When priority rules need editing, use the browser editor flow instead of hand-editing JSON directly.
- Priority rule edits are forward-looking. They affect future synced items and should not be described as retroactively rescoring historical SQLite rows unless the user explicitly asks for a rebuild/reset.
- Response drafts should not include signatures or contact blocks because Outlook adds the configured signature.
- For GitHub releases, package only the skill source/docs and exclude local artifacts like `.cache`, `data`, `records`, `sessions`, `tmp`, `log`, `__pycache__`, `*.pyc`, `.env`, and local state files.
- For release prep, update `references/release-notes-next.md` before packaging.
- Keep summaries grounded in the returned fields. Do not invent senders, actions, or deadlines.

## References

Read `README.md` for:
- standalone installation and setup
- optional Codex CLI prerequisites
- Windows/Outlook compatibility

Read `references/runbook.md` for:
- available arguments
- project file locations
- troubleshooting notes
- release packaging and validation steps

Read `references/database-design.md` for:
- the SQLite-backed history model
- grouping and sync rules
- DPAPI encryption strategy
- browser workspace behavior

Read `references/release-notes-next.md` for:
- the current draft release summary
- the pre-release validation checklist
