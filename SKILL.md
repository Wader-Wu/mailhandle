---
name: mailhandle
description: Read Outlook mail from Codex CLI on native Windows, filter messages by sender/subject/date/unread state, and summarize matching mail into prioritized todos using the runtime code under `scripts/` plus editable rules in `scripts/priority_rules.json`.
---

# Mailhandle

## Use When

- The user wants recent Outlook mail from their own desktop mailbox.
- The user asks for filtered reading like unread mail, yesterday's mail, this month's project mail, or sender/subject search.
- The user wants a todo list or prioritized summary from matching mail.
- The user wants a readable static HTML review report for personal follow-up.

## Workflow

1. Default to Codex CLI usage and use the bundled wrapper script `scripts/run_summary.py`.
2. Pass only the filters needed:
   - `--unread-only`
   - `--date-preset today|yesterday|last_7_days|this_month|last_month`
   - `--from-contains`
   - `--subject-contains`
   - `--project`
   - `--since YYYY-MM-DD`
   - `--until YYYY-MM-DD`
   - `--limit N`
   - `--include-body` only when the body is really needed
   - `--include-notifications` only when notification-style mail should stay visible
   - `--no-collapse` only when every matching mail should remain separate
   - `--verbose` only when the full raw nested message object is needed
3. Prefer `--json` for downstream handling.
4. Runtime code and runtime config live under `scripts/`.
5. `scripts/run_summary.py` now always generates a static HTML review page, tries to open it on Windows, and uses Codex CLI to build mail-body abstracts when available.
6. Use `scripts/run_review_report.py` when the user explicitly wants the HTML report flow by itself.
7. If the user asks to update priority rules, prefer `scripts/run_priority_rules_editor.py` so the JSON file opens in a browser editor and saves back directly.

## Commands

Default structured summary from Codex:
```bash
python3 skills/mailhandle/scripts/run_summary.py --json
```

This command also saves a timestamped HTML review report and opens it when possible.

For native Windows/Codex sessions, `python skills/mailhandle/scripts/run_summary.py --json`
is the expected local debug command when direct script execution is needed.

Unread mail today:
```bash
python3 skills/mailhandle/scripts/run_summary.py --unread-only --date-preset today --json
```

Yesterday's mail:
```bash
python3 skills/mailhandle/scripts/run_summary.py --date-preset yesterday --json
```

Project-specific mail:
```bash
python3 skills/mailhandle/scripts/run_summary.py --project zhuque --date-preset this_month --json
```

Static HTML review report:
```bash
python3 skills/mailhandle/scripts/run_review_report.py --date-preset today
```

Priority rules editor:
```bash
python3 skills/mailhandle/scripts/run_priority_rules_editor.py
```

## Rules

- This skill depends on local classic Outlook on Windows and Windows Python configured in `scripts/.env` for installed/runtime use.
- Run through the wrapper script instead of calling `scripts/read_outlook_win.py` directly from the skill.
- Treat Codex CLI as the primary execution mode. Use direct Python execution only for debugging.
- Mail abstracts now default to Codex-generated body summaries with a local cache; the first run for unseen emails is slower than repeat runs.
- Use the HTML report flow when the user wants a readable personal review artifact with checkboxes and exportable review state.
- Tune personal priority behavior in `scripts/priority_rules.json`, especially `owner_aliases` and `manager_senders`.
- When priority rules need editing, use the browser editor flow instead of asking the user to hand-edit JSON directly. The editor is form-based and safer than raw JSON editing.
- For GitHub releases, package only the skill source/docs and exclude local artifacts like `.cache`, `records`, `sessions`, `tmp`, `log`, `__pycache__`, `*.pyc`, `.env`, and local state files.
- Keep summaries grounded in the returned fields. Do not invent senders, actions, or deadlines.
- If the user asks to change priority logic, edit `scripts/priority_rules.json`.

## References

Read `references/runbook.md` for:
- available arguments
- project file locations
- troubleshooting notes

Read `references/example-prompts.md` for:
- ready-to-use OpenClaw prompts
- common mailbox review patterns
