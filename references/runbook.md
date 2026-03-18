# Runbook

## Source Layout

- Skill root: the installed `mailhandle` folder
- Runtime code: `scripts/`
- Runtime rule config: `scripts/priority_rules.json`
- Runtime env template: `scripts/.env.example`
- Release output from the source repo: `release/mailhandle/`
- Installed skill target in Codex: `~/.codex/skills/mailhandle`
- Installed skill target in OpenClaw: `~/.openclaw/workspace/skills/mailhandle`

## GitHub Release Packaging

When building a GitHub release, include only the skill runtime and documentation needed by another user to run it.

Keep:

- `README.md`
- `SKILL.md`
- `references/`
- `scripts/`

Ignore:

- `.cache/`
- `records/`
- `sessions/`
- `tmp/`
- `log/`
- `.sandbox/`
- `.sandbox-bin/`
- `__pycache__/`
- `*.pyc`
- `auth.json`
- `history.jsonl`
- `models_cache.json`
- `state_*.sqlite*`
- any local `.env` file

If you publish a zip or release asset manually, apply the ignore list before uploading so the bundle stays reproducible and does not leak local state.

## Wrapper Usage

Default usage is via Codex CLI with the installed skill.

From the installed skill root, use:

```bash
python3 scripts/run_summary.py --json
```

This now always:

- builds the summary result
- writes a timestamped summary JSON and static HTML review report under `records/YYYY-MM-DD/`
- tries to open the generated HTML report
- asks Codex CLI to summarize mail bodies into cleaner abstracts, then caches those abstracts locally for repeat runs

On native Windows/Codex sessions, `python scripts/run_summary.py --json` is the expected
debug command when validating the installed skill outside Codex CLI. The Windows Outlook
reader auto-detects the current Python interpreter.

For a self-contained static review page with checkboxes and exportable review state, use:

```bash
python scripts/run_review_report.py --date-preset today
```

This writes a timestamped summary JSON and HTML file under `records/YYYY-MM-DD/`.

To edit `priority_rules.json` in a browser and save it back directly, use:

```bash
python scripts/run_priority_rules_editor.py
```

This starts a local HTTP server, opens a form-based browser editor, validates the generated JSON, and writes changes back to `scripts/priority_rules.json`.

For local debugging from the source repo:

```bash
cd /path/to/mailhandle
source .venv/bin/activate
python scripts/summarize_mail.py --date-preset today
```

To build the release and install into Codex from the source repo:

```bash
cd /path/to/mailhandle
python3 .dev/publish_skill.py --install-target codex
```

To build once and install into both Codex and OpenClaw:

```bash
cd /path/to/mailhandle
python3 .dev/publish_skill.py --install-target both
```

Supported arguments:

- `--limit N`
- `--from-contains TEXT`
- `--subject-contains TEXT`
- `--project TEXT`
- `--date-preset today|yesterday|last_7_days|this_month|last_month`
- `--since YYYY-MM-DD`
- `--until YYYY-MM-DD`
- `--unread-only`
- `--include-body`
- `--include-notifications`
- `--no-collapse`
- `--verbose`
- `--json`

The HTML report builder also supports:

- `--input-json PATH`
- `--output PATH`
- `--output-dir DIR`
- `--title TEXT`

Current summary behavior:

- reads filtered inbox mail first
- checks `Sent Items` for later replies in the same thread
- marks todos with response metadata when a reply is found
- keeps replied items visible so OpenClaw can suggest monitoring instead of duplicate action
- uses Codex CLI to generate per-email abstracts from cleaned body content when Codex is available
- keeps a local abstract cache under `.cache/` so repeated runs avoid re-summarizing the same mail
- plain-text reports split results into `Needs action` and `Already replied`
- static HTML reports keep local browser-side review state for done checkboxes and can export reviewed HTML/JSON snapshots

## Priority Tuning

Edit `scripts/priority_rules.json`.

Supported rule keys:

- `priority`
- `unread`
- `importance`
- `subject_contains_any`
- `body_contains_any`
- `sender_contains_any`
- `attention_flags_any`
- `sender_matches_manager`
- `categories_any`
- `received_within_days`

Behavior flags in `priority_rules.json`:

- `suppress_low_priority_notifications`
- `collapse_similar_emails`

Priority-tuning fields in `priority_rules.json`:

- `owner_aliases`
- `manager_senders`
- `greeting_terms`

## Troubleshooting

- If Outlook access fails, verify classic Outlook is installed and signed in on Windows.
- If `Sent Items` looks empty, confirm Outlook is running in classic mode and mailbox folders have finished syncing.
- If Python path issues appear during development, check `WINDOWS_PYTHON_EXE` in the repo's `.env`.
  Native Windows runs should not need this override; it is mainly for WSL or custom Python installs.
- If abstracts are slow on the first run, that is expected; new emails are summarized through Codex CLI once and then cached under `.cache/`.
- If cached abstracts look stale or wrong after prompt changes, delete `.cache/mailhandle_abstracts.json` and rerun.
- If the installed skill behaves differently, republish from the source repo with `python3 .dev/publish_skill.py --install-target codex` or `--install-target both`.
- If Outlook is busy, rerun the command; the project already retries common COM busy errors.
