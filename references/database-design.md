# Database Design

This document defines the SQLite-backed history layer for the standalone Windows `mailhandle` runtime, which can also be packaged as a Codex skill.

## Goal

Use one local SQLite database file to centralize mailbox review history for a single Windows user.

The database should:

- preserve historical mail summaries
- keep user review state such as `status`
- avoid storing raw mail bodies
- support grouped display for threads or repeated subjects
- allow the webpage to update state and save changes immediately
- support opening a specific Outlook item from the UI

The database is a local runtime artifact, not a shared server database.

## Storage Location

Recommended path:

```text
<skill-root>\data\mailhandle.sqlite
```

Reason:

- keeps runtime state separate from `records/`
- makes the data file easy to back up or move
- keeps the database outside the release bundle when `.gitignore` is applied

If needed later, the path can be made configurable through `scripts/.env`.

## Data Rules

Store:

- `email_id` as the unique mail key
- summary metadata already present in the JSON summary output
- `abstract`
- `status`
- `notes`
- `thread_key`
- timestamps and review audit fields

Do not store:

- raw mail body text
- Codex prompt text
- Outlook COM objects
- other transient runtime caches

## Field Protection

Use DPAPI-protected keys for sensitive text fields.

Recommended split:

- plaintext fields: `email_id`, `subject`, `from`, `received_at`, `priority`, `next_action`, `thread_key`, `status`, timestamps
- encrypted fields: `abstract`, `notes`

Reason:

- plaintext fields are needed for search, sorting, and grouping
- abstract and notes may contain the most sensitive content

## Core Tables

### `mail_items`

One row per unique email.

Suggested columns:

- `email_id` TEXT primary key
- `entry_id` TEXT
- `thread_key` TEXT indexed
- `subject` TEXT
- `subject_key` TEXT indexed
- `conversation_topic` TEXT
- `from_name` TEXT
- `from_email` TEXT
- `received_at` TEXT
- `folder` TEXT
- `priority` TEXT
- `next_action` TEXT
- `abstract_enc` BLOB
- `status` TEXT
- `notes_enc` BLOB
- `responded` INTEGER
- `responded_at` TEXT
- `response_subject` TEXT
- `is_group_root` INTEGER
- `first_seen_at` TEXT
- `last_seen_at` TEXT
- `created_at` TEXT
- `updated_at` TEXT

Indexes:

- `thread_key`
- `subject_key`
- `received_at`
- `status`
- `priority`

### `mail_runs`

One row per scan run.

Suggested columns:

- `run_id` TEXT primary key
- `started_at` TEXT
- `ended_at` TEXT
- `date_preset` TEXT
- `since` TEXT
- `until` TEXT
- `folder` TEXT
- `limit_value` INTEGER
- `raw_count` INTEGER
- `stored_count` INTEGER
- `updated_count` INTEGER
- `created_at` TEXT

### `mail_item_runs`

Use this only if you want a full "seen in this run" history.

Suggested columns:

- `run_id` TEXT
- `email_id` TEXT
- `seen_at` TEXT
- `result` TEXT

This table is not required by the current release workflow.

## Grouping Logic

Use a stable `thread_key` to show related mail together in the webpage.

Priority order:

1. Use Outlook `conversation_topic` if present.
2. Otherwise use a normalized subject.
3. Normalize by removing common prefixes like `Re:`, `Fwd:`, `FW:`, and `[External]`.
4. Collapse whitespace and punctuation noise.

If multiple items share the same `thread_key`, the UI should group them together.

Do not use raw subject alone as the only grouping key, because reply prefixes and notification subjects are too noisy.

## Sync Rules

Startup behavior:

- scan the last 7 days on first run
- on later runs, refresh incrementally from the last recorded sync watermark (`until`, fallback `ended_at`) to a query upper bound captured at sync start
- fetch summary data from the existing mail pipeline
- insert only items whose `email_id` is not already in the database
- keep existing `status`
- update summary metadata if the mail is seen again and newer metadata is available
- treat priority rules as forward-looking; changing `priority_rules.json` should affect future synced items, not retroactively rescore historical rows already stored in SQLite

Update behavior:

- `status` is edited from the webpage
- changes are written back immediately
- the database is the source of truth for review state

History behavior:

- each run creates a `mail_runs` record
- the mail item row remains the canonical current-state record
- if future audit history is needed, populate `mail_item_runs`

## Web UI Behavior

The webpage should:

- list mail grouped by `thread_key`
- show current `status`
- allow filtering by priority, status, sender, and date
- support local search
- provide an `Open` action for one mail item
- provide thread-level response drafting that opens `Reply All` in Outlook
- embed the priority-rules editor in the same browser workspace

## Outlook Open Action

The simplest reliable approach is a local helper on Windows:

- the webpage sends `email_id`
- the helper looks up the matching row
- the helper opens the Outlook item by `EntryID`

This should be implemented as a local Windows action, not as a browser-only link.

## Security Model

The expected security model is:

- local single-user machine
- NTFS permissions limit access to the current Windows user
- DPAPI protects encrypted fields
- local-only webpage on `localhost`

This does not make the database secret against a machine admin, but it prevents casual file-copy disclosure.

## Implementation Plan

1. Add SQLite schema and migration helpers.
2. Add a storage layer that upserts mail rows from the JSON summary.
3. Add DPAPI key management and encrypted field helpers.
4. Add startup sync for the last 7 days.
5. Add a database-backed webpage that edits `status`.
6. Add an Outlook-open helper for `email_id`.
7. Keep the existing summary pipeline as the ingestion source.

## Runtime Entry Points

- `scripts/launch_mailhandle.ps1`: primary launcher for PowerShell and Codex sessions; starts the workspace in the background and requests browser auto-open
- `scripts/start_mailhandle.ps1`: launcher implementation that starts `run_mail_database.py --open-browser`
- `scripts/start_mailhandle.cmd`: optional CMD wrapper; direct `powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\launch_mailhandle.ps1"` is the recommended manual `cmd.exe` command from the install root
- `scripts/run_mail_database.py`: syncs the last 7 days, serves the local browser UI, and opens Outlook items
- `scripts/mailhandle_db.py`: SQLite storage and DPAPI helpers
- `data/mailhandle.sqlite`: local single-user database file
