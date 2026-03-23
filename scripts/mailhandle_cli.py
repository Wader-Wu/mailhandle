#!/usr/bin/env python3
import argparse
import json
import subprocess
import sys
from pathlib import Path

import mailhandle_db
import mailhandle_runtime
import summarize_mail


VIEW_PRESETS = ["all", *mailhandle_runtime.DATE_PRESETS]
STATUS_CHOICES = ["all_open", "all", "todo", "doing", "done"]
PRIORITY_CHOICES = ["all", "high", "medium", "low"]
PRIORITY_ORDER = {"high": 0, "medium": 1, "low": 2}


def add_sync_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--date-preset", choices=mailhandle_runtime.DATE_PRESETS, default=summarize_mail.get_default_sync_period())
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--from-contains")
    parser.add_argument("--subject-contains")
    parser.add_argument("--project")
    parser.add_argument("--unread-only", action="store_true", default=False)
    parser.add_argument("--include-body", action="store_true")
    parser.add_argument("--include-notifications", action="store_true")
    parser.add_argument("--no-collapse", action="store_true")
    parser.add_argument("--verbose", action="store_true")


def add_view_filter_arguments(
    parser: argparse.ArgumentParser,
    *,
    default_preset: str = "last_7_days",
    default_status: str = "all_open",
) -> None:
    parser.add_argument("--range", "--view-preset", dest="view_preset", choices=VIEW_PRESETS, default=default_preset)
    parser.add_argument("--since")
    parser.add_argument("--until")
    parser.add_argument("--status", choices=STATUS_CHOICES, default=default_status)
    parser.add_argument("--priority", choices=PRIORITY_CHOICES, default="all")
    parser.add_argument("--search")


def add_json_flag(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--json", action="store_true", default=False)


def add_notes_arguments(parser: argparse.ArgumentParser) -> None:
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--notes")
    group.add_argument("--notes-file")


def add_draft_arguments(parser: argparse.ArgumentParser) -> None:
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--draft")
    group.add_argument("--draft-file")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Mailhandle command-based CLI. This mode uses the local SQLite mailbox history and Outlook actions "
            "without starting the web workspace."
        )
    )
    subparsers = parser.add_subparsers(dest="command")

    overview = subparsers.add_parser("overview", help="Sync recent mail, then print a filtered mailbox overview.")
    add_sync_arguments(overview)
    overview.add_argument("--bootstrap", action="store_true", default=False)
    overview.add_argument("--skip-sync", action="store_true", default=False)
    add_view_filter_arguments(overview)
    add_json_flag(overview)

    sync = subparsers.add_parser("sync", help="Sync Outlook mail into the local SQLite database.")
    add_sync_arguments(sync)
    sync.add_argument("--bootstrap", action="store_true", default=False)
    add_json_flag(sync)

    list_parser = subparsers.add_parser("list", help="List grouped mail threads from the local database.")
    add_view_filter_arguments(list_parser)
    add_json_flag(list_parser)

    show = subparsers.add_parser("show", help="Show one grouped thread by group key or email id.")
    show.add_argument("identifier", help="Thread key, subject key, or email id.")
    show.add_argument("--include-body", action="store_true", default=False, help="Attempt to reopen Outlook items and include cleaned body text.")
    add_json_flag(show)

    status = subparsers.add_parser("status", help="Update the local review status for one email item.")
    status.add_argument("email_id")
    status.add_argument("status", choices=["todo", "doing", "done"])
    add_notes_arguments(status)
    add_json_flag(status)

    open_mail = subparsers.add_parser("open", help="Open one stored Outlook item by email id.")
    open_mail.add_argument("email_id")

    reply_draft = subparsers.add_parser("reply-draft", help="Generate a reply-all draft for a thread.")
    reply_draft.add_argument("identifier", help="Thread key, subject key, or email id.")
    add_notes_arguments(reply_draft)
    reply_draft.add_argument("--save", help="Write the generated draft JSON to a file.")
    add_json_flag(reply_draft)

    reply_open = subparsers.add_parser("reply-open", help="Open Reply All in Outlook for a thread, optionally with a prepared draft.")
    reply_open.add_argument("identifier", help="Thread key, subject key, or email id.")
    add_draft_arguments(reply_open)

    new_email_draft = subparsers.add_parser("new-email-draft", help="Generate a brand-new email draft.")
    add_notes_arguments(new_email_draft)
    new_email_draft.add_argument("--save", help="Write the generated draft JSON to a file.")
    add_json_flag(new_email_draft)

    new_email_open = subparsers.add_parser("new-email-open", help="Open a new Outlook email, optionally with a prepared draft.")
    add_draft_arguments(new_email_open)

    args_list = list(argv if argv is not None else sys.argv[1:])
    if not args_list:
        args_list = ["overview"]
    return parser.parse_args(args_list)


def _read_optional_text(*, inline_value: str | None = None, file_path: str | None = None) -> str:
    if file_path:
        return Path(file_path).read_text(encoding="utf-8")
    return str(inline_value or "")


def _write_text(path: str, text: str) -> Path:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text, encoding="utf-8")
    return output_path


def _parse_draft_payload(draft_text: str):
    try:
        return json.loads(draft_text)
    except Exception:
        return draft_text


def _build_view_filters(args: argparse.Namespace) -> dict[str, str]:
    since, until = mailhandle_runtime.resolve_db_time_window(
        date_preset=getattr(args, "view_preset", ""),
        since=getattr(args, "since", ""),
        until=getattr(args, "until", ""),
    )
    return {
        "since": since,
        "until": until,
        "status": str(getattr(args, "status", "all_open") or "all_open"),
        "priority": str(getattr(args, "priority", "all") or "all"),
        "search": str(getattr(args, "search", "") or "").strip(),
    }


def _preserve_unique(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        text = str(value or "").strip()
        key = text.lower()
        if not text or key in seen:
            continue
        seen.add(key)
        result.append(text)
    return result


def _group_summary(group: dict) -> dict:
    items = list(group.get("items") or [])
    priorities = sorted(
        _preserve_unique([str(item.get("priority") or "").lower() for item in items]),
        key=lambda value: PRIORITY_ORDER.get(value, 99),
    )
    status_counts = {"todo": 0, "doing": 0, "done": 0}
    for item in items:
        status = str(item.get("status") or "").lower()
        if status in status_counts:
            status_counts[status] += 1
    senders = _preserve_unique([str(item.get("from_name") or item.get("from") or "") for item in items])
    return {
        "group_key": group.get("group_key", ""),
        "title": group.get("title", ""),
        "count": int(group.get("count", 0) or 0),
        "latest_received": group.get("latest_received", ""),
        "oldest_received": group.get("oldest_received", ""),
        "latest_email_id": group.get("latest_email_id", ""),
        "reply_target_email_id": group.get("reply_target_email_id", ""),
        "priorities": priorities,
        "status_counts": status_counts,
        "senders": senders,
    }


def _filters_summary(filters: dict[str, str]) -> str:
    parts: list[str] = []
    if filters.get("since") or filters.get("until"):
        parts.append(f"date={filters.get('since') or '*'}..{filters.get('until') or '*'}")
    status = filters.get("status") or "all_open"
    priority = filters.get("priority") or "all"
    parts.append(f"status={status}")
    parts.append(f"priority={priority}")
    search = filters.get("search") or ""
    if search:
        parts.append(f"search={search}")
    return ", ".join(parts)


def _print_group_list(groups: list[dict], *, filters: dict[str, str], stats: dict[str, int], last_sync_end: str) -> None:
    print(f"Mailbox groups: {len(groups)}")
    if last_sync_end:
        print(f"Last sync: {last_sync_end}")
    print(f"Filters: {_filters_summary(filters)}")
    print(
        "Stats:",
        ", ".join(
            [
                f"total={stats.get('total', 0)}",
                f"high={stats.get('high', 0)}",
                f"medium={stats.get('medium', 0)}",
                f"low={stats.get('low', 0)}",
                f"todo={stats.get('todo', 0)}",
                f"doing={stats.get('doing', 0)}",
                f"done={stats.get('done', 0)}",
            ]
        ),
    )
    if not groups:
        print("No matching groups.")
        return

    for index, group in enumerate(groups, start=1):
        summary = _group_summary(group)
        priority_text = ", ".join(summary["priorities"]) if summary["priorities"] else "-"
        status_text = ", ".join(f"{key}={value}" for key, value in summary["status_counts"].items() if value)
        senders_text = ", ".join(summary["senders"]) if summary["senders"] else "-"
        print()
        print(f"[{index}] {summary['title'] or '(no subject)'}")
        print(f"  key: {summary['group_key']}")
        print(
            f"  latest: {summary['latest_received'] or '-'} | count: {summary['count']} | "
            f"priority: {priority_text} | status: {status_text or '-'}"
        )
        print(f"  senders: {senders_text}")
        print(f"  latest_email_id: {summary['latest_email_id'] or '-'}")
        print(f"  reply_target_email_id: {summary['reply_target_email_id'] or '-'}")


def _resolve_group(identifier: str) -> dict:
    group = mailhandle_db.get_group(identifier)
    if not group:
        raise KeyError(identifier)
    return group


def _build_group_payload(identifier: str, *, include_body: bool) -> dict:
    group = _resolve_group(identifier)
    payload = {
        "group_key": group.get("group_key", ""),
        "title": group.get("title", ""),
        "count": group.get("count", 0),
        "latest_received": group.get("latest_received", ""),
        "oldest_received": group.get("oldest_received", ""),
        "latest_email_id": group.get("latest_email_id", ""),
        "reply_target_email_id": group.get("reply_target_email_id", ""),
        "warnings": [],
        "items": [dict(item) for item in group.get("items", [])],
    }
    if not include_body:
        return payload

    context = mailhandle_db.load_group_context(group["group_key"])
    payload["warnings"] = list(context.get("warnings") or [])
    context_by_id = {
        str(item.get("email_id") or ""): item
        for item in context.get("items", [])
    }
    merged_items: list[dict] = []
    for item in payload["items"]:
        context_item = context_by_id.get(str(item.get("email_id") or ""))
        if context_item:
            item = item.copy()
            item["body"] = context_item.get("body", "")
            if context_item.get("body_warning"):
                item["body_warning"] = context_item["body_warning"]
        merged_items.append(item)
    payload["items"] = merged_items
    return payload


def _print_group_payload(payload: dict) -> None:
    print(payload.get("title") or "(no subject)")
    print(f"Key: {payload.get('group_key') or '-'}")
    print(
        f"Count: {payload.get('count', 0)} | Latest: {payload.get('latest_received') or '-'} | "
        f"Reply target: {payload.get('reply_target_email_id') or '-'}"
    )
    for warning in payload.get("warnings", []):
        print(f"Warning: {warning}")
    items = list(payload.get("items") or [])
    if not items:
        print("No items found.")
        return
    for index, item in enumerate(items, start=1):
        print()
        print(f"[{index}] {item.get('subject') or '(no subject)'}")
        print(f"  email_id: {item.get('email_id') or '-'}")
        print(
            f"  received: {item.get('received') or '-'} | from: {item.get('from_name') or item.get('from') or '-'} | "
            f"folder: {item.get('folder') or '-'}"
        )
        print(
            f"  priority: {item.get('priority') or '-'} | status: {item.get('status') or '-'} | "
            f"responded: {bool(item.get('responded'))}"
        )
        if item.get("notes"):
            print(f"  notes: {item.get('notes')}")
        if item.get("abstract"):
            print(f"  abstract: {item.get('abstract')}")
        if item.get("body"):
            print(f"  body: {item.get('body')}")
        if item.get("body_warning"):
            print(f"  body_warning: {item.get('body_warning')}")


def command_overview(args: argparse.Namespace) -> int:
    mailhandle_db.ensure_database()
    sync_result = None
    sync_error = ""
    if not args.skip_sync:
        try:
            sync_result = mailhandle_runtime.sync_database(args, bootstrap=args.bootstrap)
        except Exception as exc:
            sync_error = mailhandle_runtime.describe_outlook_error(exc)
    filters = _build_view_filters(args)
    items = mailhandle_db.load_items(filters)
    groups = mailhandle_db.group_items(items)
    payload = {
        "sync": sync_result,
        "sync_error": sync_error,
        "used_cached_data": bool(sync_error),
        "filters": filters,
        "last_sync_end": mailhandle_db.get_last_sync_end() or "",
        "stats": mailhandle_runtime.build_stats(items),
        "groups": [_group_summary(group) for group in groups],
    }
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    if sync_result is not None:
        counts = sync_result.get("counts", {})
        print(
            f"{sync_result.get('mode', 'unknown').capitalize()} sync stored {counts.get('stored_count', 0)} new items "
            f"and updated {counts.get('updated_count', 0)} items."
        )
    elif sync_error:
        print(sync_error)
        print("Showing cached local mailbox state from SQLite.")
    _print_group_list(groups, filters=filters, stats=payload["stats"], last_sync_end=payload["last_sync_end"])
    return 0


def command_sync(args: argparse.Namespace) -> int:
    mailhandle_db.ensure_database()
    try:
        result = mailhandle_runtime.sync_database(args, bootstrap=args.bootstrap)
    except Exception as exc:
        sync_error = mailhandle_runtime.describe_outlook_error(exc)
        if args.json:
            print(
                json.dumps(
                    {
                        "ok": False,
                        "sync_error": sync_error,
                        "last_sync_end": mailhandle_db.get_last_sync_end() or "",
                    },
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 0
        print(sync_error, file=sys.stderr)
        return 1
    if args.json:
        payload = dict(result)
        payload["ok"] = True
        payload["sync_error"] = ""
        payload["last_sync_end"] = mailhandle_db.get_last_sync_end() or ""
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    counts = result.get("counts", {})
    print(f"Mode: {result.get('mode', 'unknown')}")
    print(f"Stored new items: {counts.get('stored_count', 0)}")
    print(f"Updated items: {counts.get('updated_count', 0)}")
    print(f"Raw count: {counts.get('raw_count', 0)}")
    return 0


def command_list(args: argparse.Namespace) -> int:
    mailhandle_db.ensure_database()
    filters = _build_view_filters(args)
    items = mailhandle_db.load_items(filters)
    groups = mailhandle_db.group_items(items)
    payload = {
        "filters": filters,
        "last_sync_end": mailhandle_db.get_last_sync_end() or "",
        "stats": mailhandle_runtime.build_stats(items),
        "groups": [_group_summary(group) for group in groups],
    }
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    _print_group_list(groups, filters=filters, stats=payload["stats"], last_sync_end=payload["last_sync_end"])
    return 0


def command_show(args: argparse.Namespace) -> int:
    mailhandle_db.ensure_database()
    payload = _build_group_payload(args.identifier, include_body=args.include_body)
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    _print_group_payload(payload)
    return 0


def command_status(args: argparse.Namespace) -> int:
    mailhandle_db.ensure_database()
    notes = _read_optional_text(inline_value=args.notes, file_path=args.notes_file)
    updated = mailhandle_db.update_item(args.email_id, status=args.status, notes=notes)
    if args.json:
        print(json.dumps(updated, ensure_ascii=False, indent=2))
        return 0
    print(f"Updated {updated.get('email_id')}")
    print(f"Status: {updated.get('status')}")
    print(f"Priority: {updated.get('priority')}")
    return 0


def command_open(args: argparse.Namespace) -> int:
    mailhandle_db.open_mail(args.email_id)
    print(f"Opened Outlook item: {args.email_id}")
    return 0


def command_reply_draft(args: argparse.Namespace) -> int:
    mailhandle_db.ensure_database()
    group = _resolve_group(args.identifier)
    notes = _read_optional_text(inline_value=args.notes, file_path=args.notes_file)
    context = mailhandle_db.load_group_context(group["group_key"], latest_only=True)
    draft_text = mailhandle_runtime.request_llm_group_reply(context, notes)
    if args.json:
        payload = {"group_key": context.get("group_key", ""), "draft": _parse_draft_payload(draft_text)}
        if args.save:
            payload["saved_path"] = str(_write_text(args.save, draft_text))
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    if args.save:
        saved_path = _write_text(args.save, draft_text)
        print(f"Saved draft: {saved_path}")
    print(draft_text)
    return 0


def command_reply_open(args: argparse.Namespace) -> int:
    group = _resolve_group(args.identifier)
    draft_text = _read_optional_text(inline_value=args.draft, file_path=args.draft_file)
    email_id = mailhandle_db.open_group_reply_all(group["group_key"], draft_text)
    print(f"Opened Reply All in Outlook for: {email_id}")
    return 0


def command_new_email_draft(args: argparse.Namespace) -> int:
    notes = _read_optional_text(inline_value=args.notes, file_path=args.notes_file)
    draft_text = mailhandle_runtime.request_llm_new_email(notes)
    if args.json:
        payload = {"draft": _parse_draft_payload(draft_text)}
        if args.save:
            payload["saved_path"] = str(_write_text(args.save, draft_text))
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    if args.save:
        saved_path = _write_text(args.save, draft_text)
        print(f"Saved draft: {saved_path}")
    print(draft_text)
    return 0


def command_new_email_open(args: argparse.Namespace) -> int:
    draft_text = _read_optional_text(inline_value=args.draft, file_path=args.draft_file)
    mailhandle_db.open_new_mail(draft_text)
    print("Opened new Outlook email.")
    return 0


def main() -> int:
    mailhandle_runtime.configure_stdio()
    args = parse_args()
    command = args.command or "overview"

    handlers = {
        "overview": command_overview,
        "sync": command_sync,
        "list": command_list,
        "show": command_show,
        "status": command_status,
        "open": command_open,
        "reply-draft": command_reply_draft,
        "reply-open": command_reply_open,
        "new-email-draft": command_new_email_draft,
        "new-email-open": command_new_email_open,
    }
    return handlers[command](args)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        raise SystemExit(130)
    except subprocess.CalledProcessError as exc:
        print(exc.stderr or str(exc), file=sys.stderr)
        raise SystemExit(exc.returncode or 1)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
