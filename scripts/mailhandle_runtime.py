#!/usr/bin/env python3
import argparse
import json
import os
import subprocess
import tempfile
from datetime import datetime, timedelta
from pathlib import Path

import mailhandle_db
import summarize_mail


DATE_PRESETS = ["today", "last_1day", "last_2days", "last_7_days", "this_month", "last_month"]
PROJECT_ROOT = Path(__file__).resolve().parents[1]


def configure_stdio() -> None:
    import sys

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def now_iso() -> str:
    return datetime.now().astimezone().isoformat()


def build_summary_args(
    args: argparse.Namespace,
    *,
    date_preset: str | None,
    since: str | None,
    until: str | None,
) -> argparse.Namespace:
    return argparse.Namespace(
        limit=args.limit,
        from_contains=args.from_contains,
        subject_contains=args.subject_contains,
        project=args.project,
        date_preset=date_preset,
        since=since,
        until=until,
        unread_only=args.unread_only,
        include_body=args.include_body,
        include_notifications=args.include_notifications,
        no_collapse=args.no_collapse,
        verbose=args.verbose,
        json=True,
    )


def sync_database(args: argparse.Namespace, *, bootstrap: bool = False) -> dict:
    query_until = now_iso()
    last_sync_watermark = mailhandle_db.get_last_sync_watermark()
    if bootstrap or not last_sync_watermark:
        summary_args = build_summary_args(args, date_preset=args.date_preset, since=None, until=query_until)
        mode = "bootstrap"
    else:
        summary_args = build_summary_args(args, date_preset=None, since=last_sync_watermark, until=query_until)
        mode = "incremental"
    result = summarize_mail.build_result(summary_args)
    counts = mailhandle_db.upsert_summary(result)
    return {"mode": mode, "result": result, "counts": counts}


def apply_sync(sync_state: dict, args: argparse.Namespace, *, startup: bool = False) -> None:
    sync_state["running"] = True
    sync_state["error"] = False
    sync_state["message"] = "Starting sync..." if startup else "Refreshing from last sync..."
    try:
        result = sync_database(args)
        sync_state["result"] = result
        sync_state["message"] = f"{result['mode'].capitalize()} sync stored {result['counts']['stored_count']} new items"
    except Exception as exc:
        sync_state["error"] = True
        sync_state["message"] = describe_outlook_error(exc)
    finally:
        sync_state["running"] = False


def build_stats(items: list[dict]) -> dict[str, int]:
    stats = {
        "total": len(items),
        "inbox": 0,
        "sent_items": 0,
        "high": 0,
        "medium": 0,
        "low": 0,
        "todo": 0,
        "doing": 0,
        "done": 0,
    }
    for item in items:
        priority = str(item.get("priority") or "low")
        status = str(item.get("status") or "todo")
        folder = str(item.get("folder") or "").strip().lower()
        if folder == "inbox":
            stats["inbox"] += 1
        elif folder in {"sent", "sent items"}:
            stats["sent_items"] += 1
        if priority in stats:
            stats[priority] += 1
        if status in stats:
            stats[status] += 1
    return stats


def describe_outlook_error(exc: Exception) -> str:
    details = str(exc).strip() or exc.__class__.__name__
    return (
        "Classic Outlook is not ready. Open Outlook, wait for mailbox sync to finish, "
        f"and click Refresh. Details: {details}"
    )


def view_payload(filters: dict[str, str], sync_state: dict | None = None) -> dict:
    items = mailhandle_db.load_items(filters)
    try:
        mailbox_address = mailhandle_db.get_mailbox_address()
    except Exception:
        mailbox_address = ""
    sync_state = sync_state or {}
    return {
        "mailbox_address": mailbox_address,
        "db_path": str(mailhandle_db.get_db_path().resolve()),
        "last_sync_end": mailhandle_db.get_last_sync_end() or "",
        "count": len(items),
        "stats": build_stats(items),
        "groups": mailhandle_db.group_items(items),
        "sync_message": str(sync_state.get("message") or ""),
        "sync_error": bool(sync_state.get("error", False)),
        "sync_running": bool(sync_state.get("running", False)),
    }


def resolve_db_time_window(
    *,
    date_preset: str | None = None,
    since: str | None = None,
    until: str | None = None,
) -> tuple[str, str]:
    explicit_since = str(since or "").strip()
    explicit_until = str(until or "").strip()
    if explicit_since or explicit_until:
        return explicit_since, explicit_until

    preset = str(date_preset or "").strip().lower()
    if not preset or preset == "all":
        return "", ""

    now = datetime.now().astimezone()
    start: datetime | None = None
    end: datetime | None = None
    if preset == "today":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=1)
    elif preset == "last_1day":
        start = now - timedelta(days=1)
    elif preset == "last_2days":
        start = now - timedelta(days=2)
    elif preset == "last_7_days":
        start = now - timedelta(days=7)
    elif preset == "this_month":
        start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    elif preset == "last_month":
        first_of_this_month = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        end = first_of_this_month
        previous_month_anchor = first_of_this_month - timedelta(days=1)
        start = previous_month_anchor.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    else:
        return "", ""

    since_value = start.strftime("%Y-%m-%d %H:%M:%S") if start else ""
    until_value = end.strftime("%Y-%m-%d %H:%M:%S") if end else ""
    return since_value, until_value


def _mail_owner_name() -> str:
    owner = summarize_mail.get_mail_owner()
    return str(owner.get("name") or "").strip()


def _response_model_name() -> str:
    return os.getenv("MAILHANDLE_RESPONSE_MODEL", "").strip() or os.getenv("MAILHANDLE_ABSTRACT_MODEL", "").strip()


def _run_structured_codex(prompt: str, *, schema: dict, temp_prefix: str) -> dict:
    codex_command = summarize_mail.get_codex_command()
    if not codex_command:
        raise RuntimeError("Codex CLI is not available for draft generation.")

    with tempfile.TemporaryDirectory(prefix=temp_prefix) as temp_dir:
        temp_path = Path(temp_dir)
        schema_path = temp_path / "schema.json"
        output_path = temp_path / "result.json"
        schema_path.write_text(json.dumps(schema, ensure_ascii=False), encoding="ascii")
        command = [
            codex_command,
            "exec",
            "--skip-git-repo-check",
            "--ephemeral",
            "--sandbox",
            "read-only",
            "--color",
            "never",
            "--output-schema",
            str(schema_path),
            "-o",
            str(output_path),
            "-C",
            str(PROJECT_ROOT),
            "-",
        ]
        model_name = _response_model_name()
        if model_name:
            command[2:2] = ["-m", model_name]
        subprocess.run(command, input=prompt, check=True, capture_output=True, text=True, encoding="utf-8", timeout=180)
        return json.loads(output_path.read_text(encoding="utf-8"))


def _normalize_structured_draft(payload: dict) -> str:
    normalized = {
        "subject": str(payload.get("subject") or "").strip(),
        "greeting": str(payload.get("greeting") or "").strip(),
        "body_en": str(payload.get("body_en") or "").strip(),
        "body_local": str(payload.get("body_local") or "").strip(),
        "local_language": str(payload.get("local_language") or "").strip().lower(),
        "closing": str(payload.get("closing") or "").strip(),
    }
    return json.dumps(normalized, ensure_ascii=False, indent=2)


def _build_group_reply_payload(group_context: dict, notes: str) -> dict:
    items = list(group_context.get("items") or [])
    latest_item = dict(items[-1]) if items else {}
    return {
        "thread": {
            "group_key": group_context.get("group_key", ""),
            "title": group_context.get("title", ""),
            "count": int(group_context.get("count", len(items)) or len(items)),
            "latest_email_id": group_context.get("latest_email_id", ""),
            "reply_target_email_id": group_context.get("reply_target_email_id", ""),
            "warnings": list(group_context.get("warnings") or []),
        },
        "latest_message": latest_item,
        "user_notes": notes,
    }


def request_llm_group_reply(group_context: dict, notes: str) -> str:
    codex_command = summarize_mail.get_codex_command()
    if not codex_command:
        raise RuntimeError("Codex CLI is not available for response generation.")
    owner_name = _mail_owner_name()
    payload = _build_group_reply_payload(group_context, notes)
    prompt = (
        "Draft a professional Outlook reply-all email from the provided thread context.\n"
        "- The payload includes only the latest available message from the thread.\n"
        "- That latest message body may already include quoted history from earlier emails.\n"
        "- Keep facts grounded in the thread.\n"
        "- Do not invent dates or commitments.\n"
        "- Use a structured reply format with these fields only: subject, greeting, body_en, body_local, local_language, closing.\n"
        "- subject is a concise reply subject line.\n"
        "- greeting is the salutation only.\n"
        "- body_en is the English reply body only and must not repeat the greeting or closing.\n"
        "- body_local is an optional second-language reply body only and must not repeat the greeting or closing.\n"
        "- local_language is the language code for body_local, like th, ja, de, zh. If the user asks for a Thailand/Thai version, use th.\n"
        "- closing is a closing-only block and must include the sender name.\n"
        "- If no second language is requested, return body_local as an empty string and local_language as an empty string.\n"
        "- Return JSON only.\n\n"
        f"{json.dumps(payload, ensure_ascii=False)}\n"
    )
    if owner_name:
        prompt += (
            f'- closing must end with the sender name exactly as "{owner_name}".\n'
            f'- a valid closing example is "Best Regards,\\n{owner_name}".\n'
        )
    schema = {
        "type": "object",
        "properties": {
            "subject": {"type": "string"},
            "greeting": {"type": "string"},
            "body_en": {"type": "string"},
            "body_local": {"type": "string"},
            "local_language": {"type": "string"},
            "closing": {"type": "string"},
        },
        "required": ["subject", "greeting", "body_en", "body_local", "local_language", "closing"],
        "additionalProperties": False,
    }
    payload = _run_structured_codex(prompt, schema=schema, temp_prefix="mailhandle-reply-")
    return _normalize_structured_draft(payload)


def request_llm_new_email(notes: str) -> str:
    codex_command = summarize_mail.get_codex_command()
    if not codex_command:
        raise RuntimeError("Codex CLI is not available for new email generation.")
    owner_name = _mail_owner_name()
    payload = {"user_notes": notes}
    prompt = (
        "Draft a professional new Outlook email from the provided user notes.\n"
        "- Keep facts grounded in the notes you are given.\n"
        "- Do not invent dates, names, or commitments.\n"
        "- Use a structured reply format with these fields only: subject, greeting, body_en, body_local, local_language, closing.\n"
        "- subject is a concise email subject line.\n"
        "- greeting is the salutation only.\n"
        "- body_en is the English email body only and must not repeat the greeting or closing.\n"
        "- body_local is an optional second-language email body only and must not repeat the greeting or closing.\n"
        "- local_language is the language code for body_local, like th, ja, de, zh. If the user asks for a Thailand/Thai version, use th.\n"
        "- closing is a closing-only block and must include the sender name.\n"
        "- If no second language is requested, return body_local as an empty string and local_language as an empty string.\n"
        "- Return JSON only.\n\n"
        f"{json.dumps(payload, ensure_ascii=False)}\n"
    )
    if owner_name:
        prompt += (
            f'- closing must end with the sender name exactly as "{owner_name}".\n'
            f'- a valid closing example is "Best Regards,\\n{owner_name}".\n'
        )
    schema = {
        "type": "object",
        "properties": {
            "subject": {"type": "string"},
            "greeting": {"type": "string"},
            "body_en": {"type": "string"},
            "body_local": {"type": "string"},
            "local_language": {"type": "string"},
            "closing": {"type": "string"},
        },
        "required": ["subject", "greeting", "body_en", "body_local", "local_language", "closing"],
        "additionalProperties": False,
    }
    payload = _run_structured_codex(prompt, schema=schema, temp_prefix="mailhandle-new-email-")
    return _normalize_structured_draft(payload)
