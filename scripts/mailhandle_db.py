#!/usr/bin/env python3
import ctypes
import html
import json
import os
import re
import secrets
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
DB_PATH = DATA_DIR / "mailhandle.sqlite"

ALLOWED_STATUSES = {"todo", "doing", "done"}


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", ctypes.c_uint),
        ("pbData", ctypes.POINTER(ctypes.c_byte)),
    ]


_crypt32 = ctypes.windll.crypt32
_kernel32 = ctypes.windll.kernel32


def now_iso() -> str:
    return datetime.now().astimezone().isoformat()


def reinterpret_utc_labeled_local(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return text
    if getattr(parsed, "tzinfo", None) is None:
        return text
    offset = parsed.utcoffset()
    local_offset = datetime.now().astimezone().utcoffset()
    if offset is None or offset.total_seconds() != 0:
        return text
    if local_offset is None or local_offset.total_seconds() == 0:
        return text
    return parsed.replace(tzinfo=None).astimezone().isoformat()


def normalize_text(value: str) -> str:
    return " ".join(str(value or "").split()).strip()


def is_sent_folder(value: str) -> bool:
    text = normalize_text(value).lower()
    return text in {"sent", "sent items"}


def _draft_text_to_html(value: str) -> str:
    paragraphs: list[str] = []
    blocks = str(value or "").replace("\r\n", "\n").replace("\r", "\n").split("\n\n")
    for block in blocks:
        lines = [html.escape(line.strip()) for line in block.split("\n") if line.strip()]
        if not lines:
            continue
        paragraphs.append(f"<p>{'<br>'.join(lines)}</p>")
    return "".join(paragraphs)


def _decode_draft_body(raw: str) -> str:
    return (
        str(raw or "")
        .replace("\\r\\n", "\n")
        .replace("\\n", "\n")
        .replace("\\r", "\n")
        .replace('\\"', '"')
        .replace("\\\\", "\\")
        .strip()
    )


def _parse_structured_draft(value: str) -> dict[str, str] | None:
    text = str(value or "").strip()
    if not (text.startswith("{") and text.endswith("}")):
        return None
    try:
        payload = json.loads(text)
    except Exception:
        return None
    if not isinstance(payload, dict):
        return None
    return {
        "subject": _decode_draft_body(payload.get("subject")),
        "greeting": _decode_draft_body(payload.get("greeting")),
        "body_en": _decode_draft_body(payload.get("body_en")),
        "body_local": _decode_draft_body(payload.get("body_local")),
        "local_language": normalize_text(payload.get("local_language")).lower(),
        "closing": _decode_draft_body(payload.get("closing")),
    }


def _draft_sections(value: str) -> list[tuple[str, str]]:
    text = str(value or "").strip()
    if not text:
        return []
    structured = _parse_structured_draft(text)
    if structured is not None:
        greeting = structured["greeting"]
        body_en = structured["body_en"]
        body_local = structured["body_local"]
        local_language = structured["local_language"]
        if body_en or body_local or greeting:
            sections: list[tuple[str, str]] = []
            en_parts = [part for part in [greeting, body_en] if part]
            if en_parts:
                sections.append(("BODY", f"{os.linesep}{os.linesep}".join(en_parts)))
            if body_local:
                local_note = f"[Language_{(local_language or 'local').upper()}]"
                sections.append(("LOCAL", f"{local_note} {body_local}"))
            if sections:
                return sections
        payload = json.loads(text)
        if isinstance(payload, dict):
            sections: list[tuple[str, str]] = []
            for key, raw_value in payload.items():
                if not str(key).startswith("body_"):
                    continue
                body = _decode_draft_body(raw_value)
                if not body:
                    continue
                label = str(key)[5:].replace("_", " ").strip().upper() or "BODY"
                sections.append((label, body))
            if sections:
                return sections
            for key in ("draft", "body", "reply"):
                body = _decode_draft_body(payload.get(key))
                if body:
                    return [("BODY", body)]
        matches = re.findall(
            r'"(body_[^"]+)"\s*:\s*"(.*?)"(?=\s*,\s*"body_|\s*\}$)',
            text,
            flags=re.DOTALL,
        )
        if matches:
            sections = []
            for key, raw_value in matches:
                body = _decode_draft_body(raw_value)
                if not body:
                    continue
                label = str(key)[5:].replace("_", " ").strip().upper() or "BODY"
                sections.append((label, body))
            if sections:
                return sections
    return [("BODY", text)]


def _draft_sections_to_text(value: str) -> str:
    sections = _draft_sections(value)
    if not sections:
        return ""
    return f"{os.linesep}{os.linesep}".join(body.strip() for _, body in sections if body.strip()).strip()


def _draft_sections_to_html(value: str) -> str:
    sections = _draft_sections(value)
    if not sections:
        return ""
    if len(sections) == 1 and sections[0][0] == "BODY":
        return _draft_text_to_html(sections[0][1])

    blocks: list[str] = []
    for index, (label, body) in enumerate(sections):
        content_html = _draft_text_to_html(body)
        if not content_html:
            continue
        if index == 0 and label == "BODY":
            blocks.append(content_html)
            continue
        blocks.append(
            '<div style="margin:12px 0 0 0; padding:12px 14px; border:1px solid #c9d2dc; '
            'border-radius:12px; color:#1f2937; background:#eef2f6;">'
            f"{content_html}</div>"
        )
    return "".join(blocks)


def _build_new_mail_text(value: str) -> tuple[str, str]:
    structured = _parse_structured_draft(value)
    if structured is None:
        clean_text = _draft_sections_to_text(value)
        body = (
            f"{clean_text}{os.linesep}{os.linesep}{_codex_footer_text()}"
            if clean_text
            else _codex_footer_text()
        )
        return "", body

    sections = _draft_sections(value)
    body_parts = [section_body.strip() for _, section_body in sections if section_body.strip()]
    closing = structured["closing"].strip()
    subject = structured["subject"].strip()
    parts = list(body_parts)
    parts.append(_codex_footer_text())
    if closing:
        parts.append(closing)
    return subject, f"{os.linesep}{os.linesep}".join(parts).strip()


def _build_new_mail_html(value: str) -> tuple[str, str]:
    structured = _parse_structured_draft(value)
    if structured is None:
        body_html = _draft_sections_to_html(value)
        if body_html:
            return "", f"{body_html}{_codex_footer_html()}"
        return "", _codex_footer_html()

    subject = structured["subject"].strip()
    closing_html = _draft_text_to_html(structured["closing"])
    body_html = _draft_sections_to_html(value)
    parts: list[str] = []
    if body_html:
        parts.append(body_html)
    parts.append(_codex_footer_html())
    if closing_html:
        parts.append(closing_html)
    return subject, "".join(parts)


def _codex_footer_html() -> str:
    return (
        '<div style="margin:12px 0 10px 0; padding-top:10px; border-top:1px solid #c7ccd6; '
        "font-family:Consolas, 'Courier New', monospace; font-size:10pt; line-height:1.4; "
        'color:#5f6b7a;">[ Powered by Codex ]</div>'
    )


def _codex_footer_text() -> str:
    return f"------------------------------{os.linesep}[ Powered by Codex ]"


def normalize_subject_key(subject: str) -> str:
    text = normalize_text(subject).lower()
    text = text.replace("[external]:", "").replace("[ext]:", "")
    for prefix in ("re:", "fw:", "fwd:"):
        while text.startswith(prefix):
            text = text[len(prefix) :].strip()
    text = " ".join(text.split())
    return text


def make_thread_key(conversation_topic: str, subject: str) -> str:
    source = normalize_text(conversation_topic) or normalize_text(subject)
    return normalize_subject_key(source)


def _make_data_blob(data: bytes) -> tuple[DATA_BLOB, Any]:
    buffer = ctypes.create_string_buffer(data)
    blob = DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    return blob, buffer


def _blob_to_bytes(blob: DATA_BLOB) -> bytes:
    return ctypes.string_at(blob.pbData, blob.cbData)


def dpapi_protect_bytes(data: bytes) -> bytes:
    if not data:
        return b""
    blob_in, buffer = _make_data_blob(data)
    blob_out = DATA_BLOB()
    if not _crypt32.CryptProtectData(
        ctypes.byref(blob_in),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(blob_out),
    ):
        raise ctypes.WinError()
    try:
        return _blob_to_bytes(blob_out)
    finally:
        if blob_out.pbData:
            _kernel32.LocalFree(blob_out.pbData)
        _ = buffer


def dpapi_unprotect_bytes(data: bytes) -> bytes:
    if not data:
        return b""
    blob_in, buffer = _make_data_blob(data)
    blob_out = DATA_BLOB()
    if not _crypt32.CryptUnprotectData(
        ctypes.byref(blob_in),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(blob_out),
    ):
        raise ctypes.WinError()
    try:
        return _blob_to_bytes(blob_out)
    finally:
        if blob_out.pbData:
            _kernel32.LocalFree(blob_out.pbData)
        _ = buffer


def encrypt_text(text: str) -> bytes:
    return dpapi_protect_bytes(text.encode("utf-8"))


def decrypt_text(blob: bytes | None) -> str:
    if not blob:
        return ""
    return dpapi_unprotect_bytes(blob).decode("utf-8", errors="replace")


def get_db_path() -> Path:
    return DB_PATH


def get_daily_backup_path(reference: datetime | None = None) -> Path:
    _ = reference
    return DB_PATH.with_name(f"{DB_PATH.stem}.backup{DB_PATH.suffix}")


def _is_backup_current(backup_path: Path, reference: datetime | None = None) -> bool:
    current = reference or datetime.now().astimezone()
    try:
        if not backup_path.exists() or backup_path.stat().st_size <= 0:
            return False
        backup_day = datetime.fromtimestamp(backup_path.stat().st_mtime).astimezone().date()
    except OSError:
        return False
    return backup_day == current.date()


def _cleanup_legacy_daily_backups() -> None:
    pattern = f"{DB_PATH.stem}-*.backup{DB_PATH.suffix}"
    for path in DATA_DIR.glob(pattern):
        try:
            path.unlink()
        except OSError:
            continue


def ensure_daily_backup() -> Path | None:
    if not DB_PATH.exists():
        return None
    try:
        if DB_PATH.stat().st_size <= 0:
            return None
    except OSError:
        return None
    backup_path = get_daily_backup_path()
    _cleanup_legacy_daily_backups()
    temp_path = backup_path.with_name(f"{backup_path.name}.tmp")
    try:
        if temp_path.exists():
            temp_path.unlink()
    except OSError:
        pass
    if _is_backup_current(backup_path):
        return backup_path
    try:
        if backup_path.exists():
            backup_path.unlink()
    except OSError:
        pass
    try:
        with sqlite3.connect(DB_PATH) as source_conn:
            source_conn.execute("PRAGMA query_only=ON")
            with sqlite3.connect(backup_path) as target_conn:
                target_conn.execute("PRAGMA journal_mode=DELETE")
                source_conn.backup(target_conn)
                target_conn.commit()
        try:
            if backup_path.exists() and backup_path.stat().st_size > 0:
                return backup_path
        except OSError:
            pass
        return None
    except Exception:
        try:
            if backup_path.exists():
                backup_path.unlink()
        except OSError:
            pass
        raise


def ensure_database() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ensure_daily_backup()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS mail_items (
                email_id TEXT PRIMARY KEY,
                entry_id TEXT NOT NULL,
                store_id TEXT NOT NULL,
                thread_key TEXT NOT NULL,
                subject TEXT NOT NULL,
                subject_key TEXT NOT NULL,
                conversation_topic TEXT NOT NULL,
                from_name TEXT NOT NULL,
                from_email TEXT NOT NULL,
                received_at TEXT NOT NULL,
                folder TEXT NOT NULL,
                priority TEXT NOT NULL,
                computed_priority TEXT NOT NULL DEFAULT 'low',
                next_action TEXT NOT NULL,
                abstract_enc BLOB,
                status TEXT NOT NULL DEFAULT 'todo',
                notes_enc BLOB,
                responded INTEGER NOT NULL DEFAULT 0,
                responded_at TEXT NOT NULL DEFAULT '',
                response_subject TEXT NOT NULL DEFAULT '',
                first_seen_at TEXT NOT NULL,
                last_seen_at TEXT NOT NULL,
                seen_count INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_mail_items_thread_key ON mail_items(thread_key)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_mail_items_subject_key ON mail_items(subject_key)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_mail_items_received_at ON mail_items(received_at)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_mail_items_status ON mail_items(status)"
        )
        columns = {
            str(row[1])
            for row in conn.execute("PRAGMA table_info(mail_items)").fetchall()
        }
        if "computed_priority" not in columns:
            conn.execute(
                "ALTER TABLE mail_items ADD COLUMN computed_priority TEXT NOT NULL DEFAULT 'low'"
            )
            conn.execute("UPDATE mail_items SET computed_priority = priority")
            conn.execute("UPDATE mail_items SET priority = 'low' WHERE status = 'done'")
        conn.execute("UPDATE mail_items SET folder = 'Inbox' WHERE lower(folder) = 'inbox'")
        conn.execute(
            "UPDATE mail_items SET folder = 'Sent Items' WHERE lower(folder) IN ('sent', 'sent items')"
        )
        rows = conn.execute(
            """
            SELECT email_id, received_at, responded_at
            FROM mail_items
            WHERE received_at LIKE '%+00:00'
               OR received_at LIKE '%Z'
               OR responded_at LIKE '%+00:00'
               OR responded_at LIKE '%Z'
            """
        ).fetchall()
        for email_id, received_at, responded_at in rows:
            fixed_received = reinterpret_utc_labeled_local(received_at)
            fixed_responded = reinterpret_utc_labeled_local(responded_at)
            if fixed_received == received_at and fixed_responded == responded_at:
                continue
            conn.execute(
                """
                UPDATE mail_items
                SET received_at = ?, responded_at = ?
                WHERE email_id = ?
                """,
                (fixed_received, fixed_responded, email_id),
            )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS mail_runs (
                run_id TEXT PRIMARY KEY,
                started_at TEXT NOT NULL,
                ended_at TEXT NOT NULL,
                date_preset TEXT NOT NULL,
                since TEXT NOT NULL DEFAULT '',
                until TEXT NOT NULL DEFAULT '',
                folder TEXT NOT NULL,
                limit_value INTEGER NOT NULL,
                raw_count INTEGER NOT NULL,
                stored_count INTEGER NOT NULL,
                updated_count INTEGER NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )


def _row_to_item(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "email_id": row["email_id"],
        "entry_id": row["entry_id"],
        "store_id": row["store_id"],
        "thread_key": row["thread_key"],
        "subject_key": row["subject_key"],
        "conversation_topic": row["conversation_topic"],
        "subject": row["subject"],
        "from": row["from_name"],
        "from_name": row["from_name"],
        "from_email": row["from_email"],
        "received": row["received_at"],
        "folder": row["folder"],
        "priority": row["priority"],
        "next_action": row["next_action"],
        "abstract": decrypt_text(row["abstract_enc"]),
        "status": row["status"],
        "notes": decrypt_text(row["notes_enc"]),
        "responded": bool(row["responded"]),
        "responded_at": row["responded_at"],
        "response_subject": row["response_subject"],
        "first_seen_at": row["first_seen_at"],
        "last_seen_at": row["last_seen_at"],
        "seen_count": row["seen_count"],
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


def get_last_sync_end() -> str | None:
    ensure_database()
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT MAX(ended_at) AS ended_at FROM mail_runs"
        ).fetchone()
    if not row:
        return None
    value = row[0]
    return str(value) if value else None


def get_last_sync_watermark() -> str | None:
    ensure_database()
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT COALESCE(NULLIF(until, ''), ended_at) AS watermark
            FROM mail_runs
            ORDER BY datetime(ended_at) DESC, ended_at DESC
            LIMIT 1
            """
        ).fetchone()
    if not row:
        return None
    value = row[0]
    return str(value) if value else None


def upsert_summary(summary: dict) -> dict[str, int]:
    ensure_database()
    created_at = now_iso()
    started_at = created_at
    items = summary.get("todos", [])
    filters = summary.get("filters", {})
    stored_count = 0
    updated_count = 0

    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        for todo in items:
            email_id = str(todo.get("email_id") or "").strip()
            if not email_id:
                continue
            existing = conn.execute(
                "SELECT * FROM mail_items WHERE email_id = ?",
                (email_id,),
            ).fetchone()
            status = str(existing["status"]) if existing else str(todo.get("status") or "todo")
            notes_enc = existing["notes_enc"] if existing else None
            first_seen_at = existing["first_seen_at"] if existing else created_at
            seen_count = int(existing["seen_count"]) + 1 if existing else 1
            created_row_at = existing["created_at"] if existing else created_at
            computed_priority = str(todo.get("priority") or "low")
            effective_priority = "low" if status == "done" else computed_priority
            folder = str(todo.get("folder") or filters.get("folder") or "inbox")

            conn.execute(
                """
                INSERT INTO mail_items (
                    email_id, entry_id, store_id, thread_key, subject, subject_key,
                    conversation_topic, from_name, from_email, received_at, folder,
                    priority, computed_priority, next_action, abstract_enc, status, notes_enc, responded,
                    responded_at, response_subject, first_seen_at, last_seen_at,
                    seen_count, created_at, updated_at
                ) VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                )
                ON CONFLICT(email_id) DO UPDATE SET
                    entry_id = excluded.entry_id,
                    store_id = excluded.store_id,
                    thread_key = excluded.thread_key,
                    subject = excluded.subject,
                    subject_key = excluded.subject_key,
                    conversation_topic = excluded.conversation_topic,
                    from_name = excluded.from_name,
                    from_email = excluded.from_email,
                    received_at = excluded.received_at,
                    folder = excluded.folder,
                    priority = CASE
                        WHEN mail_items.status = 'done' THEN 'low'
                        ELSE excluded.computed_priority
                    END,
                    computed_priority = excluded.computed_priority,
                    next_action = excluded.next_action,
                    abstract_enc = excluded.abstract_enc,
                    status = mail_items.status,
                    notes_enc = mail_items.notes_enc,
                    responded = excluded.responded,
                    responded_at = excluded.responded_at,
                    response_subject = excluded.response_subject,
                    first_seen_at = mail_items.first_seen_at,
                    last_seen_at = excluded.last_seen_at,
                    seen_count = mail_items.seen_count + 1,
                    created_at = mail_items.created_at,
                    updated_at = excluded.updated_at
                """,
                (
                    email_id,
                    str(todo.get("email_id") or ""),
                    str(todo.get("store_id") or ""),
                    str(todo.get("thread_key") or ""),
                    str(todo.get("title") or ""),
                    str(todo.get("subject_key") or ""),
                    str(todo.get("conversation_topic") or ""),
                    str(todo.get("from") or ""),
                    str(todo.get("message", {}).get("sender", {}).get("email", "") or ""),
                    str(todo.get("received") or ""),
                    folder,
                    effective_priority,
                    computed_priority,
                    str(todo.get("next_action") or ""),
                    sqlite3.Binary(encrypt_text(str(todo.get("abstract") or ""))),
                    status,
                    sqlite3.Binary(notes_enc) if notes_enc else None,
                    1 if todo.get("responded") else 0,
                    str(todo.get("responded_at") or ""),
                    str(todo.get("response_subject") or ""),
                    first_seen_at,
                    created_at,
                    seen_count,
                    created_row_at,
                    created_at,
                ),
            )
            if existing:
                updated_count += 1
            else:
                stored_count += 1

        run_id = str(summary.get("report_meta", {}).get("run_id") or created_at.replace(":", "").replace("-", ""))
        conn.execute(
            """
            INSERT INTO mail_runs (
                run_id, started_at, ended_at, date_preset, since, until, folder,
                limit_value, raw_count, stored_count, updated_count, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(run_id) DO UPDATE SET
                ended_at = excluded.ended_at,
                raw_count = excluded.raw_count,
                stored_count = excluded.stored_count,
                updated_count = excluded.updated_count
            """,
            (
                run_id,
                started_at,
                created_at,
                str(filters.get("date_preset") or ""),
                str(filters.get("since") or ""),
                str(filters.get("until") or ""),
                str(filters.get("folder") or "inbox"),
                int(filters.get("limit") or 0),
                int(summary.get("count") or len(items)),
                stored_count,
                updated_count,
                created_at,
            ),
        )
        conn.commit()

    return {
        "raw_count": int(summary.get("count") or len(items)),
        "stored_count": stored_count,
        "updated_count": updated_count,
    }


def _build_item_query(filters: dict[str, Any] | None = None) -> tuple[str, list[Any]]:
    filters = filters or {}
    clauses = []
    params: list[Any] = []

    since = str(filters.get("since") or "").strip()
    until = str(filters.get("until") or "").strip()
    if since:
        clauses.append("datetime(received_at) >= datetime(?)")
        params.append(since)
    if until:
        clauses.append("datetime(received_at) < datetime(?)")
        params.append(until)

    status = str(filters.get("status") or "").strip()
    if status == "all_open":
        clauses.append("status != ?")
        params.append("done")
    elif status and status != "all":
        clauses.append("status = ?")
        params.append(status)

    priority = str(filters.get("priority") or "").strip()
    if priority and priority != "all":
        clauses.append("priority = ?")
        params.append(priority)

    where_sql = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    query = f"""
        SELECT *
        FROM mail_items
        {where_sql}
        ORDER BY datetime(received_at) DESC, email_id DESC
    """
    return query, params


def load_items(filters: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    ensure_database()
    filters = filters or {}
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        query, params = _build_item_query(filters)
        rows = conn.execute(
            query,
            params,
        ).fetchall()
    items = [_row_to_item(row) for row in rows]

    search = str(filters.get("search") or "").strip().lower()
    if search:
        def matches_search(item: dict[str, Any]) -> bool:
            haystack = " ".join(
                str(value or "")
                for value in (
                    item.get("subject"),
                    item.get("from_name"),
                    item.get("from_email"),
                    item.get("conversation_topic"),
                    item.get("thread_key"),
                    item.get("subject_key"),
                    item.get("priority"),
                    item.get("next_action"),
                    item.get("abstract"),
                    item.get("notes"),
                )
            ).lower()
            return search in haystack

        items = [item for item in items if matches_search(item)]
    return items


def group_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, dict[str, Any]] = {}
    for item in items:
        key = item.get("thread_key") or item.get("subject_key") or item.get("email_id")
        group = grouped.get(key)
        if group is None:
            group = {
                "group_key": key,
                "title": item.get("conversation_topic") or item.get("subject") or "<no subject>",
                "count": 0,
                "latest_received": item.get("received") or "",
                "oldest_received": item.get("received") or "",
                "latest_email_id": item.get("email_id") or "",
                "reply_target_email_id": item.get("email_id") or "",
                "items": [],
            }
            grouped[key] = group
        group["count"] += 1
        group["items"].append(item)
        if (item.get("received") or "") > group["latest_received"]:
            group["latest_received"] = item.get("received") or ""
            group["latest_email_id"] = item.get("email_id") or ""
        if not group["oldest_received"] or (item.get("received") or "") < group["oldest_received"]:
            group["oldest_received"] = item.get("received") or ""

    result = list(grouped.values())
    for group in result:
        group["items"].sort(key=lambda item: item.get("received") or "")
        reply_target_email_id = ""
        for item in reversed(group["items"]):
            if not is_sent_folder(str(item.get("folder") or "")):
                reply_target_email_id = str(item.get("email_id") or "")
                break
        group["reply_target_email_id"] = reply_target_email_id or str(group.get("latest_email_id") or "")
    result.sort(key=lambda group: group.get("latest_received") or "", reverse=True)
    return result


def load_group_items(group_key: str) -> list[dict[str, Any]]:
    key = normalize_text(group_key)
    if not key:
        return []
    ensure_database()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT *
            FROM mail_items
            WHERE thread_key = ? OR subject_key = ? OR email_id = ?
            ORDER BY datetime(received_at) ASC, email_id ASC
            """,
            (key, key, key),
        ).fetchall()
    return [_row_to_item(row) for row in rows]


def get_group(group_key: str) -> dict[str, Any]:
    items = load_group_items(group_key)
    grouped = group_items(items)
    if not grouped:
        raise KeyError(group_key)
    return grouped[0]


def _load_item_row(email_id: str) -> sqlite3.Row:
    ensure_database()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT * FROM mail_items WHERE email_id = ?",
            (email_id,),
        ).fetchone()
    if row is None:
        raise KeyError(email_id)
    return row


def _get_outlook_item(entry_id: str, store_id: str):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    if store_id:
        return namespace.GetItemFromID(entry_id, store_id)
    return namespace.GetItemFromID(entry_id)


def get_mailbox_address() -> str:
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        current_user = getattr(namespace, "CurrentUser", None)
        if current_user is not None:
            address = str(getattr(current_user, "Address", "") or "").strip()
            exchange_user = None
            try:
                address_entry = getattr(current_user, "AddressEntry", None)
                if address_entry is not None:
                    exchange_user = address_entry.GetExchangeUser()
            except Exception:
                exchange_user = None
            primary = str(getattr(exchange_user, "PrimarySmtpAddress", "") or "").strip()
            if primary:
                return primary
            if "@" in address:
                return address
        accounts = getattr(namespace.Session, "Accounts", None)
        if accounts:
            for index in range(1, int(accounts.Count) + 1):
                smtp = str(getattr(accounts.Item(index), "SmtpAddress", "") or "").strip()
                if smtp:
                    return smtp
        return ""
    finally:
        pythoncom.CoUninitialize()


def _clean_mail_body(text: str, limit: int = 1600) -> str:
    lines: list[str] = []
    for raw_line in str(text or "").splitlines():
        line = " ".join(raw_line.split()).strip()
        if not line:
            continue
        lower = line.lower()
        if lower.startswith(("from:", "sent:", "to:", "cc:", "subject:")):
            continue
        if line.startswith((">", "-", "*")) and len(line) < 4:
            continue
        lines.append(line)
        if len(" ".join(lines)) >= limit:
            break
    return " ".join(lines)[:limit].strip()


def _describe_outlook_item_error(exc: Exception) -> str:
    details = normalize_text(str(exc) or exc.__class__.__name__)
    if len(details) > 160:
        return f"{details[:157]}..."
    return details or exc.__class__.__name__


def _build_group_context_item(
    item: dict[str, Any],
    *,
    body: str = "",
    body_warning: str = "",
) -> dict[str, Any]:
    payload = {
        "email_id": item.get("email_id", ""),
        "subject": item.get("subject", ""),
        "from": item.get("from_name", ""),
        "folder": item.get("folder", ""),
        "received": item.get("received", ""),
        "status": item.get("status", ""),
        "abstract": item.get("abstract", ""),
        "body": body,
    }
    if body_warning:
        payload["body_warning"] = body_warning
    return payload


def _get_reply_candidate_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ordered: list[dict[str, Any]] = []
    seen: set[str] = set()
    for sent_only in (False, True):
        for item in reversed(items):
            email_id = str(item.get("email_id") or "")
            if not email_id or email_id in seen:
                continue
            if is_sent_folder(str(item.get("folder") or "")) != sent_only:
                continue
            seen.add(email_id)
            ordered.append(item)
    return ordered


def _apply_draft_to_reply(reply, draft_text: str) -> None:
    draft = normalize_text(draft_text)
    if not draft:
        return
    clean_text = _draft_sections_to_text(draft_text)
    draft_html = _draft_sections_to_html(draft_text)
    existing_html = str(getattr(reply, "HTMLBody", "") or "")
    if draft_html and existing_html:
        reply.HTMLBody = f"{draft_html}{_codex_footer_html()}{existing_html}"
        return
    reply.Body = (
        f"{clean_text}{os.linesep}{os.linesep}"
        f"{_codex_footer_text()}{os.linesep}{os.linesep}{reply.Body}"
    )


def _apply_draft_to_new_mail(mail_item, draft_text: str) -> None:
    draft = normalize_text(draft_text)
    if not draft:
        return
    subject, draft_html = _build_new_mail_html(draft_text)
    if subject:
        mail_item.Subject = subject
    if draft_html:
        mail_item.HTMLBody = draft_html
        return
    subject, clean_text = _build_new_mail_text(draft_text)
    if subject:
        mail_item.Subject = subject
    mail_item.Body = clean_text


def load_group_context(group_key: str) -> dict[str, Any]:
    group = get_group(group_key)
    items = group.get("items", [])
    if not items:
        raise KeyError(group_key)

    context_items: list[dict[str, Any]] = []
    missing_count = 0
    pythoncom.CoInitialize()
    try:
        for item in items:
            try:
                outlook_item = _get_outlook_item(
                    str(item.get("entry_id") or ""),
                    str(item.get("store_id") or ""),
                )
                body = _clean_mail_body(getattr(outlook_item, "Body", "") or "")
                context_items.append(_build_group_context_item(item, body=body))
            except Exception as exc:
                missing_count += 1
                context_items.append(
                    _build_group_context_item(
                        item,
                        body_warning=(
                            "Original Outlook item is unavailable. "
                            f"Using stored abstract only. Details: {_describe_outlook_item_error(exc)}"
                        ),
                    )
                )
    finally:
        pythoncom.CoUninitialize()

    latest_item = items[-1]
    warnings: list[str] = []
    if missing_count:
        noun = "item" if missing_count == 1 else "items"
        warnings.append(
            f"{missing_count} thread {noun} could not be reopened from Outlook. Stored abstracts are shown instead."
        )
    return {
        "group_key": group.get("group_key", ""),
        "title": group.get("title", ""),
        "count": len(context_items),
        "latest_email_id": latest_item.get("email_id", ""),
        "reply_target_email_id": group.get("reply_target_email_id", ""),
        "warnings": warnings,
        "items": context_items,
    }


def update_item(email_id: str, *, status: str, notes: str) -> dict[str, Any]:
    ensure_database()
    if status not in ALLOWED_STATUSES:
        raise ValueError(f"Invalid status: {status}")

    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT * FROM mail_items WHERE email_id = ?",
            (email_id,),
        ).fetchone()
        if row is None:
            raise KeyError(email_id)
        priority = "low" if status == "done" else str(
            row["computed_priority"] or row["priority"] or "low"
        )
        conn.execute(
            """
            UPDATE mail_items
            SET status = ?, priority = ?, notes_enc = ?, updated_at = ?
            WHERE email_id = ?
            """,
            (
                status,
                priority,
                sqlite3.Binary(encrypt_text(notes)),
                now_iso(),
                email_id,
            ),
        )
        conn.commit()
        updated = conn.execute(
            "SELECT * FROM mail_items WHERE email_id = ?",
            (email_id,),
        ).fetchone()
    return _row_to_item(updated)


def open_mail(email_id: str) -> None:
    row = _load_item_row(email_id)

    pythoncom.CoInitialize()
    try:
        entry_id = str(row["entry_id"] or "")
        store_id = str(row["store_id"] or "")
        item = _get_outlook_item(entry_id, store_id)
        item.Display()
    finally:
        pythoncom.CoUninitialize()


def open_reply_all(email_id: str, draft_text: str) -> None:
    row = _load_item_row(email_id)

    pythoncom.CoInitialize()
    try:
        item = _get_outlook_item(str(row["entry_id"] or ""), str(row["store_id"] or ""))
        reply = item.ReplyAll()
        _apply_draft_to_reply(reply, draft_text)
        reply.Display()
    finally:
        pythoncom.CoUninitialize()


def open_group_reply_all(group_key: str, draft_text: str) -> str:
    group = get_group(group_key)
    items = list(group.get("items") or [])
    if not items:
        raise KeyError(group_key)

    failures: list[str] = []
    pythoncom.CoInitialize()
    try:
        for item in _get_reply_candidate_items(items):
            try:
                outlook_item = _get_outlook_item(
                    str(item.get("entry_id") or ""),
                    str(item.get("store_id") or ""),
                )
                reply = outlook_item.ReplyAll()
                _apply_draft_to_reply(reply, draft_text)
                reply.Display()
                return str(item.get("email_id") or "")
            except Exception as exc:
                subject = str(item.get("subject") or "(no subject)")
                failures.append(f"{subject}: {_describe_outlook_item_error(exc)}")
                continue
    finally:
        pythoncom.CoUninitialize()

    details = failures[0] if failures else "No valid reply target was found."
    raise RuntimeError(f"No reply target could be opened from Outlook. Details: {details}")


def open_new_mail(draft_text: str) -> None:
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail_item = outlook.CreateItem(0)
        _apply_draft_to_new_mail(mail_item, draft_text)
        mail_item.Display()
    finally:
        pythoncom.CoUninitialize()


def get_access_token() -> str:
    return secrets.token_urlsafe(24)
