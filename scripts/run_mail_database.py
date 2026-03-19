#!/usr/bin/env python3
import argparse
import json
import os
import subprocess
import tempfile
import threading
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import edit_priority_rules
import mailhandle_db
import summarize_mail


DATE_PRESETS = ["today", "last_2days", "last_7_days", "this_month", "last_month"]
PROJECT_ROOT = Path(__file__).resolve().parents[1]
HTML_TEMPLATE = """<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Mailhandle Workspace</title><style>
:root{--bg:#efe9dc;--panel:#fffdf7;--line:#d7cdb9;--text:#1f2937;--muted:#667085;--accent:#0f766e;--soft:rgba(15,118,110,.12);--shadow:0 18px 40px rgba(15,23,42,.08)}*{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:linear-gradient(180deg,#f9f6ee 0%,var(--bg) 100%);color:var(--text)}.page{width:min(1480px,calc(100% - 28px));margin:20px auto 36px}.panel{background:var(--panel);border:1px solid var(--line);border-radius:20px;box-shadow:var(--shadow)}.hero,.section{padding:20px}.hero{margin-bottom:16px}.eyebrow{display:inline-flex;padding:6px 10px;border-radius:999px;background:var(--soft);color:var(--accent);font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase}h1{margin:14px 0 8px;font-size:clamp(28px,4vw,42px);line-height:1.05}.subtitle{margin:0;color:var(--muted);line-height:1.55;max-width:90ch}.meta,.toolbar,.filters,.actions,.tabs,.group-meta,.pills{display:flex;flex-wrap:wrap;gap:10px;align-items:center}.meta{margin-top:14px}.chip,.tab-btn,.pill{display:inline-flex;align-items:center;padding:8px 12px;border-radius:999px;border:1px solid var(--line);background:#fff;font-size:12px;font-weight:700}.pill{padding:6px 10px;background:#f3f4f6;color:#344054}.tab-btn{cursor:pointer}.tab-btn.active{background:var(--accent);border-color:var(--accent);color:#fff}.stat-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:12px;margin-top:16px}.stat{border:1px solid var(--line);border-radius:14px;padding:14px;background:#fff}.stat strong{display:block;font-size:26px}.stat span{display:block;margin-top:6px;color:var(--muted);font-size:13px}button{border:0;border-radius:999px;padding:10px 14px;font:inherit;font-weight:700;cursor:pointer;background:var(--accent);color:#fff}button.secondary{background:#fff;color:var(--text);border:1px solid var(--line)}button.item-open{padding:7px 11px;font-size:12px;line-height:1.1}button:disabled{opacity:.6;cursor:default}.toolbar{justify-content:space-between;margin-bottom:14px}.filters{margin-top:12px;align-items:end}.field{display:flex;flex-direction:column;gap:6px;min-width:170px;flex:1 1 170px}.field.wide{flex:2 1 260px}label{font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}input[type="search"],select,textarea{width:100%;padding:11px 12px;border:1px solid #d0d5dd;border-radius:12px;font:inherit;background:#fff;color:var(--text)}textarea{min-height:110px;resize:vertical}.sections{display:grid;gap:18px;margin-top:16px}.group{border:1px solid var(--line);border-radius:18px;padding:16px;background:#faf4e6}.group-head{display:grid;grid-template-columns:1fr auto;gap:12px;align-items:start;margin-bottom:14px}.group-head h2{margin:0;font-size:20px}.timeline{display:grid;gap:14px;position:relative;padding-left:18px}.timeline:before{content:"";position:absolute;left:7px;top:6px;bottom:6px;width:2px;background:linear-gradient(180deg,rgba(15,118,110,.35),rgba(15,118,110,.08))}.item{position:relative;border:1px solid var(--line);border-radius:16px;padding:16px;background:#fff}.item:before{content:"";position:absolute;left:-18px;top:24px;width:10px;height:10px;border-radius:999px;background:var(--accent);box-shadow:0 0 0 4px rgba(15,118,110,.12)}.item-top{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:12px;align-items:start}.item-controls{display:flex;flex-wrap:nowrap;gap:8px;align-items:center;justify-content:flex-end;white-space:nowrap}.item-status{width:92px;min-width:92px;padding:7px 28px 7px 10px;border-radius:10px;font-size:12px;font-weight:700;line-height:1.1}.title{margin:0;font-size:18px;line-height:1.35}.item-body{display:grid;gap:12px;margin-top:14px}.box{border:1px solid var(--line);border-radius:14px;padding:13px 14px;background:#fffdfb}.label{display:block;margin-bottom:8px;font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}.label-row{display:flex;align-items:center;justify-content:flex-start;gap:8px;margin-bottom:8px;flex-wrap:wrap}.label-row .label{margin-bottom:0;margin-right:auto}.language-select-label{font-size:12px;font-weight:600;color:var(--muted);letter-spacing:0;text-transform:none}.reply-language-select{width:180px;min-width:180px;padding:8px 12px;border-radius:10px;font-size:12px;font-weight:700}.text{margin:0;line-height:1.55;white-space:pre-wrap;word-break:break-word}.tab-panel{display:none}.tab-panel.active{display:block}.rules-frame{width:100%;height:min(86vh,1100px);border:0;border-radius:14px;background:#fff}.empty{margin-top:18px;text-align:center;color:var(--muted)}.prio-high{color:#b42318;background:rgba(180,35,24,.12)}.prio-medium{color:#b54708;background:rgba(181,71,8,.12)}.prio-low{color:#175cd3;background:rgba(23,92,211,.12)}.status-pill{background:rgba(15,118,110,.12);color:#134e4a}.modal-shell{position:fixed;inset:0;display:none;align-items:center;justify-content:center;background:rgba(15,23,42,.38);padding:18px;z-index:50}.modal-shell.open{display:flex}.modal{width:min(980px,100%);max-height:min(88vh,980px);overflow:auto;background:#fffdf8;border:1px solid var(--line);border-radius:20px;box-shadow:var(--shadow);padding:18px}.modal-head{display:grid;grid-template-columns:1fr auto;gap:12px;align-items:start;margin-bottom:14px}.modal-head h3{margin:0;font-size:22px}.modal-status{color:var(--muted);font-size:13px}.modal-grid,.context-list{display:grid;gap:14px}.context-item{border:1px solid var(--line);border-radius:14px;padding:12px;background:#fff}.context-item strong{display:block}.context-item span{display:block;margin-top:6px;color:var(--muted);font-size:13px}@media (max-width:960px){.group-head,.item-top,.modal-head{grid-template-columns:1fr}.toolbar{flex-direction:column;align-items:stretch}.item-controls{justify-content:flex-start}.label-row{align-items:flex-start;flex-direction:column}.label-row .label{margin-right:0}.reply-language-select{width:100%;min-width:0}}</style></head><body><div class="page"><section class="panel hero"><div class="eyebrow">Mailhandle workspace</div><h1>Local mail history and priority rules</h1><p class="subtitle">Grouped review view with thread timelines, inline status updates, and assisted reply drafts.</p><div class="meta" hidden><span class="chip" id="dbPath"></span><span class="chip" id="syncState"></span></div><div class="stat-grid" id="stats"></div><div class="tabs" style="margin-top:16px;"><button id="mailTabBtn" class="tab-btn active" type="button">Mailbox</button><button id="rulesTabBtn" class="tab-btn" type="button">Priority rules</button></div></section><section id="mailPanel" class="panel section tab-panel active"><div class="toolbar"><div class="actions"><button id="refreshBtn" type="button">Refresh from last sync</button><button id="reloadBtn" class="secondary" type="button">Reload view</button></div></div><div class="filters"><div class="field wide"><label for="searchText">Search</label><input id="searchText" type="search" placeholder="Search subject, sender, abstract"></div><div class="field"><label for="rangeFilter">Time range</label><select id="rangeFilter"><option value="all">All</option><option value="today">Today</option><option value="last_2days">Last 2 days</option><option value="last_7_days" selected>Last 7 days</option><option value="this_month">This month</option><option value="last_month">Last month</option></select></div><div class="field"><label for="priorityFilter">Priority</label><select id="priorityFilter"><option value="all">All</option><option value="high">High</option><option value="medium">Medium</option><option value="low">Low</option></select></div><div class="field"><label for="statusFilter">Status</label><select id="statusFilter"><option value="all">All</option><option value="todo">Todo</option><option value="doing">Doing</option><option value="done">Done</option></select></div></div><div id="groups" class="sections"></div><div id="emptyState" class="panel empty" hidden>No matching items.</div></section><section id="rulesPanel" class="panel section tab-panel"><div class="toolbar" style="margin-bottom:12px;"><div class="actions"><button id="openRulesBtn" type="button">Open rules editor</button></div><div class="chip">Embedded editor for priority_rules.json</div></div><iframe id="rulesFrame" class="rules-frame" src="/priority-editor?token=__TOKEN__"></iframe></section></div><div id="responseModalShell" class="modal-shell" aria-hidden="true"><div class="modal"><div class="modal-head"><div><h3 id="responseModalTitle">Thread response</h3><div id="responseModalStatus" class="modal-status">Load a thread to prepare a reply.</div></div><button id="closeModalBtn" class="secondary" type="button">Close</button></div><div class="modal-grid"><div class="box"><span class="label">Thread context</span><div id="responseContext" class="context-list"></div></div><div class="box"><div class="label-row"><span class="label">Additional notes</span><label for="replyLanguage" class="language-select-label">second language</label><select id="replyLanguage" class="reply-language-select"><option value="">None</option><option value="th">Thailand</option><option value="zh">Chinese</option></select></div><textarea id="responseNotes" placeholder="Add guidance for the reply."></textarea></div><div class="actions"><button id="generateReplyBtn" type="button">Generate</button></div><div class="box"><span class="label">Generated reply</span><textarea id="generatedReply" placeholder="The generated draft will appear here."></textarea></div><div class="actions"><button id="replyAllBtn" type="button" disabled>Response</button></div></div></div></div><script>window.MAILHANDLE_TOKEN="__TOKEN__";</script><script>__APP_JS__</script></body></html>"""

HTML_TEMPLATE = HTML_TEMPLATE.replace(
    ".hero{margin-bottom:16px}",
    ".hero{position:relative;margin-bottom:16px;padding-top:72px}",
)
HTML_TEMPLATE = HTML_TEMPLATE.replace(
    ".tab-btn{cursor:pointer}.tab-btn.active{background:var(--accent);border-color:var(--accent);color:#fff}",
    ".tabs{position:absolute;top:18px;right:20px;gap:8px;z-index:2}.tab-btn{cursor:pointer;background:#f5efe2;border-color:#c3b08e;color:#5a4634;box-shadow:0 6px 16px rgba(15,23,42,.08)}.tab-btn:hover{background:#fff7e8}.tab-btn.active{background:var(--accent);border-color:var(--accent);color:#fff;box-shadow:0 10px 22px rgba(15,118,110,.22)}",
)
HTML_TEMPLATE = HTML_TEMPLATE.replace(
    '@media (max-width:960px){.group-head,.item-top,.modal-head{grid-template-columns:1fr}.toolbar{flex-direction:column;align-items:stretch}.item-controls{justify-content:flex-start}}',
    '@media (max-width:960px){.group-head,.item-top,.modal-head{grid-template-columns:1fr}.toolbar{flex-direction:column;align-items:stretch}.tabs{position:static;justify-content:flex-start;margin-top:14px}.item-controls{justify-content:flex-start}}',
)
HTML_TEMPLATE = HTML_TEMPLATE.replace(
    '<div class="tabs" style="margin-top:16px;">',
    '<div class="tabs">',
)
HTML_TEMPLATE = HTML_TEMPLATE.replace(
    '<div class="eyebrow">Mailhandle workspace</div><h1>Local mail history and priority rules</h1><p class="subtitle">Grouped review view with thread timelines, inline status updates, and assisted reply drafts.</p>',
    '<div class="hero-head"><div class="eyebrow">Mailhandle workspace</div><div class="hero-identity"><div id="mailboxAddress" class="mailbox-address"></div><div id="lastSync" class="last-sync-line"></div></div></div><div id="workspaceNotice" class="workspace-notice" hidden></div>',
)
HTML_TEMPLATE = HTML_TEMPLATE.replace(
    ".hero{position:relative;margin-bottom:16px;padding-top:72px}",
    ".hero{position:relative;margin-bottom:16px;padding-top:18px}.hero-head{display:flex;align-items:center;gap:14px;flex-wrap:wrap;padding-right:250px;min-height:44px}.hero-head .eyebrow{font-size:clamp(24px,3.4vw,36px);padding:8px 14px;letter-spacing:.04em;text-transform:none}.hero-identity{display:grid;gap:4px}.mailbox-address{font-size:14px;font-weight:700;line-height:1.35;color:var(--muted)}.last-sync-line{font-size:12px;line-height:1.35;color:var(--muted)}.workspace-notice{margin-top:10px;padding:10px 12px;border:1px solid #e3d3b1;border-radius:12px;background:#fff7e8;color:#7a4b12;font-size:13px;line-height:1.45}.workspace-notice.error{border-color:#f3c1bb;background:#fff1ef;color:#8a1c12}",
)
HTML_TEMPLATE = HTML_TEMPLATE.replace(
    '<div class="meta" hidden><span class="chip" id="dbPath"></span><span class="chip" id="syncState"></span></div>',
    '',
)
HTML_TEMPLATE = HTML_TEMPLATE.replace(
    '<select id="statusFilter"><option value="all">All</option><option value="todo">Todo</option><option value="doing">Doing</option><option value="done">Done</option></select>',
    '<select id="statusFilter"><option value="all_open" selected>All_Open</option><option value="all">All</option><option value="todo">Todo</option><option value="doing">Doing</option><option value="done">Done</option></select>',
)
HTML_TEMPLATE = HTML_TEMPLATE.replace(
    '.item{position:relative;border:1px solid var(--line);border-radius:16px;padding:16px;background:#fff}.item:before{content:"";position:absolute;left:-18px;top:24px;width:10px;height:10px;border-radius:999px;background:var(--accent);box-shadow:0 0 0 4px rgba(15,118,110,.12)}',
    '.item{position:relative;border:1px solid var(--line);border-radius:16px;padding:16px;background:#fff}.item-responded{background:#eef9f0;border-color:#b7d8b7}.item:before{content:"";position:absolute;left:-18px;top:24px;width:10px;height:10px;border-radius:999px;background:var(--accent);box-shadow:0 0 0 4px rgba(15,118,110,.12)}.item-responded:before{background:#52a36b;box-shadow:0 0 0 4px rgba(82,163,107,.14)}.reply-pill{background:rgba(82,163,107,.14);color:#1f6b38}.context-item.responded{background:#eef9f0;border-color:#b7d8b7}',
)


def configure_stdio() -> None:
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--date-preset", choices=DATE_PRESETS, default="last_7_days")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--from-contains")
    parser.add_argument("--subject-contains")
    parser.add_argument("--project")
    parser.add_argument("--unread-only", action="store_true", default=False)
    parser.add_argument("--include-body", action="store_true")
    parser.add_argument("--include-notifications", action="store_true")
    parser.add_argument("--no-collapse", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--port", type=int, default=0)
    parser.add_argument("--open-browser", action="store_true")
    parser.add_argument("--no-browser", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args()


def now_iso() -> str:
    return datetime.now().astimezone().isoformat()


def build_summary_args(args: argparse.Namespace, *, date_preset: str | None, since: str | None, until: str | None) -> argparse.Namespace:
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


def open_browser_async(url: str) -> None:
    def worker() -> None:
        try:
            webbrowser.open(url)
        except Exception as exc:
            import sys

            print(f"Browser open failed: {exc}", file=sys.stderr)

    threading.Thread(target=worker, daemon=True).start()


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


def parse_item_filters(query: dict[str, list[str]]) -> dict[str, str]:
    def value(name: str, default: str = "") -> str:
        return str(query.get(name, [default])[0] or "").strip()
    return {
        "since": value("since"),
        "until": value("until"),
        "status": value("status", "all_open"),
        "priority": value("priority", "all"),
        "search": value("search"),
    }


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


def priority_editor_html(token: str) -> str:
    return edit_priority_rules.HTML_PAGE.replace("/api/rules", f"/api/rules?token={token}")


def request_llm_group_reply(group_context: dict, notes: str) -> str:
    codex_command = summarize_mail.get_codex_command()
    if not codex_command:
        raise RuntimeError("Codex CLI is not available for response generation.")
    payload = {"thread": group_context, "user_notes": notes}
    prompt = (
        "Draft a professional Outlook reply-all email from the provided thread context.\n"
        "- Keep facts grounded in the thread.\n"
        "- Do not invent dates or commitments.\n"
        "- Do not include a subject line.\n"
        "- Use a structured reply format with these fields only: greeting, body_en, body_local, local_language, closing.\n"
        "- greeting is the salutation only.\n"
        "- body_en is the English reply body only and must not repeat the greeting or closing.\n"
        "- body_local is an optional second-language reply body only and must not repeat the greeting or closing.\n"
        "- local_language is the language code for body_local, like th, ja, de, zh. If the user asks for a Thailand/Thai version, use th.\n"
        "- closing is optional closing text only.\n"
        "- Do not include a signature, sender name, contact block, or sign-off because Outlook will add the configured signature.\n"
        "- If no second language is requested, return body_local as an empty string and local_language as an empty string.\n"
        "- Return JSON only.\n\n"
        f"{json.dumps(payload, ensure_ascii=False)}\n"
    )
    schema = {
        "type": "object",
        "properties": {
            "greeting": {"type": "string"},
            "body_en": {"type": "string"},
            "body_local": {"type": "string"},
            "local_language": {"type": "string"},
            "closing": {"type": "string"},
        },
        "required": ["greeting", "body_en", "body_local", "local_language", "closing"],
        "additionalProperties": False,
    }
    with tempfile.TemporaryDirectory(prefix="mailhandle-reply-") as temp_dir:
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
        model_name = os.getenv("MAILHANDLE_RESPONSE_MODEL", "").strip() or os.getenv("MAILHANDLE_ABSTRACT_MODEL", "").strip()
        if model_name:
            command[2:2] = ["-m", model_name]
        subprocess.run(command, input=prompt, check=True, capture_output=True, text=True, encoding="utf-8", timeout=180)
        payload = json.loads(output_path.read_text(encoding="utf-8"))
    normalized = {
        "greeting": str(payload.get("greeting") or "").strip(),
        "body_en": str(payload.get("body_en") or "").strip(),
        "body_local": str(payload.get("body_local") or "").strip(),
        "local_language": str(payload.get("local_language") or "").strip().lower(),
        "closing": str(payload.get("closing") or "").strip(),
    }
    return json.dumps(normalized, ensure_ascii=False, indent=2)


APP_JS = """(function(){const token=window.MAILHANDLE_TOKEN;const el={mailboxAddress:document.getElementById(\"mailboxAddress\"),workspaceNotice:document.getElementById(\"workspaceNotice\"),dbPath:document.getElementById(\"dbPath\"),syncState:document.getElementById(\"syncState\"),lastSync:document.getElementById(\"lastSync\"),stats:document.getElementById(\"stats\"),refreshBtn:document.getElementById(\"refreshBtn\"),reloadBtn:document.getElementById(\"reloadBtn\"),searchText:document.getElementById(\"searchText\"),rangeFilter:document.getElementById(\"rangeFilter\"),priorityFilter:document.getElementById(\"priorityFilter\"),statusFilter:document.getElementById(\"statusFilter\"),groups:document.getElementById(\"groups\"),emptyState:document.getElementById(\"emptyState\"),mailTabBtn:document.getElementById(\"mailTabBtn\"),rulesTabBtn:document.getElementById(\"rulesTabBtn\"),mailPanel:document.getElementById(\"mailPanel\"),rulesPanel:document.getElementById(\"rulesPanel\"),openRulesBtn:document.getElementById(\"openRulesBtn\"),rulesFrame:document.getElementById(\"rulesFrame\"),responseModalShell:document.getElementById(\"responseModalShell\"),responseModalTitle:document.getElementById(\"responseModalTitle\"),responseModalStatus:document.getElementById(\"responseModalStatus\"),responseContext:document.getElementById(\"responseContext\"),responseNotes:document.getElementById(\"responseNotes\"),replyLanguageThai:document.getElementById(\"replyLanguageThai\"),replyLanguageChinese:document.getElementById(\"replyLanguageChinese\"),generatedReply:document.getElementById(\"generatedReply\"),generateReplyBtn:document.getElementById(\"generateReplyBtn\"),replyAllBtn:document.getElementById(\"replyAllBtn\"),closeModalBtn:document.getElementById(\"closeModalBtn\")};let data={groups:[],stats:{},db_path:\"\",last_sync_end:\"\",mailbox_address:\"\",sync_message:\"\",sync_error:false,sync_running:false,count:0};let activeGroup=null;async function requestJson(path,options){const requestOptions=Object.assign({method:\"GET\"},options||{});const url=new URL(path,window.location.origin);url.searchParams.set(\"token\",token);const headers=new Headers(requestOptions.headers||{});if(requestOptions.method!=\"GET\"&&!headers.has(\"Content-Type\")){headers.set(\"Content-Type\",\"application/json; charset=utf-8\")}requestOptions.headers=headers;const response=await fetch(url.toString(),requestOptions);const text=await response.text();let payload={};if(text){try{payload=JSON.parse(text)}catch(error){throw new Error(\"Invalid server response\")}}if(!response.ok){throw new Error(payload.error||payload.message||response.statusText||(\"HTTP \"+response.status))}return payload}function setNotice(message,isError){if(!el.workspaceNotice)return;if(message){el.workspaceNotice.hidden=false;el.workspaceNotice.textContent=message;el.workspaceNotice.classList.toggle(\"error\",!!isError)}else{el.workspaceNotice.hidden=true;el.workspaceNotice.textContent=\"\";el.workspaceNotice.classList.remove(\"error\")}}function pad2(v){return String(v).padStart(2,\"0\")}function fmt(date){return date.getFullYear()+\"-\"+pad2(date.getMonth()+1)+\"-\"+pad2(date.getDate())+\" \"+pad2(date.getHours())+\":\"+pad2(date.getMinutes())+\":\"+pad2(date.getSeconds())}function resolveRange(range){const now=new Date();let start=null,end=null;if(range===\"today\"){start=new Date(now.getFullYear(),now.getMonth(),now.getDate());end=new Date(start);end.setDate(end.getDate()+1)}else if(range===\"last_2days\"){start=new Date(now.getTime()-2*24*60*60*1000)}else if(range===\"last_7_days\"){start=new Date(now.getTime()-7*24*60*60*1000)}else if(range===\"this_month\"){start=new Date(now.getFullYear(),now.getMonth(),1)}else if(range===\"last_month\"){start=new Date(now.getFullYear(),now.getMonth()-1,1);end=new Date(now.getFullYear(),now.getMonth(),1)}return{since:start?fmt(start):\"\",until:end?fmt(end):\"\"}}function buildItemQuery(){const params=new URLSearchParams();const range=resolveRange(el.rangeFilter.value);if(range.since)params.set(\"since\",range.since);if(range.until)params.set(\"until\",range.until);if(el.statusFilter.value!==\"all\")params.set(\"status\",el.statusFilter.value);if(el.priorityFilter.value!==\"all\")params.set(\"priority\",el.priorityFilter.value);const search=el.searchText.value.trim();if(search)params.set(\"search\",search);return params.toString()}async function loadData(){if(el.syncState)el.syncState.textContent=\"Loading...\";try{const query=buildItemQuery();data=await requestJson(\"/api/items\"+(query?\"?\"+query:\"\"),{method:\"GET\"});render();el.mailboxAddress.textContent=data.mailbox_address||\"Mailbox\";el.lastSync.textContent=data.last_sync_end?\"Last sync: \"+data.last_sync_end:\"Last sync: none\";setNotice(data.sync_message||\"\",data.sync_error);if(el.dbPath)el.dbPath.textContent=data.db_path||\"\";if(el.syncState)el.syncState.textContent=\"\"}catch(error){setNotice(\"Workspace load failed: \"+error.message,true);if(el.syncState)el.syncState.textContent=\"Load failed: \"+error.message}}async function refreshNow(){setNotice(\"Refreshing from last sync...\",false);if(el.syncState)el.syncState.textContent=\"Refreshing from last sync...\";try{const payload=await requestJson(\"/api/sync\",{method:\"POST\",body:\"{}\"});if(el.syncState)el.syncState.textContent=payload.message||\"Sync complete\";await loadData()}catch(error){setNotice(\"Refresh failed. Open classic Outlook, wait for sync to finish, and try again. Details: \"+error.message,true);if(el.syncState)el.syncState.textContent=\"Refresh failed: \"+error.message}}async function saveStatus(emailId,status,notes){await requestJson(\"/api/item/\"+encodeURIComponent(emailId),{method:\"POST\",body:JSON.stringify({status:status,notes:notes||\"\"})})}async function openMail(emailId){try{await requestJson(\"/api/open/\"+encodeURIComponent(emailId),{method:\"POST\",body:\"{}\"});if(el.syncState)el.syncState.textContent=\"Opened mail \"+emailId}catch(error){if(el.syncState)el.syncState.textContent=\"Open failed: \"+error.message}}function pill(text,extraClass){const span=document.createElement(\"span\");span.className=\"pill\"+(extraClass?\" \"+extraClass:\"\");span.textContent=text;return span}function renderStats(){const stats=data.stats||{};const items=[[\"Total\",stats.total||0],[\"High\",stats.high||0],[\"Medium\",stats.medium||0],[\"Low\",stats.low||0],[\"Todo\",stats.todo||0],[\"Doing\",stats.doing||0],[\"Done\",stats.done||0]];el.stats.replaceChildren();items.forEach(function(entry){const box=document.createElement(\"div\");box.className=\"stat\";const strong=document.createElement(\"strong\");strong.textContent=String(entry[1]);const span=document.createElement(\"span\");span.textContent=entry[0];box.append(strong,span);el.stats.appendChild(box)})}function panel(labelText,bodyNodeOrText){const box=document.createElement(\"div\");box.className=\"box\";const label=document.createElement(\"span\");label.className=\"label\";label.textContent=labelText;box.appendChild(label);if(typeof bodyNodeOrText===\"string\"){const text=document.createElement(\"p\");text.className=\"text\";text.textContent=bodyNodeOrText||\"-\";box.appendChild(text)}else{box.appendChild(bodyNodeOrText)}return box}function renderStatusControl(item){const statusSelect=document.createElement(\"select\");statusSelect.className=\"item-status\";[\"todo\",\"doing\",\"done\"].forEach(function(value){const option=document.createElement(\"option\");option.value=value;option.textContent=value;if((item.status||\"todo\")===value)option.selected=true;statusSelect.appendChild(option)});statusSelect.addEventListener(\"change\",function(){saveStatus(item.email_id,statusSelect.value,item.notes||\"\").catch(function(error){if(el.syncState)el.syncState.textContent=\"Save failed: \"+error.message})});return statusSelect}function renderItem(item){const card=document.createElement(\"article\");card.className=\"item\";card.setAttribute(\"data-email-id\",item.email_id);const top=document.createElement(\"div\");top.className=\"item-top\";const info=document.createElement(\"div\");const title=document.createElement(\"h3\");title.className=\"title\";title.textContent=item.subject||\"(no subject)\";const pills=document.createElement(\"div\");pills.className=\"pills\";pills.appendChild(pill((item.priority||\"unknown\").toUpperCase(),\"prio-\"+(item.priority||\"low\")));pills.appendChild(pill(item.from_name||item.from||\"Unknown sender\",\"\"));pills.appendChild(pill(item.received||\"-\",\"\"));if(item.responded)pills.appendChild(pill(\"replied\",\"\"));info.append(title,pills);const controls=document.createElement(\"div\");controls.className=\"item-controls\";controls.appendChild(renderStatusControl(item));const openBtn=document.createElement(\"button\");openBtn.type=\"button\";openBtn.className=\"secondary item-open\";openBtn.textContent=\"Open\";openBtn.addEventListener(\"click\",function(){openMail(item.email_id)});controls.appendChild(openBtn);top.append(info,controls);const body=document.createElement(\"div\");body.className=\"item-body\";body.appendChild(panel(\"Abstract\",item.abstract||\"-\"));card.append(top,body);return card}function renderGroup(group){const section=document.createElement(\"section\");section.className=\"panel group\";const head=document.createElement(\"div\");head.className=\"group-head\";const info=document.createElement(\"div\");const title=document.createElement(\"h2\");title.textContent=group.title||\"(no subject)\";const meta=document.createElement(\"div\");meta.className=\"group-meta\";meta.appendChild(pill(group.count+\" item\"+(group.count===1?\"\":\"s\"),\"\"));if(group.oldest_received)meta.appendChild(pill(\"start \"+group.oldest_received,\"\"));if(group.latest_received)meta.appendChild(pill(\"latest \"+group.latest_received,\"\"));info.append(title,meta);const actions=document.createElement(\"div\");const responseBtn=document.createElement(\"button\");responseBtn.type=\"button\";responseBtn.textContent=\"Response\";responseBtn.addEventListener(\"click\",function(){openResponseModal(group.group_key,group.title||\"(no subject)\")});actions.appendChild(responseBtn);head.append(info,actions);const timeline=document.createElement(\"div\");timeline.className=\"timeline\";(group.items||[]).forEach(function(item){timeline.appendChild(renderItem(item))});section.append(head,timeline);return section}function render(){renderStats();const groups=(data.groups||[]).filter(function(group){return Array.isArray(group.items)&&group.items.length});el.groups.replaceChildren();el.emptyState.hidden=groups.length!==0;groups.forEach(function(group){el.groups.appendChild(renderGroup(group))})}function setTab(active){const mailActive=active===\"mail\";el.mailTabBtn.classList.toggle(\"active\",mailActive);el.rulesTabBtn.classList.toggle(\"active\",!mailActive);el.mailPanel.classList.toggle(\"active\",mailActive);el.rulesPanel.classList.toggle(\"active\",!mailActive)}function setSelectedReplyLanguage(selected){if(el.replyLanguageThai)el.replyLanguageThai.checked=selected===\"th\";if(el.replyLanguageChinese)el.replyLanguageChinese.checked=selected===\"zh\"}function getReplyNotes(){const notes=String(el.responseNotes.value||\"\").trim();let languageInstruction=\"\";if(el.replyLanguageThai&&el.replyLanguageThai.checked){languageInstruction=\"Also provide a Thai version with local_language set to th.\"}else if(el.replyLanguageChinese&&el.replyLanguageChinese.checked){languageInstruction=\"Also provide a Chinese version with local_language set to zh.\"}return [notes,languageInstruction].filter(Boolean).join(\"\\n\\n\")}function closeResponseModal(){activeGroup=null;el.responseModalShell.classList.remove(\"open\");el.responseModalShell.setAttribute(\"aria-hidden\",\"true\");el.responseContext.replaceChildren();el.responseNotes.value=\"\";setSelectedReplyLanguage(\"\");el.generatedReply.value=\"\";el.replyAllBtn.disabled=true;el.responseModalStatus.textContent=\"Load a thread to prepare a reply.\"}function renderResponseContext(group){el.responseContext.replaceChildren();(group.items||[]).forEach(function(item){const node=document.createElement(\"div\");node.className=\"context-item\";const strong=document.createElement(\"strong\");strong.textContent=item.received+\" | \"+item.from;const subject=document.createElement(\"span\");subject.textContent=item.subject||\"(no subject)\";const abstract=document.createElement(\"span\");abstract.textContent=item.abstract||item.body||\"-\";node.append(strong,subject,abstract);el.responseContext.appendChild(node)})}async function openResponseModal(groupKey,title){el.responseModalTitle.textContent=title;el.responseModalStatus.textContent=\"Loading thread context...\";el.responseContext.replaceChildren();el.responseNotes.value=\"\";setSelectedReplyLanguage(\"\");el.generatedReply.value=\"\";el.replyAllBtn.disabled=true;el.responseModalShell.classList.add(\"open\");el.responseModalShell.setAttribute(\"aria-hidden\",\"false\");try{const payload=await requestJson(\"/api/group/\"+encodeURIComponent(groupKey),{method:\"GET\"});activeGroup=payload.group;el.responseModalStatus.textContent=\"Thread loaded. Add notes and generate a reply draft.\";renderResponseContext(payload.group)}catch(error){el.responseModalStatus.textContent=\"Failed to load thread: \"+error.message}}async function generateReply(){if(!activeGroup)return;el.responseModalStatus.textContent=\"Generating reply draft...\";el.generateReplyBtn.disabled=true;try{const payload=await requestJson(\"/api/group/\"+encodeURIComponent(activeGroup.group_key)+\"/draft\",{method:\"POST\",body:JSON.stringify({notes:getReplyNotes()})});el.generatedReply.value=payload.draft||\"\";el.replyAllBtn.disabled=!el.generatedReply.value.trim();el.responseModalStatus.textContent=\"Draft ready. Review it, then use Response to open Reply All in Outlook.\"}catch(error){el.responseModalStatus.textContent=\"Draft generation failed: \"+error.message}finally{el.generateReplyBtn.disabled=false}}async function composeReplyAll(){if(!activeGroup)return;const draft=el.generatedReply.value.trim();if(!draft)return;el.responseModalStatus.textContent=\"Opening Reply All in Outlook...\";try{await requestJson(\"/api/group/\"+encodeURIComponent(activeGroup.group_key)+\"/reply\",{method:\"POST\",body:JSON.stringify({draft:draft})});el.responseModalStatus.textContent=\"Reply All opened in Outlook with the draft pasted in front of the original thread.\"}catch(error){el.responseModalStatus.textContent=\"Failed to open Reply All: \"+error.message}}if(el.replyLanguageThai){el.replyLanguageThai.addEventListener(\"change\",function(){if(el.replyLanguageThai.checked)setSelectedReplyLanguage(\"th\")})}if(el.replyLanguageChinese){el.replyLanguageChinese.addEventListener(\"change\",function(){if(el.replyLanguageChinese.checked)setSelectedReplyLanguage(\"zh\")})}el.refreshBtn.addEventListener(\"click\",refreshNow);el.reloadBtn.addEventListener(\"click\",loadData);el.searchText.addEventListener(\"input\",loadData);el.rangeFilter.addEventListener(\"change\",loadData);el.priorityFilter.addEventListener(\"change\",loadData);el.statusFilter.addEventListener(\"change\",loadData);el.mailTabBtn.addEventListener(\"click\",function(){setTab(\"mail\")});el.rulesTabBtn.addEventListener(\"click\",function(){setTab(\"rules\")});el.openRulesBtn.addEventListener(\"click\",function(){el.rulesFrame.src=\"/priority-editor?token=\"+encodeURIComponent(token);setTab(\"rules\")});el.closeModalBtn.addEventListener(\"click\",closeResponseModal);el.generateReplyBtn.addEventListener(\"click\",generateReply);el.replyAllBtn.addEventListener(\"click\",composeReplyAll);el.responseModalShell.addEventListener(\"click\",function(event){if(event.target===el.responseModalShell)closeResponseModal()});loadData()})();"""
APP_JS = APP_JS.replace(
    'function fmt(date){return date.getFullYear()+\"-\"+pad2(date.getMonth()+1)+\"-\"+pad2(date.getDate())+\" \"+pad2(date.getHours())+\":\"+pad2(date.getMinutes())+\":\"+pad2(date.getSeconds())}function resolveRange(range){',
    'function fmt(date){return date.getFullYear()+\"-\"+pad2(date.getMonth()+1)+\"-\"+pad2(date.getDate())+\" \"+pad2(date.getHours())+\":\"+pad2(date.getMinutes())+\":\"+pad2(date.getSeconds())}function parseTimestamp(value){const text=String(value||\"\").trim();if(!text)return null;const normalized=text.replace(/\\.(\\d{3})\\d+(?=(Z|[+-]\\d\\d:\\d\\d)$)/,\".$1\");const date=new Date(normalized);return Number.isNaN(date.getTime())?null:date}function formatLocalTimestamp(value){const text=String(value||\"\").trim();if(!text)return \"-\";const date=parseTimestamp(text);if(!date)return text;return date.getFullYear()+\"-\"+pad2(date.getMonth()+1)+\"-\"+pad2(date.getDate())+\" \"+pad2(date.getHours())+\":\"+pad2(date.getMinutes())+\":\"+pad2(date.getSeconds())}function resolveRange(range){',
)
APP_JS = APP_JS.replace(
    'el.lastSync.textContent=data.last_sync_end?\"Last sync: \"+data.last_sync_end:\"Last sync: none\";',
    'el.lastSync.textContent=data.last_sync_end?\"Last sync: \"+formatLocalTimestamp(data.last_sync_end):\"Last sync: none\";',
)
APP_JS = APP_JS.replace(
    'pills.appendChild(pill(item.received||\"-\",\"\"));',
    'pills.appendChild(pill(formatLocalTimestamp(item.received),\"\"));',
)
APP_JS = APP_JS.replace(
    'pills.appendChild(pill(item.from_name||item.from||\"Unknown sender\",\"\"));pills.appendChild(pill(formatLocalTimestamp(item.received),\"\"));',
    'pills.appendChild(pill((item.folder||\"Mailbox\").toUpperCase(),\"\"));pills.appendChild(pill(item.from_name||item.from||\"Unknown sender\",\"\"));pills.appendChild(pill(formatLocalTimestamp(item.received),\"\"));',
)
APP_JS = APP_JS.replace(
    'if(group.oldest_received)meta.appendChild(pill(\"start \"+group.oldest_received,\"\"));if(group.latest_received)meta.appendChild(pill(\"latest \"+group.latest_received,\"\"));',
    'if(group.oldest_received)meta.appendChild(pill(\"start \"+formatLocalTimestamp(group.oldest_received),\"\"));if(group.latest_received)meta.appendChild(pill(\"latest \"+formatLocalTimestamp(group.latest_received),\"\"));',
)
APP_JS = APP_JS.replace(
    'strong.textContent=item.received+\" | \"+item.from;',
    'strong.textContent=formatLocalTimestamp(item.received)+\" | \"+(item.folder||\"Mailbox\")+\" | \"+item.from;',
)
APP_JS = APP_JS.replace(
    'if(el.statusFilter.value!==\"all\")params.set(\"status\",el.statusFilter.value);',
    'if(el.statusFilter.value!==\"all_open\")params.set(\"status\",el.statusFilter.value);',
)
APP_JS = APP_JS.replace(
    'const items=[[\"Total\",stats.total||0],[\"High\",stats.high||0],[\"Medium\",stats.medium||0],[\"Low\",stats.low||0],[\"Todo\",stats.todo||0],[\"Doing\",stats.doing||0],[\"Done\",stats.done||0]];',
    'const items=[[\"Inbox\",stats.inbox||0],[\"Sent Items\",stats.sent_items||0],[\"High\",stats.high||0],[\"Medium\",stats.medium||0],[\"Low\",stats.low||0],[\"Todo\",stats.todo||0],[\"Doing\",stats.doing||0],[\"Done\",stats.done||0]];',
)
APP_JS = APP_JS.replace(
    'responseNotes:document.getElementById(\"responseNotes\"),replyLanguageThai:document.getElementById(\"replyLanguageThai\"),replyLanguageChinese:document.getElementById(\"replyLanguageChinese\"),generatedReply:',
    'responseNotes:document.getElementById(\"responseNotes\"),replyLanguage:document.getElementById(\"replyLanguage\"),generatedReply:',
)
APP_JS = APP_JS.replace(
    'function setTab(active){const mailActive=active===\"mail\";el.mailTabBtn.classList.toggle(\"active\",mailActive);el.rulesTabBtn.classList.toggle(\"active\",!mailActive);el.mailPanel.classList.toggle(\"active\",mailActive);el.rulesPanel.classList.toggle(\"active\",!mailActive)}function setSelectedReplyLanguage(selected){if(el.replyLanguageThai)el.replyLanguageThai.checked=selected===\"th\";if(el.replyLanguageChinese)el.replyLanguageChinese.checked=selected===\"zh\"}function getReplyNotes(){const notes=String(el.responseNotes.value||\"\").trim();let languageInstruction=\"\";if(el.replyLanguageThai&&el.replyLanguageThai.checked){languageInstruction=\"Also provide a Thai version with local_language set to th.\"}else if(el.replyLanguageChinese&&el.replyLanguageChinese.checked){languageInstruction=\"Also provide a Chinese version with local_language set to zh.\"}return [notes,languageInstruction].filter(Boolean).join(\"\\n\\n\")}',
    'function setTab(active){const mailActive=active===\"mail\";el.mailTabBtn.classList.toggle(\"active\",mailActive);el.rulesTabBtn.classList.toggle(\"active\",!mailActive);el.mailPanel.classList.toggle(\"active\",mailActive);el.rulesPanel.classList.toggle(\"active\",!mailActive)}function getReplyNotes(){const notes=String(el.responseNotes.value||\"\").trim();const selected=String(el.replyLanguage&&el.replyLanguage.value||\"\").trim();let languageInstruction=\"\";if(selected===\"th\"){languageInstruction=\"Also provide a Thai version with local_language set to th.\"}else if(selected===\"zh\"){languageInstruction=\"Also provide a Chinese version with local_language set to zh.\"}return [notes,languageInstruction].filter(Boolean).join(\"\\n\\n\")}',
)
APP_JS = APP_JS.replace(
    'el.responseNotes.value=\"\";setSelectedReplyLanguage(\"\");el.generatedReply.value=\"\";',
    'el.responseNotes.value=\"\";if(el.replyLanguage)el.replyLanguage.value=\"\";el.generatedReply.value=\"\";',
)
APP_JS = APP_JS.replace(
    'if(el.replyLanguageThai){el.replyLanguageThai.addEventListener(\"change\",function(){if(el.replyLanguageThai.checked)setSelectedReplyLanguage(\"th\")})}if(el.replyLanguageChinese){el.replyLanguageChinese.addEventListener(\"change\",function(){if(el.replyLanguageChinese.checked)setSelectedReplyLanguage(\"zh\")})}',
    '',
)
APP_JS = APP_JS.replace(
    'if(item.responded)pills.appendChild(pill(\"replied\",\"\"));',
    'if(item.responded)pills.appendChild(pill(\"replied\",\"reply-pill\"));',
)
APP_JS = APP_JS.replace(
    'function renderItem(item){const card=document.createElement(\"article\");card.className=\"item\";card.setAttribute(\"data-email-id\",item.email_id);',
    'function renderItem(item){const card=document.createElement(\"article\");const folderName=String(item.folder||\"\").toLowerCase();const respondedItem=folderName===\"sent items\"||folderName===\"sent\";card.className=\"item\"+(respondedItem?\" item-responded\":\"\");card.setAttribute(\"data-email-id\",item.email_id);',
)
APP_JS = APP_JS.replace(
    'function renderResponseContext(group){el.responseContext.replaceChildren();(group.items||[]).forEach(function(item){const node=document.createElement(\"div\");node.className=\"context-item\";',
    'function renderResponseContext(group){el.responseContext.replaceChildren();(group.items||[]).forEach(function(item){const node=document.createElement(\"div\");const folderName=String(item.folder||\"\").toLowerCase();const respondedItem=folderName===\"sent items\"||folderName===\"sent\";node.className=\"context-item\"+(respondedItem?\" responded\":\"\");',
)
APP_JS = APP_JS.replace(
    'async function saveStatus(emailId,status,notes){await requestJson(\"/api/item/\"+encodeURIComponent(emailId),{method:\"POST\",body:JSON.stringify({status:status,notes:notes||\"\"})})}',
    'async function saveStatus(emailId,status,notes){await requestJson(\"/api/item/\"+encodeURIComponent(emailId),{method:\"POST\",body:JSON.stringify({status:status,notes:notes||\"\"})});await loadData()}',
)
APP_JS = APP_JS.replace(
    'let activeGroup=null;',
    'let activeGroup=null;let syncPollTimer=null;let noticeTimer=null;let dismissedSyncMessage=\"\";',
)
APP_JS = APP_JS.replace(
    'function setNotice(message,isError){if(!el.workspaceNotice)return;if(message){el.workspaceNotice.hidden=false;el.workspaceNotice.textContent=message;el.workspaceNotice.classList.toggle(\"error\",!!isError)}else{el.workspaceNotice.hidden=true;el.workspaceNotice.textContent=\"\";el.workspaceNotice.classList.remove(\"error\")}}',
    'function setNotice(message,isError){if(!el.workspaceNotice)return;if(message){el.workspaceNotice.hidden=false;el.workspaceNotice.textContent=message;el.workspaceNotice.classList.toggle(\"error\",!!isError)}else{el.workspaceNotice.hidden=true;el.workspaceNotice.textContent=\"\";el.workspaceNotice.classList.remove(\"error\")}}function clearSyncPoll(){if(syncPollTimer){window.clearTimeout(syncPollTimer);syncPollTimer=null}}function scheduleSyncPoll(delay){clearSyncPoll();syncPollTimer=window.setTimeout(function(){loadData()},delay||1500)}function clearNoticeTimer(){if(noticeTimer){window.clearTimeout(noticeTimer);noticeTimer=null}}function updateSyncNotice(){const message=String(data.sync_message||\"\");if(data.sync_running){dismissedSyncMessage=\"\";clearNoticeTimer();setNotice(message||\"Starting sync...\",false);scheduleSyncPoll(1500);return}clearSyncPoll();if(data.sync_error){dismissedSyncMessage=\"\";clearNoticeTimer();setNotice(message,true);return}if(!message){dismissedSyncMessage=\"\";clearNoticeTimer();setNotice(\"\",false);return}if(dismissedSyncMessage===message){setNotice(\"\",false);return}setNotice(message,false);clearNoticeTimer();noticeTimer=window.setTimeout(function(){dismissedSyncMessage=message;setNotice(\"\",false)},4000)}',
)
APP_JS = APP_JS.replace(
    'async function loadData(){if(el.syncState)el.syncState.textContent=\"Loading...\";try{const query=buildItemQuery();data=await requestJson(\"/api/items\"+(query?\"?\"+query:\"\"),{method:\"GET\"});render();el.mailboxAddress.textContent=data.mailbox_address||\"Mailbox\";el.lastSync.textContent=data.last_sync_end?\"Last sync: \"+formatLocalTimestamp(data.last_sync_end):\"Last sync: none\";setNotice(data.sync_message||\"\",data.sync_error);if(el.dbPath)el.dbPath.textContent=data.db_path||\"\";if(el.syncState)el.syncState.textContent=\"\"}catch(error){setNotice(\"Workspace load failed: \"+error.message,true);if(el.syncState)el.syncState.textContent=\"Load failed: \"+error.message}}',
    'async function loadData(){if(el.syncState)el.syncState.textContent=\"Loading...\";try{const query=buildItemQuery();data=await requestJson(\"/api/items\"+(query?\"?\"+query:\"\"),{method:\"GET\"});render();el.mailboxAddress.textContent=data.mailbox_address||\"Mailbox\";el.lastSync.textContent=data.last_sync_end?\"Last sync: \"+formatLocalTimestamp(data.last_sync_end):\"Last sync: none\";updateSyncNotice();if(el.dbPath)el.dbPath.textContent=data.db_path||\"\";if(el.syncState)el.syncState.textContent=\"\"}catch(error){clearSyncPoll();clearNoticeTimer();setNotice(\"Workspace load failed: \"+error.message,true);if(el.syncState)el.syncState.textContent=\"Load failed: \"+error.message}}',
)
APP_JS = APP_JS.replace(
    'async function refreshNow(){setNotice(\"Refreshing from last sync...\",false);if(el.syncState)el.syncState.textContent=\"Refreshing from last sync...\";try{const payload=await requestJson(\"/api/sync\",{method:\"POST\",body:\"{}\"});if(el.syncState)el.syncState.textContent=payload.message||\"Sync complete\";await loadData()}catch(error){setNotice(\"Refresh failed. Open classic Outlook, wait for sync to finish, and try again. Details: \"+error.message,true);if(el.syncState)el.syncState.textContent=\"Refresh failed: \"+error.message}}',
    'async function refreshNow(){dismissedSyncMessage=\"\";clearNoticeTimer();setNotice(\"Refreshing from last sync...\",false);if(el.syncState)el.syncState.textContent=\"Refreshing from last sync...\";try{const payload=await requestJson(\"/api/sync\",{method:\"POST\",body:\"{}\"});if(el.syncState)el.syncState.textContent=payload.message||\"Sync complete\";await loadData()}catch(error){setNotice(\"Refresh failed. Open classic Outlook, wait for sync to finish, and try again. Details: \"+error.message,true);if(el.syncState)el.syncState.textContent=\"Refresh failed: \"+error.message}}',
)
APP_JS = APP_JS.replace(
    'const abstract=document.createElement(\"span\");abstract.textContent=item.abstract||item.body||\"-\";node.append(strong,subject,abstract);',
    'const abstract=document.createElement(\"span\");abstract.textContent=item.body||item.abstract||\"-\";if(item.body_warning){const warning=document.createElement(\"span\");warning.textContent=item.body_warning;node.append(strong,subject,abstract,warning)}else{node.append(strong,subject,abstract)}',
)
APP_JS = APP_JS.replace(
    'try{const payload=await requestJson(\"/api/group/\"+encodeURIComponent(groupKey),{method:\"GET\"});activeGroup=payload.group;el.responseModalStatus.textContent=\"Thread loaded. Add notes and generate a reply draft.\";renderResponseContext(payload.group)}catch(error){el.responseModalStatus.textContent=\"Failed to load thread: \"+error.message}}',
    'try{const payload=await requestJson(\"/api/group/\"+encodeURIComponent(groupKey),{method:\"GET\"});activeGroup=payload.group;const defaultStatus=\"Thread loaded. Add notes and generate a reply draft.\";const warnings=Array.isArray(payload.group.warnings)?payload.group.warnings:[];el.responseModalStatus.textContent=warnings.length?defaultStatus+\" \"+warnings.join(\" \"):defaultStatus;renderResponseContext(payload.group)}catch(error){el.responseModalStatus.textContent=\"Failed to load thread: \"+error.message}}',
)



class MailDatabaseServer(ThreadingHTTPServer):
    def __init__(self, server_address, RequestHandlerClass, *, token: str, sync_state: dict):
        super().__init__(server_address, RequestHandlerClass)
        self.token = token
        self.sync_state = sync_state


class RequestHandler(BaseHTTPRequestHandler):
    server: MailDatabaseServer

    def _query(self) -> dict[str, list[str]]:
        return parse_qs(urlparse(self.path).query)

    def _authorized(self) -> bool:
        return self._query().get("token", [""])[0] == self.server.token

    def _send_json(self, payload: dict, status: int = HTTPStatus.OK) -> None:
        encoded = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def _send_html(self, html: str) -> None:
        encoded = html.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0") or "0")
        body = self.rfile.read(length) if length > 0 else b""
        return json.loads(body.decode("utf-8")) if body else {}

    def do_GET(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            html = HTML_TEMPLATE.replace("__TOKEN__", self.server.token).replace("__APP_JS__", APP_JS)
            self._send_html(html)
            return
        if path == "/priority-editor":
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            self._send_html(priority_editor_html(self.server.token))
            return
        if path == "/api/rules":
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            self._send_json({"text": edit_priority_rules.read_rules_text(), "meta": edit_priority_rules.get_rules_meta()})
            return
        if path == "/api/items":
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            self._send_json(view_payload(parse_item_filters(self._query()), self.server.sync_state))
            return
        if path.startswith("/api/group/"):
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            encoded_key = path[len("/api/group/") :]
            if encoded_key.endswith("/draft") or encoded_key.endswith("/reply"):
                self._send_json({"error": "Not found"}, status=HTTPStatus.NOT_FOUND)
                return
            try:
                self._send_json({"group": mailhandle_db.load_group_context(unquote(encoded_key))})
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        self._send_json({"error": "Not found"}, status=HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if path == "/api/rules":
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            try:
                payload = self._read_json()
                normalized = json.dumps(payload["rules"], ensure_ascii=False, indent=2) + "\n"
                meta = edit_priority_rules.write_rules_text(normalized)
                self._send_json({"text": normalized, "meta": meta})
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if path == "/api/sync":
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            if self.server.sync_state.get("running"):
                self._send_json({"error": "A mailbox sync is already running."}, status=HTTPStatus.CONFLICT)
                return
            try:
                self.server.sync_state["running"] = True
                self.server.sync_state["error"] = False
                self.server.sync_state["message"] = "Refreshing from last sync..."
                result = sync_database(self.server.sync_state["args"])
                self.server.sync_state["result"] = result
                self.server.sync_state["message"] = f"{result['mode'].capitalize()} sync stored {result['counts']['stored_count']} new items"
                self._send_json({"mode": result["mode"], "message": self.server.sync_state["message"], "raw_count": result["counts"]["raw_count"], "stored_count": result["counts"]["stored_count"], "updated_count": result["counts"]["updated_count"]})
            except Exception as exc:
                self.server.sync_state["error"] = True
                self.server.sync_state["message"] = describe_outlook_error(exc)
                self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            finally:
                self.server.sync_state["running"] = False
            return
        if path.startswith("/api/item/"):
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            email_id = path.rsplit("/", 1)[-1]
            try:
                payload = self._read_json()
                updated = mailhandle_db.update_item(email_id, status=str(payload.get("status") or "todo"), notes=str(payload.get("notes") or ""))
                self._send_json({"item": updated})
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if path.startswith("/api/open/"):
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            email_id = path.rsplit("/", 1)[-1]
            try:
                mailhandle_db.open_mail(email_id)
                self._send_json({"ok": True})
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if path.startswith("/api/group/") and path.endswith("/draft"):
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            group_key = unquote(path[len("/api/group/") : -len("/draft")])
            try:
                payload = self._read_json()
                group_context = mailhandle_db.load_group_context(group_key)
                draft = request_llm_group_reply(group_context, str(payload.get("notes") or ""))
                self._send_json({"draft": draft, "group": group_context})
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if path.startswith("/api/group/") and path.endswith("/reply"):
            if not self._authorized():
                self._send_json({"error": "Unauthorized"}, status=HTTPStatus.FORBIDDEN)
                return
            group_key = unquote(path[len("/api/group/") : -len("/reply")])
            try:
                payload = self._read_json()
                latest_email_id = mailhandle_db.open_group_reply_all(group_key, str(payload.get("draft") or ""))
                self._send_json({"ok": True, "email_id": latest_email_id})
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        self._send_json({"error": "Not found"}, status=HTTPStatus.NOT_FOUND)

    def log_message(self, format, *args):  # noqa: A003
        return


def main() -> int:
    configure_stdio()
    args = parse_args()
    mailhandle_db.ensure_database()
    sync_state = {"running": True, "message": "Starting sync...", "error": False, "args": args, "result": {}}
    token = mailhandle_db.get_access_token()
    server = MailDatabaseServer(("127.0.0.1", args.port), RequestHandler, token=token, sync_state=sync_state)
    host, port = server.server_address
    url = f"http://{host}:{port}/?token={token}"
    threading.Thread(target=server.serve_forever, daemon=True).start()
    print(f"Mailhandle workspace: {url}")
    print(f"Database: {mailhandle_db.get_db_path().resolve()}")
    if args.open_browser:
        open_browser_async(url)
    threading.Thread(target=apply_sync, args=(sync_state, args), kwargs={"startup": True}, daemon=True).start()
    try:
        while True:
            threading.Event().wait(1.0)
    except KeyboardInterrupt:
        server.shutdown()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
