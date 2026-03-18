#!/usr/bin/env python3
import argparse
import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SUMMARY_SCRIPT = PROJECT_ROOT / "scripts" / "summarize_mail.py"
DEFAULT_RECORDS_DIR = PROJECT_ROOT / "records"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-json")
    parser.add_argument("--output")
    parser.add_argument("--output-dir")
    parser.add_argument("--title")
    parser.add_argument("--limit", type=int, default=20)
    parser.add_argument("--from-contains")
    parser.add_argument("--subject-contains")
    parser.add_argument("--project")
    parser.add_argument(
        "--date-preset",
        choices=["today", "yesterday", "last_7_days", "this_month", "last_month"],
    )
    parser.add_argument("--since")
    parser.add_argument("--until")
    parser.add_argument("--unread-only", action="store_true", default=False)
    parser.add_argument("--include-body", action="store_true")
    parser.add_argument("--include-notifications", action="store_true")
    parser.add_argument("--no-collapse", action="store_true")
    return parser.parse_args()


def build_summary_command(args: argparse.Namespace) -> list[str]:
    command = [
        sys.executable,
        str(SUMMARY_SCRIPT),
        "--json",
        "--limit",
        str(args.limit),
    ]
    if args.from_contains:
        command.extend(["--from-contains", args.from_contains])
    if args.subject_contains:
        command.extend(["--subject-contains", args.subject_contains])
    if args.project:
        command.extend(["--project", args.project])
    if args.date_preset:
        command.extend(["--date-preset", args.date_preset])
    if args.since:
        command.extend(["--since", args.since])
    if args.until:
        command.extend(["--until", args.until])
    if args.unread_only:
        command.append("--unread-only")
    if args.include_body:
        command.append("--include-body")
    if args.include_notifications:
        command.append("--include-notifications")
    if args.no_collapse:
        command.append("--no-collapse")
    return command


def load_summary(args: argparse.Namespace) -> dict:
    if args.input_json:
        return json.loads(Path(args.input_json).read_text(encoding="utf-8"))

    completed = subprocess.run(
        build_summary_command(args),
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return json.loads(completed.stdout)


def resolve_output_paths(args: argparse.Namespace, created_at: datetime) -> tuple[Path, Path]:
    stamp = created_at.strftime("%Y%m%d-%H%M%S")
    if args.output:
        html_path = Path(args.output)
        if html_path.suffix.lower() != ".html":
            html_path = html_path.with_suffix(".html")
        summary_path = html_path.with_name(f"{html_path.stem}-summary.json")
        return html_path, summary_path

    if args.output_dir:
        base_dir = Path(args.output_dir)
    elif args.input_json:
        base_dir = Path(args.input_json).resolve().parent
    else:
        base_dir = DEFAULT_RECORDS_DIR / created_at.strftime("%Y-%m-%d")

    html_path = base_dir / f"report-{stamp}.html"
    summary_path = base_dir / f"summary-{stamp}.json"
    return html_path, summary_path


def resolve_title(args: argparse.Namespace, summary: dict, created_at: datetime) -> str:
    if args.title:
        return args.title
    filters = summary.get("filters", {})
    if filters.get("date_preset"):
        return f"Mailhandle review: {filters['date_preset']}"
    if filters.get("since") or filters.get("until"):
        start = filters.get("since") or "*"
        end = filters.get("until") or "*"
        return f"Mailhandle review: {start} to {end}"
    return f"Mailhandle review: {created_at.strftime('%Y-%m-%d %H:%M')}"


def build_payload(
    title: str | None,
    summary: dict,
    created_at: datetime,
    summary_path: Path,
    html_path: Path,
) -> dict:
    run_id = created_at.strftime("%Y%m%d-%H%M%S")
    return {
        "report_meta": {
            "run_id": run_id,
            "created_at": created_at.isoformat(),
            "title": title or f"Mailhandle review: {created_at.strftime('%Y-%m-%d %H:%M')}",
            "storage_key": f"mailhandle-review:{run_id}",
            "summary_path": str(summary_path.resolve()),
            "html_path": str(html_path.resolve()),
        },
        "summary": summary,
    }


def escape_html(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>__REPORT_TITLE__</title>
  <style>
    :root{--bg:#f6f1e8;--panel:#fffdf8;--line:#ddd4c7;--text:#1e293b;--muted:#667085;--accent:#0f766e;--high:#b42318;--medium:#b54708;--low:#175cd3;--done:#475467}
    *{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;color:var(--text);background:linear-gradient(180deg,#faf6ef 0,var(--bg) 100%)}
    .page{width:min(1180px,calc(100% - 32px));margin:24px auto 40px}.panel{background:var(--panel);border:1px solid var(--line);border-radius:16px;box-shadow:0 14px 36px rgba(15,23,42,.08)}
    .hero,.controls,.section,.empty{padding:20px}.eyebrow{display:inline-block;padding:6px 10px;border-radius:999px;background:rgba(15,118,110,.1);color:var(--accent);font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.08em}
    h1{margin:14px 0 10px;font-size:clamp(28px,4vw,42px);line-height:1.05}.subtitle{margin:0;color:var(--muted);line-height:1.5}
    .grid{display:grid;gap:12px}.meta-grid{grid-template-columns:repeat(auto-fit,minmax(220px,1fr));margin-top:18px}.stats-grid{grid-template-columns:repeat(auto-fit,minmax(120px,1fr));margin-top:12px}
    .meta,.stat{border:1px solid var(--line);border-radius:14px;padding:14px;background:#fff}.meta strong,.stat strong{display:block}.meta span,.stat span{display:block;margin-top:6px;color:var(--muted);font-size:13px;line-height:1.4;word-break:break-word}
    .stat strong{font-size:28px;letter-spacing:-.03em}.controls{margin-top:18px}.bar,.filters,.actions,.card-actions,.pills{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
    .bar{justify-content:space-between}.status{font-size:13px;color:var(--muted)}.filters{margin-top:12px}.field{display:flex;flex-direction:column;gap:6px;min-width:180px;flex:1 1 180px}
    .field label,.toggle{font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}.toggle{display:flex;align-items:center;gap:8px;padding-top:20px}
    input[type="search"],select,textarea{width:100%;padding:11px 12px;border:1px solid #d0d5dd;border-radius:12px;font:inherit;background:#fff;color:var(--text)}textarea{min-height:84px;resize:vertical}
    input[type="checkbox"]{accent-color:var(--accent);width:16px;height:16px}button{border:0;border-radius:999px;padding:10px 14px;font:inherit;font-weight:700;cursor:pointer;background:var(--accent);color:#fff}
    button.secondary{background:#fff;color:var(--text);border:1px solid var(--line)}button.ghost{background:transparent;color:var(--muted);border:1px dashed #c5cbd3}
    .sections{display:grid;gap:18px;margin-top:18px}.section-head{display:flex;justify-content:space-between;gap:12px;align-items:baseline;margin-bottom:14px}.section-head h2{margin:0;font-size:20px}.section-head span{color:var(--muted);font-size:13px}
    .cards{display:grid;gap:14px}.card{border:1px solid var(--line);border-radius:16px;padding:16px;background:#fff}.card.done{opacity:.82;background:#fafafa}.card-top{display:grid;grid-template-columns:auto 1fr auto;gap:12px;align-items:start}
    .title{margin:0;font-size:18px;line-height:1.35}.pills{margin-top:10px}.pill{display:inline-flex;align-items:center;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:700;background:#f3f4f6}
    .priority-high{color:var(--high);background:rgba(180,35,24,.12)}.priority-medium{color:var(--medium);background:rgba(181,71,8,.12)}.priority-low{color:var(--low);background:rgba(23,92,211,.12)}.priority-done{color:var(--done);background:rgba(71,84,103,.12)}
    .card-body{display:grid;gap:12px;grid-template-columns:1fr;margin-top:14px}.stack{display:grid;gap:12px}.box{border:1px solid var(--line);border-radius:14px;padding:13px 14px;background:#fffdfb}
    .label{display:block;margin-bottom:8px;font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}.body{margin:0;line-height:1.55;white-space:pre-wrap;word-break:break-word}
    .empty{margin-top:18px;text-align:center;color:var(--muted)}@media (max-width:860px){.card-top,.card-body{grid-template-columns:1fr}.bar{flex-direction:column;align-items:stretch}}
  </style>
</head>
<body>
  <div class="page">
    <section class="panel hero">
      <div class="eyebrow">Mailhandle review report</div>
      <h1 id="title"></h1>
      <p id="subtitle" class="subtitle"></p>
      <div id="metaGrid" class="grid meta-grid"></div>
      <div id="statsGrid" class="grid stats-grid"></div>
    </section>
    <section class="panel controls">
      <div class="bar">
        <div id="statusLine" class="status"></div>
        <div class="actions">
          <button id="exportReviewJson" type="button">Export review JSON</button>
          <button id="exportReviewedHtml" class="secondary" type="button">Export reviewed HTML</button>
          <button id="clearReview" class="ghost" type="button">Reset local review</button>
        </div>
      </div>
      <div class="filters">
        <div class="field"><label for="searchText">Search</label><input id="searchText" type="search" placeholder="Search subject, sender, abstract, next action"></div>
        <div class="field"><label for="statusFilter">Status</label><select id="statusFilter"><option value="all">All items</option><option value="needs_action">Needs action</option><option value="already_replied">Already replied</option><option value="done">Done</option></select></div>
        <div class="field"><label for="priorityFilter">Priority</label><select id="priorityFilter"><option value="all">All priorities</option><option value="high">High only</option><option value="medium">Medium only</option><option value="low">Low only</option></select></div>
        <div class="field"><label class="toggle" for="hideDone"><input id="hideDone" type="checkbox">Hide done</label></div>
      </div>
    </section>
    <div id="sections" class="sections"></div>
    <div id="emptyState" class="panel empty" hidden>No matching items for the current filter.</div>
  </div>
  <script id="mailhandle-payload" type="application/json">__MAILHANDLE_PAYLOAD__</script>
  <script id="mailhandle-review-state" type="application/json">{}</script>
  <script>
    (function () {
      const payload = JSON.parse(document.getElementById("mailhandle-payload").textContent);
      const reviewStateEl = document.getElementById("mailhandle-review-state");
      const initialState = JSON.parse(reviewStateEl.textContent || "{}");
      const summary = payload.summary || {};
      const todos = summary.todos || [];
      const storageKey = payload.report_meta.storage_key;
      const defaultState = { items: {}, last_saved_at: "", last_exported_at: "" };
      const state = mergeState(defaultState, initialState, loadStoredState());
      const el = {
        title: document.getElementById("title"),
        subtitle: document.getElementById("subtitle"),
        metaGrid: document.getElementById("metaGrid"),
        statsGrid: document.getElementById("statsGrid"),
        statusLine: document.getElementById("statusLine"),
        searchText: document.getElementById("searchText"),
        statusFilter: document.getElementById("statusFilter"),
        priorityFilter: document.getElementById("priorityFilter"),
        hideDone: document.getElementById("hideDone"),
        exportReviewJson: document.getElementById("exportReviewJson"),
        exportReviewedHtml: document.getElementById("exportReviewedHtml"),
        clearReview: document.getElementById("clearReview"),
        sections: document.getElementById("sections"),
        emptyState: document.getElementById("emptyState")
      };
      el.searchText.addEventListener("input", render);
      el.statusFilter.addEventListener("change", render);
      el.priorityFilter.addEventListener("change", render);
      el.hideDone.addEventListener("change", render);
      el.exportReviewJson.addEventListener("click", exportReviewJson);
      el.exportReviewedHtml.addEventListener("click", exportReviewedHtml);
      el.clearReview.addEventListener("click", clearReview);
      renderHeader(); persistState(); render();

      function mergeState() {
        const merged = { items: {}, last_saved_at: "", last_exported_at: "" };
        Array.from(arguments).forEach(function (candidate) {
          if (!candidate || typeof candidate !== "object") return;
          merged.items = Object.assign(merged.items, candidate.items || {});
          merged.last_saved_at = candidate.last_saved_at || merged.last_saved_at;
          merged.last_exported_at = candidate.last_exported_at || merged.last_exported_at;
        });
        return merged;
      }
      function loadStoredState() {
        try { const raw = localStorage.getItem(storageKey); return raw ? JSON.parse(raw) : {}; } catch (error) { return {}; }
      }
      function ensureItemState(id) {
        if (!state.items[id]) state.items[id] = { done: false, updated_at: "" };
        return state.items[id];
      }
      function getTodoState(todo) { return ensureItemState(todo.email_id || todo.title); }
      function persistState() {
        state.last_saved_at = new Date().toISOString();
        reviewStateEl.textContent = JSON.stringify(state);
        try { localStorage.setItem(storageKey, JSON.stringify(state)); } catch (error) {}
        updateStatusLine();
      }
      function clearReview() {
        state.items = {}; state.last_saved_at = new Date().toISOString(); state.last_exported_at = "";
        persistState(); render();
      }
      function renderHeader() {
        const meta = payload.report_meta || {};
        const stats = summary.stats || {};
        const doneCount = todos.filter(function (todo) { return getTodoState(todo).done; }).length;
        el.title.textContent = meta.title || "Mailhandle review";
        el.subtitle.textContent = summary.summary || "Mail review report";
        const metaItems = [
          ["Created", formatDateTime(meta.created_at)],
          ["Summary JSON", meta.summary_path || "-"],
          ["HTML file", meta.html_path || "-"],
          ["Filters", formatFilters(summary.filters || {})]
        ];
        const statItems = [
          ["Total", String(summary.count || 0)],
          ["High", String((stats.by_priority || {}).high || 0)],
          ["Medium", String((stats.by_priority || {}).medium || 0)],
          ["Low", String((stats.by_priority || {}).low || 0)],
          ["Replied", String(stats.responded || 0)],
          ["Done", String(doneCount)]
        ];
        replaceChildren(el.metaGrid, metaItems.map(function (item) { return card("meta", item[0], item[1]); }));
        replaceChildren(el.statsGrid, statItems.map(function (item) { return card("stat", item[1], item[0]); }));
      }
      function card(kind, strongText, spanText) {
        const box = document.createElement("div"); box.className = kind;
        const strong = document.createElement("strong"); strong.textContent = strongText;
        const span = document.createElement("span"); span.textContent = spanText || "-";
        box.append(strong, span); return box;
      }
      function updateStatusLine() {
        const saved = state.last_saved_at ? formatDateTime(state.last_saved_at) : "not yet";
        const exported = state.last_exported_at ? formatDateTime(state.last_exported_at) : "not exported";
        el.statusLine.textContent = "Local review state saved " + saved + " | last export " + exported;
      }
      function render() {
        renderHeader();
        const filtered = todos.filter(matchesFilter);
        const groups = [
          ["Needs action", filtered.filter(function (todo) { return !getTodoState(todo).done && !todo.responded; })],
          ["Already replied", filtered.filter(function (todo) { return !getTodoState(todo).done && todo.responded; })],
          ["Done", filtered.filter(function (todo) { return getTodoState(todo).done; })]
        ];
        replaceChildren(el.sections, groups.filter(function (entry) { return entry[1].length; }).map(function (entry) { return renderSection(entry[0], entry[1]); }));
        el.emptyState.hidden = filtered.length !== 0;
      }
      function matchesFilter(todo) {
        const itemState = getTodoState(todo);
        const statusFilter = el.statusFilter.value;
        const priorityFilter = el.priorityFilter.value;
        const searchText = el.searchText.value.trim().toLowerCase();
        if (el.hideDone.checked && itemState.done) return false;
        if (statusFilter === "needs_action" && (itemState.done || todo.responded)) return false;
        if (statusFilter === "already_replied" && (itemState.done || !todo.responded)) return false;
        if (statusFilter === "done" && !itemState.done) return false;
        if (priorityFilter !== "all" && todo.priority !== priorityFilter) return false;
        if (!searchText) return true;
        const haystack = [todo.title, todo.from, todo.abstract, todo.next_action, (todo.projects || []).join(" ")].join(" ").toLowerCase();
        return haystack.includes(searchText);
      }
      function renderSection(title, items) {
        const section = document.createElement("section"); section.className = "panel section";
        const head = document.createElement("div"); head.className = "section-head";
        const h2 = document.createElement("h2"); h2.textContent = title;
        const span = document.createElement("span"); span.textContent = items.length + (items.length === 1 ? " item" : " items");
        head.append(h2, span);
        const cards = document.createElement("div"); cards.className = "cards";
        items.forEach(function (todo) { cards.appendChild(renderCard(todo)); });
        section.append(head, cards); return section;
      }
      function renderCard(todo) {
        const itemState = getTodoState(todo);
        const cardEl = document.createElement("article"); cardEl.className = "card" + (itemState.done ? " done" : "");
        const top = document.createElement("div"); top.className = "card-top";
        const checkbox = document.createElement("input"); checkbox.type = "checkbox"; checkbox.checked = itemState.done;
        checkbox.addEventListener("change", function () { itemState.done = checkbox.checked; itemState.updated_at = new Date().toISOString(); persistState(); render(); });
        const info = document.createElement("div");
        const title = document.createElement("h3"); title.className = "title"; title.textContent = todo.title || "(no subject)";
        const pills = document.createElement("div"); pills.className = "pills";
        pills.appendChild(pill(todo.priority ? todo.priority.toUpperCase() : "UNKNOWN", "priority-" + (itemState.done ? "done" : todo.priority)));
        pills.appendChild(pill(todo.from || "Unknown sender", ""));
        pills.appendChild(pill(formatDateTime(todo.received), ""));
        if (todo.collapsed_count > 1) pills.appendChild(pill("Grouped x" + todo.collapsed_count, ""));
        if (todo.responded) pills.appendChild(pill("Already replied", ""));
        if (todo.owner_attention) pills.appendChild(pill("Direct attention", ""));
        info.append(title, pills);
        const actions = document.createElement("div"); actions.className = "card-actions";
        top.append(checkbox, info, actions);
        const body = document.createElement("div"); body.className = "card-body";
        const left = document.createElement("div"); left.className = "stack";
        left.append(panel("Next action", todo.next_action || "-"));
        left.append(panel("Abstract", todo.abstract || todo.group_summary || "-"));
        if ((todo.projects || []).length) left.append(panel("Projects", todo.projects.join(", ")));
        if (todo.responded) left.append(panel("Reply tracking", formatReply(todo)));
        body.append(left);
        cardEl.append(top, body); return cardEl;
      }
      function panel(labelText, bodyText) {
        const box = document.createElement("div"); box.className = "box";
        const label = document.createElement("span"); label.className = "label"; label.textContent = labelText;
        const body = document.createElement("p"); body.className = "body"; body.textContent = bodyText || "-";
        box.append(label, body); return box;
      }
      function pill(text, extraClass) { const span = document.createElement("span"); span.className = "pill" + (extraClass ? " " + extraClass : ""); span.textContent = text; return span; }
      function formatFilters(filters) {
        const parts = [];
        if (filters.date_preset) parts.push("date=" + filters.date_preset);
        else if (filters.since || filters.until) parts.push("date=" + (filters.since || "*") + ".." + (filters.until || "*"));
        if (filters.project) parts.push("project=" + filters.project);
        if (filters.subject_contains) parts.push("subject~" + filters.subject_contains);
        if (filters.from_contains) parts.push("from~" + filters.from_contains);
        if (filters.unread_only) parts.push("unread_only");
        parts.push("limit=" + (filters.limit || "?"));
        return parts.join(", ");
      }
      function formatReply(todo) {
        const parts = [];
        if (todo.responded_at) parts.push("Replied at " + formatDateTime(todo.responded_at));
        if (todo.response_subject) parts.push("Reply subject: " + todo.response_subject);
        return parts.join(" | ") || "Reply detected";
      }
      function formatDateTime(value) {
        if (!value) return "-";
        const parsed = new Date(value);
        return Number.isNaN(parsed.getTime()) ? value : parsed.toLocaleString();
      }
      function replaceChildren(node, children) { node.replaceChildren(); children.forEach(function (child) { node.appendChild(child); }); }
      function downloadText(filename, text, mimeType) {
        const blob = new Blob([text], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a"); link.href = url; link.download = filename; document.body.appendChild(link); link.click(); link.remove(); URL.revokeObjectURL(url);
      }
      function exportReviewJson() {
        state.last_exported_at = new Date().toISOString(); persistState();
        downloadText("review-" + payload.report_meta.run_id + ".json", JSON.stringify({ report_meta: payload.report_meta, filters: summary.filters || {}, summary: summary.summary || "", todos: todos, review_state: state }, null, 2), "application/json");
      }
      function exportReviewedHtml() {
        state.last_exported_at = new Date().toISOString(); persistState();
        downloadText("report-reviewed-" + payload.report_meta.run_id + ".html", "<!DOCTYPE html>\\n" + document.documentElement.outerHTML, "text/html");
      }
    })();
  </script>
</body>
</html>"""


def render_report_html(payload: dict) -> str:
    payload_json = (
        json.dumps(payload, ensure_ascii=False)
        .replace("&", "\\u0026")
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
    )
    return (
        HTML_TEMPLATE.replace("__MAILHANDLE_PAYLOAD__", payload_json)
        .replace("__REPORT_TITLE__", escape_html(payload["report_meta"]["title"]))
    )


def write_review_report(
    summary: dict,
    *,
    created_at: datetime | None = None,
    title: str | None = None,
    output: str | None = None,
    output_dir: str | None = None,
    input_json: str | None = None,
) -> dict:
    created_at = created_at or datetime.now().astimezone()
    helper_args = argparse.Namespace(
        input_json=input_json,
        output=output,
        output_dir=output_dir,
        title=title,
    )
    html_path, summary_path = resolve_output_paths(helper_args, created_at)
    html_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.parent.mkdir(parents=True, exist_ok=True)

    summary_path.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    resolved_title = title or resolve_title(helper_args, summary, created_at)
    payload = build_payload(resolved_title, summary, created_at, summary_path, html_path)
    html_path.write_text(render_report_html(payload), encoding="utf-8")
    return payload["report_meta"]


def open_review_report(html_path: str | Path) -> bool:
    path = str(Path(html_path).resolve())
    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
            return True
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
            return True
        subprocess.Popen(["xdg-open", path])
        return True
    except Exception:
        return False


def main() -> int:
    args = parse_args()
    created_at = datetime.now().astimezone()
    summary = load_summary(args)
    report_meta = write_review_report(
        summary,
        created_at=created_at,
        title=args.title,
        output=args.output,
        output_dir=args.output_dir,
        input_json=args.input_json,
    )
    opened = open_review_report(report_meta["html_path"])

    print(f"Saved summary JSON: {report_meta['summary_path']}")
    print(f"Saved review HTML: {report_meta['html_path']}")
    print(f"Opened report: {'yes' if opened else 'no'}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except subprocess.CalledProcessError as exc:
        print(exc.stderr or str(exc), file=sys.stderr)
        raise SystemExit(exc.returncode or 1)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
