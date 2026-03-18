#!/usr/bin/env python3
import argparse
import json
import sys
import threading
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RULES_FILE = PROJECT_ROOT / "scripts" / "priority_rules.json"

HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Mailhandle Priority Rules Editor</title>
  <style>
    :root {
      --bg: #f6f0e5;
      --panel: #fffdf8;
      --panel-strong: #fffaf0;
      --line: #ddd2c2;
      --text: #1f2937;
      --muted: #667085;
      --accent: #0f766e;
      --accent-soft: rgba(15, 118, 110, 0.12);
      --warn: #b54708;
      --warn-soft: rgba(181, 71, 8, 0.12);
      --good: #027a48;
      --good-soft: rgba(2, 122, 72, 0.12);
      --bad: #b42318;
      --bad-soft: rgba(180, 35, 24, 0.12);
      --shadow: 0 18px 40px rgba(15, 23, 42, 0.08);
      --radius: 18px;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      background:
        radial-gradient(circle at top left, rgba(15, 118, 110, 0.12), transparent 28%),
        linear-gradient(180deg, #faf6ee 0%, var(--bg) 100%);
      color: var(--text);
    }

    .page {
      width: min(1180px, calc(100% - 32px));
      margin: 24px auto 36px;
    }

    .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
    }

    .hero,
    .editor {
      padding: 22px;
    }

    .hero {
      margin-bottom: 18px;
    }

    .eyebrow {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 6px 10px;
      border-radius: 999px;
      background: var(--accent-soft);
      color: var(--accent);
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }

    h1 {
      margin: 14px 0 10px;
      font-size: clamp(28px, 4vw, 40px);
      line-height: 1.05;
      letter-spacing: -0.03em;
    }

    .subtitle {
      margin: 0;
      color: var(--muted);
      line-height: 1.55;
      max-width: 80ch;
    }

    .meta {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 12px;
      margin-top: 18px;
    }

    .card {
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px;
      background: #fff;
    }

    .card strong {
      display: block;
    }

    .card span {
      display: block;
      margin-top: 6px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
      word-break: break-word;
    }

    .toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 16px;
    }

    .status {
      padding: 10px 12px;
      border-radius: 12px;
      font-size: 13px;
      border: 1px solid var(--line);
      background: rgba(255, 255, 255, 0.8);
      color: var(--muted);
    }

    .status.good {
      color: var(--good);
      background: var(--good-soft);
      border-color: rgba(2, 122, 72, 0.18);
    }

    .status.bad {
      color: var(--bad);
      background: var(--bad-soft);
      border-color: rgba(180, 35, 24, 0.18);
    }

    .status.warn {
      color: var(--warn);
      background: var(--warn-soft);
      border-color: rgba(181, 71, 8, 0.18);
    }

    .actions {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
    }

    button {
      border: 0;
      border-radius: 999px;
      padding: 10px 14px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
      background: var(--accent);
      color: #fff;
    }

    button.secondary {
      background: #fff;
      color: var(--text);
      border: 1px solid var(--line);
    }

    button.warn {
      background: var(--warn);
    }

    button.ghost {
      background: transparent;
      color: var(--muted);
      border: 1px dashed #c5cbd3;
    }

    .section-grid {
      display: grid;
      gap: 16px;
    }

    .group {
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 16px;
      background: var(--panel-strong);
    }

    .group h2 {
      margin: 0 0 12px;
      font-size: 19px;
      letter-spacing: -0.02em;
    }

    .group-note {
      margin: 0 0 14px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
    }

    .fields {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 14px;
    }

    .field {
      display: flex;
      flex-direction: column;
      gap: 7px;
    }

    .field.full {
      grid-column: 1 / -1;
    }

    label {
      font-size: 12px;
      font-weight: 700;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }

    input[type="text"],
    input[type="number"],
    select,
    textarea {
      width: 100%;
      border: 1px solid #d0d5dd;
      border-radius: 12px;
      padding: 11px 12px;
      font: inherit;
      color: var(--text);
      background: #fff;
    }

    textarea {
      min-height: 88px;
      resize: vertical;
      line-height: 1.5;
    }

    .hint {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
    }

    .toggle {
      display: inline-flex;
      align-items: center;
      gap: 10px;
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 11px 12px;
      background: #fff;
      min-height: 46px;
    }

    input[type="checkbox"] {
      accent-color: var(--accent);
      width: 16px;
      height: 16px;
    }

    .rules-head {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 14px;
    }

    .rules-list {
      display: grid;
      gap: 14px;
    }

    .rule-card {
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 16px;
      background: #fff;
    }

    .rule-head {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }

    .rule-title {
      display: flex;
      align-items: center;
      gap: 10px;
      font-weight: 700;
      font-size: 16px;
    }

    .rule-badge {
      display: inline-flex;
      align-items: center;
      padding: 5px 9px;
      border-radius: 999px;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      background: rgba(15, 118, 110, 0.1);
      color: var(--accent);
    }

    .rule-actions {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }

    details.advanced {
      margin-top: 16px;
      border: 1px dashed var(--line);
      border-radius: 16px;
      padding: 14px;
      background: rgba(255, 255, 255, 0.55);
    }

    details.advanced summary {
      cursor: pointer;
      font-weight: 700;
    }

    .preview {
      width: 100%;
      min-height: 220px;
      margin-top: 12px;
      font: 13px/1.55 Consolas, "Courier New", monospace;
    }

    .empty-rules {
      padding: 18px;
      border: 1px dashed var(--line);
      border-radius: 14px;
      color: var(--muted);
      text-align: center;
      background: rgba(255, 255, 255, 0.45);
    }

    @media (max-width: 860px) {
      .toolbar,
      .rules-head,
      .rule-head {
        align-items: stretch;
      }

      .actions,
      .rule-actions {
        width: 100%;
      }

      .actions button,
      .rule-actions button {
        flex: 1 1 160px;
      }
    }
  </style>
</head>
<body>
  <div class="page">
    <section class="panel hero">
      <div class="eyebrow">Mailhandle priority editor</div>
      <h1>Friendly editor for priority rules</h1>
      <p class="subtitle">
        Edit <code>priority_rules.json</code> with labeled fields instead of raw JSON.
        The page validates your changes, keeps a backup, and saves directly back to disk.
      </p>
      <div class="meta" id="meta"></div>
    </section>

    <section class="panel editor">
      <div class="toolbar">
        <div id="status" class="status">Loading rules...</div>
        <div class="actions">
          <button id="reloadBtn" type="button" class="secondary">Reload from file</button>
          <button id="formatBtn" type="button" class="secondary">Refresh JSON preview</button>
          <button id="saveBtn" type="button">Save back</button>
          <button id="saveCloseBtn" type="button" class="warn">Save and close</button>
        </div>
      </div>

      <div class="section-grid">
        <section class="group">
          <h2>General Settings</h2>
          <p class="group-note">These top-level settings control default behavior before individual rules are applied.</p>
          <div class="fields">
            <div class="field">
              <label for="defaultPriority">Default priority</label>
              <select id="defaultPriority">
                <option value="high">high</option>
                <option value="medium">medium</option>
                <option value="low">low</option>
              </select>
            </div>
            <div class="field">
              <label for="ownerAliases">Owner aliases</label>
              <textarea id="ownerAliases" placeholder="One alias per line"></textarea>
              <div class="hint">Use nicknames, short names, or alternate names for greeting and @ mention matching.</div>
            </div>
            <div class="field">
              <label for="managerSenders">Manager senders</label>
              <textarea id="managerSenders" placeholder="One sender fragment per line"></textarea>
              <div class="hint">Match manager display names or email fragments, for example <code>boss@company.com</code>.</div>
            </div>
            <div class="field">
              <label for="greetingTerms">Greeting terms</label>
              <textarea id="greetingTerms" placeholder="One greeting per line"></textarea>
              <div class="hint">These are checked at the start of the email, such as <code>hi</code> or <code>hello</code>.</div>
            </div>
            <div class="field">
              <label>Suppress low-priority notifications</label>
              <div class="toggle">
                <input id="suppressNotifications" type="checkbox">
                <span>Hide low-priority notification mail from todo results</span>
              </div>
            </div>
            <div class="field">
              <label>Collapse similar emails</label>
              <div class="toggle">
                <input id="collapseSimilar" type="checkbox">
                <span>Group similar threads into a single todo item</span>
              </div>
            </div>
            <div class="field">
              <label>Boost owner attention</label>
              <div class="toggle">
                <input id="boostOwnerAttention" type="checkbox">
                <span>Apply the legacy attention boost in addition to explicit rules</span>
              </div>
            </div>
          </div>
        </section>

        <section class="group">
          <div class="rules-head">
            <div>
              <h2>Rules</h2>
              <p class="group-note">Rules are applied in order. A rule can raise priority when one or more conditions match.</p>
            </div>
            <div class="actions">
              <button id="addRuleBtn" type="button" class="secondary">Add rule</button>
            </div>
          </div>
          <div id="rulesList" class="rules-list"></div>
          <div id="emptyRules" class="empty-rules" hidden>No rules yet. Click <strong>Add rule</strong> to create one.</div>
        </section>
      </div>

      <details class="advanced">
        <summary>Advanced JSON preview</summary>
        <div class="hint" style="margin-top: 10px;">This preview is generated from the form and is read-only. Use it to inspect the final JSON before saving.</div>
        <textarea id="jsonPreview" class="preview" readonly spellcheck="false"></textarea>
      </details>
    </section>
  </div>

  <template id="ruleTemplate">
    <article class="rule-card">
      <div class="rule-head">
        <div class="rule-title">
          <span class="rule-badge">Rule</span>
          <span class="rule-name-label">New rule</span>
        </div>
        <div class="rule-actions">
          <button type="button" class="secondary move-up-btn">Move up</button>
          <button type="button" class="secondary move-down-btn">Move down</button>
          <button type="button" class="ghost duplicate-btn">Duplicate</button>
          <button type="button" class="ghost delete-btn">Delete</button>
        </div>
      </div>
      <div class="fields">
        <div class="field">
          <label>Rule name</label>
          <input type="text" class="rule-name" placeholder="Direct Greeting In To">
        </div>
        <div class="field">
          <label>Priority</label>
          <select class="rule-priority">
            <option value="high">high</option>
            <option value="medium">medium</option>
            <option value="low">low</option>
          </select>
        </div>
        <div class="field">
          <label>Unread condition</label>
          <select class="rule-unread">
            <option value="">Any</option>
            <option value="true">Must be unread</option>
            <option value="false">Must be read</option>
          </select>
        </div>
        <div class="field">
          <label>Received within days</label>
          <input type="number" min="0" step="1" class="rule-received-within-days" placeholder="Example: 2">
        </div>
        <div class="field">
          <label>Importance values</label>
          <textarea class="rule-importance" placeholder="One value per line, for example high"></textarea>
        </div>
        <div class="field">
          <label>Attention flags</label>
          <textarea class="rule-attention-flags" placeholder="One flag per line, for example owner_greeted"></textarea>
        </div>
        <div class="field">
          <label>Subject contains</label>
          <textarea class="rule-subject-contains" placeholder="One keyword per line"></textarea>
        </div>
        <div class="field">
          <label>Body contains</label>
          <textarea class="rule-body-contains" placeholder="One keyword per line"></textarea>
        </div>
        <div class="field">
          <label>Sender contains</label>
          <textarea class="rule-sender-contains" placeholder="One sender fragment per line"></textarea>
        </div>
        <div class="field">
          <label>Categories</label>
          <textarea class="rule-categories" placeholder="One category per line"></textarea>
        </div>
        <div class="field full">
          <label>Sender matches configured manager list</label>
          <div class="toggle">
            <input type="checkbox" class="rule-sender-manager">
            <span>Use <code>manager_senders</code> from the top-level settings as a match list for this rule</span>
          </div>
        </div>
      </div>
    </article>
  </template>

  <script>
    const statusEl = document.getElementById("status");
    const metaEl = document.getElementById("meta");
    const rulesListEl = document.getElementById("rulesList");
    const emptyRulesEl = document.getElementById("emptyRules");
    const jsonPreviewEl = document.getElementById("jsonPreview");
    const ruleTemplate = document.getElementById("ruleTemplate");

    const form = {
      defaultPriority: document.getElementById("defaultPriority"),
      ownerAliases: document.getElementById("ownerAliases"),
      managerSenders: document.getElementById("managerSenders"),
      greetingTerms: document.getElementById("greetingTerms"),
      suppressNotifications: document.getElementById("suppressNotifications"),
      collapseSimilar: document.getElementById("collapseSimilar"),
      boostOwnerAttention: document.getElementById("boostOwnerAttention")
    };

    document.getElementById("reloadBtn").addEventListener("click", loadRules);
    document.getElementById("formatBtn").addEventListener("click", () => {
      try {
        jsonPreviewEl.value = JSON.stringify(readForm(), null, 2) + "\\n";
        setStatus("JSON preview refreshed.", "good");
      } catch (error) {
        setStatus("Cannot build JSON preview: " + error.message, "bad");
      }
    });
    document.getElementById("saveBtn").addEventListener("click", () => saveRules(false));
    document.getElementById("saveCloseBtn").addEventListener("click", () => saveRules(true));
    document.getElementById("addRuleBtn").addEventListener("click", () => {
      addRuleCard(defaultRule());
      refreshDerivedViews();
      setStatus("Rule added.", "good");
    });

    Object.values(form).forEach((element) => {
      element.addEventListener("input", onDirty);
      element.addEventListener("change", onDirty);
    });

    function setStatus(text, kind) {
      statusEl.textContent = text;
      statusEl.className = "status" + (kind ? " " + kind : "");
    }

    function onDirty() {
      refreshDerivedViews();
      setStatus("Unsaved changes.", "warn");
    }

    function renderMeta(meta) {
      metaEl.replaceChildren();
      [
        ["Rules file", meta.path || "-"],
        ["Last modified", meta.last_modified || "-"],
        ["Backup file", meta.last_backup || "Created on first save"]
      ].forEach(([label, value]) => {
        const card = document.createElement("div");
        card.className = "card";
        const strong = document.createElement("strong");
        strong.textContent = label;
        const span = document.createElement("span");
        span.textContent = value;
        card.append(strong, span);
        metaEl.appendChild(card);
      });
    }

    function splitLines(text) {
      return text
        .split(/\\r?\\n/)
        .map((line) => line.trim())
        .filter(Boolean);
    }

    function joinLines(values) {
      return (values || []).join("\\n");
    }

    function defaultRule() {
      return {
        name: "",
        priority: "medium"
      };
    }

    function fillForm(rulesConfig) {
      form.defaultPriority.value = rulesConfig.default_priority || "low";
      form.ownerAliases.value = joinLines(rulesConfig.owner_aliases || []);
      form.managerSenders.value = joinLines(rulesConfig.manager_senders || []);
      form.greetingTerms.value = joinLines(rulesConfig.greeting_terms || []);
      form.suppressNotifications.checked = !!rulesConfig.suppress_low_priority_notifications;
      form.collapseSimilar.checked = !!rulesConfig.collapse_similar_emails;
      form.boostOwnerAttention.checked = !!rulesConfig.boost_owner_attention;

      rulesListEl.replaceChildren();
      (rulesConfig.rules || []).forEach((rule) => addRuleCard(rule));
      refreshDerivedViews();
    }

    function readForm() {
      return {
        default_priority: form.defaultPriority.value,
        suppress_low_priority_notifications: form.suppressNotifications.checked,
        collapse_similar_emails: form.collapseSimilar.checked,
        boost_owner_attention: form.boostOwnerAttention.checked,
        owner_aliases: splitLines(form.ownerAliases.value),
        manager_senders: splitLines(form.managerSenders.value),
        greeting_terms: splitLines(form.greetingTerms.value),
        rules: Array.from(rulesListEl.querySelectorAll(".rule-card")).map(readRuleCard)
      };
    }

    function addRuleCard(rule) {
      const fragment = ruleTemplate.content.cloneNode(true);
      const card = fragment.querySelector(".rule-card");
      const nameInput = card.querySelector(".rule-name");
      const prioritySelect = card.querySelector(".rule-priority");
      const unreadSelect = card.querySelector(".rule-unread");
      const receivedInput = card.querySelector(".rule-received-within-days");
      const importanceInput = card.querySelector(".rule-importance");
      const attentionInput = card.querySelector(".rule-attention-flags");
      const subjectInput = card.querySelector(".rule-subject-contains");
      const bodyInput = card.querySelector(".rule-body-contains");
      const senderInput = card.querySelector(".rule-sender-contains");
      const categoriesInput = card.querySelector(".rule-categories");
      const senderManagerInput = card.querySelector(".rule-sender-manager");
      const nameLabel = card.querySelector(".rule-name-label");

      nameInput.value = rule.name || "";
      prioritySelect.value = rule.priority || "medium";
      unreadSelect.value = rule.unread === true ? "true" : rule.unread === false ? "false" : "";
      receivedInput.value = rule.received_within_days ?? "";
      importanceInput.value = joinLines(rule.importance || []);
      attentionInput.value = joinLines(rule.attention_flags_any || []);
      subjectInput.value = joinLines(rule.subject_contains_any || []);
      bodyInput.value = joinLines(rule.body_contains_any || []);
      senderInput.value = joinLines(rule.sender_contains_any || []);
      categoriesInput.value = joinLines(rule.categories_any || []);
      senderManagerInput.checked = !!rule.sender_matches_manager;
      nameLabel.textContent = rule.name || "New rule";

      card.querySelectorAll("input, select, textarea").forEach((element) => {
        element.addEventListener("input", () => {
          nameLabel.textContent = nameInput.value.trim() || "New rule";
          onDirty();
        });
        element.addEventListener("change", () => {
          nameLabel.textContent = nameInput.value.trim() || "New rule";
          onDirty();
        });
      });

      card.querySelector(".delete-btn").addEventListener("click", () => {
        card.remove();
        refreshDerivedViews();
        setStatus("Rule removed.", "warn");
      });
      card.querySelector(".duplicate-btn").addEventListener("click", () => {
        addRuleCard(readRuleCard(card));
        refreshDerivedViews();
        setStatus("Rule duplicated.", "good");
      });
      card.querySelector(".move-up-btn").addEventListener("click", () => {
        const previous = card.previousElementSibling;
        if (previous) {
          rulesListEl.insertBefore(card, previous);
          refreshDerivedViews();
          setStatus("Rule moved up.", "good");
        }
      });
      card.querySelector(".move-down-btn").addEventListener("click", () => {
        const next = card.nextElementSibling;
        if (next) {
          rulesListEl.insertBefore(next, card);
          refreshDerivedViews();
          setStatus("Rule moved down.", "good");
        }
      });

      rulesListEl.appendChild(card);
    }

    function readRuleCard(card) {
      const unreadValue = card.querySelector(".rule-unread").value;
      const receivedValue = card.querySelector(".rule-received-within-days").value.trim();
      const rule = {
        name: card.querySelector(".rule-name").value.trim() || "Unnamed rule",
        priority: card.querySelector(".rule-priority").value
      };

      if (unreadValue) {
        rule.unread = unreadValue === "true";
      }
      if (receivedValue) {
        rule.received_within_days = Number(receivedValue);
      }

      const arrayMappings = [
        ["importance", ".rule-importance"],
        ["attention_flags_any", ".rule-attention-flags"],
        ["subject_contains_any", ".rule-subject-contains"],
        ["body_contains_any", ".rule-body-contains"],
        ["sender_contains_any", ".rule-sender-contains"],
        ["categories_any", ".rule-categories"]
      ];

      arrayMappings.forEach(([key, selector]) => {
        const values = splitLines(card.querySelector(selector).value);
        if (values.length) {
          rule[key] = values;
        }
      });

      if (card.querySelector(".rule-sender-manager").checked) {
        rule.sender_matches_manager = true;
      }

      return rule;
    }

    function refreshDerivedViews() {
      emptyRulesEl.hidden = rulesListEl.children.length !== 0;
      try {
        jsonPreviewEl.value = JSON.stringify(readForm(), null, 2) + "\\n";
      } catch (error) {
        jsonPreviewEl.value = "// Invalid form state: " + error.message;
      }
    }

    async function loadRules() {
      setStatus("Loading rules...", "");
      const response = await fetch("/api/rules");
      const payload = await response.json();
      fillForm(JSON.parse(payload.text));
      renderMeta(payload.meta || {});
      setStatus("Rules loaded.", "good");
    }

    async function saveRules(closeAfterSave) {
      let rules;
      try {
        rules = readForm();
      } catch (error) {
        setStatus("Cannot save: " + error.message, "bad");
        return;
      }

      setStatus("Saving...", "");
      const response = await fetch("/api/rules", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rules, close_after_save: !!closeAfterSave })
      });
      const payload = await response.json();
      if (!response.ok) {
        setStatus(payload.error || "Save failed.", "bad");
        return;
      }

      fillForm(JSON.parse(payload.text));
      renderMeta(payload.meta || {});
      setStatus("Saved back to priority_rules.json.", "good");
      if (closeAfterSave) {
        setTimeout(() => window.close(), 250);
      }
    }

    document.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        saveRules(false);
      }
    });

    loadRules().catch((error) => setStatus("Failed to load rules: " + error.message, "bad"));
  </script>
</body>
</html>
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=0)
    parser.add_argument("--no-open", action="store_true")
    return parser.parse_args()


def configure_stdio() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def get_rules_meta() -> dict:
    stat = RULES_FILE.stat()
    backup_path = RULES_FILE.with_suffix(".json.bak")
    return {
        "path": str(RULES_FILE.resolve()),
        "last_modified": datetime.fromtimestamp(stat.st_mtime).astimezone().isoformat(timespec="seconds"),
        "last_backup": str(backup_path.resolve()) if backup_path.exists() else "",
    }


def read_rules_text() -> str:
    return RULES_FILE.read_text(encoding="utf-8")


def write_rules_text(text: str) -> dict:
    backup_path = RULES_FILE.with_suffix(".json.bak")
    if not backup_path.exists():
        backup_path.write_text(read_rules_text(), encoding="utf-8")
    RULES_FILE.write_text(text, encoding="utf-8")
    return get_rules_meta()


def make_handler(stop_server: threading.Event):
    class Handler(BaseHTTPRequestHandler):
        def _send_json(self, payload: dict, status: int = HTTPStatus.OK) -> None:
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _send_html(self, html: str) -> None:
            body = html.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self) -> None:  # noqa: N802
            if self.path in ("/", "/index.html"):
                self._send_html(HTML_PAGE)
                return
            if self.path == "/api/rules":
                self._send_json({"text": read_rules_text(), "meta": get_rules_meta()})
                return
            self._send_json({"error": "Not found"}, status=HTTPStatus.NOT_FOUND)

        def do_POST(self) -> None:  # noqa: N802
            if self.path != "/api/rules":
                self._send_json({"error": "Not found"}, status=HTTPStatus.NOT_FOUND)
                return

            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length).decode("utf-8")
            try:
                payload = json.loads(raw)
                rules = payload["rules"]
                normalized_text = json.dumps(rules, ensure_ascii=False, indent=2) + "\n"
                meta = write_rules_text(normalized_text)
                if payload.get("close_after_save"):
                    stop_server.set()
                self._send_json({"text": normalized_text, "meta": meta})
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)

        def log_message(self, format: str, *args) -> None:  # noqa: A003
            return

    return Handler


def main() -> int:
    configure_stdio()
    args = parse_args()
    stop_server = threading.Event()
    server = ThreadingHTTPServer((args.host, args.port), make_handler(stop_server))
    url = f"http://{server.server_address[0]}:{server.server_address[1]}/"
    print(f"Priority rules editor: {url}")
    print(f"Editing file: {RULES_FILE}")

    if not args.no_open:
        webbrowser.open(url)

    try:
        while not stop_server.is_set():
            server.handle_request()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
