import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

PRIORITY_ORDER = {"low": 0, "medium": 1, "high": 2}
PROJECT_KEYWORDS = {
    "zhuque": ["zhuque"],
    "nokia": ["nokia"],
    "ciena": ["ciena"],
    "eco": ["eco-"],
}
NOTIFICATION_TERMS = ["notification", "newsletter", "digest", "noreply", "updates."]
MAIL_OWNER_EMAIL = "luqing.wu@lumentum.com"
MAIL_OWNER_NAME = "Luqing Wu"
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEV_ENV_FILE = PROJECT_ROOT / ".env"
RUNTIME_ENV_FILE = PROJECT_ROOT / "scripts" / ".env"
RULES_FILE = PROJECT_ROOT / "scripts" / "priority_rules.json"
ABSTRACT_CACHE_FILE = PROJECT_ROOT / ".cache" / "mailhandle_abstracts.json"
ABSTRACT_CACHE_VERSION = 2
ABSTRACT_BODY_LIMIT = 1800


def configure_stdio() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip())


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
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
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--json", action="store_true")
    return parser.parse_args()


def load_priority_rules() -> dict:
    return json.loads(RULES_FILE.read_text(encoding="utf-8"))


def get_mail_owner() -> dict[str, str]:
    return {
        "email": os.getenv("MAIL_OWNER_EMAIL", MAIL_OWNER_EMAIL),
        "name": os.getenv("MAIL_OWNER_NAME", MAIL_OWNER_NAME),
    }


def get_owner_aliases(owner: dict[str, str], rules_config: dict) -> list[str]:
    aliases: list[str] = []
    configured = rules_config.get("owner_aliases", [])
    if isinstance(configured, list):
        aliases.extend(str(value).strip() for value in configured if str(value).strip())

    owner_name = owner.get("name", "").strip()
    owner_email = owner.get("email", "").strip()
    if owner_name:
        aliases.append(owner_name)
        aliases.extend(part for part in owner_name.split() if part)
    if owner_email:
        local_part = owner_email.split("@", 1)[0]
        aliases.append(local_part)
        aliases.extend(part for part in re.split(r"[._-]+", local_part) if part)

    deduped: list[str] = []
    seen = set()
    for value in aliases:
        normalized = normalize_person_text(value)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        deduped.append(normalized)
    return deduped


def resolve_dates(args: argparse.Namespace) -> tuple[str | None, str | None, str | None]:
    if args.date_preset != "yesterday":
        return args.date_preset, args.since, args.until

    yesterday = date.today() - timedelta(days=1)
    day = yesterday.isoformat()
    return None, day, day


def resolve_subject_filter(args: argparse.Namespace) -> str | None:
    filters = [value for value in [args.subject_contains, args.project] if value]
    if not filters:
        return None
    return " ".join(filters)


def build_reader_command(
    args: argparse.Namespace,
    *,
    folder: str,
    limit: int,
    date_preset: str | None,
    since: str | None,
    until: str | None,
    subject_filter: str | None,
    include_body: bool,
) -> list[str]:
    script_path = Path(__file__).with_name("read_outlook.py")
    command = [
        sys.executable,
        str(script_path),
        "--folder",
        folder,
        "--limit",
        str(limit),
        "--json",
    ]

    if args.from_contains:
        command.extend(["--from-contains", args.from_contains])
    if subject_filter:
        command.extend(["--subject-contains", subject_filter])
    if date_preset:
        command.extend(["--date-preset", date_preset])
    if since:
        command.extend(["--since", since])
    if until:
        command.extend(["--until", until])
    if args.unread_only:
        command.append("--unread-only")
    if include_body:
        command.append("--include-body")
    return command


def run_reader(args: argparse.Namespace, *, folder: str = "inbox", limit: int | None = None) -> dict:
    date_preset, since, until = resolve_dates(args)
    subject_filter = resolve_subject_filter(args)
    command = build_reader_command(
        args,
        folder=folder,
        limit=limit or args.limit,
        date_preset=date_preset,
        since=since,
        until=until,
        subject_filter=subject_filter,
        include_body=args.include_body or folder == "inbox",
    )
    completed = run_command_with_retry(command)
    payload = json.loads(completed.stdout)
    filters = payload.setdefault("filters", {})
    filters["folder"] = folder
    if args.date_preset == "yesterday":
        filters["date_preset"] = "yesterday"
        filters["since"] = since
        filters["until"] = until
    if args.project:
        filters["project"] = args.project
    if subject_filter:
        filters["subject_contains"] = subject_filter
    filters["include_body"] = args.include_body
    return payload


def run_command_with_retry(command: list[str], retries: int = 4) -> subprocess.CompletedProcess:
    last_error = None
    for attempt in range(retries):
        try:
            return subprocess.run(
                command,
                check=True,
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
        except subprocess.CalledProcessError as exc:
            last_error = exc
            stderr = exc.stderr or ""
            if "Call was rejected by callee" not in stderr:
                raise
            wait_seconds = 0.75 * (attempt + 1)
            import time
            time.sleep(wait_seconds)
    raise last_error


def parse_received_iso(message: dict) -> datetime | None:
    received_iso = message.get("received", {}).get("iso")
    if not received_iso:
        return None
    return datetime.fromisoformat(received_iso)


def get_message_timestamp(message: dict) -> datetime | None:
    return parse_received_iso(message)


def text_contains_any(text: str, needles: list[str]) -> bool:
    haystack = text.lower()
    return any(needle.lower() in haystack for needle in needles)


def normalize_subject(subject: str) -> str:
    cleaned = subject.strip().lower()
    cleaned = re.sub(r"^\[(external|ext)\]\s*:?", "", cleaned).strip()
    while True:
        updated = re.sub(r"^(re|fw|fwd|recall)\s*:\s*", "", cleaned).strip()
        if updated == cleaned:
            break
        cleaned = updated
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned


def normalize_person_text(value: str) -> str:
    value = value.strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value


def get_sender_match_keys(message: dict) -> set[str]:
    sender = message.get("sender", {})
    keys = set()
    for value in [
        sender.get("display", ""),
        sender.get("name", ""),
        sender.get("email", ""),
    ]:
        normalized = normalize_person_text(value)
        if normalized:
            keys.add(normalized)
    return keys


def get_sender_display_name(message: dict) -> str:
    sender = message.get("sender", {})
    return (
        sender.get("name")
        or sender.get("email")
        or sender.get("display")
        or "<unknown>"
    )


def get_recipient_match_keys(message: dict) -> set[str]:
    keys = set()
    for recipient in message.get("to", []) + message.get("cc", []):
        for value in [
            recipient.get("display", ""),
            recipient.get("name", ""),
            recipient.get("email", ""),
        ]:
            normalized = normalize_person_text(value)
            if normalized:
                keys.add(normalized)
    return keys


def get_message_thread_key(message: dict) -> str:
    return normalize_subject(message.get("conversation_topic") or message.get("subject", ""))


def get_sent_lookup_window(mail_payload: dict) -> tuple[str | None, str | None]:
    timestamps = [
        timestamp
        for message in mail_payload.get("messages", [])
        if (timestamp := get_message_timestamp(message)) is not None
    ]
    if not timestamps:
        return None, None

    earliest = min(timestamps) - timedelta(days=1)
    latest = max(timestamps) + timedelta(days=7)
    return earliest.date().isoformat(), latest.date().isoformat()


def get_reply_limit(message_count: int) -> int:
    return max(50, min(300, message_count * 8))


def fetch_sent_messages(args: argparse.Namespace, mail_payload: dict) -> list[dict]:
    since, until = get_sent_lookup_window(mail_payload)
    if since is None:
        return []

    subject_filter = resolve_subject_filter(args)
    command = build_reader_command(
        args,
        folder="sent",
        limit=get_reply_limit(len(mail_payload.get("messages", []))),
        date_preset=None,
        since=since,
        until=until,
        subject_filter=subject_filter,
        include_body=False,
    )
    completed = run_command_with_retry(command)
    payload = json.loads(completed.stdout)
    return payload.get("messages", [])


def find_response_match(message: dict, sent_messages: list[dict]) -> dict | None:
    inbound_time = get_message_timestamp(message)
    if inbound_time is None:
        return None

    inbound_thread = get_message_thread_key(message)
    inbound_sender_keys = get_sender_match_keys(message)
    candidates = []

    for sent_message in sent_messages:
        sent_time = get_message_timestamp(sent_message)
        if sent_time is None or sent_time < inbound_time:
            continue
        if get_message_thread_key(sent_message) != inbound_thread:
            continue

        recipient_keys = get_recipient_match_keys(sent_message)
        if inbound_sender_keys and recipient_keys and not inbound_sender_keys.intersection(recipient_keys):
            continue
        candidates.append(sent_message)

    if not candidates:
        return None
    return min(candidates, key=lambda candidate: get_message_timestamp(candidate) or datetime.max)


def infer_projects(message: dict) -> list[str]:
    parts = [
        message.get("subject", ""),
        message.get("conversation_topic", ""),
        message.get("preview", ""),
    ]
    haystack = " ".join(parts).lower()
    tags = []
    for project, keywords in PROJECT_KEYWORDS.items():
        if any(keyword in haystack for keyword in keywords):
            tags.append(project)
    return tags


def recipient_matches_owner(recipient: dict, owner: dict[str, str]) -> bool:
    email = recipient.get("email", "").lower()
    name = recipient.get("name", "").lower()
    owner_email = owner["email"].lower()
    owner_name = owner["name"].lower()
    return owner_email in email or (owner_name and owner_name in name)


def get_attention_flags(message: dict, owner: dict[str, str]) -> list[str]:
    return sorted(set(get_attention_flags_with_rules(message, owner, {})))


def get_message_body_text(message: dict) -> str:
    return str(message.get("body") or message.get("preview") or "").strip()


def clean_body_for_summary(body_text: str) -> str:
    stop_patterns = (
        "-----original message-----",
        "begin forwarded message",
    )
    skip_prefixes = (
        "from:",
        "sent:",
        "to:",
        "cc:",
        "subject:",
        "importance:",
        "external email",
        "this email originated",
    )
    signature_prefixes = (
        "best regards",
        "regards",
        "thanks",
        "thank you",
        "sincerely",
    )

    lines = []
    for raw_line in body_text.splitlines():
        line = " ".join(raw_line.split()).strip()
        if not line:
            if lines and lines[-1] != "":
                lines.append("")
            continue

        lowered = line.lower()
        if lowered.startswith(stop_patterns):
            break
        if re.match(r"^on .+ wrote:$", lowered):
            break
        if any(lowered.startswith(prefix) for prefix in skip_prefixes):
            continue
        if any(lowered.startswith(prefix) for prefix in signature_prefixes):
            break
        if lowered.startswith(">"):
            continue

        line = re.sub(r"^[\*\-\u2022]+\s*", "", line)
        line = re.sub(r"^\d+[.)]\s*", "", line)
        lines.append(line)

    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()

    if lines:
        first_line = lines[0].lower()
        if re.match(
            r"^(hi|hello|dear|good morning|good afternoon|good evening)\b[^.!?]{0,80}$",
            first_line,
        ):
            lines.pop(0)

    cleaned = " ".join(part for part in lines if part)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def get_message_opening_line(message: dict) -> str:
    body_text = clean_body_for_summary(get_message_body_text(message))
    for line in body_text.splitlines():
        normalized = " ".join(line.split()).strip()
        if normalized:
            return normalized
    return " ".join(body_text.split())[:240]


def text_matches_any(text: str, values: list[str]) -> bool:
    haystack = text.lower()
    return any(value.lower() in haystack for value in values if value)


def sender_matches_patterns(message: dict, patterns: list[str]) -> bool:
    sender = message.get("sender", {})
    sender_text = " ".join(
        [
            sender.get("display", ""),
            sender.get("name", ""),
            sender.get("email", ""),
        ]
    ).lower()
    return any(pattern.lower() in sender_text for pattern in patterns if pattern)


def is_owner_greeted(message: dict, owner: dict[str, str], rules_config: dict) -> bool:
    if not any(recipient_matches_owner(recipient, owner) for recipient in message.get("to", [])):
        return False

    greeting_terms = rules_config.get(
        "greeting_terms",
        ["hi", "hello", "dear", "good morning", "good afternoon", "good evening"],
    )
    opening_line = get_message_opening_line(message).lower()
    if not any(opening_line.startswith(term.lower()) for term in greeting_terms):
        return False

    aliases = get_owner_aliases(owner, rules_config)
    return any(alias in opening_line for alias in aliases)


def is_owner_tagged(message: dict, owner: dict[str, str], rules_config: dict) -> bool:
    body_text = get_message_body_text(message).lower()
    aliases = get_owner_aliases(owner, rules_config)
    for alias in aliases:
        compact = alias.replace(" ", "")
        if f"@{alias}" in body_text or (compact and f"@{compact}" in body_text):
            return True
    return False


def get_attention_flags_with_rules(
    message: dict,
    owner: dict[str, str],
    rules_config: dict,
) -> list[str]:
    flags = []
    to_recipients = message.get("to", [])
    cc_recipients = message.get("cc", [])
    body_text = get_message_body_text(message).lower()
    owner_email = owner["email"].lower()
    owner_name = owner["name"].lower()

    if any(recipient_matches_owner(recipient, owner) for recipient in to_recipients):
        flags.append("owner_in_to")
    if any(recipient_matches_owner(recipient, owner) for recipient in cc_recipients):
        flags.append("owner_in_cc")
    if owner_email and owner_email in body_text:
        flags.append("owner_email_mentioned")
    if owner_name and owner_name in body_text:
        flags.append("owner_name_mentioned")
    if is_owner_tagged(message, owner, rules_config):
        flags.append("owner_tagged")
    if is_owner_greeted(message, owner, rules_config):
        flags.append("owner_greeted")
    if sender_matches_patterns(message, rules_config.get("manager_senders", [])):
        flags.append("manager_sender")
    return sorted(set(flags))


def rule_matches(rule: dict, message: dict, rules_config: dict) -> bool:
    subject = message.get("subject", "")
    body_text = get_message_body_text(message)
    sender = message.get("sender", {})
    sender_text = " ".join(
        [
            sender.get("display", ""),
            sender.get("name", ""),
            sender.get("email", ""),
        ]
    )

    if "unread" in rule and message.get("unread") is not rule["unread"]:
        return False
    if "importance" in rule and message.get("importance") not in rule["importance"]:
        return False
    if "subject_contains_any" in rule and not text_contains_any(
        subject, rule["subject_contains_any"]
    ):
        return False
    if "body_contains_any" in rule and not text_contains_any(
        body_text, rule["body_contains_any"]
    ):
        return False
    if "sender_contains_any" in rule and not text_contains_any(
        sender_text, rule["sender_contains_any"]
    ):
        return False
    if "attention_flags_any" in rule:
        flags = set(message.get("_attention_flags", []))
        wanted = {value for value in rule["attention_flags_any"]}
        if not flags.intersection(wanted):
            return False
    if rule.get("sender_matches_manager") and not sender_matches_patterns(
        message, rules_config.get("manager_senders", [])
    ):
        return False
    if "categories_any" in rule:
        categories = {value.lower() for value in message.get("categories", [])}
        wanted = {value.lower() for value in rule["categories_any"]}
        if not categories.intersection(wanted):
            return False
    if "received_within_days" in rule:
        received_at = parse_received_iso(message)
        if received_at is None:
            return False
        now = datetime.now(received_at.tzinfo)
        if received_at < now - timedelta(days=rule["received_within_days"]):
            return False
    return True


def assign_priority(message: dict, rules_config: dict) -> tuple[str, list[str]]:
    chosen_priority = rules_config.get("default_priority", "medium")
    reasons = [f"default:{chosen_priority}"]

    for rule in rules_config.get("rules", []):
        if not rule_matches(rule, message, rules_config):
            continue
        rule_priority = rule.get("priority", chosen_priority)
        if PRIORITY_ORDER.get(rule_priority, -1) >= PRIORITY_ORDER.get(chosen_priority, -1):
            chosen_priority = rule_priority
        reasons.append(f"rule:{rule.get('name', 'unnamed')}")

    return chosen_priority, reasons


def infer_next_action(message: dict, priority: str) -> str:
    response = message.get("_response_match")
    if response:
        return "Monitor for follow-up after your reply"

    subject = message.get("subject", "").lower()
    preview = message.get("preview", "").lower()
    text = f"{subject} {preview}"

    if "review" in text:
        return "Review and send feedback"
    if "ship" in text or "shipment" in text:
        return "Check shipment status and confirm follow-up"
    if "meeting" in text:
        return "Review details and decide whether to respond"
    if "failed" in text or "error" in text:
        return "Investigate issue and respond soon"
    if message.get("unread"):
        return "Read and decide next action"
    if priority == "high":
        return "Review and respond soon"
    return "Review email"


def split_meaningful_sentences(text: str) -> list[str]:
    normalized = clean_body_for_summary(text)
    normalized = re.sub(r"\b(best regards|regards|thanks|thank you)\b.*$", "", normalized, flags=re.IGNORECASE)
    sentences = re.split(r"(?<=[.!?])\s+|(?<=;)\s+", normalized)
    result = []
    for sentence in sentences:
        sentence = sentence.strip(" -\t\r\n")
        if len(sentence) < 20:
            continue
        if sentence.lower().startswith(("from:", "sent:", "to:", "subject:")):
            continue
        result.append(sentence)
    return result


def build_body_based_abstract(message: dict, limit: int = 240) -> str:
    body_text = clean_body_for_summary(get_message_body_text(message))
    sentences = split_meaningful_sentences(body_text)
    if not sentences:
        return message.get("subject", "<no subject>")

    chosen = []
    total = 0
    for sentence in sentences:
        extra = len(sentence) + (1 if chosen else 0)
        if total + extra > limit and chosen:
            break
        chosen.append(sentence)
        total += extra
        if len(chosen) >= 2:
            break
    summary = " ".join(chosen).strip()
    return summary[:limit].rstrip()


def load_abstract_cache() -> dict:
    if not ABSTRACT_CACHE_FILE.exists():
        return {}
    try:
        payload = json.loads(ABSTRACT_CACHE_FILE.read_text(encoding="utf-8"))
        if payload.get("_version") != ABSTRACT_CACHE_VERSION:
            return {}
        cache = payload.get("items", {})
        if isinstance(cache, dict):
            return cache
        return {}
    except Exception:
        return {}


def save_abstract_cache(cache: dict) -> None:
    ABSTRACT_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    ABSTRACT_CACHE_FILE.write_text(
        json.dumps(
            {
                "_version": ABSTRACT_CACHE_VERSION,
                "items": cache,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def get_cached_abstract(message: dict, cache: dict) -> str:
    message_id = str(message.get("id") or "")
    if not message_id:
        return ""
    entry = cache.get(message_id)
    if not isinstance(entry, dict):
        return ""
    if entry.get("subject") != message.get("subject", ""):
        return ""
    if entry.get("received") != message.get("received", {}).get("iso", ""):
        return ""
    return str(entry.get("abstract") or "").strip()


def cache_abstract(message: dict, abstract: str, cache: dict) -> None:
    message_id = str(message.get("id") or "")
    if not message_id or not abstract:
        return
    cache[message_id] = {
        "subject": message.get("subject", ""),
        "received": message.get("received", {}).get("iso", ""),
        "abstract": abstract,
    }


def get_codex_command() -> str | None:
    return shutil.which("codex")


def build_llm_email_payload(message: dict) -> dict | None:
    body_text = clean_body_for_summary(get_message_body_text(message))
    if not body_text:
        return None
    return {
        "subject": message.get("subject", ""),
        "sender": get_sender_display_name(message),
        "body": body_text[:ABSTRACT_BODY_LIMIT],
    }


def sanitize_abstract_text(text: str) -> str:
    cleaned = " ".join(str(text or "").split()).strip()
    cleaned = re.sub(
        r"^(hi|hello|dear|good morning|good afternoon|good evening)\b[^a-zA-Z0-9]{0,3}\s*",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = re.sub(r"^(from|sent|to|cc|subject):\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"^[\*\-\u2022]+\s*", "", cleaned)
    return cleaned.strip(" .;,-")


def request_llm_abstract(email_payload: dict) -> str:
    codex_command = get_codex_command()
    if not codex_command:
        return ""

    prompt = (
        "You are summarizing one Outlook email for a personal mailbox review.\n"
        "Write one concise factual abstract in one sentence.\n"
        "Requirements:\n"
        "- Focus on the main request, decision, status update, or meeting purpose.\n"
        "- Ignore greetings, signatures, legal disclaimers, forwarded headers, quoted reply chains, attendee lists, and bullet numbering.\n"
        "- Do not start with Hi, Hello, Dear, or any header field.\n"
        "- Do not invent deadlines, actions, or owners.\n"
        "- Keep the abstract under 180 characters when possible.\n"
        "- Return JSON only matching the provided schema.\n\n"
        f"{json.dumps(email_payload, ensure_ascii=False)}\n"
    )
    schema = {
        "type": "object",
        "properties": {
            "abstract": {"type": "string"}
        },
        "required": ["abstract"],
        "additionalProperties": False,
    }

    with tempfile.TemporaryDirectory(prefix="mailhandle-abstracts-") as temp_dir:
        temp_dir_path = Path(temp_dir)
        schema_path = temp_dir_path / "schema.json"
        output_path = temp_dir_path / "result.json"
        schema_path.write_text(
            json.dumps(schema, ensure_ascii=False),
            encoding="ascii",
        )

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
        model_name = os.getenv("MAILHANDLE_ABSTRACT_MODEL", "").strip()
        if model_name:
            command[2:2] = ["-m", model_name]

        completed = subprocess.run(
            command,
            input=prompt,
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=120,
        )
        _ = completed
        payload = json.loads(output_path.read_text(encoding="utf-8"))

    return sanitize_abstract_text(payload.get("abstract", ""))


def generate_llm_abstracts(messages: list[dict]) -> dict[str, str]:
    cache = load_abstract_cache()
    abstracts: dict[str, str] = {}
    cache_changed = False

    for message in messages:
        message_id = str(message.get("id") or "")
        if not message_id:
            continue
        cached = get_cached_abstract(message, cache)
        if cached:
            abstracts[message_id] = cached
            continue

        payload = build_llm_email_payload(message)
        if payload is None:
            continue
        try:
            abstract = request_llm_abstract(payload)
        except Exception:
            return abstracts
        if not abstract:
            continue
        abstracts[message_id] = abstract
        cache_abstract(message, abstract, cache)
        cache_changed = True

    if cache_changed:
        save_abstract_cache(cache)
    return abstracts


def build_abstract(
    message: dict,
    priority: str,
    attention_flags: list[str],
    responded: bool,
) -> str:
    llm_abstract = " ".join(str(message.get("_llm_abstract") or "").split()).strip()
    if llm_abstract:
        return llm_abstract

    content_summary = build_body_based_abstract(message)
    if content_summary and normalize_subject(content_summary) != normalize_subject(
        message.get("subject", "")
    ):
        return content_summary

    next_action = infer_next_action(message, priority)
    if next_action:
        return next_action

    return message.get("subject", "<no subject>")


def build_group_summary(message: dict, collapsed_count: int) -> str:
    sender_name = message.get("sender", {}).get("name") or message.get("sender", {}).get("email", "unknown sender")
    subject = message.get("subject", "<no subject>")
    abstract = message.get("_abstract", "")
    if collapsed_count > 1:
        return (
            f"{collapsed_count} related emails from {sender_name} about '{subject}'. "
            f"Latest summary: {abstract}"
        )
    return f"1 email from {sender_name} about '{subject}'. Summary: {abstract}"


def apply_attention_boost(
    priority: str,
    reasons: list[str],
    attention_flags: list[str],
    rules_config: dict,
) -> tuple[str, list[str]]:
    if not attention_flags:
        return priority, reasons
    if not rules_config.get("boost_owner_attention", True):
        return priority, reasons

    boosted = priority
    if any(flag in attention_flags for flag in ["owner_in_to", "owner_email_mentioned", "owner_name_mentioned", "owner_tagged"]):
        if PRIORITY_ORDER.get(boosted, -1) < PRIORITY_ORDER["high"]:
            boosted = "high"
        reasons = reasons + ["attention:owner_direct"]
    elif "owner_in_cc" in attention_flags and PRIORITY_ORDER.get(boosted, -1) < PRIORITY_ORDER["medium"]:
        boosted = "medium"
        reasons = reasons + ["attention:owner_cc"]
    return boosted, reasons


def build_todo(
    message: dict,
    priority: str,
    reasons: list[str],
    owner: dict[str, str],
    rules_config: dict,
) -> dict:
    sender = message.get("sender", {})
    title = message.get("subject", "<no subject>")
    attention_flags = message.get("_attention_flags") or get_attention_flags_with_rules(
        message, owner, rules_config
    )
    priority, reasons = apply_attention_boost(
        priority, reasons, attention_flags, rules_config
    )
    response_match = message.get("_response_match")
    responded = response_match is not None
    if responded:
        reasons = reasons + ["response:already_sent"]
    abstract = build_abstract(message, priority, attention_flags, responded)
    message["_abstract"] = abstract

    todo = {
        "title": title,
        "priority": priority,
        "next_action": infer_next_action(message, priority),
        "reason": ", ".join(reasons),
        "email_id": message.get("id", ""),
        "from": get_sender_display_name(message),
        "received": message.get("received", {}).get("iso", ""),
        "abstract": abstract,
        "projects": infer_projects(message),
        "attention_flags": attention_flags,
        "owner_attention": bool(attention_flags),
        "responded": responded,
        "response_email_id": "",
        "responded_at": "",
        "response_subject": "",
        "collapsed_count": 1,
        "related_email_ids": [message.get("id", "")],
        "group_summary": build_group_summary(message, 1),
        "message": message,
    }
    if responded:
        todo["response_email_id"] = response_match.get("id", "")
        todo["responded_at"] = response_match.get("received", {}).get("iso", "")
        todo["response_subject"] = response_match.get("subject", "")
    return todo


def summarize_todos(todos: list[dict]) -> str:
    if not todos:
        return "No matching emails found."

    high = sum(todo["priority"] == "high" for todo in todos)
    medium = sum(todo["priority"] == "medium" for todo in todos)
    low = sum(todo["priority"] == "low" for todo in todos)
    responded = sum(todo.get("responded", False) for todo in todos)
    grouped = sum(max(todo.get("collapsed_count", 1) - 1, 0) for todo in todos)
    summary = (
        f"Found {len(todos)} todos: {high} high, {medium} medium, {low} low priority."
    )
    if responded:
        summary += f" {responded} already have a sent reply."
    if grouped:
        summary += f" Collapsed {grouped} related emails into grouped todos."
    return summary


def sort_todos(todos: list[dict]) -> list[dict]:
    def received_sort_key(todo: dict) -> float:
        received = todo.get("received", "")
        if not received:
            return float("-inf")
        return datetime.fromisoformat(received).timestamp()

    return sorted(
        todos,
        key=lambda todo: (
            PRIORITY_ORDER.get(todo["priority"], -1),
            received_sort_key(todo),
        ),
        reverse=True,
    )


def should_suppress_todo(todo: dict, rules_config: dict, include_notifications: bool) -> bool:
    if include_notifications:
        return False
    if not rules_config.get("suppress_low_priority_notifications", False):
        return False
    if todo.get("priority") != "low":
        return False

    haystack = " ".join(
        [todo.get("title", ""), todo.get("from", ""), todo.get("abstract", "")]
    ).lower()
    return any(term in haystack for term in NOTIFICATION_TERMS)


def collapse_key(todo: dict) -> tuple[str, str]:
    message = todo.get("message", {})
    sender_email = message.get("sender", {}).get("email", "").lower()
    conversation = message.get("conversation_topic") or todo.get("title", "")
    normalized = normalize_subject(conversation)
    return sender_email, normalized


def collapse_todos(todos: list[dict], collapse_enabled: bool) -> list[dict]:
    if not collapse_enabled:
        return todos

    grouped: dict[tuple[str, str], dict] = {}
    for todo in todos:
        key = collapse_key(todo)
        existing = grouped.get(key)
        if existing is None:
            grouped[key] = todo
            continue

        existing["collapsed_count"] += 1
        existing["related_email_ids"].extend(todo.get("related_email_ids", []))
        existing["responded"] = existing.get("responded", False) or todo.get("responded", False)
        if not existing.get("responded_at") and todo.get("responded_at"):
            existing["responded_at"] = todo["responded_at"]
            existing["response_email_id"] = todo.get("response_email_id", "")
            existing["response_subject"] = todo.get("response_subject", "")

        merged_projects = sorted(set(existing.get("projects", [])) | set(todo.get("projects", [])))
        existing["projects"] = merged_projects

        existing_priority = PRIORITY_ORDER.get(existing["priority"], -1)
        todo_priority = PRIORITY_ORDER.get(todo["priority"], -1)
        existing_received = existing.get("received", "")
        todo_received = todo.get("received", "")

        if todo_priority > existing_priority or (
            todo_priority == existing_priority and todo_received > existing_received
        ):
            merged = todo.copy()
            merged["collapsed_count"] = existing["collapsed_count"]
            merged["related_email_ids"] = existing["related_email_ids"]
            merged["projects"] = merged_projects
            merged["responded"] = existing.get("responded", False) or todo.get("responded", False)
            if not merged.get("responded_at"):
                merged["responded_at"] = existing.get("responded_at", "")
                merged["response_email_id"] = existing.get("response_email_id", "")
                merged["response_subject"] = existing.get("response_subject", "")
            grouped[key] = merged
        else:
            existing["group_summary"] = build_group_summary(existing["message"], existing["collapsed_count"])

    collapsed = list(grouped.values())
    for todo in collapsed:
        todo["group_summary"] = build_group_summary(todo["message"], todo["collapsed_count"])
    return collapsed


def finalize_todo(todo: dict, verbose: bool) -> dict:
    if verbose:
        return todo
    lean = todo.copy()
    lean.pop("message", None)
    return lean


def build_stats(todos: list[dict]) -> dict:
    by_priority = {"high": 0, "medium": 0, "low": 0}
    projects: dict[str, int] = {}
    for todo in todos:
        by_priority[todo["priority"]] += 1
        for project in todo.get("projects", []):
            projects[project] = projects.get(project, 0) + 1
    return {
        "by_priority": by_priority,
        "projects": projects,
        "responded": sum(1 for todo in todos if todo.get("responded")),
        "grouped_threads": sum(1 for todo in todos if todo.get("collapsed_count", 1) > 1),
    }


def format_filter_summary(filters: dict) -> str:
    parts = []
    if filters.get("date_preset"):
        parts.append(f"date={filters['date_preset']}")
    elif filters.get("since") or filters.get("until"):
        parts.append(
            f"date={filters.get('since') or '*'}..{filters.get('until') or '*'}"
        )
    if filters.get("project"):
        parts.append(f"project={filters['project']}")
    if filters.get("subject_contains"):
        parts.append(f"subject~{filters['subject_contains']}")
    if filters.get("from_contains"):
        parts.append(f"from~{filters['from_contains']}")
    if filters.get("unread_only"):
        parts.append("unread_only")
    parts.append(f"limit={filters.get('limit', '?')}")
    return ", ".join(parts)


def format_todo_line(todo: dict) -> list[str]:
    lines = [f"[{todo['priority'].upper()}] {todo['title']}"]
    meta = [f"from {todo['from']}", f"received {todo['received'] or '<unknown>'}"]
    if todo.get("projects"):
        meta.append(f"projects {', '.join(todo['projects'])}")
    if todo.get("collapsed_count", 1) > 1:
        meta.append(f"grouped x{todo['collapsed_count']}")
    lines.append("  " + " | ".join(meta))
    lines.append(f"  next: {todo['next_action']}")
    if todo.get("owner_attention"):
        lines.append(f"  attention: {', '.join(todo['attention_flags'])}")
    if todo.get("responded"):
        response_text = todo.get("responded_at") or "<unknown>"
        response_subject = todo.get("response_subject") or "<unknown>"
        lines.append(f"  replied: {response_text} | {response_subject}")
    lines.append(f"  why: {todo['reason']}")
    lines.append(f"  abstract: {todo['abstract']}")
    return lines


def print_report(result: dict) -> None:
    print(result["summary"])

    filters = result.get("filters", {})
    if filters:
        print("Filters:", format_filter_summary(filters))

    stats = result.get("stats", {})
    print(
        "Stats:",
        ", ".join(
            [
                f"high={stats.get('by_priority', {}).get('high', 0)}",
                f"medium={stats.get('by_priority', {}).get('medium', 0)}",
                f"low={stats.get('by_priority', {}).get('low', 0)}",
                f"responded={stats.get('responded', 0)}",
                f"grouped_threads={stats.get('grouped_threads', 0)}",
            ]
        ),
    )
    if stats.get("projects"):
        print(
            "Projects:",
            ", ".join(f"{name}({count})" for name, count in sorted(stats["projects"].items())),
        )

    active = [todo for todo in result.get("todos", []) if not todo.get("responded")]
    replied = [todo for todo in result.get("todos", []) if todo.get("responded")]

    if active:
        print("\nNeeds action")
        for todo in active:
            for line in format_todo_line(todo):
                print(line)
            print()

    if replied:
        print("Already replied")
        for todo in replied:
            for line in format_todo_line(todo):
                print(line)
            print()


def build_result(args: argparse.Namespace) -> dict:
    load_env_file(DEV_ENV_FILE if DEV_ENV_FILE.exists() else RUNTIME_ENV_FILE)

    mail_payload = run_reader(args)
    sent_messages = fetch_sent_messages(args, mail_payload)
    rules_config = load_priority_rules()
    owner = get_mail_owner()
    llm_abstracts = generate_llm_abstracts(mail_payload.get("messages", []))

    todos = []
    for message in mail_payload.get("messages", []):
        attention_flags = get_attention_flags_with_rules(message, owner, rules_config)
        message = message.copy()
        message["_attention_flags"] = attention_flags
        message["_llm_abstract"] = llm_abstracts.get(str(message.get("id") or ""), "")
        response_match = find_response_match(message, sent_messages)
        if response_match:
            message["_response_match"] = response_match
        priority, reasons = assign_priority(message, rules_config)
        todos.append(build_todo(message, priority, reasons, owner, rules_config))

    todos = [
        todo
        for todo in todos
        if not should_suppress_todo(todo, rules_config, args.include_notifications)
    ]
    todos = collapse_todos(
        todos,
        collapse_enabled=rules_config.get("collapse_similar_emails", False)
        and not args.no_collapse,
    )
    todos = sort_todos(todos)
    return {
        "summary": summarize_todos(todos),
        "filters": mail_payload.get("filters", {}),
        "count": len(todos),
        "stats": build_stats(todos),
        "todos": [finalize_todo(todo, args.verbose) for todo in todos],
    }


def main() -> int:
    configure_stdio()
    args = parse_args()
    result = build_result(args)

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    print_report(result)
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
