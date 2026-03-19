import argparse
from datetime import datetime, timedelta
import json
import time
import sys

import pythoncom
import win32com.client


FOLDERS = {
    "inbox": {"id": 6, "name": "Inbox", "time_field": "ReceivedTime"},
    "sent": {"id": 5, "name": "Sent Items", "time_field": "SentOn"},
}


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def get_sender(item) -> str:
    sender_email = getattr(item, "SenderEmailAddress", "") or ""
    sender_name = getattr(item, "SenderName", "") or ""
    if sender_name and sender_email:
        return f"{sender_name} <{sender_email}>"
    return sender_email or sender_name or "<unknown>"


def get_sender_fields(item) -> tuple[str, str, str]:
    sender_email = str(getattr(item, "SenderEmailAddress", "") or "")
    sender_name = str(getattr(item, "SenderName", "") or "")
    sender = get_sender(item)
    return sender, sender_name, sender_email


def get_preview(item, max_length: int = 160) -> str:
    body = str(getattr(item, "Body", "") or "")
    preview = " ".join(body.split())
    return preview[:max_length]


def get_body(item, max_length: int | None = None) -> str:
    body = " ".join(str(getattr(item, "Body", "") or "").split())
    if max_length is not None:
        return body[:max_length]
    return body


def get_mailbox_owner() -> dict[str, str]:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    current_user = getattr(namespace, "CurrentUser", None)

    name = str(getattr(current_user, "Name", "") or "").strip()
    address = str(getattr(current_user, "Address", "") or "").strip()
    email = ""

    try:
        address_entry = getattr(current_user, "AddressEntry", None)
        if address_entry is not None:
            exchange_user = address_entry.GetExchangeUser()
            exchange_name = str(getattr(exchange_user, "Name", "") or "").strip()
            exchange_email = str(getattr(exchange_user, "PrimarySmtpAddress", "") or "").strip()
            if exchange_name:
                name = exchange_name
            if exchange_email:
                email = exchange_email
    except Exception:
        pass

    if not email and "@" in address:
        email = address

    accounts = getattr(namespace.Session, "Accounts", None)
    if accounts:
        for index in range(1, int(accounts.Count) + 1):
            account = accounts.Item(index)
            if not email:
                smtp = str(getattr(account, "SmtpAddress", "") or "").strip()
                if smtp:
                    email = smtp
            if not name:
                for attr_name in ("DisplayName", "UserName"):
                    value = str(getattr(account, attr_name, "") or "").strip()
                    if value:
                        name = value
                        break
            if email and name:
                break

    return {"email": email, "name": name}


def format_received(value) -> tuple[str, str]:
    if not value:
        return "<unknown>", ""
    text = str(value)
    iso_text = ""
    try:
        iso_text = value.isoformat()
    except AttributeError:
        iso_text = text
    return text, iso_text


def get_folder_config(folder_key: str) -> dict[str, str | int]:
    return FOLDERS[folder_key]


def get_message_time(item, folder_key: str):
    field_name = str(get_folder_config(folder_key)["time_field"])
    return getattr(item, field_name, None)


def normalize_datetime(value):
    if not value:
        return None
    if getattr(value, "tzinfo", None) is not None:
        return value.replace(tzinfo=None)
    return value


def parse_date(value: str | None, end_of_day: bool = False):
    if not value:
        return None
    text = str(value).strip()
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        parsed = datetime.strptime(text[:10], "%Y-%m-%d")
        if end_of_day:
            parsed = parsed + timedelta(days=1)
    else:
        if getattr(parsed, "tzinfo", None) is not None:
            parsed = parsed.astimezone().replace(tzinfo=None)
        if end_of_day and len(text) <= 10:
            parsed = parsed + timedelta(days=1)
    return parsed


def get_date_bounds(date_preset: str, since: str | None, until: str | None):
    now = datetime.now()
    start = parse_date(since)
    end = parse_date(until, end_of_day=True)

    if date_preset == "today":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=1)
    elif date_preset == "last_2days":
        start = now - timedelta(days=2)
    elif date_preset == "last_7_days":
        start = now - timedelta(days=7)
    elif date_preset == "this_month":
        start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        end = None
    elif date_preset == "last_month":
        first_of_this_month = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        end = first_of_this_month
        previous_month_end = first_of_this_month - timedelta(days=1)
        start = previous_month_end.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

    return start, end


def format_restrict_datetime(value: datetime) -> str:
    return value.strftime("%m/%d/%Y %I:%M %p")


def build_restrict_filter(args: argparse.Namespace, start, end) -> str:
    clauses = []
    time_field = str(get_folder_config(args.folder)["time_field"])

    if start is not None:
        clauses.append(
            f"[{time_field}] >= '{format_restrict_datetime(start)}'"
        )
    if end is not None:
        clauses.append(
            f"[{time_field}] < '{format_restrict_datetime(end)}'"
        )
    if args.folder == "inbox" and args.unread_only:
        clauses.append("[Unread] = True")

    return " AND ".join(clauses)


def get_categories(item) -> list[str]:
    raw = str(getattr(item, "Categories", "") or "")
    if not raw:
        return []
    return [part.strip() for part in raw.split(",") if part.strip()]


def get_importance_label(item) -> str:
    importance = int(getattr(item, "Importance", 1) or 1)
    labels = {0: "low", 1: "normal", 2: "high"}
    return labels.get(importance, "unknown")


def get_recipients(item, recipient_type: int | None = None) -> list[dict[str, str]]:
    recipients = []
    value = getattr(item, "Recipients", None)
    if value is None:
        return recipients

    for recipient in value:
        if recipient_type is not None and int(getattr(recipient, "Type", 0) or 0) != recipient_type:
            continue
        name = str(getattr(recipient, "Name", "") or "")
        address = str(getattr(recipient, "Address", "") or "")
        recipients.append(
            {
                "name": name,
                "email": address,
                "display": f"{name} <{address}>" if name and address else name or address,
            }
        )
    return recipients


def matches_filters(item, args: argparse.Namespace, start, end) -> bool:
    subject = str(getattr(item, "Subject", "") or "")
    sender, sender_name, sender_email = get_sender_fields(item)
    received = normalize_datetime(get_message_time(item, args.folder))

    if args.folder == "inbox" and args.unread_only and not getattr(item, "UnRead", False):
        return False
    if args.subject_contains and args.subject_contains.lower() not in subject.lower():
        return False
    if args.from_contains:
        haystacks = [sender.lower(), sender_name.lower(), sender_email.lower()]
        needle = args.from_contains.lower()
        if not any(needle in haystack for haystack in haystacks):
            return False
    if start and (received is None or received < start):
        return False
    if end and (received is None or received >= end):
        return False
    return True


def fetch_messages(args: argparse.Namespace) -> list[dict]:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder_config = get_folder_config(args.folder)
    folder = namespace.GetDefaultFolder(int(folder_config["id"]))
    items = folder.Items
    start, end = get_date_bounds(args.date_preset, args.since, args.until)
    restrict_filter = build_restrict_filter(args, start, end)
    if restrict_filter:
        items = items.Restrict(restrict_filter)
    items.Sort(f"[{folder_config['time_field']}]", True)

    messages = []
    for index in range(1, items.Count + 1):
        item = get_item_with_retry(items, index)
        if not matches_filters(item, args, start, end):
            continue
        sender, sender_name, sender_email = get_sender_fields(item)
        received_text, received_iso = format_received(get_message_time(item, args.folder))
        messages.append(
            {
                "id": str(getattr(item, "EntryID", "") or ""),
                "store_id": str(getattr(item, "StoreID", "") or ""),
                "folder": str(folder_config["name"]),
                "subject": str(getattr(item, "Subject", "<no subject>")),
                "conversation_topic": str(
                    getattr(item, "ConversationTopic", "") or ""
                ),
                "sender": {
                    "display": sender,
                    "name": sender_name,
                    "email": sender_email,
                },
                "received": {
                    "display": received_text,
                    "iso": received_iso,
                },
                "to": get_recipients(item, 1),
                "cc": get_recipients(item, 2),
                "unread": bool(getattr(item, "UnRead", False)),
                "importance": get_importance_label(item),
                "categories": get_categories(item),
                "preview": get_preview(item),
                "body": get_body(item) if args.include_body else "",
                "has_attachments": bool(getattr(item, "Attachments", []).Count > 0),
            }
        )
        if len(messages) >= args.limit:
            break
    return messages


def get_item_with_retry(items, index: int, retries: int = 5):
    last_error = None
    for attempt in range(retries):
        try:
            return items.Item(index)
        except pythoncom.com_error as exc:
            last_error = exc
            error_text = str(exc)
            retryable_errors = (
                "Call was rejected by callee",
                "Server execution failed",
            )
            if not any(message in error_text for message in retryable_errors):
                raise
            time.sleep(0.5 * (attempt + 1))
    raise last_error


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", choices=sorted(FOLDERS), default="inbox")
    parser.add_argument("--limit", type=int, default=10)
    parser.add_argument("--from-contains")
    parser.add_argument("--subject-contains")
    parser.add_argument(
        "--date-preset",
        choices=["today", "last_2days", "last_7_days", "this_month", "last_month"],
    )
    parser.add_argument("--since")
    parser.add_argument("--until")
    parser.add_argument("--unread-only", action="store_true", default=False)
    parser.add_argument("--include-body", action="store_true")
    parser.add_argument("--mailbox-owner", action="store_true")
    parser.add_argument("--json", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    configure_stdout()
    pythoncom.CoInitialize()
    try:
        owner = get_mailbox_owner() if args.mailbox_owner else None
        messages = [] if args.mailbox_owner else fetch_messages(args)
    finally:
        pythoncom.CoUninitialize()

    if args.mailbox_owner:
        payload = {
            "owner": owner or {"email": "", "name": ""},
            "source": "outlook_com",
        }
        if args.json:
            print(json.dumps(payload, ensure_ascii=False))
            return 0

        print("Mailbox owner:", payload["owner"].get("name", "") or "<unknown>")
        print("Email:", payload["owner"].get("email", "") or "<unknown>")
        return 0

    if args.json:
        print(
            json.dumps(
                {
                    "messages": messages,
                    "count": len(messages),
                    "source": "outlook_com",
                    "filters": {
                        "folder": args.folder,
                        "from_contains": args.from_contains,
                        "subject_contains": args.subject_contains,
                        "date_preset": args.date_preset,
                        "since": args.since,
                        "until": args.until,
                        "unread_only": args.unread_only,
                        "include_body": args.include_body,
                        "limit": args.limit,
                    },
                },
                ensure_ascii=False,
            )
        )
        return 0

    print(f"Reading {len(messages)} messages from {get_folder_config(args.folder)['name']}")
    for message in messages:
        print("Subject:", message["subject"])
        print("Folder:", message["folder"])
        print("Conversation:", message["conversation_topic"])
        print("From:", message["sender"]["display"])
        print("Received:", message["received"]["display"])
        print("Unread:", message["unread"])
        print("Importance:", message["importance"])
        print("Categories:", ", ".join(message["categories"]) or "<none>")
        print("Attachments:", message["has_attachments"])
        print("Preview:", message["preview"])
        if args.include_body:
            print("Body:", message["body"])
        print("-" * 40)

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
