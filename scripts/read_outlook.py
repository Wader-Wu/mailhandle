import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEV_ENV_FILE = PROJECT_ROOT / ".env"
RUNTIME_ENV_FILE = PROJECT_ROOT / "scripts" / ".env"


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
    parser.add_argument("--folder", choices=["inbox", "sent"], default="inbox")
    parser.add_argument("--limit", type=int, default=10)
    parser.add_argument("--from-contains")
    parser.add_argument("--subject-contains")
    parser.add_argument(
        "--date-preset",
        choices=["today", "last_7_days", "this_month", "last_month"],
    )
    parser.add_argument("--since")
    parser.add_argument("--until")
    parser.add_argument("--unread-only", action="store_true", default=False)
    parser.add_argument("--include-body", action="store_true")
    parser.add_argument("--json", action="store_true")
    return parser.parse_args()


def get_default_windows_python() -> str:
    versions = ("Python313", "Python312", "Python311", "Python310")

    if os.name == "nt":
        local_appdata = os.getenv("LOCALAPPDATA")
        if local_appdata:
            for version in versions:
                candidate = Path(local_appdata) / "Programs" / "Python" / version / "python.exe"
                if candidate.exists():
                    return str(candidate)
        return "python"

    username = os.getenv("USERNAME") or os.getenv("USER") or Path.home().name
    for version in versions:
        candidate = (
            Path("/mnt/c/Users")
            / username
            / "AppData"
            / "Local"
            / "Programs"
            / "Python"
            / version
            / "python.exe"
        )
        if candidate.exists():
            return str(candidate)

    return "/mnt/c/Users/your-user/AppData/Local/Programs/Python/Python310/python.exe"


def get_windows_python() -> str:
    configured = os.getenv("WINDOWS_PYTHON_EXE")
    if configured:
        return configured

    if os.name == "nt":
        executable = Path(sys.executable)
        if executable.exists():
            return str(executable)

        for candidate in ("py", "python"):
            if shutil.which(candidate):
                return candidate

    return get_default_windows_python()


def run_windows_reader(args: argparse.Namespace) -> dict:
    script_path = Path(__file__).with_name("read_outlook_win.py")
    command = [
        get_windows_python(),
        str(script_path),
        "--folder",
        args.folder,
        "--limit",
        str(args.limit),
        "--json",
    ]
    if args.from_contains:
        command.extend(["--from-contains", args.from_contains])
    if args.subject_contains:
        command.extend(["--subject-contains", args.subject_contains])
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

    completed = subprocess.run(
        command,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return json.loads(completed.stdout)


def main() -> int:
    configure_stdio()
    args = parse_args()
    load_env_file(DEV_ENV_FILE if DEV_ENV_FILE.exists() else RUNTIME_ENV_FILE)

    payload = run_windows_reader(args)
    messages = payload.get("messages", [])

    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    print(f"Fetched {len(messages)} messages via Outlook COM")
    if payload.get("filters"):
        print("Filters:", json.dumps(payload["filters"], ensure_ascii=False))
    for message in messages:
        print("Subject:", message.get("subject", "<no subject>"))
        print("Folder:", message.get("folder", "<unknown>"))
        print("Conversation:", message.get("conversation_topic", ""))
        print("From:", message.get("sender", {}).get("display", "<unknown>"))
        print("Received:", message.get("received", {}).get("display", "<unknown>"))
        print("Unread:", message.get("unread", False))
        print("Importance:", message.get("importance", "unknown"))
        print("Categories:", ", ".join(message.get("categories", [])) or "<none>")
        print("Attachments:", message.get("has_attachments", False))
        print("Preview:", message.get("preview", ""))
        if args.include_body:
            print("Body:", message.get("body", ""))
        print("-" * 40)

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
