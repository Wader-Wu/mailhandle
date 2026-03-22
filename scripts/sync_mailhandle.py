#!/usr/bin/env python3
import argparse
import json

import mailhandle_db
import mailhandle_runtime
import summarize_mail


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
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
    parser.add_argument("--bootstrap", action="store_true", default=False)
    parser.add_argument("--json", action="store_true", default=False)
    return parser.parse_args()


def main() -> int:
    mailhandle_runtime.configure_stdio()
    args = parse_args()
    mailhandle_db.ensure_database()
    result = mailhandle_runtime.sync_database(args, bootstrap=args.bootstrap)

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        counts = result.get("counts", {})
        print(f"Mode: {result.get('mode', 'unknown')}")
        print(f"Stored new items: {counts.get('stored_count', 0)}")
        print(f"Updated items: {counts.get('updated_count', 0)}")
        print(f"Raw count: {counts.get('raw_count', 0)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
