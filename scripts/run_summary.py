#!/usr/bin/env python3
import json
import sys
from pathlib import Path

import build_review_report
import summarize_mail


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def resolve_python() -> Path:
    candidates = []
    if sys.platform == "win32":
        candidates.append(PROJECT_ROOT / ".venv" / "Scripts" / "python.exe")
    candidates.append(PROJECT_ROOT / ".venv" / "bin" / "python")
    candidates.append(Path(sys.executable))

    for candidate in candidates:
        if candidate.exists():
            return candidate

    return Path(sys.executable)


def main() -> int:
    summarize_mail.configure_stdio()
    args = summarize_mail.parse_args()
    result = summarize_mail.build_result(args)

    report_meta = build_review_report.write_review_report(
        result,
        title=None,
    )
    opened = build_review_report.open_review_report(report_meta["html_path"])

    if args.json:
        payload = result.copy()
        payload["report_meta"] = report_meta
        payload["report_opened"] = opened
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    summarize_mail.print_report(result)
    print()
    print(f"Review HTML: {report_meta['html_path']}")
    print(f"Summary JSON: {report_meta['summary_path']}")
    print(f"Opened report: {'yes' if opened else 'no'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
