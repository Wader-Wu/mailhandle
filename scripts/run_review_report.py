#!/usr/bin/env python3
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = PROJECT_ROOT / "scripts" / "build_review_report.py"


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
    python = resolve_python()
    command = [str(python), str(SCRIPT), *sys.argv[1:]]
    completed = subprocess.run(command)
    return completed.returncode


if __name__ == "__main__":
    raise SystemExit(main())
