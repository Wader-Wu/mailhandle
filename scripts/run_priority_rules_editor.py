#!/usr/bin/env python3
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = PROJECT_ROOT / "scripts" / "edit_priority_rules.py"


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
    command = [str(resolve_python()), str(SCRIPT), *sys.argv[1:]]
    return subprocess.run(command).returncode


if __name__ == "__main__":
    raise SystemExit(main())
