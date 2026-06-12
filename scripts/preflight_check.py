# -*- coding: utf-8 -*-
"""Run the standard local checks before packaging or handoff."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def run_step(name: str, args: list[str]) -> bool:
    print(f"\n==> {name}")
    proc = subprocess.run(args, cwd=ROOT)
    if proc.returncode == 0:
        print(f"OK: {name}")
        return True
    print(f"FAILED: {name} (exit {proc.returncode})")
    return False


def main() -> int:
    python = sys.executable
    steps = [
        ("UTF-8 source check", [python, "scripts/check_encoding.py", "--root", "."]),
        ("Python compile check", [python, "-m", "compileall", "web_app2"]),
        ("version bump check", [python, "scripts/check_version_bumps.py", "--root", "."]),
        ("deploy bundle sync check", [python, "scripts/check_deploy_bundle_sync.py", "--root", "."]),
        ("unit tests", [python, "-m", "unittest", "discover", "-s", "tests"]),
    ]
    failed = [name for name, args in steps if not run_step(name, args)]
    if failed:
        print("\nPreflight failed:")
        for name in failed:
            print(f"- {name}")
        return 1
    print("\nPreflight passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
