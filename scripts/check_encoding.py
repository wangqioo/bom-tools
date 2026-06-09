# -*- coding: utf-8 -*-
"""Check and optionally normalize source text files to UTF-8.

By default this script only checks files and reports anything that cannot be
decoded as UTF-8. With --fix, it rewrites files that are already valid UTF-8 to
UTF-8 without BOM and normalizes CRLF line endings for common source files.
"""

from __future__ import annotations

import argparse
from pathlib import Path


TEXT_EXTENSIONS = {
    ".bat",
    ".cfg",
    ".css",
    ".html",
    ".ini",
    ".js",
    ".json",
    ".md",
    ".ps1",
    ".py",
    ".txt",
    ".yaml",
    ".yml",
}

SKIP_DIRS = {
    ".git",
    ".pytest_cache",
    "__pycache__",
    "deploy_bundle",
    "manufacturer_mapping_extracts",
    "web_app2/auth_data",
    "web_app2/bug_reports",
    "web_app2/cache",
    "web_app2/feature_requests",
    "web_app2/outputs",
    "web_app2/uploads",
}


def is_skipped(path: Path, root: Path) -> bool:
    posix = path.relative_to(root).as_posix()
    return any(posix == item or posix.startswith(item + "/") for item in SKIP_DIRS)


def iter_text_files(root: Path):
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if is_skipped(path, root):
            continue
        if path.suffix.lower() in TEXT_EXTENSIONS:
            yield path


def normalize_utf8(path: Path) -> bool:
    raw = path.read_bytes()
    text = raw.decode("utf-8")
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    encoded = text.encode("utf-8")
    if raw == encoded:
        return False
    path.write_bytes(encoded)
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Check source files are valid UTF-8.")
    parser.add_argument("--fix", action="store_true", help="rewrite valid UTF-8 files as UTF-8 without BOM")
    parser.add_argument("--root", default=".", help="repository root to scan")
    args = parser.parse_args()

    root = Path(args.root).resolve()
    bad = []
    changed = []

    for path in iter_text_files(root):
        rel = path.relative_to(root)
        try:
            if args.fix and normalize_utf8(path):
                changed.append(rel)
            else:
                path.read_text(encoding="utf-8")
        except UnicodeDecodeError as exc:
            bad.append((rel, exc.start, exc.reason))

    if changed:
        print(f"normalized {len(changed)} file(s) to UTF-8 without BOM:")
        for rel in changed:
            print(f"  {rel}")

    if bad:
        print(f"found {len(bad)} non-UTF-8 file(s):")
        for rel, offset, reason in bad:
            print(f"  {rel}: byte {offset}: {reason}")
        return 1

    print("all checked source files are valid UTF-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
