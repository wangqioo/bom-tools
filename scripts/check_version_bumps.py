# -*- coding: utf-8 -*-
"""Require platform/tool version bumps for code changes.

The project has two version layers:
- PLATFORM_VERSION for the whole web platform shell.
- TOOL_VERSIONS for individual tools and sub-features.

This script compares the working tree with a git base revision and fails when
files owned by a versioned area changed but the matching version did not.
"""

from __future__ import annotations

import argparse
import ast
import fnmatch
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


VERSION_FILE = "web_app2/shared.py"


@dataclass(frozen=True)
class VersionRule:
    version_key: str
    patterns: tuple[str, ...]


VERSION_RULES = (
    VersionRule(
        "platform",
        (
            "web_app2/app.py",
            "web_app2/templates/index.html",
        ),
    ),
    VersionRule(
        "tool:bom-compare",
        (
            "web_app2/bom_compare/__init__.py",
            "web_app2/templates/partials/tools/bom-compare.html",
        ),
    ),
    VersionRule(
        "tool:bom-checklist",
        (
            "web_app2/bom_checklist/*",
            "web_app2/templates/partials/tools/bom-checklist.html",
        ),
    ),
    VersionRule(
        "tool:customer-hq-compare",
        (
            "web_app2/bom_compare/customer_hq.py",
            "web_app2/bom_compare/customer_hq_export.py",
        ),
    ),
    VersionRule(
        "tool:free-bom-compare",
        (
            "web_app2/bom_compare/generic_free.py",
        ),
    ),
)

# These files are shared by many tools, so a change must be accompanied by at
# least one platform/tool version bump. The specific owner can be hard to infer
# from a file-level diff because the current frontend is intentionally bundled.
ANY_VERSION_PATTERNS = (
    "web_app2/static/js/app.js",
    "web_app2/static/css/app.css",
)

IGNORED_PATTERNS = (
    "tests/*",
    "scripts/check_version_bumps.py",
)


def run_git(args: list[str], root: Path, check: bool = True) -> str:
    proc = subprocess.run(
        ["git", *args],
        cwd=root,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    if check and proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or "git command failed")
    return proc.stdout


def changed_files(root: Path, base: str) -> list[str]:
    output = run_git(["diff", "--name-only", base], root)
    return [line.strip().replace("\\", "/") for line in output.splitlines() if line.strip()]


def file_at_revision(root: Path, revision: str, path: str) -> str:
    proc = subprocess.run(
        ["git", "show", f"{revision}:{path}"],
        cwd=root,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    return proc.stdout if proc.returncode == 0 else ""


def parse_versions(text: str) -> dict[str, str]:
    versions: dict[str, str] = {}
    if not text.strip():
        return versions
    tree = ast.parse(text)
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id == "PLATFORM_VERSION":
                versions["platform"] = str(ast.literal_eval(node.value))
            if isinstance(target, ast.Name) and target.id == "TOOL_VERSIONS":
                tool_versions = ast.literal_eval(node.value)
                for key, value in tool_versions.items():
                    versions[f"tool:{key}"] = str(value)
    return versions


def path_matches(path: str, patterns: tuple[str, ...]) -> bool:
    return any(fnmatch.fnmatch(path, pattern) for pattern in patterns)


def version_changes(old_versions: dict[str, str], new_versions: dict[str, str]) -> set[str]:
    keys = set(old_versions) | set(new_versions)
    return {key for key in keys if old_versions.get(key) != new_versions.get(key)}


def required_version_keys(paths: list[str]) -> tuple[set[str], bool]:
    required: set[str] = set()
    requires_any = False
    for path in paths:
        if path_matches(path, IGNORED_PATTERNS) or path == VERSION_FILE:
            continue
        for rule in VERSION_RULES:
            if path_matches(path, rule.patterns):
                required.add(rule.version_key)
        if path_matches(path, ANY_VERSION_PATTERNS):
            requires_any = True
    return required, requires_any


def analyze_version_bumps(
    paths: list[str],
    old_versions: dict[str, str],
    new_versions: dict[str, str],
) -> list[str]:
    required, requires_any = required_version_keys(paths)
    changed = version_changes(old_versions, new_versions)
    errors: list[str] = []
    for key in sorted(required):
        if key not in changed:
            errors.append(f"{key} must be bumped")
    if requires_any and not changed:
        errors.append("shared frontend/style changes require at least one platform or tool version bump")
    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Check required platform/tool version bumps.")
    parser.add_argument("--root", default=".", help="repository root")
    parser.add_argument("--base", default="HEAD", help="git revision to compare against")
    args = parser.parse_args(argv)

    root = Path(args.root).resolve()
    paths = changed_files(root, args.base)
    if not paths:
        return 0

    old_versions = parse_versions(file_at_revision(root, args.base, VERSION_FILE))
    new_versions = parse_versions((root / VERSION_FILE).read_text(encoding="utf-8"))
    errors = analyze_version_bumps(paths, old_versions, new_versions)
    if errors:
        print("Version bump check failed:", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        print("\nUpdate PLATFORM_VERSION or TOOL_VERSIONS in web_app2/shared.py.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
