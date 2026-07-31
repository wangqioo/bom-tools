# -*- coding: utf-8 -*-
"""Check that the direct deploy app copy is a complete source mirror.

The formal offline package is generated from the project root, but
deploy_bundle/web_app2 is also runnable through deploy_bundle/install_and_run.bat.
This guard catches newly added tools and backend/frontend/requirements drift.
"""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
from pathlib import Path


TOOL_RE = re.compile(r'data-tool="([^"]+)"')
INCLUDE_RE = re.compile(r"\{%\s*include\s+'partials/tools/([^']+\.html)'\s*%\}")
RUNTIME_DIRS = {
    "__pycache__", "auth_data", "uploads", "outputs", "cache", "logs",
    "bug_reports", "feature_requests", "manufacturer_aliases",
}


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _tools(index_html: str) -> set[str]:
    return set(TOOL_RE.findall(index_html))


def _includes(index_html: str) -> set[str]:
    return set(INCLUDE_RE.findall(index_html))


def _source_files(base: Path) -> dict[Path, str]:
    files: dict[Path, str] = {}
    for path in base.rglob("*"):
        if not path.is_file() or any(part in RUNTIME_DIRS for part in path.relative_to(base).parts):
            continue
        relative = path.relative_to(base)
        files[relative] = hashlib.sha256(path.read_bytes()).hexdigest()
    return files


def analyze_deploy_bundle_sync(root: Path) -> list[str]:
    main_index = root / "web_app2" / "templates" / "index.html"
    deploy_index = root / "deploy_bundle" / "web_app2" / "templates" / "index.html"
    deploy_tools_dir = root / "deploy_bundle" / "web_app2" / "templates" / "partials" / "tools"

    main_text = _read(main_index)
    deploy_text = _read(deploy_index)
    main_tools = _tools(main_text)
    deploy_tools = _tools(deploy_text)
    main_includes = _includes(main_text)
    deploy_includes = _includes(deploy_text)

    errors: list[str] = []
    for tool in sorted(main_tools - deploy_tools):
        errors.append(f"deploy_bundle missing nav tool: {tool}")

    for include in sorted(main_includes - deploy_includes):
        errors.append(f"deploy_bundle missing template include: {include}")

    for include in sorted(deploy_includes):
        if not (deploy_tools_dir / include).exists():
            errors.append(f"deploy_bundle missing template file: {include}")

    main_files = _source_files(root / "web_app2")
    deploy_files = _source_files(root / "deploy_bundle" / "web_app2")
    for relative in sorted(main_files.keys() - deploy_files.keys()):
        errors.append(f"deploy_bundle missing source file: {relative.as_posix()}")
    for relative in sorted(main_files.keys() & deploy_files.keys()):
        if main_files[relative] != deploy_files[relative]:
            errors.append(f"deploy_bundle source differs: {relative.as_posix()}")
    for relative in sorted(deploy_files.keys() - main_files.keys()):
        errors.append(f"deploy_bundle has unexpected source file: {relative.as_posix()}")

    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Check deploy_bundle/web_app2 tool sync.")
    parser.add_argument("--root", default=".", help="repository root")
    args = parser.parse_args(argv)

    errors = analyze_deploy_bundle_sync(Path(args.root).resolve())
    if errors:
        print("Deploy bundle sync check failed:", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
