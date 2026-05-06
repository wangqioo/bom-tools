# -*- coding: utf-8 -*-
"""Project guidance loader for the PSTX agent runtime.

This module intentionally stays generic: it only reads small Markdown guidance
files and never imports PSTX business modules.
"""

from __future__ import annotations

from pathlib import Path
import re
from typing import Mapping, Sequence


GUIDANCE_VERSION = "pstx-agent-guidance/v1"
GUIDANCE_FILENAMES = ("AGENTS.md", "Agent.md", "CLAUDE.md")
IGNORED_PATH_PARTS = {"trash", "unused_code", "archive"}


def _repo_root(start: str | Path | None = None) -> Path:
    path = Path(start or ".").expanduser().resolve()
    if path.is_file():
        path = path.parent
    for current in (path, *path.parents):
        if (current / ".git").exists() or (current / "AGENTS.md").is_file():
            return current
    return path


def _ignored(path: Path) -> bool:
    return bool(set(part.lower() for part in path.parts) & IGNORED_PATH_PARTS)


def _clean_markdown(text: str, *, max_chars: int) -> str:
    text = re.sub(r"```.*?```", "", text, flags=re.DOTALL)
    lines = []
    for raw in text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        if len(line) > 260:
            line = line[:257] + "..."
        lines.append(line)
        if sum(len(item) + 1 for item in lines) >= max_chars:
            break
    compact = "\n".join(lines).strip()
    return compact[:max_chars]


def _section_lines(text: str, section_markers: Sequence[str], *, max_items: int = 18) -> list[str]:
    markers = tuple(marker.lower() for marker in section_markers)
    capture = False
    result: list[str] = []
    for raw in text.splitlines():
        stripped = raw.strip()
        lowered = stripped.lower()
        if stripped.startswith("#"):
            capture = any(marker in lowered for marker in markers)
            continue
        if not capture:
            continue
        if not stripped or stripped.startswith("```"):
            continue
        result.append(stripped[:220])
        if len(result) >= max_items:
            break
    return result


def find_guidance_files(root: str | Path | None = None) -> list[Path]:
    base = _repo_root(root)
    found: list[Path] = []
    for filename in GUIDANCE_FILENAMES:
        path = base / filename
        if path.is_file() and not _ignored(path):
            found.append(path)
    return found


def load_project_guidance(root: str | Path | None = None, *, max_chars_per_file: int = 6000) -> dict:
    """Load compact project guidance for agent prompts and trace metadata."""

    base = _repo_root(root)
    files = []
    hard_boundaries: list[str] = []
    quick_start: list[str] = []
    summaries: list[str] = []
    for path in find_guidance_files(base):
        try:
            raw = path.read_text(encoding="utf-8", errors="replace")
        except OSError as exc:
            files.append({"path": str(path), "ok": False, "error": str(exc)})
            continue
        compact = _clean_markdown(raw, max_chars=max_chars_per_file)
        hard_boundaries.extend(_section_lines(raw, ("硬边界", "boundaries", "guardrail"), max_items=22))
        quick_start.extend(_section_lines(raw, ("运行入口", "先读哪里", "entry", "quick"), max_items=18))
        files.append({
            "path": str(path),
            "name": path.name,
            "ok": True,
            "chars": len(raw),
            "compact_chars": len(compact),
        })
        summaries.append(f"## {path.name}\n{compact}")
    summary = "\n\n".join(summaries)
    return {
        "version": GUIDANCE_VERSION,
        "root": str(base),
        "files": files,
        "source_count": len([item for item in files if item.get("ok")]),
        "hard_boundaries": hard_boundaries[:28],
        "quick_start": quick_start[:24],
        "summary": summary[: max(12000, max_chars_per_file)],
        "truncated": len(summary) > max(12000, max_chars_per_file),
    }


def compact_guidance_for_model(guidance: Mapping[str, object] | None, *, max_chars: int = 5000) -> dict:
    payload = dict(guidance or {})
    return {
        "version": payload.get("version") or GUIDANCE_VERSION,
        "source_count": int(payload.get("source_count") or 0),
        "files": [
            {"name": item.get("name"), "path": item.get("path"), "ok": item.get("ok")}
            for item in payload.get("files") or []
            if isinstance(item, Mapping)
        ][:6],
        "hard_boundaries": list(payload.get("hard_boundaries") or [])[:18],
        "quick_start": list(payload.get("quick_start") or [])[:12],
        "summary": str(payload.get("summary") or "")[:max_chars],
        "truncated": bool(payload.get("truncated")) or len(str(payload.get("summary") or "")) > max_chars,
    }
