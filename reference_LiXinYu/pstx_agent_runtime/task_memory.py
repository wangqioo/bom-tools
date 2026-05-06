# -*- coding: utf-8 -*-
"""Markdown task memory for multi-turn PSTX Harness Agent runs."""

from __future__ import annotations

from pathlib import Path
import os
import re
import time
from typing import Mapping, Sequence


TASK_MEMORY_VERSION = "pstx-agent-task-memory/v1"


def _repo_root(start: str | Path | None = None) -> Path:
    path = Path(start or ".").expanduser().resolve()
    if path.is_file():
        path = path.parent
    for current in (path, *path.parents):
        if (current / ".git").exists() or (current / "AGENTS.md").is_file():
            return current
    return path


def _memory_root(root: str | Path | None = None) -> Path:
    env = str(os.environ.get("PSTX_AGENT_MEMORY_DIR") or "").strip()
    if env:
        return Path(env).expanduser()
    return _repo_root(root) / "agent_memory"


def _safe_id(value: object, fallback: str = "default") -> str:
    text = str(value or "").strip() or fallback
    text = re.sub(r"[^A-Za-z0-9_.-]+", "_", text)
    return text[:96] or fallback


def task_memory_path(run_id: object, *, root: str | Path | None = None) -> Path:
    return _memory_root(root) / _safe_id(run_id, "run") / "TASK.md"


def _preview(value: object, limit: int = 500) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _bullet_items(items: Sequence[object], *, limit: int = 12) -> str:
    lines = []
    for item in list(items or [])[:limit]:
        if isinstance(item, Mapping):
            text = item.get("title") or item.get("summary") or item.get("question") or item.get("id") or item
        else:
            text = item
        preview = _preview(text, 240)
        if preview:
            lines.append(f"- {preview}")
    return "\n".join(lines) if lines else "- 暂无"


def _extract_open_questions(payload: Mapping[str, object]) -> list[object]:
    needs = payload.get("needs_user_input") if isinstance(payload.get("needs_user_input"), Mapping) else {}
    return list(needs.get("questions") or []) if isinstance(needs, Mapping) else []


def _extract_evidence_ids(payload: Mapping[str, object]) -> list[str]:
    result: list[str] = []
    for key in ("citations", "final_evidence"):
        for item in payload.get(key) or []:
            if not isinstance(item, Mapping):
                continue
            evidence_id = str(item.get("id") or "").strip()
            if evidence_id and evidence_id not in result:
                result.append(evidence_id)
            if len(result) >= 40:
                return result
    return result


def read_task_memory(run_id: object, *, root: str | Path | None = None, max_chars: int = 6000) -> dict:
    path = task_memory_path(run_id, root=root)
    if not path.is_file():
        return {
            "version": TASK_MEMORY_VERSION,
            "found": False,
            "path": str(path),
            "summary": "",
            "chars": 0,
        }
    text = path.read_text(encoding="utf-8", errors="replace")
    return {
        "version": TASK_MEMORY_VERSION,
        "found": True,
        "path": str(path),
        "summary": text[:max_chars],
        "chars": len(text),
        "truncated": len(text) > max_chars,
    }


def write_task_memory(run_id: object,
                      payload: Mapping[str, object],
                      *,
                      root: str | Path | None = None) -> dict:
    path = task_memory_path(run_id, root=root)
    path.parent.mkdir(parents=True, exist_ok=True)
    trace = payload.get("trace_summary") if isinstance(payload.get("trace_summary"), Mapping) else {}
    model_meta = payload.get("model_metadata") if isinstance(payload.get("model_metadata"), Mapping) else {}
    runtime_state = payload.get("runtime_state") if isinstance(payload.get("runtime_state"), Mapping) else {}
    task_ledger = runtime_state.get("task_ledger") if isinstance(runtime_state.get("task_ledger"), Mapping) else {}
    next_actions = task_ledger.get("next_actions") if isinstance(task_ledger, Mapping) else []
    selected_skills = payload.get("selected_skills") if isinstance(payload.get("selected_skills"), Mapping) else {}
    selected_skill_items = selected_skills.get("selected_skills") if isinstance(selected_skills, Mapping) else []
    guidance = payload.get("guidance_summary") if isinstance(payload.get("guidance_summary"), Mapping) else {}
    lines = [
        "# PSTX Agent Task Memory",
        "",
        f"- version: `{TASK_MEMORY_VERSION}`",
        f"- run_id: `{_safe_id(run_id, 'run')}`",
        f"- agent_run_id: `{payload.get('agent_run_id', '')}`",
        f"- profile: `{payload.get('profile', '')}`",
        f"- status: `{payload.get('status', '')}`",
        f"- stopped_reason: `{model_meta.get('stopped_reason', '')}`",
        f"- updated_at: `{time.strftime('%Y-%m-%dT%H:%M:%S')}`",
        "",
        "## Goal",
        "",
        _preview((runtime_state.get("goal") if isinstance(runtime_state, Mapping) else "") or "", 800) or "-",
        "",
        "## Answer Summary",
        "",
        _preview(payload.get("answer") or "", 1200) or "-",
        "",
        "## Guidance",
        "",
        f"- source_count: `{guidance.get('source_count', 0)}`",
        _bullet_items(guidance.get("hard_boundaries") or [], limit=8),
        "",
        "## Selected Skills",
        "",
        _bullet_items([
            f"{item.get('id')}: {item.get('description') or item.get('title')}"
            for item in selected_skill_items or []
            if isinstance(item, Mapping)
        ], limit=8),
        "",
        "## Evidence IDs",
        "",
        _bullet_items(_extract_evidence_ids(payload), limit=40),
        "",
        "## Open Questions",
        "",
        _bullet_items(_extract_open_questions(payload), limit=12),
        "",
        "## Next Actions",
        "",
        _bullet_items(next_actions if isinstance(next_actions, list) else [], limit=16),
        "",
        "## Trace",
        "",
        f"- tool_call_count: `{trace.get('tool_call_count', 0)}`",
        f"- evidence_node_count: `{trace.get('evidence_node_count', 0)}`",
        f"- input_truncated: `{trace.get('input_truncated', False)}`",
    ]
    text = "\n".join(lines).strip() + "\n"
    path.write_text(text, encoding="utf-8")
    return {
        "version": TASK_MEMORY_VERSION,
        "found": True,
        "path": str(path),
        "chars": len(text),
        "summary": text[:6000],
        "truncated": len(text) > 6000,
    }
