# -*- coding: utf-8 -*-
"""Local workspace for durable PSTX agent runs and generated artifacts."""

from __future__ import annotations

from pathlib import Path
import json
import os
import re
import time
from typing import Mapping


AGENT_WORKSPACE_VERSION = "pstx-agent-workspace/v1"
AGENT_SCRATCH_FILES_VERSION = "pstx-agent-scratch-files/v1"
WORKSPACE_DIRNAME = "agent_workspace"
MAX_SCRATCH_FILES = 12
MAX_SCRATCH_FILE_CHARS = 200_000


def _repo_root(start: str | Path | None = None) -> Path:
    path = Path(start or ".").expanduser().resolve()
    if path.is_file():
        path = path.parent
    for current in (path, *path.parents):
        if (current / ".git").exists() or (current / "AGENTS.md").is_file():
            return current
    return path


def safe_workspace_id(value: object, fallback: str = "default") -> str:
    text = str(value or "").strip() or fallback
    text = re.sub(r"[^A-Za-z0-9_.-]+", "_", text)
    return (text[:96].strip("._-") or fallback)


def agent_workspace_root(root: str | Path | None = None) -> Path:
    env = str(os.environ.get("PSTX_AGENT_WORKSPACE_DIR") or "").strip()
    if env:
        return Path(env).expanduser().resolve()
    if root is not None:
        return Path(root).expanduser().resolve()
    return _repo_root(root) / WORKSPACE_DIRNAME


def scope_workspace(scope_id: object, *, root: str | Path | None = None) -> Path:
    return agent_workspace_root(root) / safe_workspace_id(scope_id, "scope")


def ensure_scope_workspace(scope_id: object, *, root: str | Path | None = None) -> dict:
    base = scope_workspace(scope_id, root=root)
    paths = {
        "root": base,
        "runs": base / "runs",
        "artifacts": base / "artifacts",
        "drafts": base / "drafts",
        "scratch": base / "scratch",
        "logs": base / "logs",
    }
    for path in paths.values():
        path.mkdir(parents=True, exist_ok=True)
    return {
        "version": AGENT_WORKSPACE_VERSION,
        "scope_id": safe_workspace_id(scope_id, "scope"),
        "root": str(paths["root"]),
        "runs_dir": str(paths["runs"]),
        "artifacts_dir": str(paths["artifacts"]),
        "drafts_dir": str(paths["drafts"]),
        "scratch_dir": str(paths["scratch"]),
        "logs_dir": str(paths["logs"]),
    }


def _safe_join(base: Path, name: object) -> Path:
    filename = safe_workspace_id(name, "artifact")
    target = (base / filename).resolve()
    base_resolved = base.resolve()
    if target != base_resolved and base_resolved not in target.parents:
        raise ValueError("artifact path escaped agent workspace")
    return target


def write_workspace_artifact(scope_id: object,
                             agent_run_id: object,
                             filename: object,
                             content: object,
                             *,
                             root: str | Path | None = None,
                             content_type: str = "text/plain") -> dict:
    workspace = ensure_scope_workspace(scope_id, root=root)
    run_dir = Path(workspace["artifacts_dir"]) / safe_workspace_id(agent_run_id, "run")
    run_dir.mkdir(parents=True, exist_ok=True)
    target = _safe_join(run_dir, filename)
    if isinstance(content, (dict, list)):
        text = json.dumps(content, ensure_ascii=False, indent=2)
    else:
        text = "" if content is None else str(content)
    tmp = target.with_suffix(target.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(target)
    return {
        "version": AGENT_WORKSPACE_VERSION,
        "scope_id": workspace["scope_id"],
        "agent_run_id": safe_workspace_id(agent_run_id, "run"),
        "name": target.name,
        "path": str(target),
        "rel_path": str(target.relative_to(Path(workspace["root"]))),
        "content_type": content_type,
        "size": target.stat().st_size,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }


def write_workspace_draft(scope_id: object,
                          agent_run_id: object,
                          filename: object,
                          content: object,
                          *,
                          root: str | Path | None = None,
                          content_type: str = "text/markdown") -> dict:
    workspace = ensure_scope_workspace(scope_id, root=root)
    run_dir = Path(workspace["drafts_dir"]) / safe_workspace_id(agent_run_id, "run")
    run_dir.mkdir(parents=True, exist_ok=True)
    target = _safe_join(run_dir, filename)
    text = "" if content is None else str(content)
    tmp = target.with_suffix(target.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(target)
    return {
        "version": AGENT_WORKSPACE_VERSION,
        "scope_id": workspace["scope_id"],
        "agent_run_id": safe_workspace_id(agent_run_id, "run"),
        "name": target.name,
        "path": str(target),
        "rel_path": str(target.relative_to(Path(workspace["root"]))),
        "content_type": content_type,
        "size": target.stat().st_size,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }


def _scratch_file_text(content: object, *, max_chars: int = MAX_SCRATCH_FILE_CHARS) -> tuple[str, bool]:
    if isinstance(content, (dict, list)):
        text = json.dumps(content, ensure_ascii=False, indent=2)
    else:
        text = "" if content is None else str(content)
    if len(text) > max_chars:
        return text[:max_chars], True
    return text, False


def write_workspace_scratch_files(scope_id: object,
                                  agent_run_id: object,
                                  files: object,
                                  *,
                                  root: str | Path | None = None,
                                  max_files: int = MAX_SCRATCH_FILES,
                                  max_chars_per_file: int = MAX_SCRATCH_FILE_CHARS) -> dict:
    """Write model-declared temporary files into the run-scoped scratch area.

    The model never receives arbitrary filesystem write access. It may only
    declare small text artifacts; the local runtime sanitizes names, bounds
    content, and writes them under ``agent_workspace/<scope>/scratch/<run>/``.
    """

    workspace = ensure_scope_workspace(scope_id, root=root)
    run_key = safe_workspace_id(agent_run_id, "run")
    run_dir = Path(workspace["scratch_dir"]) / run_key
    run_dir.mkdir(parents=True, exist_ok=True)
    raw_items = files if isinstance(files, list) else []
    written = []
    warnings = []
    for index, item in enumerate(raw_items[:max(1, min(int(max_files or MAX_SCRATCH_FILES), MAX_SCRATCH_FILES))], start=1):
        if not isinstance(item, Mapping):
            warnings.append(f"scratch_files[{index}] ignored: item is not an object")
            continue
        name = item.get("filename") or item.get("name") or f"scratch-{index}.txt"
        target = _safe_join(run_dir, name)
        content = item.get("content")
        if content is None:
            content = item.get("text")
        if content is None:
            content = item.get("body")
        text, truncated = _scratch_file_text(content, max_chars=max(1, min(int(max_chars_per_file or MAX_SCRATCH_FILE_CHARS), MAX_SCRATCH_FILE_CHARS)))
        tmp = target.with_suffix(target.suffix + ".tmp")
        tmp.write_text(text, encoding="utf-8")
        tmp.replace(target)
        if truncated:
            warnings.append(f"{target.name} truncated to {max_chars_per_file} characters")
        written.append({
            "version": AGENT_SCRATCH_FILES_VERSION,
            "scope_id": workspace["scope_id"],
            "agent_run_id": run_key,
            "name": target.name,
            "path": str(target),
            "rel_path": str(target.relative_to(Path(workspace["root"]))),
            "content_type": str(item.get("content_type") or item.get("mime_type") or "text/plain")[:120],
            "size": target.stat().st_size,
            "temporary": True,
            "truncated": truncated,
            "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        })
    if len(raw_items) > MAX_SCRATCH_FILES:
        warnings.append(f"scratch_files limited to {MAX_SCRATCH_FILES} files")
    return {
        "version": AGENT_SCRATCH_FILES_VERSION,
        "scope_id": workspace["scope_id"],
        "agent_run_id": run_key,
        "scratch_dir": str(run_dir),
        "file_count": len(written),
        "files": written,
        "warnings": warnings,
        "temporary": True,
    }


def append_workspace_log(scope_id: object,
                         agent_run_id: object,
                         event: Mapping[str, object],
                         *,
                         root: str | Path | None = None) -> dict:
    workspace = ensure_scope_workspace(scope_id, root=root)
    target = Path(workspace["logs_dir"]) / f"{safe_workspace_id(agent_run_id, 'run')}.jsonl"
    base_resolved = Path(workspace["logs_dir"]).resolve()
    target_resolved = target.resolve()
    if target_resolved != base_resolved and base_resolved not in target_resolved.parents:
        raise ValueError("log path escaped agent workspace")
    payload = dict(event or {})
    payload.setdefault("ts", time.strftime("%Y-%m-%dT%H:%M:%S"))
    with target.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")
    return {
        "version": AGENT_WORKSPACE_VERSION,
        "scope_id": workspace["scope_id"],
        "agent_run_id": safe_workspace_id(agent_run_id, "run"),
        "name": target.name,
        "path": str(target),
        "rel_path": str(target.relative_to(Path(workspace["root"]))),
        "content_type": "application/x-jsonlines",
        "size": target.stat().st_size,
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }


def list_workspace_artifacts(scope_id: object,
                             agent_run_id: object,
                             *,
                             root: str | Path | None = None,
                             limit: int = 100) -> dict:
    workspace = ensure_scope_workspace(scope_id, root=root)
    run_key = safe_workspace_id(agent_run_id, "run")
    search_dirs = [
        Path(workspace["artifacts_dir"]) / run_key,
        Path(workspace["drafts_dir"]) / run_key,
        Path(workspace["scratch_dir"]) / run_key,
    ]
    log_path = Path(workspace["logs_dir"]) / f"{run_key}.jsonl"
    items = []
    for run_dir in search_dirs:
        if not run_dir.is_dir():
            continue
        for path in sorted(run_dir.iterdir(), key=lambda item: item.stat().st_mtime, reverse=True):
            if not path.is_file():
                continue
            items.append({
                "name": path.name,
                "path": str(path),
                "rel_path": str(path.relative_to(Path(workspace["root"]))),
                "size": path.stat().st_size,
                "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(path.stat().st_mtime)),
                "temporary": str(path.relative_to(Path(workspace["root"]))).startswith("scratch/"),
            })
    if log_path.is_file():
        items.append({
            "name": log_path.name,
            "path": str(log_path),
            "rel_path": str(log_path.relative_to(Path(workspace["root"]))),
            "size": log_path.stat().st_size,
            "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(log_path.stat().st_mtime)),
            "temporary": False,
        })
    items = sorted(items, key=lambda item: item.get("updated_at", ""), reverse=True)[:max(1, int(limit or 100))]
    return {
        "version": AGENT_WORKSPACE_VERSION,
        "scope_id": workspace["scope_id"],
        "agent_run_id": run_key,
        "workspace": workspace,
        "artifacts": items,
        "artifact_count": len(items),
    }


def write_task_markdown(scope_id: object,
                        content: object,
                        *,
                        root: str | Path | None = None) -> dict:
    workspace = ensure_scope_workspace(scope_id, root=root)
    target = Path(workspace["root"]) / "TASK.md"
    text = "" if content is None else str(content)
    tmp = target.with_suffix(".md.tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(target)
    return {
        "version": AGENT_WORKSPACE_VERSION,
        "scope_id": workspace["scope_id"],
        "path": str(target),
        "rel_path": "TASK.md",
        "size": target.stat().st_size,
    }


def workspace_status(scope_id: object, *, root: str | Path | None = None) -> dict:
    workspace = ensure_scope_workspace(scope_id, root=root)
    task_path = Path(workspace["root"]) / "TASK.md"
    scratch_dir = Path(workspace["scratch_dir"])
    return {
        "version": AGENT_WORKSPACE_VERSION,
        "workspace": workspace,
        "task_md": {
            "exists": task_path.is_file(),
            "path": str(task_path),
            "size": task_path.stat().st_size if task_path.is_file() else 0,
        },
        "scratch": {
            "exists": scratch_dir.is_dir(),
            "path": str(scratch_dir),
        },
    }
