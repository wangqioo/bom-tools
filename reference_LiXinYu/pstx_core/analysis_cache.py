# -*- coding: utf-8 -*-
"""Local derived-result cache for expensive analysis helpers."""

from __future__ import annotations

import hashlib
import json
import os
import time
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, Mapping, Tuple


ANALYSIS_CACHE_SCHEMA_VERSION = "pstx-analysis-derived-cache.v1"
ANALYSIS_CACHE_VERSION = "2026-05-06.1"
ANALYSIS_CACHE_DIR_ENV = "PSTX_ANALYSIS_CACHE_DIR"
DISABLE_ANALYSIS_CACHE_ENV = "PSTX_DISABLE_ANALYSIS_CACHE"
DEFAULT_ANALYSIS_CACHE_DIR = Path("output") / "analysis_cache"


def analysis_cache_enabled() -> bool:
    raw = str(os.environ.get(DISABLE_ANALYSIS_CACHE_ENV) or "").strip().lower()
    return raw not in {"1", "true", "yes", "on"}


def analysis_cache_dir() -> Path:
    raw = str(os.environ.get(ANALYSIS_CACHE_DIR_ENV) or "").strip()
    return Path(raw).expanduser() if raw else DEFAULT_ANALYSIS_CACHE_DIR


def _stable_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _stable_hash(value: Any) -> str:
    return hashlib.sha256(_stable_json(value).encode("utf-8")).hexdigest()


def _page_file_fingerprint(project_root: str | Path,
                           extensions: Iterable[str]) -> Dict[str, Any]:
    root = Path(project_root).expanduser()
    sch_dir = root / "sch_1"
    normalized_exts = {str(ext).lower() for ext in extensions}
    files = []
    if sch_dir.is_dir():
        for path in sorted(sch_dir.iterdir(), key=lambda item: item.name.lower()):
            if not path.is_file() or path.suffix.lower() not in normalized_exts:
                continue
            try:
                stat = path.stat()
            except OSError:
                continue
            files.append({
                "rel_path": str(path.relative_to(root)),
                "size": int(stat.st_size),
                "mtime_ns": int(stat.st_mtime_ns),
            })
    return {
        "project_root": str(root.resolve()) if root.exists() else str(root),
        "sch_dir": str(sch_dir.resolve()) if sch_dir.exists() else str(sch_dir),
        "file_count": len(files),
        "files": files,
    }


def build_analysis_cache_identity(kind: str,
                                  project_root: str | Path,
                                  *,
                                  params: Mapping[str, Any] | None = None,
                                  extensions: Iterable[str] = (".csa", ".csv")) -> Dict[str, Any]:
    return {
        "schema_version": ANALYSIS_CACHE_SCHEMA_VERSION,
        "cache_version": ANALYSIS_CACHE_VERSION,
        "kind": str(kind or ""),
        "params": dict(params or {}),
        "fingerprint": _page_file_fingerprint(project_root, extensions),
    }


def _status(kind: str,
            identity: Mapping[str, Any],
            *,
            enabled: bool,
            cache_key: str = "",
            path: str = "",
            status: str = "miss",
            reason: str = "",
            elapsed_s: float = 0.0) -> Dict[str, Any]:
    fingerprint = identity.get("fingerprint") if isinstance(identity, Mapping) else {}
    return {
        "kind": kind,
        "enabled": enabled,
        "status": status,
        "hit": status == "hit",
        "cache_key": cache_key,
        "path": path,
        "reason": reason,
        "file_count": int((fingerprint or {}).get("file_count", 0) or 0) if isinstance(fingerprint, Mapping) else 0,
        "elapsed_ms": round(float(elapsed_s) * 1000.0, 3),
    }


def get_or_compute_cached_result(kind: str,
                                 project_root: str | Path,
                                 *,
                                 params: Mapping[str, Any] | None = None,
                                 extensions: Iterable[str] = (".csa", ".csv"),
                                 compute: Callable[[], Dict[str, Any]]) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    started = time.perf_counter()
    identity = build_analysis_cache_identity(
        kind,
        project_root,
        params=params,
        extensions=extensions,
    )
    cache_key = _stable_hash(identity)
    cache_path = analysis_cache_dir() / str(kind or "analysis") / f"{cache_key}.json"
    if not analysis_cache_enabled():
        result = compute()
        return result, _status(
            kind,
            identity,
            enabled=False,
            cache_key=cache_key,
            path=str(cache_path),
            status="disabled",
            reason=DISABLE_ANALYSIS_CACHE_ENV,
            elapsed_s=time.perf_counter() - started,
        )

    try:
        if cache_path.is_file():
            cached = json.loads(cache_path.read_text(encoding="utf-8"))
            if cached.get("identity") == identity and isinstance(cached.get("result"), dict):
                return cached["result"], _status(
                    kind,
                    identity,
                    enabled=True,
                    cache_key=cache_key,
                    path=str(cache_path),
                    status="hit",
                    elapsed_s=time.perf_counter() - started,
                )
    except Exception as exc:
        cache_read_error = str(exc)
    else:
        cache_read_error = ""

    result = compute()
    status_value = "miss"
    reason = cache_read_error
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(
            json.dumps(
                {
                    "schema_version": ANALYSIS_CACHE_SCHEMA_VERSION,
                    "identity": identity,
                    "result": result,
                    "written_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
                },
                ensure_ascii=False,
                sort_keys=True,
            ),
            encoding="utf-8",
        )
    except Exception as exc:
        status_value = "write_error"
        reason = str(exc)

    return result, _status(
        kind,
        identity,
        enabled=True,
        cache_key=cache_key,
        path=str(cache_path),
        status=status_value,
        reason=reason,
        elapsed_s=time.perf_counter() - started,
    )


__all__ = [
    "ANALYSIS_CACHE_DIR_ENV",
    "ANALYSIS_CACHE_SCHEMA_VERSION",
    "ANALYSIS_CACHE_VERSION",
    "DEFAULT_ANALYSIS_CACHE_DIR",
    "DISABLE_ANALYSIS_CACHE_ENV",
    "analysis_cache_dir",
    "analysis_cache_enabled",
    "build_analysis_cache_identity",
    "get_or_compute_cached_result",
]
