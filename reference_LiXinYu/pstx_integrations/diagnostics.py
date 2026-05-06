# -*- coding: utf-8 -*-
"""General-purpose local diagnostics logging for PSTX tooling."""

from __future__ import annotations

import hashlib
import io
import json
import os
import platform
import re
import sys
import time
import traceback
import uuid
import zipfile
from pathlib import Path
from typing import Iterable, List, Optional, Tuple


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_LOG_FILE = BASE_DIR / "logs" / "pstx_diagnostics.log"
ASTER_LOG_FILE = BASE_DIR / "logs" / "aster_debug.log"
FEISHU_LOG_FILE = BASE_DIR / "logs" / "feishu_bom_debug.log"
FEISHU_PARSE_LOG_FILE = BASE_DIR / "logs" / "feishu_bom_parse_debug.log"
LOG_FILE_ENV = "PSTX_DIAGNOSTICS_LOG_FILE"
FEISHU_LOG_FILE_ENV = "PSTX_FEISHU_LOG_FILE"
FEISHU_PARSE_LOG_FILE_ENV = "PSTX_FEISHU_PARSE_LOG_FILE"
ENABLED_ENV = "PSTX_DIAGNOSTICS_ENABLED"
MAX_TEXT_PREVIEW = 1200
MAX_LIST_ITEMS = 80
MAX_DICT_ITEMS = 120

SENSITIVE_KEY_PARTS = (
    "secret",
    "token",
    "apikey",
    "api_key",
    "appsecret",
    "app_secret",
    "authorization",
    "password",
    "passwd",
    "ciphertext",
    "access_token",
    "accesskey",
)

EMBEDDED_SECRET_PATTERNS = (
    re.compile(
        r"(?i)\b(authorization\s*[:=]\s*)(bearer\s+)?([^\s,;\"'<>]+)"
    ),
    re.compile(
        r"(?i)\b(api[_-]?key|app[_-]?secret|access[_-]?token|token|secret|password|passwd|ciphertext)\s*[:=]\s*([^\s,;&\"'<>]+)"
    ),
)


def diagnostics_enabled(environ: Optional[dict] = None) -> bool:
    env = environ if environ is not None else os.environ
    value = str(env.get(ENABLED_ENV, "1")).strip().lower()
    return value not in {"0", "false", "no", "off"}


def diagnostics_log_path(log_file: str = "", environ: Optional[dict] = None) -> str:
    env = environ if environ is not None else os.environ
    raw = str(log_file or env.get(LOG_FILE_ENV) or "").strip().strip('"')
    if raw:
        return str(Path(raw).expanduser())
    return str(DEFAULT_LOG_FILE)


def new_diagnostic_request_id() -> str:
    return uuid.uuid4().hex[:12]


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())


def _is_sensitive_key(key: object) -> bool:
    normalized = str(key or "").replace("-", "_").lower()
    return any(part in normalized for part in SENSITIVE_KEY_PARTS)


def _hash_preview(value: object) -> dict:
    text = "" if value is None else str(value)
    return {
        "redacted": True,
        "length": len(text),
        "sha256_12": hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()[:12],
    }


def _truncate_text(value: object, limit: int = MAX_TEXT_PREVIEW) -> str:
    text = "" if value is None else str(value)
    text = _redact_embedded_secrets(text)
    text = text.replace("\r", "\\r")
    return text if len(text) <= limit else text[:limit] + "…"


def _redact_embedded_secrets(text: str) -> str:
    """Redact common secret forms inside generic message/error strings."""
    result = str(text or "")
    result = EMBEDDED_SECRET_PATTERNS[0].sub(r"\1<redacted>", result)
    result = EMBEDDED_SECRET_PATTERNS[1].sub(r"\1=<redacted>", result)
    return result


def sanitize_for_diagnostics(value: object, *, parent_key: str = "") -> object:
    """Return a JSON-safe and secret-redacted copy of value."""
    if _is_sensitive_key(parent_key):
        return _hash_preview(value)
    if isinstance(value, dict):
        result = {}
        for index, (key, child) in enumerate(value.items()):
            if index >= MAX_DICT_ITEMS:
                result["__truncated__"] = True
                result["__remaining__"] = len(value) - MAX_DICT_ITEMS
                break
            result[str(key)] = sanitize_for_diagnostics(child, parent_key=str(key))
        return result
    if isinstance(value, (list, tuple)):
        result = [
            sanitize_for_diagnostics(item, parent_key=parent_key)
            for item in list(value)[:MAX_LIST_ITEMS]
        ]
        if len(value) > MAX_LIST_ITEMS:
            result.append({"__truncated__": True, "__remaining__": len(value) - MAX_LIST_ITEMS})
        return result
    if isinstance(value, (str, int, float, bool)) or value is None:
        if isinstance(value, str):
            return _truncate_text(value)
        return value
    if isinstance(value, Path):
        return str(value)
    return _truncate_text(value)


def summarize_mapping(value: object, *, include_payload: bool = False) -> dict:
    if not isinstance(value, dict):
        return {"present": value is not None, "type": type(value).__name__}
    summary = {
        "present": True,
        "keys": sorted(str(key) for key in value.keys())[:80],
        "key_count": len(value),
        "json_chars": len(json.dumps(value, ensure_ascii=False, default=str)),
    }
    if include_payload:
        summary["payload"] = sanitize_for_diagnostics(value)
    return summary


def summarize_text(value: object, *, include_payload: bool = False) -> dict:
    text = "" if value is None else str(value)
    summary = {
        "present": bool(text),
        "chars": len(text),
        "sha256_12": hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()[:12] if text else "",
    }
    if include_payload:
        summary["preview"] = _truncate_text(text)
    return summary


def format_exception(exc: Optional[BaseException], *, limit: int = 8) -> dict:
    if exc is None:
        return {}
    return {
        "type": exc.__class__.__name__,
        "message": _truncate_text(str(exc), 800),
        "traceback": _truncate_text("".join(traceback.format_exception(type(exc), exc, exc.__traceback__, limit=limit)), 4000),
    }


def write_diagnostic_event(event: str,
                           details: Optional[dict] = None,
                           *,
                           level: str = "info",
                           request_id: str = "",
                           log_file: str = "",
                           environ: Optional[dict] = None) -> dict:
    """Append one sanitized JSONL diagnostics event. Never raises."""
    path = Path(diagnostics_log_path(log_file, environ=environ))
    record = {
        "ts": _now(),
        "level": str(level or "info").lower(),
        "event": str(event or "event"),
        "request_id": request_id or new_diagnostic_request_id(),
        "log_file": str(path),
    }
    record.update(sanitize_for_diagnostics(details or {}))
    if not diagnostics_enabled(environ):
        return {**record, "skipped": True}
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False, sort_keys=True, default=str) + "\n")
    except Exception:
        # Diagnostics must never break the main review workflow.
        pass
    return record


def _file_info(path: Path) -> dict:
    exists = path.exists()
    data = {
        "path": str(path),
        "exists": exists,
        "size_bytes": path.stat().st_size if exists else 0,
        "modified_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(path.stat().st_mtime)) if exists else "",
    }
    return data


def build_diagnostics_status(*, environ: Optional[dict] = None, log_file: str = "") -> dict:
    path = Path(diagnostics_log_path(log_file, environ=environ))
    aster_path = Path(str((environ or os.environ).get("PSTX_ASTER_LOG_FILE") or ASTER_LOG_FILE)).expanduser()
    feishu_path = Path(str((environ or os.environ).get(FEISHU_LOG_FILE_ENV) or FEISHU_LOG_FILE)).expanduser()
    feishu_parse_path = Path(str((environ or os.environ).get(FEISHU_PARSE_LOG_FILE_ENV) or FEISHU_PARSE_LOG_FILE)).expanduser()
    return {
        "ok": True,
        "mode": "diagnostics",
        "enabled": diagnostics_enabled(environ),
        "log_file": _file_info(path),
        "aster_log_file": _file_info(aster_path),
        "feishu_log_file": _file_info(feishu_path),
        "feishu_parse_log_file": _file_info(feishu_parse_path),
        "python": platform.python_version(),
        "platform": platform.platform(),
        "executable": sys.executable,
        "capabilities": [
            "JSONL 通用诊断日志",
            "Web/API 请求生命周期记录",
            "诊断日志 tail 和 zip 导出",
            "飞书在线库 JSONL 调试日志",
            "飞书在线库字段解析专属 JSONL 日志",
        ],
        "safeguards": [
            "secret/token/apiKey/appSecret/Authorization/ciphertext/password 等字段会脱敏。",
            "默认不记录原始 PSTX 文件全文。",
            "日志保存在本地 logs/，不会自动上传。",
        ],
    }


def tail_diagnostics(limit: int = 200, *, log_file: str = "", environ: Optional[dict] = None) -> dict:
    limit = max(1, min(int(limit or 200), 2000))
    path = Path(diagnostics_log_path(log_file, environ=environ))
    if not path.exists():
        return {"ok": True, "path": str(path), "count": 0, "records": [], "raw_lines": []}
    with path.open("r", encoding="utf-8", errors="replace") as handle:
        lines = handle.readlines()[-limit:]
    records = []
    for line in lines:
        text = line.strip()
        if not text:
            continue
        try:
            records.append(json.loads(text))
        except json.JSONDecodeError:
            records.append({"raw": _truncate_text(text)})
    return {
        "ok": True,
        "path": str(path),
        "count": len(records),
        "records": records,
        "raw_lines": [line.rstrip("\n") for line in lines],
    }


def diagnostics_export_bytes(*, include_aster: bool = True, environ: Optional[dict] = None) -> Tuple[bytes, str]:
    env = environ if environ is not None else os.environ
    buffer = io.BytesIO()
    diag_path = Path(diagnostics_log_path(environ=env))
    aster_path = Path(str(env.get("PSTX_ASTER_LOG_FILE") or ASTER_LOG_FILE)).expanduser()
    feishu_path = Path(str(env.get(FEISHU_LOG_FILE_ENV) or FEISHU_LOG_FILE)).expanduser()
    feishu_parse_path = Path(str(env.get(FEISHU_PARSE_LOG_FILE_ENV) or FEISHU_PARSE_LOG_FILE)).expanduser()
    files: List[Tuple[str, Path]] = [("pstx_diagnostics.log", diag_path)]
    if include_aster:
        files.append(("aster_debug.log", aster_path))
    files.append(("feishu_bom_debug.log", feishu_path))
    files.append(("feishu_bom_parse_debug.log", feishu_parse_path))
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("diagnostics_status.json", json.dumps(build_diagnostics_status(environ=env), ensure_ascii=False, indent=2))
        for arcname, path in files:
            if path.exists() and path.is_file():
                archive.write(path, arcname=f"logs/{arcname}")
    filename = "pstx_diagnostics_bundle_" + time.strftime("%Y%m%d_%H%M%S") + ".zip"
    return buffer.getvalue(), filename


def summarize_project_specs(projects: Iterable[object]) -> list:
    summary = []
    for item in projects:
        root = getattr(item, "project_root", "")
        name = getattr(item, "project_name", "")
        summary.append({
            "project_root": str(root or ""),
            "project_name": str(name or ""),
        })
    return summary
