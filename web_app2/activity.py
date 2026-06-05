# -*- coding: utf-8 -*-
"""Activity tracking helpers."""

import json
from functools import wraps
from urllib.parse import unquote


def _payload_from_response(response):
    try:
        return response.get_json(silent=True) or {}
    except Exception:
        return {}


def _filename_from_download(download):
    value = str(download or "")
    if "/download/" in value:
        value = value.split("/download/", 1)[1]
    return unquote(value.rsplit("/", 1)[-1]) if value else ""


def track_tool_activity(tool_name):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            from auth import record_activity

            response = func(*args, **kwargs)
            payload = _payload_from_response(response)
            detail = {
                "tool": tool_name,
                "endpoint": getattr(func, "__name__", ""),
                "success": bool(payload.get("success")),
            }
            if payload.get("download"):
                detail["download"] = payload.get("download")
                detail["filename"] = _filename_from_download(payload.get("download"))
            if payload.get("filename"):
                detail["filename"] = payload.get("filename")
            if payload.get("files"):
                detail["files"] = payload.get("files")
            for key in (
                "total", "matched", "unmatched", "changed", "same", "added", "removed",
                "count", "skipped", "customer_only", "hq_only", "left_only", "right_only",
            ):
                if key in payload:
                    detail[key] = payload.get(key)
            if not payload.get("success") and payload.get("error"):
                detail["error"] = str(payload.get("error"))[:300]
            record_activity(
                "tool_export" if payload.get("success") and payload.get("download") else "tool_run",
                "tool",
                tool_name,
                detail,
            )
            return response
        return wrapper
    return decorator
