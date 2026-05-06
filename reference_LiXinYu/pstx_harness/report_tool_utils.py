# -*- coding: utf-8 -*-
"""Shared helper functions for report harness tools."""

from __future__ import annotations

import re
from typing import Iterable, List

from pstx_harness.tool_core import HarnessToolError


BATCH_MAX_ITEMS = 20
BATCH_DEFAULT_PER_ITEM_LIMIT = 10


def _as_int(value, default: int = 0) -> int:
    try:
        return int(value or default)
    except (TypeError, ValueError):
        return default


def _safe_text(value, limit: int = 220) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _compact_rows(rows: Iterable[dict], limit: int) -> List[dict]:
    compacted = []
    for row in list(rows or [])[:max(0, limit)]:
        if isinstance(row, dict):
            compacted.append({str(key): _safe_text(value) for key, value in row.items()})
    return compacted


def _compact_mapping(mapping: dict, limit: int = 24) -> dict:
    output = {}
    for index, (key, value) in enumerate((mapping or {}).items()):
        if index >= limit:
            output["__omitted_count"] = len(mapping) - limit
            break
        output[str(key)] = _safe_text(value, 320)
    return output


def _natural_value_key(value) -> list:
    text = str(value or "")
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


def _batch_input_items(args: dict, key: str, *, max_items: int = BATCH_MAX_ITEMS) -> tuple[List[object], bool]:
    raw = args.get(key)
    if not isinstance(raw, list):
        raise HarnessToolError(f"批量工具需要数组参数：{key}。")
    items = list(raw)
    truncated = len(items) > max_items
    return items[:max_items], truncated


def _batch_limit(value, default: int = BATCH_DEFAULT_PER_ITEM_LIMIT) -> int:
    return max(1, min(_as_int(value, default), BATCH_DEFAULT_PER_ITEM_LIMIT))


def _batch_summary(title: str, items: List[dict], *, truncated: bool = False) -> str:
    found = sum(1 for item in items if item.get("status") == "found")
    missing = sum(1 for item in items if item.get("status") == "missing")
    needs_context = sum(1 for item in items if item.get("status") == "needs_context")
    errors = sum(1 for item in items if item.get("status") == "error")
    suffix = "，输入已按上限截断" if truncated else ""
    return (
        f"{title}完成：{len(items)} 项，命中 {found} 项，缺失 {missing} 项，"
        f"需补充 {needs_context} 项，错误 {errors} 项{suffix}。"
    )


def _sanitize_feishu_row(row: dict) -> dict:
    if not isinstance(row, dict):
        return {}
    raw_fields = row.get("raw_fields") if isinstance(row.get("raw_fields"), dict) else {}
    extra_fields = row.get("extra_field_values") if isinstance(row.get("extra_field_values"), dict) else {}
    return {
        "id": _as_int(row.get("id"), 0),
        "lib_id": _safe_text(row.get("lib_id", ""), 120),
        "lib_name": _safe_text(row.get("lib_name", ""), 160),
        "sheet_name": _safe_text(row.get("sheet_name", ""), 160),
        "key_value": _safe_text(row.get("key_value", ""), 220),
        "hq_no": _safe_text(row.get("hq_no", ""), 120),
        "brand": _safe_text(row.get("brand", ""), 160),
        "spec": _safe_text(row.get("spec", ""), 260),
        "description": _safe_text(row.get("description", ""), 320),
        "pi": _safe_text(row.get("pi", ""), 160),
        "selection_order": _safe_text(row.get("selection_order", ""), 120),
        "extra_field_values": _compact_mapping(extra_fields),
        "raw_fields": _compact_mapping(raw_fields),
        "synced_at": _safe_text(row.get("synced_at", ""), 120),
    }


def _feishu_row_summary(row: dict) -> str:
    hq_no = row.get("hq_no") or row.get("HQ料号") or "UNKNOWN"
    spec = row.get("spec") or row.get("key_value") or ""
    pi = row.get("pi") or ""
    order = row.get("selection_order") or ""
    parts = [f"HQ={hq_no}"]
    if spec:
        parts.append(f"规格={spec}")
    if pi:
        parts.append(f"PI={pi}")
    if order:
        parts.append(f"选型顺序={order}")
    return "；".join(parts)
