"""Compare page request/view helpers.

Heavy compare payload construction lives in ``pstx_webapp.compare_payload``;
this module owns lightweight request coercion and display-oriented summary
helpers shared by routes and payload builders.
"""

from __future__ import annotations

import re
from typing import Any


DEFAULT_COMPARE_DETAIL_LIMIT = 500
MAX_COMPARE_DETAIL_LIMIT = 5000
PASSIVE_REFDES_PREFIXES = {"R", "C", "L"}
CHIP_REFDES_PREFIXES = {"U", "PU", "XU"}
CONNECTOR_REFDES_PREFIXES = {"J", "P", "CN", "CON", "X", "XS", "XP"}
POWER_PASSIVE_REFDES_RE = re.compile(r"^P([RCL])\d+(?:A\d+)?$", re.I)


def coerce_compare_detail_limit(
    value: Any,
    *,
    default_limit: int = DEFAULT_COMPARE_DETAIL_LIMIT,
    max_limit: int = MAX_COMPARE_DETAIL_LIMIT,
) -> int:
    if value in (None, ""):
        return default_limit
    try:
        limit = int(value)
    except (TypeError, ValueError):
        raise ValueError("detail_limit 必须是整数。")
    if limit <= 0:
        raise ValueError("detail_limit 必须大于 0。")
    return min(limit, max_limit)


def build_compare_scalar_metrics(left: dict, right: dict) -> list[dict]:
    labels = list(dict.fromkeys(
        list(left.get("metric_map", {}).keys())
        + list(right.get("metric_map", {}).keys())
        + ["component_count", "net_count", "drc_count"]
    ))
    rows = []
    for label in labels:
        if label == "component_count":
            left_value, right_value, display_label = left.get("component_count", 0), right.get("component_count", 0), "元件数"
        elif label == "net_count":
            left_value, right_value, display_label = left.get("net_count", 0), right.get("net_count", 0), "网络数"
        elif label == "drc_count":
            left_value, right_value, display_label = left.get("drc_count", 0), right.get("drc_count", 0), "DRC 问题数"
        else:
            left_value = left.get("metric_map", {}).get(label, "")
            right_value = right.get("metric_map", {}).get(label, "")
            display_label = label
        if left_value == right_value:
            continue
        delta = ""
        if isinstance(left_value, (int, float)) and isinstance(right_value, (int, float)):
            diff = right_value - left_value
            delta = f"{diff:+g}"
        rows.append({
            "指标": display_label,
            "左侧": left_value,
            "右侧": right_value,
            "变化": delta or "changed",
        })
    return rows


def refdes_prefix(refdes: str) -> str:
    text = str(refdes or "").strip().upper()
    power_passive = POWER_PASSIVE_REFDES_RE.match(text)
    if power_passive:
        return power_passive.group(1).upper()
    prefix = []
    for char in text:
        if char.isdigit():
            break
        if char.isalpha():
            prefix.append(char)
            continue
        break
    return "".join(prefix) or text


def refdes_category(refdes: str) -> str:
    prefix = refdes_prefix(refdes)
    if prefix in PASSIVE_REFDES_PREFIXES:
        return "passive"
    if prefix in CHIP_REFDES_PREFIXES:
        return "chip"
    if prefix in CONNECTOR_REFDES_PREFIXES:
        return "connector"
    return "key_other"


def refdes_category_label(category: str) -> str:
    return {
        "chip": "芯片",
        "connector": "连接器",
        "passive": "R/C/L",
        "key_other": "关键器件",
    }.get(str(category or ""), "关键器件")


def is_key_refdes(refdes: str) -> bool:
    return refdes_category(refdes) != "passive"


def is_passive_refdes(refdes: str) -> bool:
    return refdes_category(refdes) == "passive"
