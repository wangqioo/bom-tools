# -*- coding: utf-8 -*-
"""Shared result metadata helpers for analyzer rules and reports."""

from __future__ import annotations

from collections import Counter
from typing import Dict, List


RESULT_KIND_LABELS = {
    "confirmed": "确定结论",
    "candidate": "候选判断",
    "indeterminate": "无法判断",
}
SEVERITY_LABELS = {
    "high": "高",
    "medium": "中",
    "low": "低",
}
CONFIDENCE_LABELS = {
    "high": "高",
    "medium": "中",
    "low": "低",
}

DRC_ISSUE_KEYS = [
    "missing_hq_code",
    "missing_value",
    "missing_package",
    "tbd_attrs",
    "single_pin_nets",
    "unnamed_nets",
    "bom_option_typos",
    "bom_option_circle_issues",
]


def meta_fields(result_kind: str, severity: str, confidence: str, reason_code: str) -> Dict[str, str]:
    return {
        "结论类型": RESULT_KIND_LABELS[result_kind],
        "严重级别": SEVERITY_LABELS[severity],
        "置信度": CONFIDENCE_LABELS[confidence],
        "原因代码": reason_code,
    }


def with_meta(row: Dict[str, str], result_kind: str, severity: str,
              confidence: str, reason_code: str) -> Dict[str, str]:
    merged = dict(row)
    merged.update(meta_fields(result_kind, severity, confidence, reason_code))
    return merged


def count_result_kinds(rows: List[dict]) -> Counter:
    counter: Counter = Counter()
    for row in rows:
        kind = row.get("结论类型")
        if kind:
            counter[kind] += 1
    return counter


def iter_list_rows(mapping: dict, keys: List[str]) -> List[dict]:
    rows: List[dict] = []
    for key in keys:
        value = mapping.get(key, [])
        if isinstance(value, list):
            rows.extend([row for row in value if isinstance(row, dict)])
    return rows
