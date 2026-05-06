"""Small request/form parsing helpers for Web routes."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple


def parse_voltage_map_text(text: str) -> Tuple[Optional[Dict[str, float]], List[str]]:
    mapping: Dict[str, float] = {}
    warnings: List[str] = []
    for idx, raw_line in enumerate((text or "").splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            warnings.append(f"电压映射第 {idx} 行缺少 \"=\"：{raw_line.strip()}")
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip()
        if not key:
            warnings.append(f"电压映射第 {idx} 行前缀为空：{raw_line.strip()}")
            continue
        try:
            mapping[key] = float(value)
        except ValueError:
            warnings.append(f"电压映射第 {idx} 行电压不是有效数字：{raw_line.strip()}")
    return mapping or None, warnings


def parse_checkbox_flag(value: object) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on"}
