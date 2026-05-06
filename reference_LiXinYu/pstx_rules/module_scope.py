# -*- coding: utf-8 -*-
"""Module-level review scope derived from module_order and component paths."""

from __future__ import annotations

import re
from collections import defaultdict
from typing import Dict, List, Optional

from pstx_core import pages as page_logic
from pstx_core.page_resolution import MAIN_MODULE_PAGE_LABEL, USER_VISIBLE_REAL_PAGE_LABEL, component_user_visible_page
from pstx_rules.common import _display_bom_option, _is_depop_option, _natural_sort_key

_MODULE_SEGMENT_RE = re.compile(
    r"^(?P<head>.+?)\((?P<view>[^)]+)\)(?::(?P<tail>.+))?$",
    re.IGNORECASE,
)
_INSTANCE_RE = re.compile(r"PAGE[_\-/ ]*\d+[A-Z]?[_\-/ ]*(?P<instance>I\d+)", re.IGNORECASE)


def _normalize_module_key(value: str) -> str:
    return str(value or "").strip().upper()


def _format_page_number(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if text.upper().startswith("PAGE"):
        return page_logic.normalize_page_label(text)
    return page_logic.normalize_page_label(f"PAGE{text}")


def _page_end(start_page: str, page_count: object) -> str:
    match = re.search(r"(\d+)", str(start_page or ""))
    if not match:
        return ""
    try:
        count = int(page_count or 0)
    except (TypeError, ValueError):
        count = 0
    if count <= 0:
        return ""
    return _format_page_number(int(match.group(1)) + count - 1)


def _parse_module_order_path(path_text: str) -> List[Dict[str, str]]:
    segments: List[Dict[str, str]] = []
    for chunk in [part.strip() for part in str(path_text or "").split("@") if part.strip()]:
        match = _MODULE_SEGMENT_RE.match(chunk)
        if not match:
            continue
        head = match.group("head").strip()
        lib, _, cell = head.rpartition(".")
        tail = str(match.group("tail") or "").strip()
        instance_match = _INSTANCE_RE.search(tail)
        segments.append({
            "raw": chunk,
            "head": head,
            "lib": lib.strip(),
            "cell": (cell or head).strip(),
            "view": match.group("view").strip(),
            "tail": tail,
            "page": page_logic.normalize_page_label(tail) if tail else "",
            "instance": instance_match.group("instance").upper() if instance_match else "",
        })
    return segments


def _module_name_from_component(comp: Dict[str, object], project_name: str = "") -> str:
    context = str(comp.get("page_context", "") or "")
    if ":" in context:
        name = context.split(":", 1)[0].strip()
        if name:
            return name
    raw_path = str(comp.get("page_path_logical_raw", "") or comp.get("drawing", "") or "")
    segments = page_logic.extract_path_segments(raw_path)
    if segments:
        return segments[0].get("cell", "") or project_name
    return project_name or "主模块"


def _module_entry_row(entry: Dict[str, object]) -> Dict[str, object]:
    segments = _parse_module_order_path(str(entry.get("path", "")))
    child = segments[-1] if segments else {}
    parent_segments = segments[:-1]
    symbol_segment = parent_segments[-1] if parent_segments else {}
    start_page = _format_page_number(entry.get("start_real_page", ""))
    page_count = int(entry.get("page_count", 0) or 0)
    symbol_locator = "可定位父级Symbol页/实例" if symbol_segment.get("page") and symbol_segment.get("instance") else "仅有module_order路径"
    return {
        "模块ID": _normalize_module_key(str(entry.get("path_key", "") or entry.get("path", ""))),
        "模块类型": "子模块",
        "模块名": child.get("cell", ""),
        "模块库": child.get("lib", ""),
        "父级模块路径": " / ".join(segment.get("cell", "") for segment in parent_segments if segment.get("cell", "")),
        "父级Symbol页码": symbol_segment.get("page", ""),
        "父级Symbol实例": symbol_segment.get("instance", ""),
        "父级Symbol定位": symbol_locator,
        "起始页码": start_page,
        "结束页码": _page_end(start_page, page_count),
        "页数": page_count,
        "module_order路径": str(entry.get("path", "")),
        "module_order来源": str(entry.get("source_file", "")),
    }


def _flatten_module_order_entries(module_order_index: Dict[str, object]) -> List[Dict[str, object]]:
    entries: List[Dict[str, object]] = []
    for bucket in (module_order_index or {}).get("by_key", {}).values():
        if isinstance(bucket, list):
            entries.extend([entry for entry in bucket if isinstance(entry, dict)])
    return sorted(entries, key=lambda item: _natural_sort_key(item.get("path", "")))


def build_module_review(components: Dict[str, dict],
                        module_order_index: Optional[Dict[str, object]],
                        *,
                        project_name: str = "") -> Dict[str, object]:
    """Build an incremental main/submodule review index.

    The existing project-wide analysis remains unchanged. This helper only
    adds a module lens so UI/Harness can filter or summarize by main module and
    module_order child instances.
    """

    module_rows_by_id: Dict[str, Dict[str, object]] = {}
    for entry in _flatten_module_order_entries(module_order_index or {}):
        row = _module_entry_row(entry)
        if row["模块ID"]:
            module_rows_by_id[str(row["模块ID"])] = row

    main_module_name = next(
        (_module_name_from_component(comp, project_name) for comp in components.values()),
        project_name or "主模块",
    )
    module_rows_by_id["__MAIN__"] = {
        "模块ID": "__MAIN__",
        "模块类型": "主模块",
        "模块名": main_module_name,
        "模块库": "",
        "父级模块路径": "",
        "父级Symbol页码": "",
        "父级Symbol实例": "",
        "父级Symbol定位": "主模块自身",
        "起始页码": "",
        "结束页码": "",
        "页数": "",
        "module_order路径": "",
        "module_order来源": "",
    }

    counters: Dict[str, Dict[str, object]] = defaultdict(lambda: {
        "元件数量": 0,
        "芯片数量": 0,
        "R/C/L数量": 0,
        "DEPOP数量": 0,
        "示例位号": [],
    })
    component_rows: List[Dict[str, object]] = []
    for refdes, comp in sorted(components.items(), key=lambda item: _natural_sort_key(item[0])):
        module_key = _normalize_module_key(str(comp.get("module_order_key", "") or ""))
        module_id = module_key if module_key in module_rows_by_id else "__MAIN__"
        module_row = module_rows_by_id[module_id]
        comp_type = str(comp.get("comp_type", "") or "")
        bom_option = str(comp.get("bom_option", "") or "")
        counter = counters[module_id]
        counter["元件数量"] = int(counter["元件数量"]) + 1
        if comp_type == "IC":
            counter["芯片数量"] = int(counter["芯片数量"]) + 1
        if comp_type in {"RES", "CAP", "IND"}:
            counter["R/C/L数量"] = int(counter["R/C/L数量"]) + 1
        if _is_depop_option(bom_option):
            counter["DEPOP数量"] = int(counter["DEPOP数量"]) + 1
        samples = counter["示例位号"]
        if isinstance(samples, list) and len(samples) < 8:
            samples.append(refdes)

        component_rows.append({
            "模块类型": module_row.get("模块类型", ""),
            "模块名": module_row.get("模块名", ""),
            "模块ID": module_id,
            "位号": refdes,
            "器件类型": comp_type,
            "VALUE": comp.get("value", ""),
            "HQ料号": comp.get("hq_code", ""),
            "BOM_OPTION": _display_bom_option(bom_option),
            USER_VISIBLE_REAL_PAGE_LABEL: component_user_visible_page(comp),
            MAIN_MODULE_PAGE_LABEL: comp.get("page_logical", "") or comp.get("page_raw", ""),
            "子模块本地页": comp.get("module_order_local_page", "") or comp.get("page_submodule_real", ""),
            "父级Symbol页码": module_row.get("父级Symbol页码", ""),
            "父级Symbol实例": module_row.get("父级Symbol实例", ""),
            "module_order状态": comp.get("module_order_state", ""),
        })

    module_rows: List[Dict[str, object]] = []
    for module_id, row in sorted(module_rows_by_id.items(), key=lambda item: (item[0] != "__MAIN__", _natural_sort_key(item[1].get("模块名", "")))):
        counts = counters.get(module_id, {})
        enriched = dict(row)
        enriched.update({
            "元件数量": int(counts.get("元件数量", 0) or 0),
            "芯片数量": int(counts.get("芯片数量", 0) or 0),
            "R/C/L数量": int(counts.get("R/C/L数量", 0) or 0),
            "DEPOP数量": int(counts.get("DEPOP数量", 0) or 0),
            "示例位号": ", ".join(counts.get("示例位号", []) or []),
        })
        module_rows.append(enriched)

    return {
        "module_rows": module_rows,
        "component_rows": component_rows,
        "summary": {
            "module_count": len(module_rows),
            "submodule_count": sum(1 for row in module_rows if row.get("模块类型") == "子模块"),
            "component_count": len(component_rows),
        },
        "warnings": [],
    }


def filter_module_review(module_review: Dict[str, object],
                         *,
                         module_id: str = "",
                         module_name: str = "",
                         module_type: str = "") -> Dict[str, object]:
    """Return a filtered module review payload without mutating the source.

    This is intentionally a view-level filter. The underlying project analysis
    still runs once for the full project; callers can use this helper to hand
    external processes or agents a precise main/submodule scope.
    """

    raw_module_rows = [dict(row) for row in (module_review or {}).get("module_rows", []) if isinstance(row, dict)]
    raw_component_rows = [
        dict(row)
        for row in (module_review or {}).get("component_rows", [])
        if isinstance(row, dict)
    ]
    target_id = _normalize_module_key(module_id)
    target_name = str(module_name or "").strip().lower()
    target_type = str(module_type or "").strip()
    if target_type.lower() == "all":
        target_type = ""

    def module_matches(row: Dict[str, object]) -> bool:
        if target_id and _normalize_module_key(str(row.get("模块ID", ""))) != target_id:
            return False
        if target_name and target_name not in str(row.get("模块名", "")).strip().lower():
            return False
        if target_type and str(row.get("模块类型", "")).strip() != target_type:
            return False
        return True

    filtered_modules = [row for row in raw_module_rows if module_matches(row)]
    if not (target_id or target_name or target_type):
        filtered_modules = raw_module_rows
    selected_ids = {str(row.get("模块ID", "")) for row in filtered_modules}
    filtered_components = [
        row for row in raw_component_rows if str(row.get("模块ID", "")) in selected_ids
    ]
    warnings = list((module_review or {}).get("warnings", []) or [])
    if (target_id or target_name or target_type) and not filtered_modules:
        warnings.append("模块过滤未命中，请检查 module_id/module_name/module_type。")

    return {
        "module_rows": filtered_modules,
        "component_rows": filtered_components,
        "summary": {
            "module_count": len(filtered_modules),
            "submodule_count": sum(1 for row in filtered_modules if row.get("模块类型") == "子模块"),
            "component_count": len(filtered_components),
            "filtered": bool(target_id or target_name or target_type),
            "filter": {
                "module_id": module_id,
                "module_name": module_name,
                "module_type": module_type,
            },
        },
        "warnings": warnings,
    }
