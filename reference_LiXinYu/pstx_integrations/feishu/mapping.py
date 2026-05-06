# -*- coding: utf-8 -*-
"""Feishu sheet header, row alignment, and field mapping helpers."""

from __future__ import annotations

import re
from typing import List, Optional, Sequence

from pstx_integrations.feishu.common import (
    FeishuBomError,
    _default_column_range,
    _normalized_name,
    _pad_row,
    _safe_cell_str,
)


def _header_index(headers: Sequence[str], column_name: str) -> int:
    target = _safe_cell_str(column_name)
    if not target:
        return -1
    try:
        return list(headers).index(target)
    except ValueError:
        return -1


def _looks_like_hq_code(value: object) -> bool:
    return bool(re.fullmatch(r"HQ[A-Za-z0-9]{6,}", _safe_cell_str(value)))


def _alignment_score_for_row(row: Sequence[object],
                             *,
                             hq_index: int,
                             spec_index: int,
                             pi_index: int,
                             selection_order_index: int) -> int:
    def get(index: int) -> str:
        return _safe_cell_str(row[index]) if 0 <= index < len(row) else ""

    score = 0
    hq_value = get(hq_index)
    spec_value = get(spec_index)
    if hq_index >= 0:
        score += 8 if _looks_like_hq_code(hq_value) else (-3 if hq_value else 0)
    if spec_index >= 0:
        score += 4 if spec_value and not _looks_like_hq_code(spec_value) else 0
    if pi_index >= 0 and get(pi_index):
        score += 1
    if selection_order_index >= 0 and get(selection_order_index):
        score += 1
    return score


def _align_data_row_to_headers(row: Sequence[object],
                               *,
                               header_width: int,
                               hq_index: int,
                               spec_index: int,
                               pi_index: int = -1,
                               selection_order_index: int = -1) -> tuple[List[object], dict]:
    """Return a row padded to header width, repairing rows whose leading empty cells were omitted."""

    original = list(row or [])
    base = _pad_row(original, header_width)
    best_offset = 0
    best_score = _alignment_score_for_row(
        base,
        hq_index=hq_index,
        spec_index=spec_index,
        pi_index=pi_index,
        selection_order_index=selection_order_index,
    )
    # Some Feishu gateway responses may omit leading blank cells in individual
    # rows even when the requested range starts at A. Try conservative left-pad
    # repairs and keep the one that best places HQ/spec fields under headers.
    max_offset = min(5, max(0, header_width - 1))
    for offset in range(1, max_offset + 1):
        candidate = _pad_row([""] * offset + original, header_width)
        score = _alignment_score_for_row(
            candidate,
            hq_index=hq_index,
            spec_index=spec_index,
            pi_index=pi_index,
            selection_order_index=selection_order_index,
        )
        if score > best_score + 3:
            best_offset = offset
            best_score = score
            base = candidate
    return base, {
        "applied": bool(best_offset),
        "left_pad": best_offset,
        "original_width": len(original),
        "aligned_width": len(base),
        "score": best_score,
    }


def _to_text_rows(values: Sequence[Sequence[object]]) -> List[List[str]]:
    output: List[List[str]] = []
    for row in values or []:
        if isinstance(row, (list, tuple)):
            output.append([_safe_cell_str(cell) for cell in row])
    return output


def _column_by_keywords(headers: Sequence[str], keywords: Sequence[str], *, avoid: Sequence[str] = ()) -> str:
    scored: List[tuple[int, int, str]] = []
    keyword_words = [_normalized_name(keyword) for keyword in keywords if _normalized_name(keyword)]
    avoid_words = [_normalized_name(word) for word in avoid if _normalized_name(word)]
    for position, header in enumerate(headers):
        normalized = _normalized_name(header)
        if not header or any(word and word in normalized for word in avoid_words):
            continue
        score = 0
        for index, keyword in enumerate(keyword_words):
            if normalized == keyword:
                score += max(1, 1000 - index)
            elif keyword in normalized:
                score += max(1, 20 - index)
        if score:
            scored.append((score, position, header))
    scored.sort(key=lambda item: (-item[0], item[1]))
    return scored[0][2] if scored else ""


def _ordered_unique(values: Sequence[object]) -> List[str]:
    output: List[str] = []
    seen = set()
    for value in values or []:
        text = _safe_cell_str(value)
        key = text.lower()
        if not text or key in seen:
            continue
        seen.add(key)
        output.append(text)
    return output


def _normalize_optional_fields(raw_optional: object, headers: Optional[Sequence[str]] = None) -> List[dict]:
    header_lookup = {_safe_cell_str(header).lower(): _safe_cell_str(header) for header in (headers or []) if _safe_cell_str(header)}
    if isinstance(raw_optional, str):
        raw_items: Sequence[object] = [part.strip() for part in re.split(r"[,，;；\n]+", raw_optional) if part.strip()]
    elif isinstance(raw_optional, list):
        raw_items = raw_optional
    else:
        raw_items = []

    output: List[dict] = []
    seen = set()
    for item in raw_items:
        if isinstance(item, dict):
            label = _safe_cell_str(item.get("label") or item.get("title") or item.get("name") or item.get("column"))
            column = _safe_cell_str(item.get("column") or item.get("title") or item.get("label") or item.get("name"))
            source = _safe_cell_str(item.get("source") or "manual")
        else:
            label = _safe_cell_str(item)
            column = label
            source = "manual"
        if headers:
            column = header_lookup.get(column.lower(), "")
        if not column:
            continue
        label = label or column
        key = column.lower()
        if key in seen:
            continue
        seen.add(key)
        output.append({
            "label": label,
            "column": column,
            "source": source,
        })
    return output


def _collect_optional_fields(headers: Sequence[str],
                             used_columns: Sequence[str],
                             *,
                             saved_optional_order: Optional[Sequence[str]] = None,
                             optional_titles: Optional[Sequence[str]] = None) -> List[dict]:
    used = {_safe_cell_str(column).lower() for column in used_columns if _safe_cell_str(column)}
    header_by_key = {header.lower(): header for header in headers if header and header.lower() not in used}
    if optional_titles:
        ordered_titles = _ordered_unique(optional_titles)
    else:
        saved = [title for title in _ordered_unique(saved_optional_order or []) if title.lower() in header_by_key]
        remaining = [header for header in headers if header.lower() in header_by_key and header not in saved]
        ordered_titles = saved + remaining

    optional_fields: List[dict] = []
    seen = set()
    for title in ordered_titles:
        column = header_by_key.get(title.lower())
        if not column or column.lower() in seen:
            continue
        seen.add(column.lower())
        source = "saved-order" if saved_optional_order and title.lower() in {v.lower() for v in saved_optional_order} else "agent-or-heuristic"
        optional_fields.append({
            "label": column,
            "column": column,
            "source": source,
        })
    return optional_fields


def build_feishu_mapping_from_headers(headers: Sequence[object],
                                      *,
                                      header_row: int = 1,
                                      sheet_title: str = "",
                                      provider: str = "local-heuristic",
                                      notes: Optional[Sequence[str]] = None,
                                      saved_optional_order: Optional[Sequence[str]] = None,
                                      optional_titles: Optional[Sequence[str]] = None) -> dict:
    cleaned_headers = [_safe_cell_str(header) for header in headers if _safe_cell_str(header)]
    spec_model_col = _column_by_keywords(
        cleaned_headers,
        ["规格型号", "part number", "part no", "p/n", "厂家型号", "制造商型号", "mpn", "manufacturer part", "器件型号", "型号"],
        avoid=["hq", "料号", "编码"],
    )
    hq_code_col = _column_by_keywords(cleaned_headers, ["hq料号", "hq编码", "hq_code", "hq code", "hq no", "物料编码", "物料号", "料号", "编码"])
    pi_col = _column_by_keywords(cleaned_headers, ["pi", "p.i.", "p.i", "owner", "负责人"])
    selection_order_col = _column_by_keywords(
        cleaned_headers,
        ["选型顺序", "选型优先级", "优选顺序", "priority", "rank", "order", "顺序"],
        avoid=["是否调整选型顺序", "是否调整"],
    )
    brand_col = _column_by_keywords(cleaned_headers, ["制造商", "manufacturer", "厂家", "品牌", "brand", "vendor"])
    desc_col = _column_by_keywords(cleaned_headers, ["描述", "description", "desc", "说明", "备注"])
    optional_fields = _collect_optional_fields(
        cleaned_headers,
        [hq_code_col, spec_model_col, pi_col, selection_order_col],
        saved_optional_order=saved_optional_order,
        optional_titles=optional_titles,
    )

    output_notes = list(notes or ["字段识别是草稿，建议同步前人工确认。"])
    if not spec_model_col:
        output_notes.append("未可靠识别规格型号 / Part Number 列。")
    if not hq_code_col:
        output_notes.append("未可靠识别 HQ 料号列。")
    if not selection_order_col:
        output_notes.append("未识别选型顺序列；芯片类表格可能允许为空。")

    return {
        "ok": True,
        "provider": provider,
        "sheet_title": _safe_cell_str(sheet_title),
        "header_row": max(1, int(header_row or 1)),
        "headers": cleaned_headers,
        "mapping": {
            "hq_code_col": hq_code_col,
            "spec_model_col": spec_model_col,
            "pi_col": pi_col,
            "selection_order_col": selection_order_col,
            "optional_fields": optional_fields,
            "key_col": spec_model_col,
            "hq_no_col": hq_code_col,
            "brand_col": brand_col,
            "spec_col": spec_model_col,
            "desc_col": desc_col,
        },
        "confidence": "medium" if spec_model_col and hq_code_col else "low",
        "notes": output_notes,
    }


def suggest_feishu_mapping_from_preview(values: Sequence[Sequence[object]],
                                        *,
                                        sheet_title: str = "",
                                        saved_optional_order: Optional[Sequence[str]] = None,
                                        optional_titles: Optional[Sequence[str]] = None,
                                        header_row: Optional[int] = None) -> dict:
    """Detect the header row locally and build a conservative draft mapping."""

    text_rows = _to_text_rows(values)
    if not text_rows:
        return {
            "ok": True,
            "provider": "local-heuristic",
            "sheet_title": _safe_cell_str(sheet_title),
            "header_row": 1,
            "headers": [],
            "mapping": {},
            "confidence": "low",
            "notes": ["没有可分析的预览行。"],
        }

    if header_row is not None:
        try:
            safe_header_row = max(1, int(header_row or 1))
        except (TypeError, ValueError):
            safe_header_row = 1
        headers = text_rows[safe_header_row - 1] if len(text_rows) >= safe_header_row else []
        headers = [header for header in headers if header]
        return build_feishu_mapping_from_headers(
            headers,
            header_row=safe_header_row,
            sheet_title=sheet_title,
            provider="local-heuristic",
            notes=["使用用户指定表头行生成字段草稿；建议同步前人工确认。"],
            saved_optional_order=saved_optional_order,
            optional_titles=optional_titles,
        )

    header_candidates: List[tuple[int, int, List[str]]] = []
    header_keywords = [
        "规格型号", "厂家型号", "制造商型号", "mpn", "part number", "pn", "p/n",
        "hq料号", "hq编码", "hq_code", "物料编码", "料号", "pi", "选型顺序",
        "制造商", "厂家", "品牌", "规格", "描述", "description",
    ]
    for index, row in enumerate(text_rows[:12], start=1):
        non_empty = [cell for cell in row if cell]
        if not non_empty:
            continue
        keyword_score = sum(
            1
            for cell in non_empty
            for keyword in header_keywords
            if keyword.lower() in cell.lower()
        )
        score = len(non_empty) + keyword_score * 5
        header_candidates.append((score, index, row))
    header_candidates.sort(key=lambda item: (-item[0], item[1]))
    _score, header_row, headers = header_candidates[0] if header_candidates else (0, 1, text_rows[0])
    headers = [header for header in headers if header]
    return build_feishu_mapping_from_headers(
        headers,
        header_row=header_row,
        sheet_title=sheet_title,
        provider="local-heuristic",
        notes=["本地启发式先识别表头，再给出标准字段和扩展字段草稿，建议同步前人工确认。"],
        saved_optional_order=saved_optional_order,
        optional_titles=optional_titles,
    )


def _normalize_sheet_cfgs(sheets: Sequence[dict]) -> List[dict]:
    normalized: List[dict] = []
    for raw in sheets or []:
        if not isinstance(raw, dict):
            continue
        sheet_id = _safe_cell_str(raw.get("sheet_id") or raw.get("sheetId"))
        if not sheet_id:
            continue
        cfg = dict(raw)
        cfg["sheet_id"] = sheet_id
        cfg["title"] = _safe_cell_str(raw.get("title") or raw.get("sheet_title") or sheet_id)
        enabled_value = raw.get("enabled", True)
        cfg["enabled"] = str(enabled_value).strip().lower() not in {"0", "false", "no", "off"}
        try:
            cfg["row_count"] = min(max(int(raw.get("row_count") or 5000), 50), 10000)
        except (TypeError, ValueError) as exc:
            raise FeishuBomError(f"{sheet_id} 的 row_count 必须是数字。") from exc
        try:
            cfg["header_row"] = max(1, int(raw.get("header_row") or 1))
        except (TypeError, ValueError) as exc:
            raise FeishuBomError(f"{sheet_id} 的 header_row 必须是数字。") from exc
        cfg["column_range"] = _safe_cell_str(
            raw.get("column_range")
            or raw.get("col_range")
            or _default_column_range(raw.get("column_count") or raw.get("columnCount"))
        ).upper()
        cfg["hq_code_col"] = _safe_cell_str(raw.get("hq_code_col") or raw.get("hq_no_col"))
        cfg["spec_model_col"] = _safe_cell_str(
            raw.get("spec_model_col")
            or raw.get("part_number_col")
            or raw.get("key_col")
            or raw.get("spec_col")
        )
        cfg["pi_col"] = _safe_cell_str(raw.get("pi_col") or raw.get("PI_col") or raw.get("pi"))
        cfg["selection_order_col"] = _safe_cell_str(
            raw.get("selection_order_col")
            or raw.get("select_order_col")
            or raw.get("priority_col")
            or raw.get("order_col")
        )
        cfg["optional_fields"] = _normalize_optional_fields(raw.get("optional_fields") or raw.get("extra_fields"))
        cfg["key_col"] = cfg["spec_model_col"]
        cfg["hq_no_col"] = cfg["hq_code_col"]
        normalized.append(cfg)
    return normalized
