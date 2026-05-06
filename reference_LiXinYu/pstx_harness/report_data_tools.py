# -*- coding: utf-8 -*-
"""Report table, project file, batch, and memory tools for report agents."""

from __future__ import annotations

from collections import Counter
import fnmatch
from pathlib import Path
import re
from typing import List, Mapping, Optional, Sequence

from pstx_agent_runtime import get_project_evidence_memory_card, search_project_evidence_memory
from pstx_core.page_resolution import summarize_module_order_page_extent
from pstx_harness.report_evidence_tools import (
    _bom_depop_tool,
    _csa_tool,
    _derating_tool,
    _drc_tool,
    _feishu_tool,
    _find_table,
    _max_rows,
    _page_mapping_tool,
    _resistor_tool,
)
from pstx_harness.report_knowledge_tools import (
    _datasheet_match_preview,
    _identity_card_preview,
)
from pstx_harness.report_tool_utils import (
    BATCH_DEFAULT_PER_ITEM_LIMIT,
    BATCH_MAX_ITEMS,
    _as_int,
    _batch_input_items,
    _batch_limit,
    _batch_summary,
    _compact_mapping,
    _compact_rows,
    _natural_value_key,
    _safe_text,
    _sanitize_feishu_row,
)
from pstx_harness.tool_core import HarnessToolContext, HarnessToolError
from pstx_knowledge.component_identity import build_component_identity_cards
from pstx_knowledge.datasheets import match_component_datasheets
from pstx_knowledge.feishu_cache import get_feishu_cache_rows


SENSITIVE_TEXT_RE = re.compile(
    r"(?i)(secret|token|apikey|api_key|appsecret|app_secret|authorization|ciphertext|password|passwd)"
    r"(\s*[:=]\s*)"
    r"([^\s,;'\"]+)"
)
SOURCE_TRACE_SCHEMA_VERSION = "pstx-source-trace.v1"
PROJECT_TEXT_SEARCH_SCHEMA_VERSION = "pstx-project-text-search.v1"
SOURCE_TRACE_TEXT_KEYS = (
    "位号", "芯片位号", "元件位号", "refdes", "component", "网络", "net", "信号", "signal",
    "HQ", "HQ_CODE", "HQ料号", "料号", "key_value",
)
SOURCE_TRACE_PAGE_KEYS = (
    "页码", "真实页", "用户页码", "主模块页", "CSA页名", "PAGE_NUMBER", "page", "page_number",
)
SOURCE_TRACE_PAGE_RE = re.compile(r"(?i)\bpage\s*([0-9]+)\b|^([0-9]+)$")


def _redact_sensitive_project_text(text: str) -> str:
    return SENSITIVE_TEXT_RE.sub(lambda match: f"{match.group(1)}{match.group(2)}<redacted>", text)


def _decode_project_text(raw: bytes) -> tuple[str, str]:
    encoding = "utf-8"
    for candidate in ["utf-8", "gb18030", "latin-1"]:
        try:
            return raw.decode(candidate), candidate
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace"), encoding


def _source_query_terms(*values: object, limit: int = 12) -> List[str]:
    terms: List[str] = []
    for value in values:
        if value is None:
            continue
        if isinstance(value, Mapping):
            for item in value.values():
                terms.extend(_source_query_terms(item, limit=limit))
                if len(terms) >= limit:
                    return terms[:limit]
            continue
        if isinstance(value, (list, tuple, set)):
            for item in value:
                terms.extend(_source_query_terms(item, limit=limit))
                if len(terms) >= limit:
                    return terms[:limit]
            continue
        text = str(value).strip().strip("'\"")
        if not text:
            continue
        pieces = [text]
        pieces.extend(part for part in re.split(r"[^A-Za-z0-9_./:+-]+", text) if part)
        for piece in pieces:
            normalized = piece.strip().strip("'\"")
            if len(normalized) < 2:
                continue
            if normalized.lower() in {"page", "none", "null", "true", "false"}:
                continue
            if normalized not in terms:
                terms.append(normalized[:160])
            if len(terms) >= limit:
                return terms[:limit]
    return terms[:limit]


def _page_numbers_from_value(value: object) -> List[int]:
    numbers: List[int] = []
    if value is None:
        return numbers
    for token in re.split(r"[,;\s]+", str(value)):
        token = token.strip().strip("'\"")
        if not token:
            continue
        match = SOURCE_TRACE_PAGE_RE.search(token)
        if not match:
            continue
        raw = match.group(1) or match.group(2)
        try:
            number = int(raw)
        except Exception:
            continue
        if number > 0 and number not in numbers:
            numbers.append(number)
    return numbers


def _int_arg(args: Mapping[str, object], key: str, default: int) -> int:
    if key not in args or args.get(key) is None:
        return default
    try:
        return int(args.get(key))
    except Exception:
        return default


def _iter_report_tables(report: dict) -> List[dict]:
    tables = []
    for section in report.get("sections", []) or []:
        for table in section.get("tables", []) or []:
            tables.append({
                "section_id": section.get("id", ""),
                "section_title": section.get("title", ""),
                "table_id": table.get("id", ""),
                "title": table.get("title", ""),
                "count": _as_int(table.get("count", len(table.get("rows", []) or []))),
                "columns": list(table.get("columns", []) or []),
            })
    return tables


def _list_report_tables_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    tables = _iter_report_tables(context.report)
    return {
        "id": "list_report_tables",
        "title": "报告表格清单",
        "target": "summary",
        "summary": f"当前报告包含 {len(tables)} 个表格。",
        "issue_count": 0,
        "tables": tables,
        "readonly": True,
    }


def _get_table_rows_tool(context: HarnessToolContext, args: dict) -> dict:
    table_id = str(args.get("table_id") or "").strip()
    limit = _as_int(args.get("limit", _max_rows(context)), _max_rows(context))
    offset = _as_int(args.get("offset", 0), 0)
    table = _find_table(context.report, table_id)
    if not table:
        raise HarnessToolError(f"未找到报告表格：{table_id}")
    rows = list(table.get("rows", []) or [])
    selected_rows = rows[offset:offset + limit]
    has_more = offset + len(selected_rows) < len(rows)
    return {
        "id": "get_table_rows",
        "title": table.get("title") or table_id,
        "target": "summary",
        "summary": (
            f"表格 {table_id} 共 {len(rows)} 行，返回 {len(selected_rows)} 行。"
            + ("当前返回不是完整表；统计唯一页码或列值请调用 summarize_table_column_values；统计原理图总页数请调用 summarize_schematic_page_count。" if has_more else "")
        ),
        "table_id": table_id,
        "offset": offset,
        "limit": limit,
        "total_rows": len(rows),
        "has_more": has_more,
        "next_offset": offset + len(selected_rows) if has_more else None,
        "truncated": has_more,
        "aggregation_hint": (
            "如需统计某列唯一值、页码覆盖范围或 top values，请调用 summarize_table_column_values；原理图总页数请调用 summarize_schematic_page_count。"
            if has_more else ""
        ),
        "columns": list(table.get("columns", []) or []),
        "rows": _compact_rows(selected_rows, limit),
        "readonly": True,
    }


def _summarize_table_column_values_tool(context: HarnessToolContext, args: dict) -> dict:
    table_id = str(args.get("table_id") or "").strip()
    column = str(args.get("column") or "").strip()
    if not column:
        raise HarnessToolError("summarize_table_column_values 需要 column。")
    table = _find_table(context.report, table_id)
    if not table:
        raise HarnessToolError(f"未找到报告表格：{table_id}")
    rows = [row for row in list(table.get("rows", []) or []) if isinstance(row, dict)]
    include_empty = bool(args.get("include_empty", False))
    limit_values = max(1, min(_as_int(args.get("limit_values", 200), 200), 1000))
    sample_per_value = max(0, min(_as_int(args.get("sample_per_value", 0), 0), 5))
    operation = str(args.get("operation") or "top").strip().lower()
    if operation not in {"top", "count", "unique"}:
        raise HarnessToolError("summarize_table_column_values 的 operation 仅支持 top/count/unique。")
    known_columns = list(table.get("columns", []) or [])
    row_columns = {str(key) for row in rows for key in row.keys()}
    available_columns = known_columns or sorted(row_columns)
    column_aliases = {
        "页面": "页码",
        "真实页": "页码",
        "用户看到的真实页": "页码",
        "总体真实页码": "页码",
        "page": "页码",
        "page_number": "页码",
    }
    normalized_columns = {re.sub(r"\s+", "", item).lower(): item for item in available_columns}
    resolved_column = column
    normalized_request = re.sub(r"\s+", "", column).lower()
    if resolved_column not in row_columns:
        resolved_column = normalized_columns.get(normalized_request) or column_aliases.get(normalized_request, column)
    if rows and resolved_column not in row_columns:
        available = known_columns or sorted({str(key) for row in rows for key in row.keys()})
        raise HarnessToolError(
            f"表格 {table_id} 不存在列 `{column}`；可用列包括："
            + "、".join(_safe_text(item, 80) for item in available[:30])
        )
    values = []
    empty_count = 0
    samples_by_value = {}
    for row_index, row in enumerate(rows):
        value = row.get(resolved_column)
        text = "" if value is None else str(value).strip()
        if not text:
            empty_count += 1
            if not include_empty:
                continue
        values.append(text)
        if sample_per_value:
            bucket = samples_by_value.setdefault(text, [])
            if len(bucket) < sample_per_value:
                bucket.append({
                    "row_index": row_index,
                    "row_number": row_index + 1,
                    "row": _compact_mapping(row),
                })
    counter = Counter(values)
    unique_values = sorted(counter.keys(), key=_natural_value_key)
    returned_values = unique_values[:limit_values]
    ordered_by_count = sorted(counter.items(), key=lambda item: (-item[1], _natural_value_key(item[0])))
    top_values = [
        {"value": _safe_text(value, 220), "count": count}
        for value, count in ordered_by_count[:limit_values]
    ]
    if operation == "top":
        selected_pairs = ordered_by_count[:limit_values]
    else:
        selected_pairs = [(value, counter[value]) for value in returned_values]
    value_counts = [
        {"value": _safe_text(value, 220), "count": count}
        for value, count in selected_pairs
    ]
    sample_rows_by_value = [
        {
            "value": _safe_text(value, 220),
            "samples": samples_by_value.get(value, []),
        }
        for value, _count in selected_pairs
        if sample_per_value and samples_by_value.get(value)
    ]
    selected_total = len(selected_pairs)
    full_total = len(counter)
    truncated = full_total > selected_total
    non_empty_count = len(rows) - empty_count
    return {
        "id": "summarize_table_column_values",
        "title": f"{table.get('title') or table_id} / {column} 列聚合",
        "target": "summary",
        "summary": (
            f"表格 {table_id} 的 `{resolved_column}` 列共 {len(rows)} 行，"
            f"非空 {non_empty_count} 行，唯一值 {len(unique_values)} 个，"
            f"按 {operation} 返回 {selected_total} 个聚合值。"
        ),
        "table_id": table_id,
        "table_title": table.get("title") or table_id,
        "column": column,
        "resolved_column": resolved_column,
        "operation": operation,
        "total_rows": len(rows),
        "non_empty_count": non_empty_count,
        "empty_count": empty_count,
        "include_empty": include_empty,
        "unique_count": len(unique_values),
        "values": [_safe_text(value, 220) for value in returned_values],
        "top_values": top_values,
        "value_counts": value_counts,
        "sample_per_value": sample_per_value,
        "sample_rows_by_value": sample_rows_by_value,
        "limit_values": limit_values,
        "truncated": truncated,
        "completeness": "partial" if truncated else "complete",
        "detail_tool": {
            "name": "get_table_rows",
            "args": {
                "table_id": table_id,
                "offset": 0,
                "limit": min(20, max(1, len(rows))),
            },
        },
        "recommended_next_tools": [] if not truncated else ["summarize_table_column_values", "get_table_rows"],
        "scope_summary": (
            f"table_id={table_id}; column={column}; resolved_column={resolved_column}; total_rows={len(rows)}; "
            f"unique_count={len(unique_values)}; operation={operation}; truncated={truncated}"
        ),
        "readonly": True,
    }


def _summarize_schematic_page_count_tool(context: HarnessToolContext, args: dict) -> dict:
    root = _project_root(context)
    summary = summarize_module_order_page_extent(str(root))
    if not summary.get("available"):
        return {
            "id": "summarize_schematic_page_count",
            "title": "原理图总页数",
            "target": "summary",
            "summary": "无法从 module_order(.dat) 计算原理图总页数；不要用 page_rows 行数替代总页数。",
            "source": "module_order",
            "project_root": str(root),
            "available": False,
            "total_pages": 0,
            "last_page": "",
            "entry_count": summary.get("entry_count", 0),
            "files": summary.get("files", []),
            "warnings": summary.get("warnings", []),
            "reason": summary.get("reason", ""),
            "completeness": "error",
            "readonly": True,
        }
    last_entry = dict(summary.get("last_entry") or {})
    return {
        "id": "summarize_schematic_page_count",
        "title": "原理图总页数",
        "target": "summary",
        "summary": (
            f"按 module_order(.dat) 计算，原理图页码总数为 {summary.get('total_pages')} 页，"
            f"最后一页为 {summary.get('last_page')}。计算口径：最后模块起始页 "
            f"{last_entry.get('start_real_page')} + 页数 {last_entry.get('page_count')} - 1。"
        ),
        "source": "module_order",
        "project_root": str(root),
        "available": True,
        "total_pages": summary.get("total_pages", 0),
        "last_page": summary.get("last_page", ""),
        "entry_count": summary.get("entry_count", 0),
        "file_count": summary.get("file_count", 0),
        "files": summary.get("files", []),
        "warnings": summary.get("warnings", []),
        "last_entry": last_entry,
        "top_entries": summary.get("top_entries", []),
        "scope_note": "page_rows 只统计有记录/有元件的页码，不代表总原理图页数；空白尾页必须以 module_order 页范围为准。",
        "readonly": True,
    }


def _query_report_entity_tool(context: HarnessToolContext, args: dict) -> dict:
    from pstx_queries.project_query import query_project_data

    mode = str(args.get("mode") or "位号").strip()
    keyword = str(args.get("keyword") or "").strip()
    result = query_project_data(
        context.bundle.get("components", {}) or {},
        context.bundle.get("nets", {}) or {},
        mode,
        keyword,
    )
    return {
        "id": "query_report_entity",
        "title": result.get("title") or keyword,
        "target": "summary",
        "summary": result.get("summary", {}),
        "query_result": result,
        "readonly": True,
    }


def _report_row_matches(report: dict, query: str, *, limit: int) -> List[dict]:
    query_lower = str(query or "").lower()
    matches = []
    if not query_lower:
        return matches
    for section in report.get("sections", []) or []:
        for table in section.get("tables", []) or []:
            table_id = str(table.get("id") or "")
            title = str(table.get("title") or table_id)
            for index, row in enumerate(table.get("rows", []) or []):
                if not isinstance(row, dict):
                    continue
                row_text = " ".join(str(value) for value in row.values()).lower()
                if query_lower not in row_text:
                    continue
                matches.append({
                    "kind": "table_row",
                    "section_id": _safe_text(section.get("id", ""), 120),
                    "section_title": _safe_text(section.get("title", ""), 160),
                    "table_id": _safe_text(table_id, 120),
                    "table_title": _safe_text(title, 160),
                    "row_index": index,
                    "row_number": index + 1,
                    "row": _compact_mapping(row, 32),
                })
                if len(matches) >= limit:
                    return matches
    return matches


def _normalize_entity_query(item, default_mode: str) -> dict:
    if isinstance(item, dict):
        keyword = str(
            item.get("keyword")
            or item.get("query")
            or item.get("refdes")
            or item.get("net")
            or item.get("hq_no")
            or item.get("page")
            or ""
        ).strip()
        mode = str(item.get("mode") or default_mode or "auto").strip()
        if not item.get("mode"):
            if item.get("net"):
                mode = "网络"
            elif item.get("refdes"):
                mode = "位号"
        return {"keyword": keyword, "mode": mode}
    return {"keyword": str(item or "").strip(), "mode": default_mode or "auto"}


def _batch_query_report_entities_tool(context: HarnessToolContext, args: dict) -> dict:
    from pstx_queries.project_query import query_project_data

    raw_items, input_truncated = _batch_input_items(args, "queries")
    default_mode = str(args.get("mode") or "auto").strip() or "auto"
    per_query_limit = _batch_limit(args.get("limit_per_query", args.get("limit", BATCH_DEFAULT_PER_ITEM_LIMIT)))
    components = context.bundle.get("components", {}) or {}
    nets = context.bundle.get("nets", {}) or {}
    items = []
    for index, raw_item in enumerate(raw_items, start=1):
        query = _normalize_entity_query(raw_item, default_mode)
        keyword = query["keyword"]
        mode = query["mode"]
        if not keyword:
            items.append({
                "index": index,
                "query": "",
                "mode": mode,
                "status": "error",
                "summary": "查询关键词为空。",
                "matches": [],
                "missing_reason": "empty_query",
            })
            continue
        try:
            modes = ["位号", "网络"] if mode in {"", "auto", "自动"} else [mode]
            matches = []
            for one_mode in modes:
                if one_mode not in {"位号", "网络"}:
                    continue
                result = query_project_data(components, nets, one_mode, keyword)
                match_type = str(result.get("match_type") or "")
                if match_type == "missing":
                    continue
                matches.append({
                    "kind": "component_query" if one_mode == "位号" else "network_query",
                    "mode": one_mode,
                    "match_type": match_type,
                    "title": _safe_text(result.get("title", keyword), 160),
                    "summary": _compact_mapping(result.get("summary") or {}, 12),
                    "query_result": {
                        "view": result.get("view", ""),
                        "entity_type": result.get("entity_type", ""),
                        "cards": list(result.get("cards") or [])[:4],
                        "items": list(result.get("items") or [])[:per_query_limit],
                    },
                })
            row_matches = _report_row_matches(context.report, keyword, limit=per_query_limit)
            matches.extend(row_matches)
            status = "found" if matches else "missing"
            items.append({
                "index": index,
                "query": _safe_text(keyword, 220),
                "mode": mode,
                "status": status,
                "match_count": len(matches),
                "matches": matches[:per_query_limit],
                "truncated": len(matches) > per_query_limit,
                "summary": f"`{keyword}` 命中 {len(matches)} 个结果。" if matches else f"`{keyword}` 未命中元件、网络或报告表。",
                "missing_reason": "" if matches else "no_report_entity_or_row_match",
            })
        except Exception as exc:
            items.append({
                "index": index,
                "query": _safe_text(keyword, 220),
                "mode": mode,
                "status": "error",
                "summary": str(exc),
                "matches": [],
                "missing_reason": str(exc),
            })
    return {
        "id": "batch_query_report_entities",
        "title": "批量查询报告实体",
        "target": "summary",
        "summary": _batch_summary("批量报告实体查询", items, truncated=input_truncated),
        "input_count": len(raw_items),
        "input_truncated": input_truncated,
        "limit_per_query": per_query_limit,
        "items": items,
        "readonly": True,
    }


def _batch_get_table_rows_tool(context: HarnessToolContext, args: dict) -> dict:
    raw_items, input_truncated = _batch_input_items(args, "requests")
    default_limit = _batch_limit(args.get("limit_per_request", BATCH_DEFAULT_PER_ITEM_LIMIT))
    items = []
    items_truncated = False
    for index, raw_item in enumerate(raw_items, start=1):
        if not isinstance(raw_item, dict):
            items.append({
                "index": index,
                "status": "error",
                "summary": "批量表格请求必须是对象。",
                "missing_reason": "bad_request_item",
            })
            continue
        table_id = str(raw_item.get("table_id") or "").strip()
        offset = _as_int(raw_item.get("offset", 0), 0)
        limit = _batch_limit(raw_item.get("limit", default_limit), default_limit)
        try:
            result = _get_table_rows_tool(context, {"table_id": table_id, "offset": offset, "limit": limit})
            rows = list(result.get("rows") or [])
            item_truncated = bool(result.get("has_more") or result.get("truncated"))
            items_truncated = items_truncated or item_truncated
            items.append({
                "index": index,
                "table_id": table_id,
                "status": "found" if rows else "missing",
                "summary": result.get("summary", ""),
                "offset": result.get("offset", offset),
                "limit": result.get("limit", limit),
                "total_rows": result.get("total_rows", 0),
                "has_more": bool(result.get("has_more", False)),
                "next_offset": result.get("next_offset"),
                "truncated": item_truncated,
                "aggregation_hint": result.get("aggregation_hint", ""),
                "columns": list(result.get("columns") or [])[:40],
                "rows": rows,
                "missing_reason": "" if rows else "table_empty_or_offset_out_of_range",
            })
        except Exception as exc:
            items.append({
                "index": index,
                "table_id": table_id,
                "status": "error",
                "summary": str(exc),
                "rows": [],
                "missing_reason": str(exc),
            })
    return {
        "id": "batch_get_table_rows",
        "title": "批量读取报告表格行",
        "target": "summary",
        "summary": _batch_summary("批量读取报告表格", items, truncated=input_truncated),
        "input_count": len(raw_items),
        "input_truncated": input_truncated,
        "items_truncated": items_truncated,
        "truncated": bool(input_truncated or items_truncated),
        "has_more": items_truncated,
        "limit_per_request": default_limit,
        "items": items,
        "readonly": True,
    }


def _batch_search_feishu_cache_rows_tool(context: HarnessToolContext, args: dict) -> dict:
    raw_items, input_truncated = _batch_input_items(args, "queries")
    default_limit = _batch_limit(args.get("limit_per_query", args.get("limit", BATCH_DEFAULT_PER_ITEM_LIMIT)))
    global_lib_id = str(args.get("lib_id") or "").strip()
    global_sheet_name = str(args.get("sheet_name") or "").strip()
    items = []
    for index, raw_item in enumerate(raw_items, start=1):
        if isinstance(raw_item, dict):
            query = str(raw_item.get("query") or raw_item.get("keyword") or raw_item.get("hq_no") or raw_item.get("spec") or "").strip()
            lib_id = str(raw_item.get("lib_id") or global_lib_id).strip()
            sheet_name = str(raw_item.get("sheet_name") or global_sheet_name).strip()
            limit = _batch_limit(raw_item.get("limit", default_limit), default_limit)
        else:
            query = str(raw_item or "").strip()
            lib_id = global_lib_id
            sheet_name = global_sheet_name
            limit = default_limit
        if not query:
            items.append({
                "index": index,
                "query": "",
                "status": "error",
                "summary": "飞书缓存查询关键词为空。",
                "rows": [],
                "missing_reason": "empty_query",
            })
            continue
        try:
            result = get_feishu_cache_rows(lib_id=lib_id, sheet_name=sheet_name, query=query, limit=limit, offset=0)
            rows = [_sanitize_feishu_row(row) for row in result.get("rows", []) or []]
            ok = bool(result.get("ok", False))
            total = _as_int(result.get("total"), 0)
            status = "found" if ok and total > 0 else ("missing" if ok else "error")
            items.append({
                "index": index,
                "query": _safe_text(query, 220),
                "lib_id": _safe_text(lib_id, 120),
                "sheet_name": _safe_text(sheet_name, 160),
                "status": status,
                "summary": (
                    f"`{query}` 命中 {total} 条，返回 {len(rows)} 条。"
                    if status == "found"
                    else (f"`{query}` 无命中。" if status == "missing" else str(result.get("error") or "飞书缓存搜索失败。"))
                ),
                "total_rows": total,
                "limit": _as_int(result.get("limit"), limit),
                "rows": rows,
                "missing_reason": "" if status == "found" else ("no_feishu_cache_match" if status == "missing" else str(result.get("error") or "")),
            })
        except Exception as exc:
            items.append({
                "index": index,
                "query": _safe_text(query, 220),
                "status": "error",
                "summary": str(exc),
                "rows": [],
                "missing_reason": str(exc),
            })
    return {
        "id": "batch_search_feishu_cache_rows",
        "title": "批量搜索飞书缓存物料",
        "target": "bom",
        "summary": _batch_summary("批量飞书缓存搜索", items, truncated=input_truncated),
        "input_count": len(raw_items),
        "input_truncated": input_truncated,
        "limit_per_query": default_limit,
        "items": items,
        "readonly": True,
    }


def _batch_get_component_identity_cards_tool(context: HarnessToolContext, args: dict) -> dict:
    raw_items, input_truncated = _batch_input_items(args, "refdes_list")
    cards = build_component_identity_cards(context.report, context.bundle)
    by_refdes = {str(card.get("refdes") or "").upper(): card for card in cards}
    items = []
    for index, raw_item in enumerate(raw_items, start=1):
        refdes = str(raw_item.get("refdes") if isinstance(raw_item, dict) else raw_item or "").strip()
        if not refdes:
            items.append({
                "index": index,
                "refdes": "",
                "status": "error",
                "summary": "refdes 为空。",
                "missing_reason": "empty_refdes",
            })
            continue
        card = by_refdes.get(refdes.upper())
        if not card:
            items.append({
                "index": index,
                "refdes": _safe_text(refdes, 120),
                "status": "missing",
                "summary": f"未找到元件身份卡：{refdes}",
                "missing_reason": "identity_card_not_found",
            })
            continue
        missing_fields = list(card.get("missing_fields") or [])
        status = "needs_context" if missing_fields else "found"
        preview = _identity_card_preview(card)
        items.append({
            "index": index,
            "refdes": preview.get("refdes", refdes),
            "status": status,
            "summary": _component_identity_batch_summary(preview, missing_fields),
            "card": preview,
            "missing_fields": missing_fields,
            "missing_reason": ",".join(str(item) for item in missing_fields) if missing_fields else "",
        })
    return {
        "id": "batch_get_component_identity_cards",
        "title": "批量读取元件身份卡",
        "target": "dfmea",
        "summary": _batch_summary("批量读取元件身份卡", items, truncated=input_truncated),
        "input_count": len(raw_items),
        "input_truncated": input_truncated,
        "items": items,
        "readonly": True,
    }


def _component_identity_batch_summary(card: dict, missing_fields: List[str]) -> str:
    refdes = card.get("refdes") or "UNKNOWN"
    base = f"{refdes} 分类={card.get('category') or 'unknown'}，HQ={card.get('hq_no') or '缺失'}，规格={card.get('spec') or '缺失'}。"
    if missing_fields:
        base += f" 缺失字段：{', '.join(str(item) for item in missing_fields[:8])}。"
    return base


def _batch_match_component_datasheets_tool(context: HarnessToolContext, args: dict) -> dict:
    raw_items, input_truncated = _batch_input_items(args, "refdes_list")
    default_limit = _batch_limit(args.get("limit_per_component", args.get("limit", 5)), 5)
    cards = build_component_identity_cards(context.report, context.bundle)
    by_refdes = {str(card.get("refdes") or "").upper(): card for card in cards}
    items = []
    for index, raw_item in enumerate(raw_items, start=1):
        refdes = str(raw_item.get("refdes") if isinstance(raw_item, dict) else raw_item or "").strip()
        limit = _batch_limit(raw_item.get("limit", default_limit), default_limit) if isinstance(raw_item, dict) else default_limit
        if not refdes:
            items.append({
                "index": index,
                "refdes": "",
                "status": "error",
                "summary": "refdes 为空。",
                "missing_reason": "empty_refdes",
                "matches": [],
            })
            continue
        card = by_refdes.get(refdes.upper())
        if not card:
            items.append({
                "index": index,
                "refdes": _safe_text(refdes, 120),
                "status": "missing",
                "summary": f"未找到元件身份卡：{refdes}",
                "missing_reason": "identity_card_not_found",
                "matches": [],
            })
            continue
        try:
            result = match_component_datasheets(card, limit=limit)
            matches = [_datasheet_match_preview(match) for match in result.get("matches", []) or []]
            items.append({
                "index": index,
                "refdes": _safe_text(refdes, 120),
                "status": "found" if matches else "missing",
                "summary": f"{refdes} 命中 {len(matches)} 个规格书候选。" if matches else f"{refdes} 暂未命中规格书候选。",
                "card": _identity_card_preview(card),
                "query": _safe_text(result.get("query", ""), 320),
                "matches": matches,
                "missing_reason": "" if matches else _safe_text(result.get("missing_reason") or "no_datasheet_match", 260),
            })
        except Exception as exc:
            items.append({
                "index": index,
                "refdes": _safe_text(refdes, 120),
                "status": "error",
                "summary": str(exc),
                "matches": [],
                "missing_reason": str(exc),
            })
    return {
        "id": "batch_match_component_datasheets",
        "title": "批量匹配元件规格书",
        "target": "dfmea",
        "summary": _batch_summary("批量匹配元件规格书", items, truncated=input_truncated),
        "input_count": len(raw_items),
        "input_truncated": input_truncated,
        "limit_per_component": default_limit,
        "items": items,
        "readonly": True,
    }


def _get_evidence_pack_tool(context: HarnessToolContext, args: dict) -> dict:
    pack_id = str(args.get("pack_id") or "").strip()
    handlers = {
        "bom_depop": _bom_depop_tool,
        "page_mapping": _page_mapping_tool,
        "drc": _drc_tool,
        "resistor": _resistor_tool,
        "derating": _derating_tool,
        "csa": _csa_tool,
        "feishu_bom": _feishu_tool,
    }
    handler = handlers.get(pack_id)
    if handler is None:
        raise HarnessToolError(f"未知证据包：{pack_id}")
    return handler(context, {})


ALLOWED_PROJECT_DIRS = {"packaged", "sch_1"}
ALLOWED_ROOT_FILES = {"module_order", "module_order.dat", "page.map"}
ALLOWED_TEXT_SUFFIXES = {".dat", ".csv", ".csa", ".map", ".txt"}


def _project_root(context: HarnessToolContext) -> Path:
    raw = str(context.bundle.get("project_root") or "").strip()
    if not raw:
        raise HarnessToolError("当前报告缺少 project_root，无法读取项目文件。")
    root = Path(raw).expanduser().resolve()
    if not root.is_dir():
        raise HarnessToolError(f"project_root 不存在或不是目录：{root}")
    return root


def _resolve_project_file(context: HarnessToolContext, raw_path: str) -> Path:
    root = _project_root(context)
    if not str(raw_path or "").strip():
        raise HarnessToolError("缺少文件路径。")
    candidate = Path(str(raw_path).strip().strip('"')).expanduser()
    if not candidate.is_absolute():
        candidate = root / candidate
    resolved = candidate.resolve()
    try:
        rel = resolved.relative_to(root)
    except ValueError as exc:
        raise HarnessToolError("禁止读取项目根目录之外的文件。") from exc
    if not _is_allowed_project_file(rel):
        raise HarnessToolError(f"文件不在 harness 允许读取范围内：{rel.as_posix()}")
    if not resolved.is_file():
        raise HarnessToolError(f"文件不存在：{rel.as_posix()}")
    return resolved


def _is_allowed_project_file(rel: Path) -> bool:
    parts = rel.parts
    if not parts:
        return False
    name = rel.name
    suffix = rel.suffix.lower()
    if len(parts) == 1 and name in ALLOWED_ROOT_FILES:
        return True
    if name == "page.map" and (len(parts) == 1 or parts[0] == "sch_1"):
        return True
    if parts[0] in ALLOWED_PROJECT_DIRS and suffix in ALLOWED_TEXT_SUFFIXES:
        return True
    return False


def _iter_allowed_project_files(root: Path, *, limit: int = 500) -> List[Path]:
    candidates = []
    for dirname in sorted(ALLOWED_PROJECT_DIRS):
        folder = root / dirname
        if folder.is_dir():
            candidates.extend(path for path in folder.rglob("*") if path.is_file())
    for name in sorted(ALLOWED_ROOT_FILES):
        path = root / name
        if path.is_file():
            candidates.append(path)
    files = []
    for path in sorted(set(candidates), key=lambda item: item.as_posix()):
        try:
            rel = path.resolve().relative_to(root)
        except ValueError:
            continue
        if _is_allowed_project_file(rel):
            files.append(path)
        if len(files) >= limit:
            break
    return files


def _normalize_project_path_prefix(value: object) -> str:
    raw = str(value or "").strip().strip('"').strip("/")
    if not raw:
        return ""
    path = Path(raw)
    if path.is_absolute() or ".." in path.parts:
        raise HarnessToolError("path_prefix 必须是项目根内的相对路径。")
    return path.as_posix().rstrip("/")


def _normalize_suffixes(value: object) -> List[str]:
    if value is None or value == "":
        return []
    raw_items = value if isinstance(value, list) else [value]
    suffixes: List[str] = []
    for item in raw_items[:12]:
        suffix = str(item or "").strip().lower()
        if not suffix:
            continue
        if not suffix.startswith("."):
            suffix = "." + suffix
        if suffix not in ALLOWED_TEXT_SUFFIXES:
            raise HarnessToolError(f"suffixes 只能使用允许的文本后缀：{', '.join(sorted(ALLOWED_TEXT_SUFFIXES))}")
        if suffix not in suffixes:
            suffixes.append(suffix)
    return suffixes


def _compile_project_text_matcher(query: str,
                                  *,
                                  mode: str,
                                  term_mode: str,
                                  case_sensitive: bool) -> tuple[List[str], object]:
    query = str(query or "").strip()
    if not query:
        raise HarnessToolError("search_project_text 需要 query。")
    normalized_mode = str(mode or "literal").strip().lower()
    if normalized_mode not in {"literal", "regex"}:
        raise HarnessToolError("search_project_text.mode 仅支持 literal/regex。")
    normalized_term_mode = str(term_mode or "any").strip().lower()
    if normalized_term_mode not in {"any", "all", "phrase"}:
        raise HarnessToolError("search_project_text.term_mode 仅支持 any/all/phrase。")
    if normalized_mode == "regex":
        flags = 0 if case_sensitive else re.IGNORECASE
        try:
            pattern = re.compile(query, flags)
        except re.error as exc:
            raise HarnessToolError(f"正则表达式无效：{exc}") from exc
        return [query], pattern
    if normalized_term_mode == "phrase":
        terms = [query]
    else:
        terms = _source_query_terms(query, limit=12)
    if not terms:
        raise HarnessToolError("search_project_text 未能从 query 提取有效搜索词。")
    return terms, None


def _line_matches_project_search(line: str,
                                 terms: Sequence[str],
                                 *,
                                 pattern: object = None,
                                 term_mode: str = "any",
                                 case_sensitive: bool = False) -> tuple[bool, List[str]]:
    if pattern is not None:
        match = pattern.search(line)  # type: ignore[attr-defined]
        return (bool(match), [match.group(0)] if match else [])
    haystack = line if case_sensitive else line.lower()
    matched: List[str] = []
    for term in terms:
        needle = str(term)
        if not case_sensitive:
            needle = needle.lower()
        if needle and needle in haystack:
            matched.append(str(term))
    if term_mode == "all":
        ok = len(matched) == len([term for term in terms if str(term or "").strip()])
    else:
        ok = bool(matched)
    return ok, list(dict.fromkeys(matched))


def _search_project_text_tool(context: HarnessToolContext, args: dict) -> dict:
    root = _project_root(context)
    query = str(args.get("query") or "").strip()
    mode = str(args.get("mode") or "literal").strip().lower()
    term_mode = str(args.get("term_mode") or "any").strip().lower()
    case_sensitive = bool(args.get("case_sensitive", False))
    terms, pattern = _compile_project_text_matcher(
        query,
        mode=mode,
        term_mode=term_mode,
        case_sensitive=case_sensitive,
    )
    context_lines = max(0, min(_int_arg(args, "context_lines", 2), 20))
    limit = max(1, min(_as_int(args.get("limit", 20), 20), 100))
    offset = max(0, _as_int(args.get("offset", 0), 0))
    max_files = max(1, min(_as_int(args.get("max_files", 500), 500), 2000))
    max_file_bytes = max(1024, min(_as_int(args.get("max_file_bytes", 8 * 1024 * 1024), 8 * 1024 * 1024), 50 * 1024 * 1024))
    path_prefix = _normalize_project_path_prefix(args.get("path_prefix"))
    suffixes = _normalize_suffixes(args.get("suffixes"))
    file_glob = str(args.get("file_glob") or "").strip()
    if file_glob and (".." in Path(file_glob).parts or Path(file_glob).is_absolute()):
        raise HarnessToolError("file_glob 必须是相对 glob，不能包含上级目录。")

    candidates: List[Path] = []
    warnings: List[str] = []
    allowed_files = _iter_allowed_project_files(root, limit=2000)
    if len(allowed_files) >= 2000:
        warnings.append("允许文件枚举已限制为前 2000 个项目文本文件。")
    candidate_limit_hit = False
    for path in allowed_files:
        rel = path.resolve().relative_to(root).as_posix()
        if path_prefix and not (rel == path_prefix or rel.startswith(path_prefix + "/")):
            continue
        if suffixes and path.suffix.lower() not in suffixes:
            continue
        if file_glob and not (fnmatch.fnmatch(rel, file_glob) or fnmatch.fnmatch(path.name, file_glob)):
            continue
        candidates.append(path)
        if len(candidates) >= max_files:
            candidate_limit_hit = True
            break
    if candidate_limit_hit:
        warnings.append(f"候选文件已限制为前 {max_files} 个允许文件。")

    all_hits: List[dict] = []
    skipped_files: List[dict] = []
    total_match_lines = 0
    for path in candidates:
        rel = path.resolve().relative_to(root).as_posix()
        try:
            size = path.stat().st_size
        except OSError as exc:
            skipped_files.append({"path": rel, "reason": f"stat failed: {exc}"})
            continue
        if size > max_file_bytes:
            skipped_files.append({"path": rel, "reason": f"file too large: {size} bytes"})
            continue
        try:
            text, encoding = _decode_project_text(path.read_bytes())
        except OSError as exc:
            skipped_files.append({"path": rel, "reason": f"read failed: {exc}"})
            continue
        lines = text.splitlines()
        seen_windows = set()
        for index, line in enumerate(lines, start=1):
            ok, matched_terms = _line_matches_project_search(
                line,
                terms,
                pattern=pattern,
                term_mode=term_mode,
                case_sensitive=case_sensitive,
            )
            if not ok:
                continue
            total_match_lines += 1
            start = max(1, index - context_lines)
            end = min(len(lines), index + context_lines)
            key = (rel, start, end)
            if key in seen_windows:
                continue
            seen_windows.add(key)
            all_hits.append(_source_hit_from_window(
                path,
                root,
                encoding,
                lines,
                start=start,
                end=end,
                match_lines=[index],
                matched_terms=matched_terms,
            ))

    selected_hits = all_hits[offset:offset + limit]
    truncated = offset + limit < len(all_hits)
    return {
        "id": "search_project_text",
        "title": "项目原始文本 grep",
        "target": "source",
        "source_schema_version": PROJECT_TEXT_SEARCH_SCHEMA_VERSION,
        "summary": (
            f"在 {len(candidates)} 个允许文件中搜索 `{query}`，命中 {len(all_hits)} 个窗口，返回 {len(selected_hits)} 个。"
            if selected_hits else
            f"在 {len(candidates)} 个允许文件中搜索 `{query}`，未找到匹配窗口。"
        ),
        "query": query,
        "query_terms": terms,
        "mode": mode,
        "term_mode": term_mode,
        "case_sensitive": case_sensitive,
        "path_prefix": path_prefix,
        "suffixes": suffixes,
        "file_glob": file_glob,
        "source_files_considered": len(candidates),
        "candidate_files": [
            path.resolve().relative_to(root).as_posix()
            for path in candidates[:80]
        ],
        "total_match_lines": total_match_lines,
        "total_hits": len(all_hits),
        "hit_count": len(selected_hits),
        "offset": offset,
        "limit": limit,
        "truncated": truncated,
        "source_hits": selected_hits,
        "skipped_files": skipped_files[:20],
        "warnings": warnings,
        "detail_tool": selected_hits[0].get("detail_tool") if selected_hits else None,
        "recommended_next_tools": ["read_project_text"] if selected_hits else ["list_project_files", "trace_project_source"],
        "completeness": "complete" if selected_hits and not truncated else ("truncated" if truncated else "missing"),
        "readonly": True,
    }


def _list_project_files_tool(context: HarnessToolContext, args: dict) -> dict:
    root = _project_root(context)
    limit = _as_int(args.get("limit", 200), 200)
    files = []
    for path in _iter_allowed_project_files(root, limit=limit):
        rel = path.resolve().relative_to(root)
        try:
            size = path.stat().st_size
        except OSError:
            size = 0
        files.append({"path": rel.as_posix(), "name": path.name, "size": size})
    return {
        "id": "list_project_files",
        "title": "项目只读文件清单",
        "target": "summary",
        "summary": f"返回 {len(files)} 个允许读取的项目文本文件。",
        "project_root": str(root),
        "files": files,
        "readonly": True,
    }


def _line_window(lines: Sequence[str], *, start_line: int, line_count: int) -> tuple[int, int, List[dict]]:
    total = len(lines)
    start = max(1, min(int(start_line or 1), max(1, total)))
    count = max(1, int(line_count or 80))
    end = min(total, start + count - 1)
    excerpt = [
        {"line": number, "text": _redact_sensitive_project_text(lines[number - 1])}
        for number in range(start, end + 1)
    ]
    return start, end, excerpt


def _match_line_windows(lines: Sequence[str],
                        terms: Sequence[str],
                        *,
                        context_lines: int,
                        limit: int) -> List[tuple[int, int, List[int], List[str]]]:
    normalized_terms = [str(term).lower() for term in terms if str(term or "").strip()]
    if not normalized_terms:
        return []
    windows: List[tuple[int, int, List[int], List[str]]] = []
    seen_ranges = set()
    for index, line in enumerate(lines, start=1):
        lower = line.lower()
        matched = [terms[pos] for pos, term in enumerate(normalized_terms) if term in lower]
        if not matched:
            continue
        start = max(1, index - context_lines)
        end = min(len(lines), index + context_lines)
        key = (start, end)
        if key in seen_ranges:
            continue
        seen_ranges.add(key)
        windows.append((start, end, [index], list(dict.fromkeys(str(item) for item in matched))))
        if len(windows) >= limit:
            break
    return windows


def _source_hit_from_window(path: Path,
                            root: Path,
                            encoding: str,
                            lines: Sequence[str],
                            *,
                            start: int,
                            end: int,
                            match_lines: Sequence[int] = (),
                            matched_terms: Sequence[str] = ()) -> dict:
    rel = path.resolve().relative_to(root).as_posix()
    line_numbers = set(int(item) for item in match_lines or [])
    excerpt = []
    for number in range(start, end + 1):
        excerpt.append({
            "line": number,
            "text": _redact_sensitive_project_text(lines[number - 1]),
            "matched": number in line_numbers,
        })
    return {
        "path": rel,
        "encoding": encoding,
        "line_start": start,
        "line_end": end,
        "match_lines": list(match_lines or []),
        "matched_terms": list(matched_terms or []),
        "excerpt": excerpt,
        "detail_tool": {
            "name": "read_project_text",
            "args": {
                "path": rel,
                "line_start": start,
                "line_count": max(1, end - start + 1),
            },
        },
    }


def _read_project_text_tool(context: HarnessToolContext, args: dict) -> dict:
    path = _resolve_project_file(context, str(args.get("path") or ""))
    root = _project_root(context)
    max_chars = _as_int(args.get("max_chars", 12000), 12000)
    text, encoding = _decode_project_text(path.read_bytes())
    redacted_text = _redact_sensitive_project_text(text)
    rel = path.resolve().relative_to(root).as_posix()
    lines = text.splitlines()
    line_start = _as_int(args.get("line_start", 0), 0)
    line_count = _as_int(args.get("line_count", 120), 120)
    query = str(args.get("query") or "").strip()
    context_lines = max(0, min(_int_arg(args, "context_lines", 4), 20))
    selected_line_start = 1
    selected_line_end = len(lines)
    excerpts: List[dict] = []
    if query:
        terms = _source_query_terms(query, limit=8)
        windows = _match_line_windows(lines, terms, context_lines=context_lines, limit=1)
        if windows:
            start, end, match_lines, matched_terms = windows[0]
            selected_line_start = start
            selected_line_end = end
            excerpts = [{
                "line_start": start,
                "line_end": end,
                "match_lines": match_lines,
                "matched_terms": matched_terms,
                "lines": _line_window(lines, start_line=start, line_count=end - start + 1)[2],
            }]
            redacted_text = "\n".join(item["text"] for item in excerpts[0]["lines"])
        else:
            redacted_text = ""
            selected_line_end = 0
            excerpts = []
    elif line_start > 0:
        selected_line_start, selected_line_end, excerpt_lines = _line_window(lines, start_line=line_start, line_count=line_count)
        redacted_text = "\n".join(item["text"] for item in excerpt_lines)
        excerpts = [{
            "line_start": selected_line_start,
            "line_end": selected_line_end,
            "match_lines": [],
            "matched_terms": [],
            "lines": excerpt_lines,
        }]
    truncated = len(redacted_text) > max_chars
    content = redacted_text[:max_chars]
    return {
        "id": "read_project_text",
        "title": rel,
        "target": "summary",
        "summary": (
            f"读取 {rel}，编码 {encoding}，返回 {len(content)} 字符。"
            + (f" 行范围 {selected_line_start}-{selected_line_end}。" if query or line_start > 0 else "")
        ),
        "path": rel,
        "encoding": encoding,
        "line_start": selected_line_start if (query or line_start > 0) else None,
        "line_end": selected_line_end if (query or line_start > 0) else None,
        "total_lines": len(lines),
        "chars": len(redacted_text),
        "truncated": truncated,
        "query": query,
        "excerpts": excerpts,
        "content": content,
        "readonly": True,
    }


def _report_row_for_source_trace(context: HarnessToolContext, args: Mapping[str, object]) -> tuple[dict, dict]:
    table_id = str(args.get("table_id") or "").strip()
    if not table_id:
        return {}, {}
    table = _find_table(context.report, table_id)
    if not table:
        raise HarnessToolError(f"未找到报告表格：{table_id}")
    rows = [row for row in list(table.get("rows") or []) if isinstance(row, dict)]
    if args.get("row_index") is not None:
        index = _as_int(args.get("row_index"), 0)
    elif args.get("row_number") is not None:
        index = max(0, _as_int(args.get("row_number"), 1) - 1)
    else:
        raise HarnessToolError("按报告行追溯时需要 row_index 或 row_number。")
    if index < 0 or index >= len(rows):
        raise HarnessToolError(f"表格 {table_id} 不存在行 index={index}。")
    return rows[index], {
        "table_id": table_id,
        "table_title": table.get("title") or table_id,
        "row_index": index,
        "row_number": index + 1,
    }


def _source_trace_page_numbers(row: Mapping[str, object], query: str, kind: str) -> List[int]:
    numbers: List[int] = []
    for key in SOURCE_TRACE_PAGE_KEYS:
        if key in row:
            for number in _page_numbers_from_value(row.get(key)):
                if number not in numbers:
                    numbers.append(number)
    if kind == "page" or re.search(r"(?i)\bpage\s*\d+\b|^\s*\d+\s*$", query or ""):
        for number in _page_numbers_from_value(query):
            if number not in numbers:
                numbers.append(number)
    return numbers[:12]


def _source_trace_terms(row: Mapping[str, object], query: str, kind: str) -> List[str]:
    values: List[object] = [query]
    for key in SOURCE_TRACE_TEXT_KEYS:
        if key in row:
            values.append(row.get(key))
    if kind == "page":
        values = [query]
    return _source_query_terms(*values, limit=12)


def _existing_allowed_candidate(root: Path, rel_path: str) -> Optional[Path]:
    path = (root / rel_path).resolve()
    try:
        rel = path.relative_to(root)
    except ValueError:
        return None
    if not _is_allowed_project_file(rel) or not path.is_file():
        return None
    return path


def _source_trace_candidates(root: Path,
                             *,
                             explicit_path: str = "",
                             kind: str = "auto",
                             query_terms: Sequence[str] = (),
                             page_numbers: Sequence[int] = (),
                             max_files: int = 120) -> tuple[List[Path], List[str]]:
    warnings: List[str] = []
    candidates: List[Path] = []

    def add(rel_path: str) -> None:
        path = _existing_allowed_candidate(root, rel_path)
        if path and path not in candidates:
            candidates.append(path)

    if explicit_path:
        add(explicit_path)
        return candidates, warnings

    if kind in {"auto", "refdes", "text"} and query_terms:
        add("packaged/pstxprt.dat")
    if kind in {"auto", "net", "refdes", "text"} and query_terms:
        add("packaged/pstxnet.dat")
    for page_number in page_numbers:
        add(f"sch_1/page{page_number}.csv")
        add(f"sch_1/page{page_number}.csa")
    if page_numbers or kind == "page":
        add("sch_1/page.map")
        add("page.map")
        add("module_order")
        add("module_order.dat")
    if not candidates and query_terms:
        for path in _iter_allowed_project_files(root, limit=max_files):
            if path not in candidates:
                candidates.append(path)
        if len(candidates) >= max_files:
            warnings.append(f"候选文件已限制为前 {max_files} 个允许文件。")
    return candidates[:max_files], warnings


def _trace_project_source_tool(context: HarnessToolContext, args: dict) -> dict:
    root = _project_root(context)
    kind = str(args.get("kind") or "auto").strip().lower() or "auto"
    if kind not in {"auto", "refdes", "net", "page", "text"}:
        raise HarnessToolError("trace_project_source.kind 仅支持 auto/refdes/net/page/text。")
    limit = max(1, min(_as_int(args.get("limit", 8), 8), 30))
    line_count = max(1, min(_as_int(args.get("line_count", 11), 11), 160))
    context_lines = max(0, min(_int_arg(args, "context_lines", 3), 20))
    explicit_path = str(args.get("path") or "").strip()
    query = str(args.get("query") or "").strip()
    row, row_locator = _report_row_for_source_trace(context, args)
    page_numbers = _source_trace_page_numbers(row, query, kind)
    terms = _source_trace_terms(row, query, kind)
    warnings: List[str] = []
    candidates, candidate_warnings = _source_trace_candidates(
        root,
        explicit_path=explicit_path,
        kind=kind,
        query_terms=terms,
        page_numbers=page_numbers,
    )
    warnings.extend(candidate_warnings)
    hits: List[dict] = []
    if explicit_path and not candidates:
        _resolve_project_file(context, explicit_path)
    if explicit_path and args.get("line_start") is not None:
        path = _resolve_project_file(context, explicit_path)
        text, encoding = _decode_project_text(path.read_bytes())
        lines = text.splitlines()
        start, end, excerpt_lines = _line_window(lines, start_line=_as_int(args.get("line_start"), 1), line_count=line_count)
        hit = _source_hit_from_window(path, root, encoding, lines, start=start, end=end)
        hit["excerpt"] = [{"line": item["line"], "text": item["text"], "matched": False} for item in excerpt_lines]
        hits.append(hit)
    else:
        for path in candidates:
            try:
                text, encoding = _decode_project_text(path.read_bytes())
            except OSError as exc:
                warnings.append(f"读取失败 {path.name}: {exc}")
                continue
            lines = text.splitlines()
            if terms:
                windows = _match_line_windows(lines, terms, context_lines=context_lines, limit=max(1, limit - len(hits)))
            else:
                start = max(1, _as_int(args.get("line_start", 1), 1))
                end = min(len(lines), start + line_count - 1)
                windows = [(start, end, [], [])]
            for start, end, match_lines, matched_terms in windows:
                if end - start + 1 > line_count:
                    end = min(end, start + line_count - 1)
                hits.append(_source_hit_from_window(
                    path,
                    root,
                    encoding,
                    lines,
                    start=start,
                    end=end,
                    match_lines=match_lines,
                    matched_terms=matched_terms,
                ))
                if len(hits) >= limit:
                    break
            if len(hits) >= limit:
                break
    truncated = len(hits) >= limit and any(path not in {Path(root / hit.get("path", "")).resolve() for hit in hits} for path in candidates)
    candidate_files = []
    for path in candidates[:80]:
        try:
            rel = path.resolve().relative_to(root).as_posix()
        except ValueError:
            continue
        candidate_files.append(rel)
    return {
        "id": "trace_project_source",
        "title": "分析到原始文件追溯",
        "target": "source",
        "source_schema_version": SOURCE_TRACE_SCHEMA_VERSION,
        "summary": (
            f"从 {len(candidate_files)} 个候选原始文件中返回 {len(hits)} 个片段。"
            if hits else f"从 {len(candidate_files)} 个候选原始文件中未找到匹配片段。"
        ),
        "kind": kind,
        "query": query,
        "query_terms": terms,
        "page_numbers": page_numbers,
        "derived_from": row_locator,
        "row": _compact_mapping(row, 24) if row else {},
        "candidate_files": candidate_files,
        "source_files_considered": len(candidates),
        "source_hits": hits,
        "hit_count": len(hits),
        "limit": limit,
        "truncated": truncated,
        "warnings": warnings,
        "detail_tool": hits[0].get("detail_tool") if hits else None,
        "recommended_next_tools": ["read_project_text"] if hits else ["list_project_files", "read_project_text"],
        "completeness": "complete" if hits and not truncated else ("truncated" if truncated else "missing"),
        "readonly": True,
    }


def _project_memory_cards(context: HarnessToolContext) -> List[dict]:
    if not isinstance(context.project_context, dict):
        return []
    return [dict(item) for item in context.project_context.get("evidence_memory_cards") or [] if isinstance(item, dict)]


def _list_project_memory_evidence_tool(context: HarnessToolContext, args: dict) -> dict:
    result = search_project_evidence_memory(
        _project_memory_cards(context),
        query=args.get("query", ""),
        evidence_type=args.get("evidence_type", ""),
        limit=int(args.get("limit") or 20),
        offset=int(args.get("offset") or 0),
    )
    cards = result.get("cards") or []
    return {
        "id": "list_project_memory_evidence",
        "title": "项目证据记忆",
        "target": "memory",
        "summary": result.get("summary", ""),
        "query": result.get("query", ""),
        "evidence_type": result.get("evidence_type", ""),
        "total_matches": result.get("total_matches", 0),
        "offset": result.get("offset", 0),
        "limit": result.get("limit", 20),
        "has_more": result.get("has_more", False),
        "cards": cards,
        "readonly": True,
    }


def _get_project_memory_evidence_tool(context: HarnessToolContext, args: dict) -> dict:
    evidence_id = str(args.get("evidence_id") or "").strip()
    if not evidence_id:
        raise HarnessToolError("evidence_id 不能为空。")
    result = get_project_evidence_memory_card(_project_memory_cards(context), evidence_id)
    return {
        "id": "get_project_memory_evidence",
        "title": "项目证据记忆详情",
        "target": "memory",
        "summary": result.get("summary", ""),
        "found": bool(result.get("found")),
        "evidence_id": evidence_id,
        "card": result.get("card") or {},
        "readonly": True,
    }


def _batch_get_project_memory_evidence_tool(context: HarnessToolContext, args: dict) -> dict:
    raw_items, input_truncated = _batch_input_items(args, "evidence_ids")
    cards = _project_memory_cards(context)
    items = []
    for index, raw_item in enumerate(raw_items, start=1):
        evidence_id = str(raw_item.get("evidence_id") if isinstance(raw_item, dict) else raw_item or "").strip()
        if not evidence_id:
            items.append({"index": index, "evidence_id": "", "status": "error", "summary": "evidence_id 为空。"})
            continue
        result = get_project_evidence_memory_card(cards, evidence_id)
        items.append({
            "index": index,
            "evidence_id": evidence_id,
            "status": "found" if result.get("found") else "missing",
            "summary": result.get("summary", ""),
            "card": result.get("card") or {},
        })
    return {
        "id": "batch_get_project_memory_evidence",
        "title": "批量读取项目证据记忆",
        "target": "memory",
        "summary": _batch_summary("批量读取项目证据记忆", items, truncated=input_truncated),
        "input_count": len(raw_items),
        "input_truncated": input_truncated,
        "items": items,
        "readonly": True,
    }
