# -*- coding: utf-8 -*-
"""Read-only tools for the project compare agent."""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional

from pstx_harness.compare_datasheet_tools import (
    batch_search_datasheet_chunks_tool,
    get_datasheet_chunk_tool,
    get_datasheet_page_excerpt_tool,
    get_datasheet_parameter_tool,
    get_datasheet_review_template_tool,
    list_datasheet_documents_tool,
    list_datasheet_review_templates_tool,
    search_datasheet_chunks_tool,
    search_datasheet_parameters_tool,
)
from pstx_harness.compare_project_tools import (
    batch_get_cadence_page_objects_tool,
    compare_cadence_page_semantics_tool,
    get_cadence_page_object_tool,
    get_cadence_page_raw_excerpt_tool,
    list_compare_project_files_tool,
    read_compare_project_text_tool,
    resolve_compare_page_range_tool,
)
from pstx_harness.skill_tools import (
    _get_harness_skill_tool,
    _list_harness_skills_tool,
    _select_harness_skills_tool,
)
from pstx_harness.tool_core import HarnessTool, HarnessToolError, HarnessToolRegistry


BATCH_MAX_ITEMS = 20
BATCH_DEFAULT_PER_ITEM_LIMIT = 10


@dataclass(frozen=True)
class CompareToolContext:
    compare_payload: dict
    left_payload: dict
    right_payload: dict
    request: object


def _as_int(value, default: int = 0) -> int:
    try:
        return int(value if value is not None else default)
    except (TypeError, ValueError):
        return default


def _safe_text(value, limit: int = 260) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _compact_row(row: dict, limit: int = 320) -> dict:
    if not isinstance(row, dict):
        return {}
    return {str(key): _safe_text(value, limit) for key, value in row.items()}


def _compact_rows(rows: Iterable[dict], limit: int) -> List[dict]:
    return [_compact_row(row) for row in list(rows or [])[:max(0, limit)] if isinstance(row, dict)]


def _schema(properties: dict, required: Optional[List[str]] = None, additional: bool = False) -> dict:
    return {
        "type": "object",
        "properties": properties,
        "required": required or [],
        "additionalProperties": additional,
    }


def _batch_input_items(args: dict, key: str, *, max_items: int = BATCH_MAX_ITEMS) -> tuple[List[object], bool]:
    raw = args.get(key)
    if not isinstance(raw, list):
        raise HarnessToolError(f"批量 compare 工具需要数组参数：{key}。")
    items = list(raw)
    return items[:max_items], len(items) > max_items


def _batch_limit(value, default: int = BATCH_DEFAULT_PER_ITEM_LIMIT) -> int:
    return max(1, min(_as_int(value, default), BATCH_DEFAULT_PER_ITEM_LIMIT))


def _batch_summary(title: str, items: List[dict], *, truncated: bool = False) -> str:
    found = sum(1 for item in items if item.get("status") == "found")
    missing = sum(1 for item in items if item.get("status") == "missing")
    errors = sum(1 for item in items if item.get("status") == "error")
    suffix = "，输入已按上限截断" if truncated else ""
    return f"{title}完成：{len(items)} 项，命中 {found} 项，缺失 {missing} 项，错误 {errors} 项{suffix}。"


def _sections(context: CompareToolContext) -> List[dict]:
    return list(context.compare_payload.get("compare_sections") or [])


def _section_by_id(context: CompareToolContext, section_id: str) -> dict:
    wanted = str(section_id or "").strip()
    for section in _sections(context):
        if str(section.get("id") or "") == wanted:
            return section
    raise HarnessToolError(f"未找到对比分区：{wanted}")


def _section_rows(section: dict) -> List[dict]:
    table = section.get("table") if isinstance(section.get("table"), dict) else {}
    rows = table.get("rows")
    if isinstance(rows, list):
        return rows
    diff = section.get("diff") if isinstance(section.get("diff"), dict) else {}
    return list(diff.get("rows") or [])


def _row_text(row: dict) -> str:
    try:
        return json.dumps(row, ensure_ascii=False, sort_keys=True, default=str)
    except (TypeError, ValueError):
        return str(row)


def _list_compare_sections_tool(context: CompareToolContext, args: dict) -> dict:
    include_empty = bool(args.get("include_empty", True))
    sections = []
    for section in _sections(context):
        diff = section.get("diff") if isinstance(section.get("diff"), dict) else {}
        table = section.get("table") if isinstance(section.get("table"), dict) else {}
        total = (
            _as_int(diff.get("added_count"))
            + _as_int(diff.get("removed_count"))
            + _as_int(diff.get("changed_count"))
        )
        rows = _section_rows(section)
        if not include_empty and total <= 0 and not rows:
            continue
        sections.append({
            "id": _safe_text(section.get("id", ""), 120),
            "title": _safe_text(section.get("title", ""), 180),
            "lead": _safe_text(section.get("lead", ""), 260),
            "priority": _safe_text(section.get("priority", ""), 80),
            "added_count": _as_int(diff.get("added_count")),
            "removed_count": _as_int(diff.get("removed_count")),
            "changed_count": _as_int(diff.get("changed_count")),
            "total_rows": _as_int(diff.get("total_rows"), len(rows)),
            "displayed_count": len(rows),
            "truncated": bool(diff.get("truncated")),
            "columns": list(table.get("columns") or [])[:40],
        })
    return {
        "id": "list_compare_sections",
        "title": "对比分区清单",
        "target": "compare",
        "summary": f"当前 A/B 对比包含 {len(sections)} 个可审查分区。",
        "sections": sections,
        "readonly": True,
    }


def _get_compare_section_rows_tool(context: CompareToolContext, args: dict) -> dict:
    section_id = str(args.get("section_id") or "").strip()
    section = _section_by_id(context, section_id)
    limit = _as_int(args.get("limit", 30), 30)
    offset = _as_int(args.get("offset", 0), 0)
    diff_type = str(args.get("diff_type") or "").strip()
    rows = list(_section_rows(section))
    if diff_type:
        rows = [row for row in rows if isinstance(row, dict) and str(row.get("类型") or row.get("type") or "") == diff_type]
    selected = rows[offset:offset + limit]
    has_more = offset + len(selected) < len(rows)
    return {
        "id": "get_compare_section_rows",
        "title": section.get("title") or section_id,
        "target": "compare",
        "summary": (
            f"分区 {section_id} 共 {len(rows)} 行，返回 {len(selected)} 行。"
            + ("当前返回不是完整分区；如需完整判断请继续读取下一页或使用批量/聚合工具。" if has_more else "")
        ),
        "section_id": section_id,
        "section_title": section.get("title") or section_id,
        "offset": offset,
        "limit": limit,
        "total_rows": len(rows),
        "has_more": has_more,
        "next_offset": offset + len(selected) if has_more else None,
        "truncated": has_more,
        "rows": [
            {
                "__row_index": offset + index,
                "__row_number": offset + index + 1,
                **_compact_row(row),
            }
            for index, row in enumerate(selected)
            if isinstance(row, dict)
        ],
        "readonly": True,
    }


def _query_compare_diff_tool(context: CompareToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("query_compare_diff 需要 query。")
    section_id = str(args.get("section_id") or "").strip()
    limit = _as_int(args.get("limit", 30), 30)
    haystack = [_section_by_id(context, section_id)] if section_id else _sections(context)
    query_lower = query.lower()
    matches = []
    for section in haystack:
        sid = str(section.get("id") or "")
        title = section.get("title") or sid
        for index, row in enumerate(_section_rows(section)):
            if not isinstance(row, dict):
                continue
            text = _row_text(row)
            if query_lower not in text.lower():
                continue
            matches.append({
                "section_id": sid,
                "section_title": _safe_text(title, 180),
                "row_index": index,
                "row_number": index + 1,
                "row": _compact_row(row),
            })
            if len(matches) >= limit:
                break
        if len(matches) >= limit:
            break
    return {
        "id": "query_compare_diff",
        "title": f"搜索对比差异：{query}",
        "target": "compare",
        "summary": f"搜索 `{query}` 命中 {len(matches)} 条对比差异。",
        "query": _safe_text(query, 220),
        "section_id": section_id,
        "limit": limit,
        "matches": matches,
        "readonly": True,
    }


def _get_compare_row_tool(context: CompareToolContext, args: dict) -> dict:
    section_id = str(args.get("section_id") or "").strip()
    row_index = _as_int(args.get("row_index"), -1)
    section = _section_by_id(context, section_id)
    rows = _section_rows(section)
    if row_index < 0 or row_index >= len(rows):
        raise HarnessToolError(f"分区 {section_id} 不存在 row_index={row_index}。")
    row = rows[row_index]
    if not isinstance(row, dict):
        raise HarnessToolError(f"分区 {section_id} 的 row_index={row_index} 不是对象行。")
    return {
        "id": "get_compare_row",
        "title": f"{section.get('title') or section_id} #{row_index + 1}",
        "target": "compare",
        "summary": f"读取分区 {section_id} 第 {row_index + 1} 行。",
        "section_id": section_id,
        "section_title": section.get("title") or section_id,
        "row_index": row_index,
        "row_number": row_index + 1,
        "row": _compact_row(row, 520),
        "readonly": True,
    }


def _batch_query_compare_diff_tool(context: CompareToolContext, args: dict) -> dict:
    raw_items, input_truncated = _batch_input_items(args, "queries")
    global_section_id = str(args.get("section_id") or "").strip()
    default_limit = _batch_limit(args.get("limit_per_query", args.get("limit", BATCH_DEFAULT_PER_ITEM_LIMIT)))
    items = []
    for index, raw_item in enumerate(raw_items, start=1):
        if isinstance(raw_item, dict):
            query = str(raw_item.get("query") or raw_item.get("keyword") or raw_item.get("refdes") or raw_item.get("net") or raw_item.get("hq_no") or "").strip()
            section_id = str(raw_item.get("section_id") or global_section_id).strip()
            limit = _batch_limit(raw_item.get("limit", default_limit), default_limit)
        else:
            query = str(raw_item or "").strip()
            section_id = global_section_id
            limit = default_limit
        if not query:
            items.append({
                "index": index,
                "query": "",
                "section_id": section_id,
                "status": "error",
                "summary": "对比差异查询关键词为空。",
                "matches": [],
                "missing_reason": "empty_query",
            })
            continue
        try:
            result = _query_compare_diff_tool(context, {"query": query, "section_id": section_id, "limit": limit})
            matches = list(result.get("matches") or [])
            items.append({
                "index": index,
                "query": _safe_text(query, 220),
                "section_id": _safe_text(section_id, 120),
                "status": "found" if matches else "missing",
                "summary": result.get("summary", ""),
                "match_count": len(matches),
                "limit": limit,
                "matches": matches,
                "missing_reason": "" if matches else "no_compare_diff_match",
            })
        except Exception as exc:
            items.append({
                "index": index,
                "query": _safe_text(query, 220),
                "section_id": _safe_text(section_id, 120),
                "status": "error",
                "summary": str(exc),
                "matches": [],
                "missing_reason": str(exc),
            })
    return {
        "id": "batch_query_compare_diff",
        "title": "批量搜索对比差异",
        "target": "compare",
        "summary": _batch_summary("批量搜索对比差异", items, truncated=input_truncated),
        "input_count": len(raw_items),
        "input_truncated": input_truncated,
        "limit_per_query": default_limit,
        "items": items,
        "readonly": True,
    }


def _batch_get_compare_rows_tool(context: CompareToolContext, args: dict) -> dict:
    raw_items, input_truncated = _batch_input_items(args, "items")
    items = []
    for index, raw_item in enumerate(raw_items, start=1):
        if not isinstance(raw_item, dict):
            items.append({
                "index": index,
                "status": "error",
                "summary": "批量 diff row 请求必须是对象。",
                "missing_reason": "bad_request_item",
            })
            continue
        section_id = str(raw_item.get("section_id") or "").strip()
        row_index = _as_int(raw_item.get("row_index"), -1)
        try:
            result = _get_compare_row_tool(context, {"section_id": section_id, "row_index": row_index})
            items.append({
                "index": index,
                "section_id": section_id,
                "row_index": row_index,
                "status": "found",
                "summary": result.get("summary", ""),
                "row": result.get("row") or {},
                "section_title": result.get("section_title", ""),
                "row_number": result.get("row_number", row_index + 1),
                "missing_reason": "",
            })
        except Exception as exc:
            items.append({
                "index": index,
                "section_id": section_id,
                "row_index": row_index,
                "status": "error",
                "summary": str(exc),
                "row": {},
                "missing_reason": str(exc),
            })
    return {
        "id": "batch_get_compare_rows",
        "title": "批量读取对比差异行",
        "target": "compare",
        "summary": _batch_summary("批量读取对比差异行", items, truncated=input_truncated),
        "input_count": len(raw_items),
        "input_truncated": input_truncated,
        "items": items,
        "readonly": True,
    }


def _risk_priority(section: dict) -> str:
    priority = str(section.get("priority") or "normal")
    if priority == "critical":
        return "critical"
    if priority == "high":
        return "high"
    if priority == "report":
        return "medium"
    return "normal"


def _summarize_compare_risks_tool(context: CompareToolContext, args: dict) -> dict:
    limit = _as_int(args.get("limit", 12), 12)
    risks = []
    for section in _sections(context):
        diff = section.get("diff") if isinstance(section.get("diff"), dict) else {}
        total = (
            _as_int(diff.get("added_count"))
            + _as_int(diff.get("removed_count"))
            + _as_int(diff.get("changed_count"))
        )
        if total <= 0:
            continue
        rows = _section_rows(section)
        risks.append({
            "section_id": section.get("id", ""),
            "title": section.get("title", ""),
            "priority": _risk_priority(section),
            "total": total,
            "added": _as_int(diff.get("added_count")),
            "removed": _as_int(diff.get("removed_count")),
            "changed": _as_int(diff.get("changed_count")),
            "sample_rows": _compact_rows(rows, 3),
        })
    order = {"critical": 0, "high": 1, "medium": 2, "normal": 3}
    risks.sort(key=lambda item: (order.get(str(item.get("priority")), 4), -_as_int(item.get("total"))))
    diff_totals = context.compare_payload.get("diff_totals") or {}
    return {
        "id": "summarize_compare_risks",
        "title": "对比风险摘要",
        "target": "compare",
        "summary": f"汇总 {len(risks)} 个有差异的分区，优先关注关键器件和 Pin/Net。",
        "diff_totals": dict(diff_totals),
        "risk_items": risks[:limit],
        "left": context.compare_payload.get("left", {}),
        "right": context.compare_payload.get("right", {}),
        "readonly": True,
    }


def build_compare_tool_registry() -> HarnessToolRegistry:
    registry = HarnessToolRegistry()
    for tool in [
        HarnessTool(
            "list_harness_skills",
            "Harness Skill 清单",
            "列出当前仓库内 Compare Agent 可读取的技能卡；技能卡只提供打法指导，不授权新工具。",
            "skill",
            _list_harness_skills_tool,
            input_schema=_schema({
                "include_body": {"type": "boolean"},
                "max_body_chars": {"type": "integer", "minimum": 200, "maximum": 20000},
                "limit": {"type": "integer", "minimum": 1, "maximum": 200},
            }),
            evidence_kind="harness_skill",
        ),
        HarnessTool(
            "select_harness_skills",
            "选择 Harness Skill",
            "按用户问题、Compare profile、playbook 或工具名选择相关技能卡，帮助 Compare Agent 补充取证打法。",
            "skill",
            _select_harness_skills_tool,
            input_schema=_schema({
                "query": {"type": "string", "maxLength": 400},
                "capability_profiles": {"type": "array", "maxItems": 12},
                "playbooks": {"type": "array", "maxItems": 12},
                "tools": {"type": "array", "maxItems": 24},
                "include_body": {"type": "boolean"},
                "max_body_chars": {"type": "integer", "minimum": 200, "maximum": 20000},
                "limit": {"type": "integer", "minimum": 1, "maximum": 24},
            }),
            evidence_kind="harness_skill",
        ),
        HarnessTool(
            "get_harness_skill",
            "读取 Harness Skill",
            "按 skill_id 读取单张技能卡详情；用于在 A/B 对比运行中查看推荐工具、证据顺序和输出约束。",
            "skill",
            _get_harness_skill_tool,
            input_schema=_schema({
                "skill_id": {"type": "string", "maxLength": 120},
                "include_body": {"type": "boolean"},
                "max_body_chars": {"type": "integer", "minimum": 200, "maximum": 20000},
            }, required=["skill_id"]),
            evidence_kind="harness_skill",
        ),
        HarnessTool(
            "list_compare_sections",
            "对比分区清单",
            "列出当前 A/B 对比的分区、优先级、行数和列信息。",
            "compare",
            _list_compare_sections_tool,
            input_schema=_schema({"include_empty": {"type": "boolean"}}),
        ),
        HarnessTool(
            "get_compare_section_rows",
            "读取对比分区行",
            "按 section_id 分页读取对比差异行。",
            "compare",
            _get_compare_section_rows_tool,
            input_schema=_schema({
                "section_id": {"type": "string", "maxLength": 120},
                "limit": {"type": "integer", "minimum": 1, "maximum": 200},
                "offset": {"type": "integer", "minimum": 0, "maximum": 100000},
                "diff_type": {"type": "string", "maxLength": 80},
            }, required=["section_id"]),
        ),
        HarnessTool(
            "query_compare_diff",
            "搜索对比差异",
            "按位号、网络、HQ 料号、PI、选型顺序或任意关键词搜索对比差异。",
            "compare",
            _query_compare_diff_tool,
            input_schema=_schema({
                "query": {"type": "string", "maxLength": 200},
                "section_id": {"type": "string", "maxLength": 120},
                "limit": {"type": "integer", "minimum": 1, "maximum": 100},
            }, required=["query"]),
        ),
        HarnessTool(
            "batch_query_compare_diff",
            "批量搜索对比差异",
            "按多个位号、网络、HQ 料号、PI、选型顺序或页码关键词批量搜索对比差异。",
            "compare",
            _batch_query_compare_diff_tool,
            input_schema=_schema({
                "queries": {"type": "array", "minItems": 1, "maxItems": BATCH_MAX_ITEMS},
                "section_id": {"type": "string", "maxLength": 120},
                "limit_per_query": {"type": "integer", "minimum": 1, "maximum": BATCH_DEFAULT_PER_ITEM_LIMIT},
                "limit": {"type": "integer", "minimum": 1, "maximum": BATCH_DEFAULT_PER_ITEM_LIMIT},
            }, required=["queries"]),
        ),
        HarnessTool(
            "get_compare_row",
            "读取单条对比差异",
            "按 section_id 和 row_index 读取单条差异详情。",
            "compare",
            _get_compare_row_tool,
            input_schema=_schema({
                "section_id": {"type": "string", "maxLength": 120},
                "row_index": {"type": "integer", "minimum": 0, "maximum": 1000000},
            }, required=["section_id", "row_index"]),
        ),
        HarnessTool(
            "batch_get_compare_rows",
            "批量读取对比差异行",
            "按多个 section_id + row_index 批量读取对比差异详情。",
            "compare",
            _batch_get_compare_rows_tool,
            input_schema=_schema({
                "items": {"type": "array", "minItems": 1, "maxItems": BATCH_MAX_ITEMS},
            }, required=["items"]),
        ),
        HarnessTool(
            "summarize_compare_risks",
            "对比风险摘要",
            "本地汇总 A/B 对比中最高优先级的差异分区和样例。",
            "compare",
            _summarize_compare_risks_tool,
            input_schema=_schema({
                "limit": {"type": "integer", "minimum": 1, "maximum": 50},
            }),
        ),
        HarnessTool(
            "list_datasheet_documents",
            "本地规格书文档清单",
            "复用报告 Harness 的 datasheet SQLite 证据库，列出已索引 PDF、页数、chunk 数和状态。",
            "datasheet",
            list_datasheet_documents_tool,
            input_schema=_schema({
                "limit": {"type": "integer", "minimum": 1, "maximum": 1000},
                "offset": {"type": "integer", "minimum": 0, "maximum": 100000},
            }),
        ),
        HarnessTool(
            "list_datasheet_review_templates",
            "Datasheet 审查模板清单",
            "复用报告 Harness 的 LLM 可读 datasheet 审查模板，用于 A/B 规格书差异取证规划。",
            "datasheet",
            list_datasheet_review_templates_tool,
            input_schema=_schema({
                "category": {"type": "string", "maxLength": 80},
                "include_questions": {"type": "boolean"},
            }),
        ),
        HarnessTool(
            "get_datasheet_review_template",
            "读取 Datasheet 审查模板",
            "按 template_id 读取完整审查模板，帮助 Compare Agent 判断哪些参数/原理图证据必须二次读取。",
            "datasheet",
            get_datasheet_review_template_tool,
            input_schema=_schema({
                "template_id": {"type": "string", "maxLength": 80},
            }, required=["template_id"]),
        ),
        HarnessTool(
            "search_datasheet_chunks",
            "搜索本地规格书片段",
            "复用报告 Harness 的 datasheet chunk 索引，按料号、规格型号、芯片型号或参数关键词搜索 PDF evidence。",
            "datasheet",
            search_datasheet_chunks_tool,
            input_schema=_schema({
                "query": {"type": "string", "maxLength": 300},
                "limit": {"type": "integer", "minimum": 1, "maximum": 100},
                "offset": {"type": "integer", "minimum": 0, "maximum": 100000},
            }, required=["query"]),
        ),
        HarnessTool(
            "batch_search_datasheet_chunks",
            "批量搜索本地规格书片段",
            "按多个 HQ、规格、芯片型号或对比关键词一次性检索 PDF chunk。",
            "datasheet",
            batch_search_datasheet_chunks_tool,
            input_schema=_schema({
                "queries": {"type": "array", "minItems": 1, "maxItems": BATCH_MAX_ITEMS},
                "limit_per_query": {"type": "integer", "minimum": 1, "maximum": BATCH_DEFAULT_PER_ITEM_LIMIT},
                "limit": {"type": "integer", "minimum": 1, "maximum": BATCH_DEFAULT_PER_ITEM_LIMIT},
            }, required=["queries"]),
        ),
        HarnessTool(
            "search_datasheet_parameters",
            "搜索规格书参数卡",
            "搜索确定性抽取出的 datasheet 参数卡，如电压、电流、热、时序和环境参数。",
            "datasheet",
            search_datasheet_parameters_tool,
            input_schema=_schema({
                "query": {"type": "string", "maxLength": 240},
                "parameter_key": {"type": "string", "maxLength": 120},
                "doc_id": {"type": "integer", "minimum": 1, "maximum": 100000000},
                "limit": {"type": "integer", "minimum": 1, "maximum": 200},
                "offset": {"type": "integer", "minimum": 0, "maximum": 100000},
            }),
        ),
        HarnessTool(
            "get_datasheet_parameter",
            "读取规格书参数卡",
            "按 parameter_id 读取完整参数卡和来源文本，供 A/B 定量参数差异复核。",
            "datasheet",
            get_datasheet_parameter_tool,
            input_schema=_schema({
                "parameter_id": {"type": "integer", "minimum": 1, "maximum": 100000000},
                "max_chars": {"type": "integer", "minimum": 1, "maximum": 12000},
            }, required=["parameter_id"]),
        ),
        HarnessTool(
            "get_datasheet_chunk",
            "读取规格书 chunk",
            "按 doc_id 和 chunk_id 读取完整 PDF chunk；Compare Agent 下定量/电气参数结论前应二次取证。",
            "datasheet",
            get_datasheet_chunk_tool,
            input_schema=_schema({
                "doc_id": {"type": "integer", "minimum": 1, "maximum": 100000000},
                "chunk_id": {"type": "string", "maxLength": 120},
                "max_chars": {"type": "integer", "minimum": 1, "maximum": 12000},
            }, required=["doc_id", "chunk_id"]),
        ),
        HarnessTool(
            "get_datasheet_page_excerpt",
            "读取规格书页片段",
            "兼容按 doc_id 和 page 读取受限长度的本地规格书文本页片段。",
            "datasheet",
            get_datasheet_page_excerpt_tool,
            input_schema=_schema({
                "doc_id": {"type": "integer", "minimum": 1, "maximum": 100000000},
                "page": {"type": "integer", "minimum": 1, "maximum": 100000},
                "max_chars": {"type": "integer", "minimum": 1, "maximum": 12000},
            }, required=["doc_id", "page"]),
        ),
        HarnessTool(
            "list_compare_project_files",
            "A/B 项目只读文件清单",
            "列出 A/B 项目根目录下 compare agent 允许读取的文本文件。",
            "compare_file",
            list_compare_project_files_tool,
            input_schema=_schema({
                "side": {"type": "string", "enum": ["left", "right", "both"]},
                "limit": {"type": "integer", "minimum": 1, "maximum": 500},
            }),
            file_access=True,
        ),
        HarnessTool(
            "read_compare_project_text",
            "读取 A/B 项目文本文件",
            "读取当前 A/B 项目根目录内白名单范围的文本文件片段。",
            "compare_file",
            read_compare_project_text_tool,
            input_schema=_schema({
                "side": {"type": "string", "enum": ["left", "right"]},
                "path": {"type": "string", "maxLength": 1000},
                "max_chars": {"type": "integer", "minimum": 1, "maximum": 50000},
            }, required=["side", "path"]),
            file_access=True,
        ),
        HarnessTool(
            "resolve_compare_page_range",
            "解析 Cadence 页范围",
            "将用户输入的第 X-Y 页解析为页码，对应 sch_1/pageX.csv|csa。",
            "cadence_page",
            resolve_compare_page_range_tool,
            input_schema=_schema({
                "page_start": {"type": "integer", "minimum": 1, "maximum": 100000},
                "page_end": {"type": "integer", "minimum": 1, "maximum": 100000},
                "page_range": {"type": "string", "maxLength": 120},
            }),
            file_access=True,
        ),
        HarnessTool(
            "compare_cadence_page_semantics",
            "Cadence 页级语义比对",
            "构建 A/B sch_1/pageX.csv|csa 的 Cadence 页面语义模型并输出页级差异。",
            "cadence_page",
            compare_cadence_page_semantics_tool,
            input_schema=_schema({
                "page_start": {"type": "integer", "minimum": 1, "maximum": 100000},
                "page_end": {"type": "integer", "minimum": 1, "maximum": 100000},
                "page_range": {"type": "string", "maxLength": 120},
                "include_raw_unknown": {"type": "boolean"},
                "coordinate_tolerance": {"type": "integer", "minimum": 0, "maximum": 100000},
                "max_diff_items": {"type": "integer", "minimum": 1, "maximum": 500},
            }),
            file_access=True,
        ),
        HarnessTool(
            "get_cadence_page_object",
            "读取 Cadence 页对象详情",
            "按 side/page/object_id 读取图形对象、连接组件或 unknown 原始对象详情。",
            "cadence_page",
            get_cadence_page_object_tool,
            input_schema=_schema({
                "side": {"type": "string", "enum": ["left", "right"]},
                "page": {"type": "integer", "minimum": 1, "maximum": 100000},
                "object_id": {"type": "string", "maxLength": 160},
            }, required=["side", "page", "object_id"]),
            file_access=True,
        ),
        HarnessTool(
            "batch_get_cadence_page_objects",
            "批量读取 Cadence 页对象详情",
            "按多个 side/page/object_id 批量读取图形对象、连接组件或 unknown 原始对象详情。",
            "cadence_page",
            batch_get_cadence_page_objects_tool,
            input_schema=_schema({
                "objects": {"type": "array", "minItems": 1, "maxItems": BATCH_MAX_ITEMS},
            }, required=["objects"]),
            file_access=True,
        ),
        HarnessTool(
            "get_cadence_page_raw_excerpt",
            "读取 Cadence 页原始片段",
            "按 side/page/file_type 分页读取 sch_1/pageX.csv|csa 原始片段。",
            "cadence_page",
            get_cadence_page_raw_excerpt_tool,
            input_schema=_schema({
                "side": {"type": "string", "enum": ["left", "right"]},
                "page": {"type": "integer", "minimum": 1, "maximum": 100000},
                "file_type": {"type": "string", "enum": ["csv", "csa"]},
                "offset": {"type": "integer", "minimum": 0, "maximum": 100000000},
                "max_chars": {"type": "integer", "minimum": 1, "maximum": 50000},
            }, required=["side", "page", "file_type"]),
            file_access=True,
        ),
    ]:
        registry.register(tool)
    return registry
