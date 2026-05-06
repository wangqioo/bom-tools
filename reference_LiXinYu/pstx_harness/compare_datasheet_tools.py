# -*- coding: utf-8 -*-
"""Datasheet-backed read-only tools for the compare agent."""

from __future__ import annotations

from pstx_knowledge.datasheets import (
    batch_search_datasheet_chunks,
    get_datasheet_chunk,
    get_datasheet_page_excerpt,
    get_datasheet_parameter,
    list_datasheet_documents,
    search_datasheet_chunks,
    search_datasheet_parameters,
)
from pstx_knowledge.datasheet_review_templates import (
    get_datasheet_review_template,
    list_datasheet_review_templates,
)
from pstx_harness.tool_core import HarnessToolError


def _as_int(value, default: int = 0) -> int:
    try:
        return int(value if value is not None else default)
    except (TypeError, ValueError):
        return default


def _safe_text(value, limit: int = 260) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _datasheet_match_preview(match: dict) -> dict:
    if not isinstance(match, dict):
        return {}
    return {
        "doc_id": _as_int(match.get("doc_id"), 0),
        "title": _safe_text(match.get("title", ""), 220),
        "page": _as_int(match.get("page"), 1),
        "chunk_id": _safe_text(match.get("chunk_id", ""), 80),
        "section_title": _safe_text(match.get("section_title", ""), 180),
        "score": _as_int(match.get("score"), 0),
        "matched_terms": list(match.get("matched_terms") or [])[:8],
        "keywords": _safe_text(match.get("keywords", ""), 240),
        "char_range": list(match.get("char_range") or [])[:2],
        "snippet": _safe_text(match.get("snippet", ""), 420),
    }


def _datasheet_parameter_preview(parameter: dict) -> dict:
    return {
        "parameter_id": _as_int(parameter.get("parameter_id"), 0),
        "evidence_id": _safe_text(parameter.get("evidence_id", ""), 120),
        "doc_id": _as_int(parameter.get("doc_id"), 0),
        "title": _safe_text(parameter.get("title", ""), 220),
        "parameter_key": _safe_text(parameter.get("parameter_key", ""), 120),
        "parameter_name": _safe_text(parameter.get("parameter_name", ""), 220),
        "value_text": _safe_text(parameter.get("value_text", ""), 220),
        "unit": _safe_text(parameter.get("unit", ""), 60),
        "condition": _safe_text(parameter.get("condition", ""), 260),
        "page": _as_int(parameter.get("page"), 1),
        "chunk_id": _safe_text(parameter.get("chunk_id", ""), 100),
        "confidence": _safe_text(parameter.get("confidence", ""), 80),
        "source_text": _safe_text(parameter.get("source_text", ""), 420),
        "detail_locator": dict(parameter.get("detail_locator") or {}),
    }


def list_datasheet_review_templates_tool(context, args: dict) -> dict:
    result = list_datasheet_review_templates(
        str(args.get("category") or ""),
        include_questions=bool(args.get("include_questions", True)),
    )
    return {
        "id": "list_datasheet_review_templates",
        "title": "Datasheet 审查模板清单",
        "target": "datasheet",
        **result,
        "readonly": True,
    }


def get_datasheet_review_template_tool(context, args: dict) -> dict:
    template_id = str(args.get("template_id") or "").strip()
    if not template_id:
        raise HarnessToolError("get_datasheet_review_template 需要 template_id。")
    result = get_datasheet_review_template(template_id)
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取 datasheet 审查模板失败。"))
    return {
        "id": "get_datasheet_review_template",
        "title": result.get("template", {}).get("title") or "Datasheet 审查模板",
        "target": "datasheet",
        **result,
        "readonly": True,
    }


def list_datasheet_documents_tool(context, args: dict) -> dict:
    result = list_datasheet_documents(
        limit=_as_int(args.get("limit", 200), 200),
        offset=_as_int(args.get("offset", 0), 0),
    )
    return {
        "id": "list_datasheet_documents",
        "title": "本地规格书文档清单",
        "target": "datasheet",
        **result,
        "readonly": True,
    }


def search_datasheet_chunks_tool(context, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("search_datasheet_chunks 需要 query。")
    result = search_datasheet_chunks(
        query,
        limit=_as_int(args.get("limit", 20), 20),
        offset=_as_int(args.get("offset", 0), 0),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "规格书 chunk 检索失败。"))
    return {
        "id": "search_datasheet_chunks",
        "title": f"搜索规格书片段：{query}",
        "target": "datasheet",
        "summary": result.get("summary", ""),
        "query": _safe_text(query, 220),
        "terms": result.get("terms", []),
        "total_matches": result.get("total_matches", 0),
        "limit": result.get("limit", 20),
        "offset": result.get("offset", 0),
        "matches": [_datasheet_match_preview(match) for match in result.get("matches", [])],
        "readonly": True,
    }


def search_datasheet_parameters_tool(context, args: dict) -> dict:
    result = search_datasheet_parameters(
        str(args.get("query") or ""),
        parameter_key=str(args.get("parameter_key") or ""),
        doc_id=_as_int(args.get("doc_id"), 0) or None,
        limit=_as_int(args.get("limit", 30), 30),
        offset=_as_int(args.get("offset", 0), 0),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "规格书参数卡检索失败。"))
    return {
        "id": "search_datasheet_parameters",
        "title": "搜索规格书参数卡",
        "target": "datasheet",
        "summary": result.get("summary", ""),
        "query": _safe_text(result.get("query", ""), 220),
        "parameter_key": _safe_text(result.get("parameter_key", ""), 120),
        "doc_id": result.get("doc_id"),
        "total_matches": result.get("total_matches", 0),
        "limit": result.get("limit", 30),
        "offset": result.get("offset", 0),
        "parameters": [_datasheet_parameter_preview(item) for item in result.get("parameters", [])],
        "readonly": True,
    }


def get_datasheet_parameter_tool(context, args: dict) -> dict:
    result = get_datasheet_parameter(
        _as_int(args.get("parameter_id"), 0),
        max_chars=_as_int(args.get("max_chars", 2400), 2400),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取规格书参数卡失败。"))
    return {
        "id": "get_datasheet_parameter",
        "title": result.get("parameter_name") or "规格书参数卡",
        "target": "datasheet",
        **result,
        "readonly": True,
    }


def get_datasheet_chunk_tool(context, args: dict) -> dict:
    result = get_datasheet_chunk(
        _as_int(args.get("doc_id"), 0),
        str(args.get("chunk_id") or ""),
        max_chars=_as_int(args.get("max_chars", 4000), 4000),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取规格书 chunk 失败。"))
    return {
        "id": "get_datasheet_chunk",
        "title": result.get("title") or "规格书 chunk",
        "target": "datasheet",
        **result,
        "readonly": True,
    }


def get_datasheet_page_excerpt_tool(context, args: dict) -> dict:
    result = get_datasheet_page_excerpt(
        _as_int(args.get("doc_id"), 0),
        _as_int(args.get("page"), 1),
        max_chars=_as_int(args.get("max_chars", 2400), 2400),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取规格书页片段失败。"))
    return {
        "id": "get_datasheet_page_excerpt",
        "title": result.get("title") or "规格书页片段",
        "target": "datasheet",
        **result,
        "readonly": True,
    }


def batch_search_datasheet_chunks_tool(context, args: dict) -> dict:
    result = batch_search_datasheet_chunks(
        args.get("queries") or [],
        limit_per_query=_as_int(args.get("limit_per_query", args.get("limit", 8)), 8),
    )
    items = []
    for item in result.get("items", []) or []:
        if not isinstance(item, dict):
            continue
        compact = dict(item)
        compact["query"] = _safe_text(item.get("query", ""), 220)
        compact["matches"] = [_datasheet_match_preview(match) for match in item.get("matches", []) or []]
        compact["missing_reason"] = _safe_text(item.get("missing_reason", ""), 260)
        if item.get("error"):
            compact["error"] = _safe_text(item.get("error", ""), 260)
        items.append(compact)
    return {
        "id": "batch_search_datasheet_chunks",
        "title": "批量搜索规格书片段",
        "target": "datasheet",
        "summary": result.get("summary", ""),
        "query_count": result.get("query_count", 0),
        "limit_per_query": result.get("limit_per_query", 8),
        "truncated": bool(result.get("truncated")),
        "items": items,
        "readonly": True,
    }
