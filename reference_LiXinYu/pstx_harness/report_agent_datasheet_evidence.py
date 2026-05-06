# -*- coding: utf-8 -*-
"""Datasheet evidence node builders for the report harness agent."""

from __future__ import annotations

from typing import List, Optional, Sequence

from pstx_harness.report_agent_observation import preview as _preview


DATASHEET_EVIDENCE_TOOLS = {
    "list_datasheet_sources",
    "list_datasheet_review_templates",
    "get_datasheet_review_template",
    "list_datasheet_documents",
    "search_datasheet_chunks",
    "batch_search_datasheet_chunks",
    "search_datasheet_parameters",
    "get_datasheet_parameter",
    "search_datasheets",
    "match_component_datasheets",
    "batch_match_component_datasheets",
    "get_datasheet_chunk",
    "get_datasheet_excerpt",
    "get_datasheet_page_excerpt",
    "summarize_dfmea_datasheet_coverage",
}


def _row_summary(row: dict) -> str:
    parts = []
    for key in ["title", "page", "chunk_id", "section_title", "score", "snippet", "refdes", "query"]:
        value = row.get(key)
        if value not in {None, ""}:
            parts.append(f"{key}={_preview(value, 100)}")
    return "；".join(parts)


def _safe_evidence_fragment(value: str) -> str:
    fragment = "".join(char if char.isalnum() else "-" for char in str(value or "").strip())
    fragment = "-".join(part for part in fragment.split("-") if part)
    return fragment[:80] or "item"


def _datasheet_match_summary(match: dict) -> str:
    parts = []
    for key, label in [
        ("title", "文档"),
        ("page", "页"),
        ("chunk_id", "chunk"),
        ("section_title", "章节"),
        ("score", "分数"),
        ("snippet", "片段"),
    ]:
        value = match.get(key)
        if value not in {None, ""}:
            parts.append(f"{label}={_preview(value, 120)}")
    terms = list(match.get("matched_terms") or [])
    if terms:
        parts.append(f"命中={','.join(str(item) for item in terms[:6])}")
    return "；".join(parts) or _row_summary(match)


def _node(evidence_id: str,
          evidence_type: str,
          title: str,
          summary: str,
          *,
          tool_name: str,
          call_index: int,
          locator: Optional[dict] = None,
          payload_preview=None,
          missing_fields: Optional[Sequence[str]] = None,
          detail_tool: Optional[dict] = None) -> dict:
    node = {
        "id": evidence_id,
        "type": evidence_type,
        "title": _preview(title, 160),
        "summary": _preview(summary, 260),
        "source": {
            "tool": tool_name,
            "tool_call_index": call_index,
        },
        "locator": locator or {},
        "payload_preview": _preview(payload_preview if payload_preview is not None else {}),
    }
    if missing_fields:
        node["missing_fields"] = [str(item) for item in list(missing_fields)[:16]]
    if detail_tool:
        node["detail_tool"] = detail_tool
    return node


def datasheet_evidence_nodes_from_tool_result(tool_name: str,
                                              result: dict,
                                              *,
                                              call_index: int,
                                              args: Optional[dict] = None) -> Optional[List[dict]]:
    if tool_name not in DATASHEET_EVIDENCE_TOOLS:
        return None

    args = args or {}
    nodes: List[dict] = []
    base = f"ev-{call_index}"
    if tool_name == "list_datasheet_sources":
        nodes.append(_node(
            f"{base}-datasheet-status",
            "datasheet_document",
            result.get("title") or "本地规格书索引状态",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"target": "dfmea", "db_path": result.get("db_path", "")},
            payload_preview={
                "configured": result.get("configured", False),
                "document_count": result.get("document_count", 0),
                "indexed_count": result.get("indexed_count", 0),
                "page_count": result.get("page_count", 0),
                "chunk_count": result.get("chunk_count", 0),
                "section_count": result.get("section_count", 0),
                "failed_count": result.get("failed_count", 0),
            },
        ))
        return nodes
    if tool_name == "list_datasheet_review_templates":
        for index, template in enumerate(list(result.get("templates") or [])[:12], start=1):
            if not isinstance(template, dict):
                continue
            template_id = str(template.get("template_id") or f"template-{index}")
            nodes.append(_node(
                f"{base}-datasheet-template-{_safe_evidence_fragment(template_id)}",
                "datasheet_review_template",
                template.get("title") or template_id,
                template.get("llm_goal") or f"category={template.get('category', '')}",
                tool_name=tool_name,
                call_index=call_index,
                locator={"template_id": template_id, "category": template.get("category", "")},
                payload_preview=template,
                detail_tool={"name": "get_datasheet_review_template", "args": {"template_id": template_id}},
            ))
        return nodes if nodes else None
    if tool_name == "get_datasheet_review_template":
        template = dict(result.get("template") or {})
        template_id = str(template.get("template_id") or args.get("template_id") or "template")
        nodes.append(_node(
            f"{base}-datasheet-template-{_safe_evidence_fragment(template_id)}",
            "datasheet_review_template",
            template.get("title") or template_id,
            template.get("llm_goal") or result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"template_id": template_id, "category": template.get("category", "")},
            payload_preview={
                "applies_to": template.get("applies_to", []),
                "section_count": len(template.get("extraction_sections", []) or []),
                "required_evidence": template.get("required_evidence", []),
                "red_flags": list(template.get("red_flags") or [])[:8],
            },
            detail_tool={"name": "get_datasheet_review_template", "args": {"template_id": template_id}},
        ))
        return nodes
    if tool_name == "list_datasheet_documents":
        for index, document in enumerate(list(result.get("documents") or [])[:24], start=1):
            doc_id = int(document.get("doc_id") or 0)
            nodes.append(_node(
                f"{base}-datasheet-document-{doc_id or index}",
                "datasheet_document",
                document.get("title") or f"规格书文档 {doc_id or index}",
                (
                    f"状态={document.get('status') or ''}；页数={document.get('page_count') or 0}；"
                    f"chunks={document.get('chunk_count') or 0}；错误={document.get('error') or ''}"
                ),
                tool_name=tool_name,
                call_index=call_index,
                locator={"doc_id": doc_id, "path": document.get("path", "")},
                payload_preview=document,
                detail_tool={"name": "search_datasheet_chunks", "args": {"query": document.get("title") or "", "limit": 10}} if doc_id else None,
            ))
        return nodes if nodes else None
    if tool_name in {"search_datasheet_chunks", "batch_search_datasheet_chunks"}:
        if tool_name == "search_datasheet_chunks":
            items = [{"query": result.get("query", ""), "matches": result.get("matches", [])}]
        else:
            items = list(result.get("items") or [])
        for item_index, item in enumerate(items[:24], start=1):
            if not isinstance(item, dict):
                continue
            query = str(item.get("query") or result.get("query") or "")
            matches = list(item.get("matches") or [])
            if not matches:
                nodes.append(_node(
                    f"{base}-datasheet-gap-{item_index}-{_safe_evidence_fragment(query or 'query')}",
                    "datasheet_gap",
                    f"规格书检索无命中：{query or '空查询'}",
                    item.get("missing_reason") or item.get("summary") or "本地规格书 chunk 索引未命中。",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"query": query},
                    payload_preview=item,
                    missing_fields=["datasheet_chunk"],
                ))
                continue
            for match_index, match in enumerate(matches[:8], start=1):
                doc_id = int(match.get("doc_id") or 0)
                page = int(match.get("page") or 1)
                chunk_id = str(match.get("chunk_id") or "")
                nodes.append(_node(
                    f"{base}-datasheet-chunk-{doc_id}-{_safe_evidence_fragment(chunk_id or match_index)}",
                    "datasheet_chunk",
                    f"{match.get('title') or '规格书'} 第 {page} 页 {chunk_id or 'chunk'}",
                    _datasheet_match_summary(match),
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"doc_id": doc_id, "page": page, "chunk_id": chunk_id, "query": query},
                    payload_preview=match,
                    detail_tool={"name": "get_datasheet_chunk", "args": {"doc_id": doc_id, "chunk_id": chunk_id, "max_chars": 4000}} if doc_id and chunk_id else None,
                ))
        return nodes if nodes else None
    if tool_name == "search_datasheet_parameters":
        parameters = list(result.get("parameters") or [])
        if not parameters:
            query = str(result.get("query") or result.get("parameter_key") or "")
            nodes.append(_node(
                f"{base}-datasheet-parameter-gap-{_safe_evidence_fragment(query or 'query')}",
                "datasheet_gap",
                f"规格书参数卡无命中：{query or '空查询'}",
                result.get("summary") or "本地规格书参数卡未命中。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"query": query, "parameter_key": result.get("parameter_key", "")},
                payload_preview=result,
                missing_fields=["datasheet_parameter"],
            ))
            return nodes
        for index, parameter in enumerate(parameters[:24], start=1):
            if not isinstance(parameter, dict):
                continue
            parameter_id = int(parameter.get("parameter_id") or index)
            doc_id = int(parameter.get("doc_id") or 0)
            page = int(parameter.get("page") or 1)
            nodes.append(_node(
                f"{base}-datasheet-param-{parameter_id}",
                "datasheet_parameter",
                parameter.get("parameter_name") or parameter.get("parameter_key") or f"参数卡 {parameter_id}",
                (
                    f"{parameter.get('value_text') or ''} {parameter.get('unit') or ''}；"
                    f"条件={parameter.get('condition') or ''}；置信度={parameter.get('confidence') or ''}"
                ),
                tool_name=tool_name,
                call_index=call_index,
                locator={"parameter_id": parameter_id, "doc_id": doc_id, "page": page, "chunk_id": parameter.get("chunk_id", "")},
                payload_preview=parameter,
                detail_tool={"name": "get_datasheet_parameter", "args": {"parameter_id": parameter_id, "max_chars": 2400}},
            ))
        return nodes
    if tool_name == "get_datasheet_parameter":
        parameter_id = int(result.get("parameter_id") or args.get("parameter_id") or 0)
        doc_id = int(result.get("doc_id") or 0)
        page = int(result.get("page") or 1)
        nodes.append(_node(
            f"{base}-datasheet-param-{parameter_id or 'parameter'}",
            "datasheet_parameter",
            result.get("parameter_name") or result.get("parameter_key") or "规格书参数卡",
            result.get("summary") or f"{result.get('value_text') or ''} {result.get('unit') or ''}",
            tool_name=tool_name,
            call_index=call_index,
            locator={"parameter_id": parameter_id, "doc_id": doc_id, "page": page, "chunk_id": result.get("chunk_id", "")},
            payload_preview={
                "value_text": result.get("value_text", ""),
                "unit": result.get("unit", ""),
                "condition": result.get("condition", ""),
                "source_text": str(result.get("source_text") or "")[:700],
                "source_truncated": result.get("source_truncated", False),
            },
            detail_tool={"name": "get_datasheet_parameter", "args": {"parameter_id": parameter_id, "max_chars": 2400}} if parameter_id else None,
        ))
        return nodes
    if tool_name in {"search_datasheets", "match_component_datasheets"}:
        refdes = str(result.get("refdes") or "")
        matches = list(result.get("matches") or [])
        if not matches and tool_name == "match_component_datasheets":
            nodes.append(_node(
                f"{base}-datasheet-gap-{_safe_evidence_fragment(refdes or result.get('query') or 'component')}",
                "datasheet_gap",
                f"{refdes or '元件'} 缺规格书证据",
                result.get("missing_reason") or result.get("summary") or "本地规格书索引未命中。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"refdes": refdes, "query": result.get("query", "")},
                payload_preview={"query": result.get("query", ""), "card": result.get("card", {})},
                missing_fields=["datasheet_match"],
            ))
            return nodes
        for index, match in enumerate(matches[:24], start=1):
            doc_id = int(match.get("doc_id") or 0)
            page = int(match.get("page") or 1)
            evidence_type = "datasheet_match" if tool_name == "match_component_datasheets" else "datasheet_excerpt"
            title_prefix = f"{refdes} " if refdes else ""
            nodes.append(_node(
                f"{base}-datasheet-{doc_id}-{page}-{index}",
                evidence_type,
                f"{title_prefix}{match.get('title') or '规格书'} 第 {page} 页",
                _datasheet_match_summary(match),
                tool_name=tool_name,
                call_index=call_index,
                locator={"doc_id": doc_id, "page": page, "refdes": refdes, "query": result.get("query", "")},
                payload_preview=match,
                detail_tool={"name": "get_datasheet_chunk", "args": {"doc_id": doc_id, "chunk_id": match.get("chunk_id", ""), "max_chars": 4000}} if doc_id and match.get("chunk_id") else (
                    {"name": "get_datasheet_excerpt", "args": {"doc_id": doc_id, "page": page, "max_chars": 2400}} if doc_id else None
                ),
            ))
        return nodes if nodes else None
    if tool_name == "batch_match_component_datasheets":
        for item_index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            refdes = str(item.get("refdes") or f"component-{item_index}")
            matches = list(item.get("matches") or [])
            if not matches:
                nodes.append(_node(
                    f"{base}-batch-datasheet-gap-{_safe_evidence_fragment(refdes)}",
                    "datasheet_gap",
                    f"{refdes} 缺规格书证据",
                    item.get("missing_reason") or item.get("summary") or "本地规格书索引未命中。",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"refdes": refdes, "query": item.get("query", "")},
                    payload_preview=item,
                    missing_fields=["datasheet_match"],
                ))
                continue
            for match_index, match in enumerate(matches[:4], start=1):
                doc_id = int(match.get("doc_id") or 0)
                page = int(match.get("page") or 1)
                nodes.append(_node(
                    f"{base}-batch-datasheet-{_safe_evidence_fragment(refdes)}-{doc_id}-{page}-{match_index}",
                    "datasheet_match",
                    f"{refdes} {match.get('title') or '规格书'} 第 {page} 页",
                    _datasheet_match_summary(match),
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"refdes": refdes, "doc_id": doc_id, "page": page, "query": item.get("query", "")},
                    payload_preview={"item": item, "match": match},
                    detail_tool={"name": "get_datasheet_chunk", "args": {"doc_id": doc_id, "chunk_id": match.get("chunk_id", ""), "max_chars": 4000}} if doc_id and match.get("chunk_id") else (
                        {"name": "get_datasheet_excerpt", "args": {"doc_id": doc_id, "page": page, "max_chars": 2400}} if doc_id else None
                    ),
                ))
        return nodes if nodes else None
    if tool_name == "get_datasheet_chunk":
        doc_id = int(result.get("doc_id") or args.get("doc_id") or 0)
        page = int(result.get("page") or 1)
        chunk_id = str(result.get("chunk_id") or args.get("chunk_id") or "")
        nodes.append(_node(
            f"{base}-datasheet-chunk-{doc_id}-{_safe_evidence_fragment(chunk_id or 'chunk')}",
            "datasheet_chunk",
            f"{result.get('title') or '规格书'} 第 {page} 页 {chunk_id or 'chunk'}",
            result.get("summary") or "规格书 chunk 片段。",
            tool_name=tool_name,
            call_index=call_index,
            locator={"doc_id": doc_id, "page": page, "chunk_id": chunk_id, "path": result.get("path", "")},
            payload_preview={
                "section_title": result.get("section_title", ""),
                "keywords": result.get("keywords", ""),
                "content_preview": str(result.get("content") or "")[:700],
                "truncated": result.get("truncated"),
            },
            detail_tool={"name": "get_datasheet_chunk", "args": {"doc_id": doc_id, "chunk_id": chunk_id, "max_chars": 4000}} if doc_id and chunk_id else None,
        ))
        return nodes
    if tool_name in {"get_datasheet_excerpt", "get_datasheet_page_excerpt"}:
        doc_id = int(result.get("doc_id") or args.get("doc_id") or 0)
        page = int(result.get("page") or args.get("page") or 1)
        nodes.append(_node(
            f"{base}-datasheet-excerpt-{doc_id}-{page}",
            "datasheet_excerpt",
            f"{result.get('title') or '规格书'} 第 {page} 页",
            result.get("summary") or "规格书页级片段。",
            tool_name=tool_name,
            call_index=call_index,
            locator={"doc_id": doc_id, "page": page, "path": result.get("path", "")},
            payload_preview={"content_preview": str(result.get("content") or "")[:600], "truncated": result.get("truncated")},
            detail_tool={"name": tool_name, "args": {"doc_id": doc_id, "page": page, "max_chars": 2400}} if doc_id else None,
        ))
        return nodes
    if tool_name == "summarize_dfmea_datasheet_coverage":
        nodes.append(_node(
            f"{base}-datasheet-coverage",
            "dfmea_readiness",
            result.get("title") or "DFMEA 规格书覆盖摘要",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"target": "dfmea"},
            payload_preview={
                "total_key_components": result.get("total_key_components", 0),
                "matched_count": result.get("matched_count", 0),
                "gap_count": result.get("gap_count", 0),
            },
            detail_tool={"name": "summarize_dfmea_datasheet_coverage", "args": {}},
        ))
        for index, card in enumerate(list(result.get("matched_cards") or [])[:12], start=1):
            if not isinstance(card, dict):
                continue
            refdes = str(card.get("refdes") or f"matched-{index}")
            for match_index, match in enumerate(list(card.get("matches") or [])[:3], start=1):
                doc_id = int(match.get("doc_id") or 0)
                page = int(match.get("page") or 1)
                nodes.append(_node(
                    f"{base}-datasheet-match-{_safe_evidence_fragment(refdes)}-{doc_id}-{page}-{match_index}",
                    "datasheet_match",
                    f"{refdes} 规格书匹配",
                    _datasheet_match_summary(match),
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"refdes": refdes, "doc_id": doc_id, "page": page},
                    payload_preview={"card": card, "match": match},
                    detail_tool={"name": "get_datasheet_excerpt", "args": {"doc_id": doc_id, "page": page, "max_chars": 2400}} if doc_id else None,
                ))
        for index, card in enumerate(list(result.get("gap_cards") or [])[:12], start=1):
            if not isinstance(card, dict):
                continue
            refdes = str(card.get("refdes") or f"gap-{index}")
            nodes.append(_node(
                f"{base}-datasheet-gap-{_safe_evidence_fragment(refdes)}",
                "datasheet_gap",
                f"{refdes} 缺规格书证据",
                card.get("missing_reason") or "本地规格书索引未命中。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"refdes": refdes, "query": card.get("query", "")},
                payload_preview=card,
                missing_fields=["datasheet_match"],
            ))
        return nodes
    return None
