# -*- coding: utf-8 -*-
"""Evidence node construction for the compare agent."""

from __future__ import annotations

from typing import List, Optional, Sequence, Tuple

from pstx_agent_runtime import (
    normalize_citations as runtime_normalize_citations,
    normalize_proposed_actions as runtime_normalize_proposed_actions,
)
from pstx_harness.compare_agent_observation import preview as _preview


def _node(evidence_id: str,
          evidence_type: str,
          title: str,
          summary: str,
          *,
          tool_name: str,
          call_index: int,
          locator: Optional[dict] = None,
          payload_preview=None,
          detail_tool: Optional[dict] = None) -> dict:
    node = {
        "id": evidence_id,
        "type": evidence_type,
        "title": _preview(title, 160),
        "summary": _preview(summary, 280),
        "source": {"tool": tool_name, "tool_call_index": call_index},
        "locator": locator or {},
        "payload_preview": _preview(payload_preview if payload_preview is not None else {}),
    }
    if detail_tool:
        node["detail_tool"] = detail_tool
    return node


def _row_summary(row: dict) -> str:
    if not isinstance(row, dict):
        return ""
    preferred = ["类型", "位号", "网络名", "器件类别", "引脚", "左侧网络", "右侧网络", "变化字段", "左侧", "右侧"]
    parts = []
    for key in preferred:
        value = row.get(key)
        if value not in (None, ""):
            parts.append(f"{key}={_preview(value, 80)}")
        if len(parts) >= 5:
            break
    if not parts:
        for key, value in list(row.items())[:5]:
            parts.append(f"{key}={_preview(value, 80)}")
    return "；".join(parts)


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
        value = match.get(key) if isinstance(match, dict) else None
        if value not in {None, ""}:
            parts.append(f"{label}={_preview(value, 120)}")
    terms = list(match.get("matched_terms") or []) if isinstance(match, dict) else []
    if terms:
        parts.append(f"命中={','.join(str(item) for item in terms[:6])}")
    return "；".join(parts) or _row_summary(match if isinstance(match, dict) else {})


def _evidence_type_for_section(section_id: str) -> str:
    if section_id in {"key_components", "components"}:
        return "compare_component"
    if section_id in {"nets", "key_pin_nets", "passive_pin_nets"}:
        return "compare_net"
    if "feishu" in section_id or "bom" in section_id:
        return "compare_feishu_material"
    return "compare_diff"


def evidence_nodes_from_tool_result(tool_name: str, result: dict, *, call_index: int, args: Optional[dict] = None) -> List[dict]:
    args = args or {}
    nodes: List[dict] = []
    base = f"ev-{call_index}"
    if tool_name == "list_compare_sections":
        for index, section in enumerate(list(result.get("sections") or [])[:24], start=1):
            sid = str(section.get("id") or "")
            nodes.append(_node(
                f"{base}-section-{index}",
                "compare_diff",
                section.get("title") or sid,
                f"分区 {sid}，差异 {section.get('total_rows', 0)} 行。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"section_id": sid},
                payload_preview=section,
            ))
        return nodes
    if tool_name in {"get_compare_section_rows", "query_compare_diff"}:
        rows = result.get("rows") if tool_name == "get_compare_section_rows" else result.get("matches")
        for index, item in enumerate(list(rows or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            row = item.get("row") if isinstance(item.get("row"), dict) else item
            sid = str(item.get("section_id") or result.get("section_id") or args.get("section_id") or "")
            row_index = int(item.get("row_index", item.get("__row_index", index - 1)) or 0)
            nodes.append(_node(
                f"{base}-compare-{sid or 'all'}-{row_index + 1}",
                _evidence_type_for_section(sid),
                item.get("section_title") or result.get("section_title") or sid or "对比差异",
                _row_summary(row),
                tool_name=tool_name,
                call_index=call_index,
                locator={"section_id": sid, "row_index": row_index, "row_number": row_index + 1},
                payload_preview=row,
            ))
        return nodes
    if tool_name == "batch_query_compare_diff":
        for item_index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            query = str(item.get("query") or f"query-{item_index}")
            matches = list(item.get("matches") or [])
            if not matches:
                nodes.append(_node(
                    f"{base}-batch-query-{item_index}",
                    "compare_diff",
                    f"对比搜索 {query}",
                    item.get("summary") or item.get("missing_reason") or "",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"query": query, "section_id": item.get("section_id", ""), "status": item.get("status", "")},
                    payload_preview=item,
                ))
                continue
            for match_index, match in enumerate(matches[:8], start=1):
                if not isinstance(match, dict):
                    continue
                row = match.get("row") if isinstance(match.get("row"), dict) else match
                sid = str(match.get("section_id") or item.get("section_id") or "")
                row_index = int(match.get("row_index", match_index - 1) or 0)
                nodes.append(_node(
                    f"{base}-batch-query-{item_index}-{sid or 'all'}-{row_index + 1}",
                    _evidence_type_for_section(sid),
                    match.get("section_title") or sid or f"对比搜索 {query}",
                    _row_summary(row),
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"query": query, "section_id": sid, "row_index": row_index, "row_number": row_index + 1},
                    payload_preview=row,
                ))
        return nodes
    if tool_name == "get_compare_row":
        row = result.get("row") if isinstance(result.get("row"), dict) else {}
        sid = str(result.get("section_id") or args.get("section_id") or "")
        row_index = int(result.get("row_index", args.get("row_index", 0)) or 0)
        return [_node(
            f"{base}-compare-{sid}-{row_index + 1}",
            _evidence_type_for_section(sid),
            result.get("title") or sid,
            _row_summary(row),
            tool_name=tool_name,
            call_index=call_index,
            locator={"section_id": sid, "row_index": row_index, "row_number": row_index + 1},
            payload_preview=row,
        )]
    if tool_name == "batch_get_compare_rows":
        for item_index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            sid = str(item.get("section_id") or "")
            row_index = int(item.get("row_index", 0) or 0)
            row = item.get("row") if isinstance(item.get("row"), dict) else {}
            nodes.append(_node(
                f"{base}-batch-row-{item_index}-{sid}-{row_index + 1}",
                _evidence_type_for_section(sid),
                item.get("section_title") or sid or f"对比差异行 {item_index}",
                _row_summary(row) if row else item.get("summary", ""),
                tool_name=tool_name,
                call_index=call_index,
                locator={"section_id": sid, "row_index": row_index, "row_number": row_index + 1, "status": item.get("status", "")},
                payload_preview=item,
            ))
        return nodes
    if tool_name == "summarize_compare_risks":
        for index, item in enumerate(list(result.get("risk_items") or [])[:24], start=1):
            sid = str(item.get("section_id") or "")
            nodes.append(_node(
                f"{base}-risk-{index}",
                _evidence_type_for_section(sid),
                item.get("title") or sid or f"风险 {index}",
                f"{item.get('priority', 'normal')} 优先级，差异 {item.get('total', 0)} 行。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"section_id": sid},
                payload_preview=item,
            ))
        return nodes
    if tool_name == "list_datasheet_documents":
        for index, document in enumerate(list(result.get("documents") or [])[:24], start=1):
            if not isinstance(document, dict):
                continue
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
        return nodes
    if tool_name == "list_datasheet_review_templates":
        for index, template in enumerate(list(result.get("templates") or [])[:12], start=1):
            if not isinstance(template, dict):
                continue
            template_id = str(template.get("template_id") or f"template-{index}")
            nodes.append(_node(
                f"{base}-datasheet-template-{template_id}",
                "datasheet_review_template",
                template.get("title") or template_id,
                template.get("llm_goal") or f"category={template.get('category', '')}",
                tool_name=tool_name,
                call_index=call_index,
                locator={"template_id": template_id, "category": template.get("category", "")},
                payload_preview=template,
                detail_tool={"name": "get_datasheet_review_template", "args": {"template_id": template_id}},
            ))
        return nodes
    if tool_name == "get_datasheet_review_template":
        template = dict(result.get("template") or {})
        template_id = str(template.get("template_id") or args.get("template_id") or "template")
        return [_node(
            f"{base}-datasheet-template-{template_id}",
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
        )]
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
                    f"{base}-datasheet-gap-{item_index}",
                    "datasheet_gap",
                    f"规格书检索无命中：{query or '空查询'}",
                    item.get("missing_reason") or item.get("summary") or "本地规格书 chunk 索引未命中。",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"query": query},
                    payload_preview=item,
                ))
                continue
            for match_index, match in enumerate(matches[:8], start=1):
                if not isinstance(match, dict):
                    continue
                doc_id = int(match.get("doc_id") or 0)
                page = int(match.get("page") or 1)
                chunk_id = str(match.get("chunk_id") or "")
                nodes.append(_node(
                    f"{base}-datasheet-chunk-{doc_id}-{chunk_id or match_index}",
                    "datasheet_chunk",
                    f"{match.get('title') or '规格书'} 第 {page} 页 {chunk_id or 'chunk'}",
                    _datasheet_match_summary(match),
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"doc_id": doc_id, "page": page, "chunk_id": chunk_id, "query": query},
                    payload_preview=match,
                    detail_tool={"name": "get_datasheet_chunk", "args": {"doc_id": doc_id, "chunk_id": chunk_id, "max_chars": 4000}} if doc_id and chunk_id else None,
                ))
        return nodes
    if tool_name == "search_datasheet_parameters":
        parameters = list(result.get("parameters") or [])
        if not parameters:
            query = str(result.get("query") or result.get("parameter_key") or "")
            return [_node(
                f"{base}-datasheet-parameter-gap",
                "datasheet_gap",
                f"规格书参数卡无命中：{query or '空查询'}",
                result.get("summary") or "本地规格书参数卡未命中。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"query": query, "parameter_key": result.get("parameter_key", "")},
                payload_preview=result,
            )]
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
                    f"条件={parameter.get('condition') or ''}"
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
        return [_node(
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
            },
            detail_tool={"name": "get_datasheet_parameter", "args": {"parameter_id": parameter_id, "max_chars": 2400}} if parameter_id else None,
        )]
    if tool_name == "get_datasheet_chunk":
        doc_id = int(result.get("doc_id") or args.get("doc_id") or 0)
        page = int(result.get("page") or 1)
        chunk_id = str(result.get("chunk_id") or args.get("chunk_id") or "")
        return [_node(
            f"{base}-datasheet-chunk-{doc_id}-{chunk_id or 'chunk'}",
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
        )]
    if tool_name == "get_datasheet_page_excerpt":
        doc_id = int(result.get("doc_id") or args.get("doc_id") or 0)
        page = int(result.get("page") or args.get("page") or 1)
        return [_node(
            f"{base}-datasheet-excerpt-{doc_id}-{page}",
            "datasheet_excerpt",
            f"{result.get('title') or '规格书'} 第 {page} 页",
            result.get("summary") or "规格书页级片段。",
            tool_name=tool_name,
            call_index=call_index,
            locator={"doc_id": doc_id, "page": page, "path": result.get("path", "")},
            payload_preview={"content_preview": str(result.get("content") or "")[:600], "truncated": result.get("truncated")},
            detail_tool={"name": "get_datasheet_page_excerpt", "args": {"doc_id": doc_id, "page": page, "max_chars": 2400}} if doc_id else None,
        )]
    if tool_name == "list_compare_project_files":
        for index, item in enumerate(list(result.get("files") or [])[:24], start=1):
            nodes.append(_node(
                f"{base}-file-{index}",
                "compare_file_excerpt",
                f"{item.get('side', '')}:{item.get('path', '')}",
                f"允许读取文件，大小 {item.get('size', 0)} bytes。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"side": item.get("side", ""), "path": item.get("path", "")},
                payload_preview=item,
            ))
        return nodes
    if tool_name == "read_compare_project_text":
        return [_node(
            f"{base}-excerpt",
            "compare_file_excerpt",
            result.get("title") or f"{result.get('side', '')}:{result.get('path', '')}",
            result.get("summary") or "A/B 项目文本片段。",
            tool_name=tool_name,
            call_index=call_index,
            locator={"side": result.get("side", ""), "path": result.get("path", ""), "encoding": result.get("encoding", "")},
            payload_preview={"content_preview": str(result.get("content") or "")[:500], "truncated": result.get("truncated")},
        )]
    if tool_name == "resolve_compare_page_range":
        return [_node(
            f"{base}-page-range",
            "compare_page_range",
            result.get("title") or "Cadence 页范围",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"page_start": result.get("page_start"), "page_end": result.get("page_end")},
            payload_preview={
                "page_start": result.get("page_start"),
                "page_end": result.get("page_end"),
                "page_count": result.get("page_count"),
                "page_semantics": result.get("page_semantics"),
            },
        )]
    if tool_name == "compare_cadence_page_semantics":
        for index, page_result in enumerate(list(result.get("page_results") or [])[:12], start=1):
            page = page_result.get("page")
            status = page_result.get("status")
            nodes.append(_node(
                f"{base}-cadence-page-{page}",
                "cadence_page_model",
                f"PAGE{page} Cadence 语义模型",
                f"PAGE{page} 状态 {status}，差异 {page_result.get('diff_count', 0)} 项。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"page": page, "status": status},
                payload_preview={
                    "left_digest": page_result.get("left_digest"),
                    "right_digest": page_result.get("right_digest"),
                    "diff_count": page_result.get("diff_count"),
                },
            ))
            for diff_index, diff in enumerate(list(page_result.get("diffs") or [])[:4], start=1):
                item_type = str(diff.get("item_type") or "")
                evidence_type = "cadence_topology_diff" if item_type == "CONNECTIVITY" else (
                    "cadence_property_diff" if item_type in {"CSV_PROPERTY", "PAGE_NUMBER"} else "cadence_graphic_object"
                )
                nodes.append(_node(
                    f"{base}-cadence-page-{page}-diff-{diff_index}",
                    evidence_type,
                    f"PAGE{page} {item_type} {diff.get('type', '')}",
                    f"{diff.get('type', '')} {item_type}",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"page": page, "diff_index": diff_index - 1, "item_type": item_type},
                    payload_preview=diff,
                ))
        return nodes
    if tool_name == "get_cadence_page_object":
        object_kind = str(result.get("object_kind") or "")
        evidence_type = "cadence_topology_diff" if object_kind == "connectivity" else "cadence_graphic_object"
        return [_node(
            f"{base}-cadence-object",
            evidence_type,
            result.get("title") or "Cadence 对象详情",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"side": result.get("side"), "page": result.get("page"), "object_id": result.get("object_id")},
            payload_preview=result.get("object") or {},
        )]
    if tool_name == "batch_get_cadence_page_objects":
        for item_index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            object_kind = str(item.get("object_kind") or "")
            evidence_type = "cadence_topology_diff" if object_kind == "connectivity" else "cadence_graphic_object"
            nodes.append(_node(
                f"{base}-batch-cadence-object-{item_index}",
                evidence_type,
                item.get("object_id") or f"Cadence 对象 {item_index}",
                item.get("summary") or "",
                tool_name=tool_name,
                call_index=call_index,
                locator={
                    "side": item.get("side"),
                    "page": item.get("page"),
                    "object_id": item.get("object_id"),
                    "status": item.get("status", ""),
                },
                payload_preview=item.get("object") or item,
            ))
        return nodes
    if tool_name == "get_cadence_page_raw_excerpt":
        return [_node(
            f"{base}-cadence-raw",
            "cadence_raw_excerpt",
            result.get("title") or "Cadence 原始片段",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"side": result.get("side"), "page": result.get("page"), "file_type": result.get("file_type"), "offset": result.get("offset")},
            payload_preview={"content_preview": str(result.get("content") or "")[:500], "truncated": result.get("truncated")},
        )]
    return [_node(
        f"{base}-result",
        "compare_diff",
        result.get("title") or tool_name,
        result.get("summary") or "",
        tool_name=tool_name,
        call_index=call_index,
        locator={"target": result.get("target", "")},
        payload_preview=result,
    )]


def _citation_items(raw: dict) -> List[dict]:
    items = raw.get("citations")
    if items is None:
        items = raw.get("evidence")
    result = []
    if isinstance(items, list):
        for item in items[:24]:
            if isinstance(item, dict):
                evidence_id = str(item.get("id") or item.get("evidence_id") or "").strip()
                if evidence_id:
                    result.append({"id": evidence_id, "note": str(item.get("note") or item.get("reason") or "")})
            else:
                evidence_id = str(item or "").strip()
                if evidence_id:
                    result.append({"id": evidence_id, "note": ""})
    return result


def normalize_citations(raw: dict, evidence_nodes: Sequence[dict]) -> Tuple[List[dict], dict]:
    return runtime_normalize_citations(raw, evidence_nodes, fallback_when_empty=False)


def normalize_proposed_actions(raw: dict) -> List[dict]:
    return runtime_normalize_proposed_actions(raw)
