# -*- coding: utf-8 -*-
"""Evidence node builders for the report harness agent."""

from __future__ import annotations

from typing import List, Optional, Sequence

from pstx_harness.report_agent_datasheet_evidence import datasheet_evidence_nodes_from_tool_result
from pstx_harness.report_agent_document_evidence import document_evidence_nodes_from_tool_result
from pstx_harness.report_agent_material_evidence import material_evidence_nodes_from_tool_result
from pstx_harness.report_agent_observation import preview as _preview
from pstx_harness.report_agent_reference_evidence import reference_evidence_nodes_from_tool_result
from pstx_harness.report_agent_topology_evidence import topology_evidence_nodes_from_tool_result


def _row_summary(row: dict) -> str:
    priority_keys = [
        "位号", "refdes", "网络", "net", "真实页", "page", "问题", "kind", "结论", "status",
        "hq_no", "key_value", "spec", "pi", "selection_order",
    ]
    parts = []
    for key in priority_keys:
        if key in row and row.get(key) not in {None, ""}:
            parts.append(f"{key}={_preview(row.get(key), 80)}")
    if not parts:
        for key, value in list(row.items())[:3]:
            parts.append(f"{key}={_preview(value, 80)}")
    return "；".join(parts)


def _evidence_type_for_table(table_id: str) -> str:
    table_id = str(table_id or "")
    if "page" in table_id:
        return "page"
    if "chip" in table_id or "component" in table_id or "bom" in table_id:
        return "component"
    if "net" in table_id:
        return "net"
    return "table_row"


def _safe_evidence_fragment(value: str) -> str:
    fragment = "".join(char if char.isalnum() else "-" for char in str(value or "").strip())
    fragment = "-".join(part for part in fragment.split("-") if part)
    return fragment[:80] or "item"


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


def evidence_nodes_from_tool_result(tool_name: str,
                                     result: dict,
                                     *,
                                     call_index: int,
                                     args: Optional[dict] = None) -> List[dict]:
    args = args or {}
    nodes: List[dict] = []
    base = f"ev-{call_index}"
    if tool_name == "get_table_rows":
        table_id = str(result.get("table_id") or args.get("table_id") or "")
        offset = int(result.get("offset") or 0)
        rows = list(result.get("rows") or [])[:24]
        for row_index, row in enumerate(rows, start=1):
            if not isinstance(row, dict):
                continue
            row_number = offset + row_index
            nodes.append(_node(
                f"{base}-row-{row_number}",
                _evidence_type_for_table(table_id),
                f"{result.get('title') or table_id} #{row_number}",
                _row_summary(row),
                tool_name=tool_name,
                call_index=call_index,
                locator={"table_id": table_id, "row_index": row_number - 1, "row_number": row_number},
                payload_preview=row,
            ))
        if nodes:
            return nodes
    if tool_name == "summarize_table_column_values":
        table_id = str(result.get("table_id") or args.get("table_id") or "")
        column = str(result.get("column") or args.get("column") or "")
        nodes.append(_node(
            f"{base}-table-aggregate-{_safe_evidence_fragment(table_id)}-{_safe_evidence_fragment(column)}",
            "table_column_summary",
            result.get("title") or f"{table_id} / {column} 列聚合",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"table_id": table_id, "column": column},
            payload_preview={
                "total_rows": result.get("total_rows"),
                "non_empty_count": result.get("non_empty_count"),
                "empty_count": result.get("empty_count"),
                "unique_count": result.get("unique_count"),
                "values": result.get("values", []),
                "top_values": result.get("top_values", []),
                "truncated": result.get("truncated", False),
            },
            detail_tool={
                "name": "summarize_table_column_values",
                "args": {
                    "table_id": table_id,
                    "column": column,
                    "limit_values": min(int(result.get("unique_count") or 200), 1000),
                },
            },
        ))
        return nodes
    if tool_name == "summarize_schematic_page_count":
        nodes.append(_node(
            f"{base}-schematic-page-count",
            "schematic_page_count",
            result.get("title") or "原理图总页数",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"source": "module_order", "last_page": result.get("last_page", "")},
            payload_preview={
                "available": result.get("available", False),
                "total_pages": result.get("total_pages", 0),
                "last_page": result.get("last_page", ""),
                "last_entry": result.get("last_entry", {}),
                "scope_note": result.get("scope_note", ""),
            },
            detail_tool={"name": "summarize_schematic_page_count", "args": {}},
        ))
        return nodes
    if tool_name in {"list_project_memory_evidence", "get_project_memory_evidence", "batch_get_project_memory_evidence"}:
        def append_memory_card(card: dict, index: int, *, missing: bool = False, summary: str = "") -> None:
            evidence_id = str(card.get("id") or args.get("evidence_id") or f"memory-{index}")
            fragment = _safe_evidence_fragment(evidence_id)
            original_type = str(card.get("type") or "").strip()
            node_type = "missing_context" if missing else (original_type or "project_memory_evidence")
            title = card.get("title") or evidence_id
            original_detail_tool = card.get("detail_tool") if isinstance(card.get("detail_tool"), dict) else {}
            detail_tool = (
                dict(original_detail_tool)
                if original_detail_tool
                else {"name": "get_project_memory_evidence", "args": {"evidence_id": evidence_id}}
            )
            nodes.append(_node(
                f"{base}-memory-{fragment}",
                node_type,
                str(title),
                summary or str(card.get("summary") or f"项目证据记忆：{evidence_id}"),
                tool_name=tool_name,
                call_index=call_index,
                locator={
                    "evidence_id": evidence_id,
                    "original_type": card.get("type", ""),
                    "agent_run_id": card.get("agent_run_id", ""),
                    "refdes": (card.get("locator") or {}).get("refdes", "") if isinstance(card.get("locator"), dict) else "",
                },
                payload_preview=card,
                missing_fields=["evidence_memory_card"] if missing else [],
                detail_tool=detail_tool,
            ))

        if tool_name == "list_project_memory_evidence":
            for index, card in enumerate(list(result.get("cards") or [])[:24], start=1):
                if isinstance(card, dict):
                    append_memory_card(card, index)
            if not nodes:
                nodes.append(_node(
                    f"{base}-memory-gap",
                    "missing_context",
                    "项目证据记忆无命中",
                    result.get("summary") or "项目证据记忆没有命中本次查询。",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"query": result.get("query", ""), "evidence_type": result.get("evidence_type", "")},
                    payload_preview=result,
                    missing_fields=["project_memory_evidence"],
                ))
            return nodes
        if tool_name == "get_project_memory_evidence":
            card = result.get("card") if isinstance(result.get("card"), dict) else {}
            if card and result.get("found"):
                append_memory_card(card, 1)
            else:
                append_memory_card(
                    {"id": result.get("evidence_id") or args.get("evidence_id") or "missing"},
                    1,
                    missing=True,
                    summary=result.get("summary") or "未找到该项目证据记忆。",
                )
            return nodes
        for index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            card = item.get("card") if isinstance(item.get("card"), dict) else {}
            if card and item.get("status") == "found":
                append_memory_card(card, index)
            else:
                append_memory_card(
                    {"id": item.get("evidence_id") or f"memory-{index}"},
                    index,
                    missing=True,
                    summary=item.get("summary") or "未找到该项目证据记忆。",
                )
        if nodes:
            return nodes
    if tool_name == "list_report_tables":
        for index, table in enumerate(list(result.get("tables") or [])[:24], start=1):
            table_id = str(table.get("table_id") or "")
            nodes.append(_node(
                f"{base}-table-{index}",
                "table_row",
                table.get("title") or table_id or f"表格 {index}",
                f"表格 {table_id}，记录 {table.get('count', 0)} 行。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"table_id": table_id, "section_id": table.get("section_id", "")},
                payload_preview=table,
            ))
        if nodes:
            return nodes
    if tool_name == "query_report_entity":
        mode = str(args.get("mode") or "")
        keyword = str(args.get("keyword") or "")
        nodes.append(_node(
            f"{base}-entity",
            "net" if "网络" in mode else "component",
            result.get("title") or keyword or "查询结果",
            result.get("summary") or f"{mode} {keyword} 查询结果。",
            tool_name=tool_name,
            call_index=call_index,
            locator={"mode": mode, "keyword": keyword},
            payload_preview=result.get("query_result") or {},
        ))
        return nodes
    if tool_name == "batch_query_report_entities":
        for item_index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            query = str(item.get("query") or f"query-{item_index}")
            status = str(item.get("status") or "")
            matches = list(item.get("matches") or [])
            if not matches:
                nodes.append(_node(
                    f"{base}-batch-query-{item_index}",
                    "missing_context" if status in {"missing", "needs_context"} else "rule_result",
                    f"批量查询 {query}",
                    item.get("summary") or item.get("missing_reason") or "",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"query": query, "status": status, "mode": item.get("mode", "")},
                    payload_preview=item,
                    missing_fields=[item.get("missing_reason")] if item.get("missing_reason") else [],
                ))
                continue
            for match_index, match in enumerate(matches[:6], start=1):
                kind = str(match.get("kind") or "")
                table_id = str(match.get("table_id") or "")
                evidence_type = _evidence_type_for_table(table_id) if table_id else (
                    "net" if kind == "network_query" else "component"
                )
                title = match.get("table_title") or match.get("title") or query
                nodes.append(_node(
                    f"{base}-batch-query-{item_index}-{match_index}",
                    evidence_type,
                    title,
                    match.get("summary") if isinstance(match.get("summary"), str) else item.get("summary", ""),
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={
                        "query": query,
                        "kind": kind,
                        "table_id": table_id,
                        "row_index": match.get("row_index", ""),
                        "mode": match.get("mode", item.get("mode", "")),
                    },
                    payload_preview=match,
                ))
        if nodes:
            return nodes
    if tool_name == "batch_get_table_rows":
        for item_index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            table_id = str(item.get("table_id") or f"table-{item_index}")
            offset = int(item.get("offset") or 0)
            rows = list(item.get("rows") or [])
            if not rows:
                nodes.append(_node(
                    f"{base}-batch-table-{item_index}",
                    "table_row",
                    table_id,
                    item.get("summary") or item.get("missing_reason") or "",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"table_id": table_id, "status": item.get("status", "")},
                    payload_preview=item,
                    missing_fields=[item.get("missing_reason")] if item.get("missing_reason") else [],
                ))
                continue
            for row_index, row in enumerate(rows[:8], start=1):
                if not isinstance(row, dict):
                    continue
                row_number = offset + row_index
                nodes.append(_node(
                    f"{base}-batch-table-{item_index}-row-{row_number}",
                    _evidence_type_for_table(table_id),
                    f"{table_id} #{row_number}",
                    _row_summary(row),
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"table_id": table_id, "row_index": row_number - 1, "row_number": row_number},
                    payload_preview=row,
                ))
        if nodes:
            return nodes
    material_nodes = material_evidence_nodes_from_tool_result(
        tool_name,
        result,
        call_index=call_index,
        args=args,
    )
    if material_nodes is not None:
        return material_nodes
    topology_nodes = topology_evidence_nodes_from_tool_result(
        tool_name,
        result,
        call_index=call_index,
        args=args,
    )
    if topology_nodes is not None:
        return topology_nodes
    document_nodes = document_evidence_nodes_from_tool_result(
        tool_name,
        result,
        call_index=call_index,
        args=args,
    )
    if document_nodes is not None:
        return document_nodes
    datasheet_nodes = datasheet_evidence_nodes_from_tool_result(
        tool_name,
        result,
        call_index=call_index,
        args=args,
    )
    if datasheet_nodes is not None:
        return datasheet_nodes
    reference_nodes = reference_evidence_nodes_from_tool_result(
        tool_name,
        result,
        call_index=call_index,
        args=args,
    )
    if reference_nodes is not None:
        return reference_nodes
    if tool_name == "list_project_files":
        for index, item in enumerate(list(result.get("files") or [])[:24], start=1):
            nodes.append(_node(
                f"{base}-file-{index}",
                "file_excerpt",
                item.get("path") or item.get("name") or f"文件 {index}",
                f"允许读取文件，大小 {item.get('size', 0)} bytes。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"path": item.get("path", "")},
                payload_preview=item,
            ))
        if nodes:
            return nodes
    if tool_name in {"trace_project_source", "search_project_text"}:
        hits = [item for item in list(result.get("source_hits") or [])[:24] if isinstance(item, dict)]
        for index, hit in enumerate(hits, start=1):
            path = str(hit.get("path") or "")
            line_start = hit.get("line_start")
            line_end = hit.get("line_end")
            matched_terms = list(hit.get("matched_terms") or result.get("query_terms") or [])[:8]
            excerpt_lines = list(hit.get("excerpt") or [])[:8]
            nodes.append(_node(
                f"{base}-source-{index}",
                "source_trace",
                f"{path}:{line_start or 1}",
                (
                    f"原始文件片段 {path}:{line_start}-{line_end}"
                    + (f"，匹配 {'/'.join(str(item) for item in matched_terms)}。" if matched_terms else "。")
                ),
                tool_name=tool_name,
                call_index=call_index,
                locator={
                    "path": path,
                    "line_start": line_start,
                    "line_end": line_end,
                    "match_lines": list(hit.get("match_lines") or []),
                    "query_terms": list(result.get("query_terms") or []),
                    "derived_from": result.get("derived_from") or {},
                },
                payload_preview={
                    "matched_terms": matched_terms,
                    "excerpt": excerpt_lines,
                    "warnings": result.get("warnings") or [],
                },
                detail_tool=hit.get("detail_tool") if isinstance(hit.get("detail_tool"), dict) else None,
            ))
        if nodes:
            return nodes
        nodes.append(_node(
            f"{base}-source-missing",
            "missing_context",
            "原始文件搜索无命中" if tool_name == "search_project_text" else "原始文件追溯无命中",
            result.get("summary") or "未在允许的项目原始文件中找到匹配片段。",
            tool_name=tool_name,
            call_index=call_index,
            locator={
                "query": result.get("query", ""),
                "query_terms": result.get("query_terms") or [],
                "candidate_files": result.get("candidate_files") or [],
                "path_prefix": result.get("path_prefix", ""),
                "file_glob": result.get("file_glob", ""),
            },
            payload_preview=result,
            missing_fields=["source_trace"],
            detail_tool={"name": "list_project_files", "args": {}},
        ))
        return nodes
    if tool_name == "read_project_text":
        nodes.append(_node(
            f"{base}-excerpt",
            "file_excerpt",
            result.get("title") or result.get("path") or "文件片段",
            result.get("summary") or "项目文本文件片段。",
            tool_name=tool_name,
            call_index=call_index,
            locator={
                "path": result.get("path", ""),
                "encoding": result.get("encoding", ""),
                "line_start": result.get("line_start"),
                "line_end": result.get("line_end"),
                "query": result.get("query", ""),
            },
            payload_preview={
                "content_preview": str(result.get("content") or "")[:500],
                "excerpts": list(result.get("excerpts") or [])[:2],
                "truncated": result.get("truncated"),
            },
        ))
        return nodes

    pack_id = str(result.get("id") or tool_name)
    nodes.append(_node(
        f"{base}-pack",
        "rule_result",
        result.get("title") or pack_id,
        result.get("summary") or "",
        tool_name=tool_name,
        call_index=call_index,
        locator={"pack_id": pack_id, "target": result.get("target", "")},
        payload_preview={
            "issue_count": result.get("issue_count"),
            "severity": result.get("severity"),
            "metrics": result.get("metrics", []),
            "notes": result.get("notes", []),
        },
    ))
    for index, table in enumerate(list(result.get("tables") or [])[:24], start=1):
        table_id = str(table.get("id") or table.get("table_id") or "")
        nodes.append(_node(
            f"{base}-table-{index}",
            _evidence_type_for_table(table_id),
            table.get("title") or table_id or f"{pack_id} 表格 {index}",
            f"表格 {table_id}，记录 {table.get('count', 0)} 行。",
            tool_name=tool_name,
            call_index=call_index,
            locator={"pack_id": pack_id, "table_id": table_id},
            payload_preview={
                "count": table.get("count"),
                "kind_counts": table.get("kind_counts", {}),
                "sample_rows": table.get("sample_rows", []),
            },
        ))
    return nodes
