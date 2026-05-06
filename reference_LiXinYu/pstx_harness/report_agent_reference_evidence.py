# -*- coding: utf-8 -*-
"""Reference document evidence node builders for the report harness agent."""

from __future__ import annotations

from typing import List, Optional, Sequence

from pstx_harness.report_agent_observation import preview as _preview


REFERENCE_EVIDENCE_TOOLS = {
    "list_agent_ref_sources",
    "search_agent_ref_pdfs",
    "get_agent_ref_pdf_excerpt",
    "list_review_checklist_sources",
    "search_review_checklists",
    "get_review_checklist_excerpt",
}


def _row_summary(row: dict) -> str:
    parts = []
    for key in ["title", "page", "section_title", "score", "snippet", "rel_path", "query"]:
        value = row.get(key)
        if value not in {None, ""}:
            parts.append(f"{key}={_preview(value, 100)}")
    return "；".join(parts)


def _match_summary(match: dict) -> str:
    parts = []
    for key, label in [
        ("title", "文档"),
        ("page", "页"),
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


def reference_evidence_nodes_from_tool_result(tool_name: str,
                                             result: dict,
                                             *,
                                             call_index: int,
                                             args: Optional[dict] = None) -> Optional[List[dict]]:
    if tool_name not in REFERENCE_EVIDENCE_TOOLS:
        return None

    args = args or {}
    nodes: List[dict] = []
    base = f"ev-{call_index}"
    if tool_name == "list_agent_ref_sources":
        nodes.append(_node(
            f"{base}-agent-ref-status",
            "agent_ref_index",
            result.get("title") or "Agent Lab ref PDF 索引状态",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"target": "agent_ref", "ref_dir": result.get("ref_dir", ""), "db_path": result.get("db_path", "")},
            payload_preview={
                "pdf_count": result.get("pdf_count", 0),
                "document_count": result.get("document_count", 0),
                "indexed_count": result.get("indexed_count", 0),
                "page_count": result.get("page_count", 0),
                "failed_count": result.get("failed_count", 0),
            },
            detail_tool={"name": "list_agent_ref_sources", "args": {}},
        ))
        return nodes
    if tool_name == "search_agent_ref_pdfs":
        for index, match in enumerate(list(result.get("matches") or [])[:24], start=1):
            if not isinstance(match, dict):
                continue
            doc_id = int(match.get("doc_id") or 0)
            page = int(match.get("page") or 1)
            nodes.append(_node(
                f"{base}-agent-ref-{doc_id}-{page}-{index}",
                "agent_ref_excerpt",
                f"{match.get('title') or 'ref PDF'} 第 {page} 页",
                _match_summary(match),
                tool_name=tool_name,
                call_index=call_index,
                locator={"doc_id": doc_id, "page": page, "query": result.get("query", ""), "rel_path": match.get("rel_path", "")},
                payload_preview=match,
                detail_tool={"name": "get_agent_ref_pdf_excerpt", "args": {"doc_id": doc_id, "page": page, "max_chars": 2400}} if doc_id else None,
            ))
        return nodes if nodes else None
    if tool_name == "get_agent_ref_pdf_excerpt":
        doc_id = int(result.get("doc_id") or args.get("doc_id") or 0)
        page = int(result.get("page") or args.get("page") or 1)
        nodes.append(_node(
            f"{base}-agent-ref-excerpt-{doc_id}-{page}",
            "agent_ref_excerpt",
            f"{result.get('title') or 'ref PDF'} 第 {page} 页",
            result.get("summary") or "Agent Lab ref PDF 页级片段。",
            tool_name=tool_name,
            call_index=call_index,
            locator={"doc_id": doc_id, "page": page, "rel_path": result.get("rel_path", ""), "path": result.get("path", "")},
            payload_preview={"content_preview": str(result.get("content") or "")[:600], "truncated": result.get("truncated")},
            detail_tool={"name": "get_agent_ref_pdf_excerpt", "args": {"doc_id": doc_id, "page": page, "max_chars": 2400}} if doc_id else None,
        ))
        return nodes
    if tool_name == "list_review_checklist_sources":
        nodes.append(_node(
            f"{base}-review-checklist-status",
            "review_checklist_index",
            result.get("title") or "Review checklist 索引状态",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"target": "review_checklist", "ref_dir": result.get("ref_dir", ""), "db_path": result.get("db_path", "")},
            payload_preview={
                "file_count": result.get("file_count", 0),
                "document_count": result.get("document_count", 0),
                "indexed_count": result.get("indexed_count", 0),
                "page_count": result.get("page_count", 0),
                "failed_count": result.get("failed_count", 0),
            },
            detail_tool={"name": "list_review_checklist_sources", "args": {}},
        ))
        return nodes
    if tool_name == "search_review_checklists":
        for index, match in enumerate(list(result.get("matches") or [])[:24], start=1):
            if not isinstance(match, dict):
                continue
            doc_id = int(match.get("doc_id") or 0)
            page = int(match.get("page") or 1)
            nodes.append(_node(
                f"{base}-review-checklist-{doc_id}-{page}-{index}",
                "review_checklist_excerpt",
                f"{match.get('title') or 'review checklist'} 片段 {page}",
                _match_summary(match),
                tool_name=tool_name,
                call_index=call_index,
                locator={"doc_id": doc_id, "page": page, "query": result.get("query", ""), "rel_path": match.get("rel_path", "")},
                payload_preview=match,
                detail_tool={"name": "get_review_checklist_excerpt", "args": {"doc_id": doc_id, "page": page, "max_chars": 2400}} if doc_id else None,
            ))
        return nodes if nodes else None
    if tool_name == "get_review_checklist_excerpt":
        doc_id = int(result.get("doc_id") or args.get("doc_id") or 0)
        page = int(result.get("page") or args.get("page") or 1)
        nodes.append(_node(
            f"{base}-review-checklist-excerpt-{doc_id}-{page}",
            "review_checklist_excerpt",
            f"{result.get('title') or 'review checklist'} 片段 {page}",
            result.get("summary") or "Review checklist 审查经验片段。",
            tool_name=tool_name,
            call_index=call_index,
            locator={"doc_id": doc_id, "page": page, "rel_path": result.get("rel_path", ""), "path": result.get("path", "")},
            payload_preview={"content_preview": str(result.get("content") or "")[:600], "truncated": result.get("truncated")},
            detail_tool={"name": "get_review_checklist_excerpt", "args": {"doc_id": doc_id, "page": page, "max_chars": 2400}} if doc_id else None,
        ))
        return nodes
    return None
