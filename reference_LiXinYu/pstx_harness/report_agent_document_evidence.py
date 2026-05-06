# -*- coding: utf-8 -*-
"""Local document search evidence for the report agent."""

from __future__ import annotations

from typing import List, Optional, Sequence

from pstx_harness.report_agent_observation import preview as _preview


DOCUMENT_EVIDENCE_TOOLS = {
    "list_document_search_sources",
    "search_documents",
    "batch_search_documents",
    "get_document_excerpt",
}


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


def document_evidence_nodes_from_tool_result(tool_name: str,
                                             result: dict,
                                             *,
                                             call_index: int,
                                             args: Optional[dict] = None) -> Optional[List[dict]]:
    """Build document-search evidence nodes for a matching tool result."""
    if tool_name not in DOCUMENT_EVIDENCE_TOOLS:
        return None
    args = args or {}
    nodes: List[dict] = []
    base = f"ev-{call_index}"

    if tool_name == "list_document_search_sources":
        nodes.append(_node(
            f"{base}-document-search-status",
            "document_search_source",
            result.get("title") or "本地文档搜索状态",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"target": "document_search"},
            payload_preview={
                "configured_roots": result.get("configured_roots", []),
                "document_count": result.get("document_count", 0),
                "suffix_counts": result.get("suffix_counts", {}),
            },
        ))
        return nodes

    if tool_name in {"search_documents", "batch_search_documents"}:
        items = (
            [{"query": result.get("query", ""), "matches": result.get("matches", [])}]
            if tool_name == "search_documents" else list(result.get("items") or [])
        )
        for item_index, item in enumerate(items[:24], start=1):
            if not isinstance(item, dict):
                continue
            query = str(item.get("query") or result.get("query") or f"query-{item_index}")
            matches = list(item.get("matches") or [])
            if not matches:
                nodes.append(_node(
                    f"{base}-document-gap-{item_index}-{_safe_evidence_fragment(query)}",
                    "document_gap",
                    f"文档搜索无命中：{query}",
                    item.get("missing_reason") or item.get("summary") or "本地文档未命中。",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"query": query},
                    payload_preview=item,
                    missing_fields=["document_match"],
                ))
                continue
            for match_index, match in enumerate(matches[:8], start=1):
                if not isinstance(match, dict):
                    continue
                doc_id = str(match.get("doc_id") or f"{item_index}-{match_index}")
                nodes.append(_node(
                    f"{base}-document-match-{_safe_evidence_fragment(doc_id)}-{match_index}",
                    "document_match",
                    match.get("title") or f"文档命中 {doc_id}",
                    (
                        f"{match.get('rel_path') or ''} 第 {match.get('line_number') or '?'} 行命中 "
                        f"`{match.get('matched_term') or query}`；{match.get('snippet') or ''}"
                    ),
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={
                        "query": query,
                        "doc_id": doc_id,
                        "rel_path": match.get("rel_path", ""),
                        "line_number": match.get("line_number", ""),
                        "char_start": match.get("char_start", 0),
                    },
                    payload_preview=match,
                    detail_tool={
                        "name": "get_document_excerpt",
                        "args": {"doc_id": doc_id, "char_start": int(match.get("char_start") or 0), "max_chars": 5000},
                    },
                ))
        return nodes or None

    if tool_name == "get_document_excerpt":
        doc_id = str(result.get("doc_id") or args.get("doc_id") or "document")
        nodes.append(_node(
            f"{base}-document-excerpt-{_safe_evidence_fragment(doc_id)}",
            "document_excerpt",
            result.get("title") or f"文档片段 {doc_id}",
            result.get("summary") or _preview(result.get("excerpt", ""), 240),
            tool_name=tool_name,
            call_index=call_index,
            locator={
                "doc_id": doc_id,
                "rel_path": result.get("rel_path", ""),
                "line_number": result.get("line_number", ""),
                "char_start": result.get("char_start", 0),
                "char_end": result.get("char_end", 0),
            },
            payload_preview={
                "rel_path": result.get("rel_path", ""),
                "line_number": result.get("line_number", ""),
                "excerpt": result.get("excerpt", ""),
                "truncated": result.get("truncated", False),
            },
            detail_tool={"name": "get_document_excerpt", "args": {"doc_id": doc_id, "char_start": int(result.get("char_start") or 0), "max_chars": 8000}},
        ))
        return nodes

    return None
