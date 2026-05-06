# -*- coding: utf-8 -*-
"""Feishu material and component identity evidence for the report agent."""

from __future__ import annotations

from typing import List, Optional, Sequence

from pstx_harness.report_agent_observation import preview as _preview


MATERIAL_EVIDENCE_TOOLS = {
    "search_feishu_cache_rows",
    "batch_search_feishu_cache_rows",
    "get_feishu_cache_row",
    "list_feishu_cache_libraries",
    "list_component_identity_cards",
    "search_component_identity_cards",
    "batch_get_component_identity_cards",
    "get_component_identity_card",
    "summarize_dfmea_readiness",
}


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


def _feishu_material_summary(row: dict) -> str:
    parts = []
    for key, label in [
        ("hq_no", "HQ"),
        ("key_value", "规格/关键值"),
        ("spec", "规格"),
        ("pi", "PI"),
        ("selection_order", "选型顺序"),
        ("lib_name", "库"),
        ("sheet_name", "Sheet"),
    ]:
        value = row.get(key)
        if value not in {None, ""}:
            parts.append(f"{label}={_preview(value, 100)}")
    return "；".join(parts) or _row_summary(row)


def _component_identity_summary(card: dict) -> str:
    parts = []
    for key, label in [
        ("refdes", "位号"),
        ("category", "类别"),
        ("candidate_chip_type", "候选类型"),
        ("hq_no", "HQ"),
        ("spec", "规格"),
        ("pi", "PI"),
        ("selection_order", "选型顺序"),
        ("user_visible_page", "页码"),
    ]:
        value = card.get(key)
        if value not in {None, ""}:
            parts.append(f"{label}={_preview(value, 100)}")
    missing = list(card.get("missing_fields") or [])
    if missing:
        parts.append(f"缺失={','.join(str(item) for item in missing[:8])}")
    return "；".join(parts) or _row_summary(card)


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


def _append_material_match_node(nodes: List[dict],
                                *,
                                base: str,
                                card: dict,
                                tool_name: str,
                                call_index: int) -> None:
    match = card.get("feishu_match") if isinstance(card.get("feishu_match"), dict) else {}
    if match.get("status") != "matched":
        return
    refdes = str(card.get("refdes") or "component")
    row_id = str(match.get("row_id") or refdes)
    detail_tool = None
    try:
        row_id_int = int(match.get("row_id") or 0)
    except (TypeError, ValueError):
        row_id_int = 0
    if row_id_int > 0:
        detail_tool = {"name": "get_feishu_cache_row", "args": {"row_id": row_id_int}}
    nodes.append(_node(
        f"{base}-material-{_safe_evidence_fragment(refdes)}-{_safe_evidence_fragment(row_id)}",
        "material_match",
        f"{refdes} 飞书物料匹配",
        _feishu_material_summary({
            "hq_no": match.get("hq_no") or card.get("hq_no"),
            "spec": match.get("spec") or card.get("spec"),
            "pi": match.get("pi") or card.get("pi"),
            "selection_order": match.get("selection_order") or card.get("selection_order"),
            "lib_name": match.get("lib_name", ""),
            "sheet_name": match.get("sheet_name", ""),
        }),
        tool_name=tool_name,
        call_index=call_index,
        locator={
            "refdes": refdes,
            "row_id": match.get("row_id", ""),
            "hq_no": match.get("hq_no") or card.get("hq_no", ""),
        },
        payload_preview=match,
        detail_tool=detail_tool,
    ))


def material_evidence_nodes_from_tool_result(tool_name: str,
                                             result: dict,
                                             *,
                                             call_index: int,
                                             args: Optional[dict] = None) -> Optional[List[dict]]:
    """Build material/identity evidence nodes for a matching tool result.

    Returns ``None`` when the tool is not handled here, or when the previous
    monolithic dispatcher would have fallen through to the generic pack node.
    """
    if tool_name not in MATERIAL_EVIDENCE_TOOLS:
        return None
    args = args or {}
    nodes: List[dict] = []
    base = f"ev-{call_index}"

    if tool_name == "search_feishu_cache_rows":
        offset = int(result.get("offset") or 0)
        for row_index, row in enumerate(list(result.get("rows") or [])[:24], start=1):
            if not isinstance(row, dict):
                continue
            row_id = str(row.get("id") or offset + row_index)
            nodes.append(_node(
                f"{base}-feishu-row-{row_id}",
                "feishu_material",
                row.get("hq_no") or row.get("key_value") or f"飞书物料 {row_id}",
                _feishu_material_summary(row),
                tool_name=tool_name,
                call_index=call_index,
                locator={
                    "row_id": row.get("id", ""),
                    "lib_id": row.get("lib_id", ""),
                    "lib_name": row.get("lib_name", ""),
                    "sheet_name": row.get("sheet_name", ""),
                    "query": result.get("query", ""),
                },
                payload_preview=row,
            ))
        return nodes or None

    if tool_name == "batch_search_feishu_cache_rows":
        for item_index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            query = str(item.get("query") or f"query-{item_index}")
            rows = list(item.get("rows") or [])
            if not rows:
                nodes.append(_node(
                    f"{base}-batch-feishu-{item_index}",
                    "missing_context" if item.get("status") == "missing" else "rule_result",
                    f"飞书缓存搜索 {query}",
                    item.get("summary") or item.get("missing_reason") or "",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"query": query, "status": item.get("status", "")},
                    payload_preview=item,
                    missing_fields=[item.get("missing_reason")] if item.get("missing_reason") else [],
                ))
                continue
            for row_index, row in enumerate(rows[:8], start=1):
                if not isinstance(row, dict):
                    continue
                row_id = str(row.get("id") or f"{item_index}-{row_index}")
                nodes.append(_node(
                    f"{base}-batch-feishu-{item_index}-{_safe_evidence_fragment(row_id)}",
                    "feishu_material",
                    row.get("hq_no") or row.get("key_value") or f"飞书物料 {row_id}",
                    _feishu_material_summary(row),
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={
                        "query": query,
                        "row_id": row.get("id", ""),
                        "lib_id": row.get("lib_id", ""),
                        "sheet_name": row.get("sheet_name", ""),
                    },
                    payload_preview=row,
                    detail_tool={"name": "get_feishu_cache_row", "args": {"row_id": int(row.get("id") or 0)}} if int(row.get("id") or 0) else None,
                ))
        return nodes or None

    if tool_name == "get_feishu_cache_row":
        row = result.get("row")
        if isinstance(row, dict):
            row_id = str(row.get("id") or result.get("row_id") or "1")
            nodes.append(_node(
                f"{base}-feishu-row-{row_id}",
                "feishu_material",
                row.get("hq_no") or row.get("key_value") or f"飞书物料 {row_id}",
                _feishu_material_summary(row),
                tool_name=tool_name,
                call_index=call_index,
                locator={
                    "row_id": row.get("id", ""),
                    "lib_id": row.get("lib_id", ""),
                    "lib_name": row.get("lib_name", ""),
                    "sheet_name": row.get("sheet_name", ""),
                },
                payload_preview=row,
            ))
            return nodes
        return None

    if tool_name == "list_feishu_cache_libraries":
        for index, library in enumerate(list(result.get("libraries") or [])[:24], start=1):
            if not isinstance(library, dict):
                continue
            nodes.append(_node(
                f"{base}-feishu-lib-{index}",
                "rule_result",
                library.get("lib_name") or library.get("lib_id") or f"飞书库 {index}",
                f"飞书缓存库 {library.get('lib_name', '')}，缓存 {library.get('cache_count', 0)} 行。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"lib_id": library.get("lib_id", ""), "lib_name": library.get("lib_name", "")},
                payload_preview=library,
            ))
        return nodes or None

    if tool_name in {"list_component_identity_cards", "search_component_identity_cards"}:
        offset = int(result.get("offset") or 0)
        for card_index, card in enumerate(list(result.get("cards") or [])[:24], start=1):
            if not isinstance(card, dict):
                continue
            refdes = str(card.get("refdes") or offset + card_index)
            nodes.append(_node(
                f"{base}-identity-{_safe_evidence_fragment(refdes)}",
                "component_identity",
                f"{refdes} 身份卡",
                _component_identity_summary(card),
                tool_name=tool_name,
                call_index=call_index,
                locator={
                    "refdes": card.get("refdes", ""),
                    "category": card.get("category", ""),
                    "query": result.get("query", ""),
                },
                payload_preview=card,
                missing_fields=card.get("missing_fields") or [],
                detail_tool={"name": "get_component_identity_card", "args": {"refdes": refdes}},
            ))
            _append_material_match_node(nodes, base=base, card=card, tool_name=tool_name, call_index=call_index)
        return nodes or None

    if tool_name == "batch_get_component_identity_cards":
        for item_index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            card = item.get("card") if isinstance(item.get("card"), dict) else {}
            refdes = str(item.get("refdes") or card.get("refdes") or f"component-{item_index}")
            evidence_type = "component_identity" if card else "missing_context"
            nodes.append(_node(
                f"{base}-batch-identity-{_safe_evidence_fragment(refdes)}",
                evidence_type,
                f"{refdes} 身份卡",
                _component_identity_summary(card) if card else item.get("summary", ""),
                tool_name=tool_name,
                call_index=call_index,
                locator={"refdes": refdes, "status": item.get("status", "")},
                payload_preview=item,
                missing_fields=item.get("missing_fields") or ([item.get("missing_reason")] if item.get("missing_reason") else []),
                detail_tool={"name": "get_component_identity_card", "args": {"refdes": refdes}} if card else None,
            ))
            if card:
                _append_material_match_node(nodes, base=base, card=card, tool_name=tool_name, call_index=call_index)
        return nodes or None

    if tool_name == "get_component_identity_card":
        card = result.get("card")
        if isinstance(card, dict):
            refdes = str(card.get("refdes") or args.get("refdes") or "component")
            nodes.append(_node(
                f"{base}-identity-{_safe_evidence_fragment(refdes)}",
                "component_identity",
                f"{refdes} 身份卡",
                _component_identity_summary(card),
                tool_name=tool_name,
                call_index=call_index,
                locator={"refdes": refdes, "category": card.get("category", "")},
                payload_preview=card,
                missing_fields=card.get("missing_fields") or [],
                detail_tool={"name": "get_component_identity_card", "args": {"refdes": refdes}},
            ))
            _append_material_match_node(nodes, base=base, card=card, tool_name=tool_name, call_index=call_index)
            return nodes
        return None

    if tool_name == "summarize_dfmea_readiness":
        nodes.append(_node(
            f"{base}-dfmea-readiness",
            "dfmea_readiness",
            result.get("title") or "DFMEA 准备度摘要",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"target": "dfmea", "total_components": result.get("total_components", 0)},
            payload_preview={
                "total_components": result.get("total_components", 0),
                "category_counts": result.get("category_counts", {}),
                "missing_counts": result.get("missing_counts", {}),
                "ready_count": result.get("ready_count", 0),
                "needs_context_count": result.get("needs_context_count", 0),
            },
            detail_tool={"name": "summarize_dfmea_readiness", "args": {}},
        ))
        for index, card in enumerate(list(result.get("ready_cards") or [])[:12], start=1):
            if not isinstance(card, dict):
                continue
            refdes = str(card.get("refdes") or f"ready-{index}")
            nodes.append(_node(
                f"{base}-ready-{_safe_evidence_fragment(refdes)}",
                "component_identity",
                f"{refdes} DFMEA 输入已就绪",
                _component_identity_summary(card),
                tool_name=tool_name,
                call_index=call_index,
                locator={"refdes": refdes, "category": card.get("category", ""), "readiness": "ready"},
                payload_preview=card,
                detail_tool={"name": "get_component_identity_card", "args": {"refdes": refdes}},
            ))
            _append_material_match_node(nodes, base=base, card=card, tool_name=tool_name, call_index=call_index)
        for index, card in enumerate(list(result.get("needs_context_cards") or [])[:12], start=1):
            if not isinstance(card, dict):
                continue
            refdes = str(card.get("refdes") or f"needs-context-{index}")
            nodes.append(_node(
                f"{base}-missing-{_safe_evidence_fragment(refdes)}",
                "missing_context",
                f"{refdes} 需补充 DFMEA 上下文",
                _component_identity_summary(card),
                tool_name=tool_name,
                call_index=call_index,
                locator={"refdes": refdes, "category": card.get("category", ""), "readiness": "needs_context"},
                payload_preview=card,
                missing_fields=card.get("missing_fields") or [],
                detail_tool={"name": "get_component_identity_card", "args": {"refdes": refdes}},
            ))
            _append_material_match_node(nodes, base=base, card=card, tool_name=tool_name, call_index=call_index)
        return nodes

    return None
