# -*- coding: utf-8 -*-
"""Chip-level topology evidence for the report agent."""

from __future__ import annotations

from typing import List, Optional, Sequence

from pstx_harness.report_agent_observation import preview as _preview


TOPOLOGY_EVIDENCE_TOOLS = {
    "summarize_llm_topology_netlist",
    "query_llm_topology_netlist",
    "batch_query_llm_topology_netlist",
    "get_llm_topology_node",
    "get_llm_topology_edge",
    "summarize_topology_review_tasks",
    "get_topology_review_task",
    "batch_expand_topology_review_tasks",
    "summarize_chip_topology",
    "query_chip_topology",
    "batch_query_chip_topology",
    "get_chip_topology_edge",
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


def topology_evidence_nodes_from_tool_result(tool_name: str,
                                             result: dict,
                                             *,
                                             call_index: int,
                                             args: Optional[dict] = None) -> Optional[List[dict]]:
    """Build chip topology evidence nodes for a matching tool result."""
    if tool_name not in TOPOLOGY_EVIDENCE_TOOLS:
        return None
    args = args or {}
    nodes: List[dict] = []
    base = f"ev-{call_index}"
    is_llm = tool_name.startswith(("summarize_llm_", "query_llm_", "batch_query_llm_", "get_llm_"))
    summary_type = "llm_topology_summary" if is_llm else "chip_topology_summary"
    edge_type = "llm_topology_edge" if is_llm else "chip_topology_edge"
    node_type = "llm_topology_node" if is_llm else "chip_topology_node"
    gap_missing_field = edge_type
    summarize_tool = "summarize_llm_topology_netlist" if is_llm else "summarize_chip_topology"
    query_tool = "query_llm_topology_netlist" if is_llm else "query_chip_topology"
    edge_detail_tool = "get_llm_topology_edge" if is_llm else "get_chip_topology_edge"

    def _review_task_node(task: dict, *, suffix: str = "") -> dict:
        task_id = str(task.get("task_id") or task.get("id") or "task")
        missing = list(task.get("missing_signals") or [])
        return _node(
            f"{base}-llm-review-task-{_safe_evidence_fragment(task_id)}{suffix}",
            "llm_topology_review_task",
            task.get("title") or task_id,
            task.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={
                "task_id": task_id,
                "source_kind": task.get("source_kind", ""),
                "source_id": task.get("source_id", ""),
                "refdes": list(task.get("refdes") or [])[:4],
                "pages": list(task.get("pages") or [])[:6],
                "review_priority": task.get("review_priority", ""),
            },
            payload_preview=task,
            missing_fields=missing,
            detail_tool=task.get("detail_tool") or {"name": "get_topology_review_task", "args": {"task_id": task_id}},
        )

    if tool_name == "summarize_topology_review_tasks":
        nodes.append(_node(
            f"{base}-llm-topology-review-task-summary",
            "llm_topology_review_task_summary",
            result.get("title") or "拓扑 review 任务队列",
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={
                "target": "topology",
                "total_count": result.get("total_count", 0),
                "returned_count": result.get("returned_count", 0),
                "truncated": result.get("truncated", False),
            },
            payload_preview={
                "filters": result.get("filters", {}),
                "total_count": result.get("total_count", 0),
                "returned_count": result.get("returned_count", 0),
                "tasks": list(result.get("tasks") or [])[:6],
            },
            detail_tool={"name": "summarize_topology_review_tasks", "args": {"limit": 100}},
        ))
        for task in list(result.get("tasks") or [])[:30]:
            if isinstance(task, dict):
                nodes.append(_review_task_node(task))
        return nodes

    if tool_name == "get_topology_review_task":
        task = result.get("task")
        if isinstance(task, dict):
            return [_review_task_node(task, suffix="-detail")]
        return None

    if tool_name == "batch_expand_topology_review_tasks":
        for item_index, item in enumerate(list(result.get("items") or [])[:24], start=1):
            if not isinstance(item, dict):
                continue
            task = item.get("task") if isinstance(item.get("task"), dict) else {}
            if task:
                nodes.append(_review_task_node(task, suffix=f"-{item_index}"))
            else:
                task_id = str(item.get("task_id") or f"task-{item_index}")
                nodes.append(_node(
                    f"{base}-llm-review-task-gap-{_safe_evidence_fragment(task_id)}",
                    "missing_context",
                    f"拓扑 review task 无命中：{task_id}",
                    item.get("missing_reason") or item.get("summary") or "未找到该拓扑 review task。",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"task_id": task_id, "status": item.get("status", "")},
                    payload_preview=item,
                    missing_fields=["llm_topology_review_task"],
                ))
        return nodes or None

    def _edge_type_for(edge: dict) -> str:
        if is_llm and str(edge.get("edge_kind") or "") == "supply":
            return "llm_topology_supply_edge"
        return edge_type

    if tool_name in {"summarize_chip_topology", "summarize_llm_topology_netlist"}:
        nodes.append(_node(
            f"{base}-{'llm' if is_llm else 'chip'}-topology-summary",
            summary_type,
            result.get("title") or ("LLM 拓扑网表摘要" if is_llm else "芯片级连接拓扑摘要"),
            result.get("summary") or "",
            tool_name=tool_name,
            call_index=call_index,
            locator={"target": "topology", "node_count": result.get("node_count", 0), "edge_count": result.get("edge_count", 0)},
            payload_preview={
                "node_count": result.get("node_count", 0),
                "edge_count": result.get("edge_count", 0),
                "hubs": result.get("hubs", []),
                "role_links": result.get("role_links", []),
                "scope_note": result.get("scope_note", ""),
            },
            detail_tool={"name": summarize_tool, "args": {"limit": 100}},
        ))
        for edge_index, edge in enumerate(list(result.get("edges") or [])[:24], start=1):
            if not isinstance(edge, dict):
                continue
            edge_id = str(edge.get("edge_id") or f"edge-{edge_index}")
            nodes.append(_node(
                f"{base}-{'llm' if is_llm else 'chip'}-edge-{_safe_evidence_fragment(edge_id)}",
                edge_type,
                edge.get("relation_label") or edge_id,
                edge.get("summary") or "",
                tool_name=tool_name,
                call_index=call_index,
                locator={
                    "edge_id": edge_id,
                    "source_refdes": edge.get("source_refdes", ""),
                    "target_refdes": edge.get("target_refdes", ""),
                    "shared_net_count": edge.get("shared_net_count", 0),
                },
                payload_preview=edge,
                detail_tool={"name": edge_detail_tool, "args": {"edge_id": edge_id}},
            ))
        for edge_index, edge in enumerate(list(result.get("supply_edges") or [])[:16], start=1):
            if not isinstance(edge, dict):
                continue
            edge_id = str(edge.get("edge_id") or f"supply-{edge_index}")
            nodes.append(_node(
                f"{base}-{'llm' if is_llm else 'chip'}-supply-edge-{_safe_evidence_fragment(edge_id)}",
                "llm_topology_supply_edge" if is_llm else edge_type,
                edge.get("relation_label") or edge_id,
                edge.get("summary") or "",
                tool_name=tool_name,
                call_index=call_index,
                locator={
                    "edge_id": edge_id,
                    "source_refdes": edge.get("source_refdes", ""),
                    "target_refdes": edge.get("target_refdes", ""),
                    "supply_net": edge.get("supply_net", ""),
                    "voltage_domain": edge.get("voltage_domain", ""),
                },
                payload_preview=edge,
                detail_tool=edge.get("detail_tool") or {"name": query_tool, "args": {"query": edge.get("supply_net", ""), "limit": 20}},
            ))
        for node_index, node in enumerate(list(result.get("nodes") or [])[:24], start=1):
            if not isinstance(node, dict):
                continue
            refdes = str(node.get("refdes") or f"node-{node_index}")
            nodes.append(_node(
                f"{base}-{'llm' if is_llm else 'chip'}-node-{_safe_evidence_fragment(refdes)}",
                node_type,
                f"{refdes} {node.get('role') or '芯片节点'}",
                (
                    f"{refdes} 角色={node.get('role') or ''}；"
                    f"信号网络={node.get('signal_net_count', 0)}；"
                    f"页码={node.get('user_visible_page') or node.get('页码') or node.get('用户看到的真实页') or ''}。"
                ),
                tool_name=tool_name,
                call_index=call_index,
                locator={"refdes": refdes, "role": node.get("role", ""), "node_id": node.get("node_id", "")},
                payload_preview=node,
                detail_tool={"name": "get_llm_topology_node" if is_llm else query_tool, "args": {"refdes": refdes} if is_llm else {"query": refdes, "limit": 20}},
            ))
        return nodes

    if tool_name in {"query_chip_topology", "batch_query_chip_topology", "query_llm_topology_netlist", "batch_query_llm_topology_netlist"}:
        items = (
            [{"query": result.get("query", ""), "items": result.get("items", [])}]
            if tool_name in {"query_chip_topology", "query_llm_topology_netlist"} else list(result.get("items") or [])
        )
        for item_index, item in enumerate(items[:24], start=1):
            if not isinstance(item, dict):
                continue
            query = str(item.get("query") or result.get("query") or f"query-{item_index}")
            matches = list(item.get("items") or [])
            if not matches:
                nodes.append(_node(
                    f"{base}-chip-topology-gap-{item_index}",
                    "missing_context",
                    f"芯片级拓扑无命中：{query}",
                    item.get("missing_reason") or item.get("summary") or "芯片级拓扑未命中该关键词。",
                    tool_name=tool_name,
                    call_index=call_index,
                    locator={"query": query, "status": item.get("status", "")},
                    payload_preview=item,
                    missing_fields=[gap_missing_field],
                ))
                continue
            for match_index, match in enumerate(matches[:8], start=1):
                if not isinstance(match, dict):
                    continue
                edge = match.get("edge") if isinstance(match.get("edge"), dict) else {}
                node = match.get("node") if isinstance(match.get("node"), dict) else {}
                if edge:
                    edge_id = str(edge.get("edge_id") or f"{item_index}-{match_index}")
                    evidence_kind = "supply-edge" if _edge_type_for(edge) == "llm_topology_supply_edge" else "edge"
                    nodes.append(_node(
                        f"{base}-{'llm' if is_llm else 'chip'}-query-{evidence_kind}-{_safe_evidence_fragment(edge_id)}",
                        _edge_type_for(edge),
                        edge.get("relation_label") or edge_id,
                        edge.get("summary") or match.get("summary") or "",
                        tool_name=tool_name,
                        call_index=call_index,
                        locator={
                            "query": query,
                            "edge_id": edge_id,
                            "source_refdes": edge.get("source_refdes", ""),
                            "target_refdes": edge.get("target_refdes", ""),
                        },
                        payload_preview=edge,
                        detail_tool={"name": edge_detail_tool, "args": {"edge_id": edge_id}},
                    ))
                elif node:
                    refdes = str(node.get("refdes") or f"{item_index}-{match_index}")
                    nodes.append(_node(
                        f"{base}-{'llm' if is_llm else 'chip'}-query-node-{_safe_evidence_fragment(refdes)}",
                        node_type,
                        f"{refdes} {node.get('role') or '芯片节点'}",
                        match.get("summary") or f"{refdes} 芯片级拓扑节点。",
                        tool_name=tool_name,
                        call_index=call_index,
                        locator={"query": query, "refdes": refdes, "role": node.get("role", "")},
                        payload_preview=node,
                        detail_tool={"name": "get_llm_topology_node" if is_llm else query_tool, "args": {"refdes": refdes} if is_llm else {"query": refdes, "limit": 20}},
                    ))
                else:
                    task = match.get("review_task") if isinstance(match.get("review_task"), dict) else {}
                    if task:
                        nodes.append(_review_task_node(task, suffix=f"-query-{item_index}-{match_index}"))
        return nodes or None

    if tool_name in {"get_chip_topology_edge", "get_llm_topology_edge"}:
        edge = result.get("edge")
        if isinstance(edge, dict):
            edge_id = str(edge.get("edge_id") or args.get("edge_id") or "edge")
            evidence_kind = "supply-edge" if _edge_type_for(edge) == "llm_topology_supply_edge" else "edge"
            nodes.append(_node(
                f"{base}-{'llm' if is_llm else 'chip'}-{evidence_kind}-detail-{_safe_evidence_fragment(edge_id)}",
                _edge_type_for(edge),
                edge.get("relation_label") or edge_id,
                edge.get("summary") or result.get("summary") or "",
                tool_name=tool_name,
                call_index=call_index,
                locator={
                    "edge_id": edge_id,
                    "source_refdes": edge.get("source_refdes", ""),
                    "target_refdes": edge.get("target_refdes", ""),
                    "shared_net_count": edge.get("shared_net_count", 0),
                },
                payload_preview=edge,
                detail_tool={"name": edge_detail_tool, "args": {"edge_id": edge_id}},
            ))
            return nodes
        return None

    if tool_name == "get_llm_topology_node":
        node = result.get("node")
        if isinstance(node, dict):
            refdes = str(node.get("refdes") or args.get("refdes") or "node")
            nodes.append(_node(
                f"{base}-llm-node-detail-{_safe_evidence_fragment(refdes)}",
                "llm_topology_node",
                f"{refdes} {node.get('role') or '拓扑节点'}",
                result.get("summary") or f"{refdes} LLM 拓扑节点详情。",
                tool_name=tool_name,
                call_index=call_index,
                locator={"refdes": refdes, "node_id": node.get("node_id", ""), "page": node.get("user_visible_page", "")},
                payload_preview={
                    "node": node,
                    "pin_nets": list(result.get("pin_nets") or [])[:24],
                    "edges": list(result.get("edges") or [])[:12],
                },
                detail_tool={"name": "get_llm_topology_node", "args": {"refdes": refdes}},
            ))
            return nodes
        return None

    return None
