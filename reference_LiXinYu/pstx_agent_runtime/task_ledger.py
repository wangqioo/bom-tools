# -*- coding: utf-8 -*-
"""Deterministic task ledger for Codex/Claude-Code style agent runtime state."""

from __future__ import annotations

from collections.abc import Mapping, Sequence

from .protocol import AgentTodoItem, AgentTodoList, PROTOCOL_VERSION


TASK_LEDGER_VERSION = "agent-task-ledger/v1"
INCOMPLETE_COMPLETENESS = {"preview", "partial", "truncated"}
ERROR_COMPLETENESS = {"error"}
CONNECTION_REVIEW_PLAYBOOK_ID = "schematic_datasheet_connection_review"
CONNECTION_REVIEW_SCHEMATIC_EVIDENCE = {
    "llm_topology_edge",
    "llm_topology_node",
    "llm_topology_summary",
    "llm_topology_review_task",
    "source_trace",
    "file_excerpt",
    "component",
    "net",
    "table_row",
}
CONNECTION_REVIEW_IDENTITY_EVIDENCE = {"component_identity"}
CONNECTION_REVIEW_DATASHEET_LOCATOR_EVIDENCE = {
    "datasheet_document",
    "datasheet_match",
    "datasheet_gap",
}
CONNECTION_REVIEW_DATASHEET_DETAIL_EVIDENCE = {
    "datasheet_chunk",
    "datasheet_excerpt",
    "datasheet_parameter",
}
CONNECTION_REVIEW_DETAIL_TOOLS = {
    "get_datasheet_chunk",
    "get_datasheet_excerpt",
    "get_datasheet_page_excerpt",
    "get_datasheet_parameter",
}


def _text(value: object, limit: int = 240) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _dedupe(items: Sequence[object], *, limit: int = 40, text_limit: int = 160) -> tuple[str, ...]:
    output: list[str] = []
    seen = set()
    for item in items or []:
        text = _text(item, text_limit)
        key = text.lower()
        if not text or key in seen:
            continue
        seen.add(key)
        output.append(text)
        if len(output) >= limit:
            break
    return tuple(output)


def _mapping_items(items: Sequence[object]) -> list[Mapping[str, object]]:
    return [item for item in items or [] if isinstance(item, Mapping)]


def _evidence_ids_from_observations(observations: Sequence[Mapping[str, object]], *, limit: int = 80) -> tuple[str, ...]:
    ids: list[object] = []
    for observation in _mapping_items(observations):
        ids.extend(observation.get("evidence_node_ids") or [])
        for node in observation.get("evidence_nodes") or []:
            if isinstance(node, Mapping):
                ids.append(node.get("id"))
        layers = observation.get("evidence_layers")
        cards = layers.get("evidence_card_layer") if isinstance(layers, Mapping) else []
        for card in cards or []:
            if isinstance(card, Mapping):
                ids.append(card.get("id"))
    return _dedupe(ids, limit=limit, text_limit=120)


def _evidence_nodes_from_observations(observations: Sequence[Mapping[str, object]]) -> tuple[Mapping[str, object], ...]:
    nodes: list[Mapping[str, object]] = []
    for observation in _mapping_items(observations):
        nodes.extend(_mapping_items(observation.get("evidence_nodes") or []))
        layers = observation.get("evidence_layers")
        cards = layers.get("evidence_card_layer") if isinstance(layers, Mapping) else []
        nodes.extend(_mapping_items(cards or []))
    return tuple(nodes)


def _present_evidence_types(evidence_nodes: Sequence[Mapping[str, object]]) -> set[str]:
    return {
        _text(node.get("type"), 120)
        for node in _mapping_items(evidence_nodes)
        if _text(node.get("type"), 120)
    }


def _node_source_tool(node: Mapping[str, object]) -> str:
    source = node.get("source") if isinstance(node.get("source"), Mapping) else {}
    return _text(source.get("tool"), 120)


def _has_detail_datasheet_node(evidence_nodes: Sequence[Mapping[str, object]]) -> bool:
    for node in _mapping_items(evidence_nodes):
        evidence_type = _text(node.get("type"), 120)
        if evidence_type not in CONNECTION_REVIEW_DATASHEET_DETAIL_EVIDENCE:
            continue
        if _node_source_tool(node) in CONNECTION_REVIEW_DETAIL_TOOLS:
            return True
    return False


def _detail_tools_from_nodes(evidence_nodes: Sequence[Mapping[str, object]], *, limit: int = 4) -> tuple[dict, ...]:
    tools: list[dict] = []
    seen = set()
    for node in _mapping_items(evidence_nodes):
        detail = node.get("detail_tool") if isinstance(node.get("detail_tool"), Mapping) else {}
        name = _text(detail.get("name"), 120)
        args = detail.get("args") if isinstance(detail.get("args"), Mapping) else {}
        if not name or not args:
            continue
        key = (name, str(dict(args)))
        if key in seen:
            continue
        seen.add(key)
        tools.append({"name": name, "args": dict(args)})
        if len(tools) >= limit:
            break
    return tuple(tools)


def _used_tools(observations: Sequence[Mapping[str, object]]) -> tuple[str, ...]:
    return _dedupe([item.get("tool") for item in _mapping_items(observations)], limit=80, text_limit=100)


def _progress(items: Sequence[Mapping[str, object]]) -> dict:
    counts = {"completed": 0, "in_progress": 0, "pending": 0, "blocked": 0}
    for item in items:
        status = str(item.get("status") or "pending")
        if status not in counts:
            status = "pending"
        counts[status] += 1
    total = sum(counts.values())
    counts["total"] = total
    counts["open"] = counts["pending"] + counts["in_progress"] + counts["blocked"]
    return counts


def _append_item(items: list[dict],
                 *,
                 item_id: str,
                 title: str,
                 status: str,
                 source: str,
                 note: str = "",
                 evidence_ids: Sequence[object] = (),
                 recommended_tools: Sequence[object] = (),
                 detail_tools: Sequence[object] = (),
                 blocking_reason: str = "") -> None:
    if not item_id or any(existing.get("id") == item_id for existing in items):
        return
    payload = {
        "id": _text(item_id, 100),
        "title": _text(title, 180),
        "status": status if status in {"pending", "in_progress", "completed", "blocked"} else "pending",
        "source": _text(source, 80),
        "evidence_ids": list(_dedupe(evidence_ids, limit=24, text_limit=120)),
        "recommended_tools": list(_dedupe(recommended_tools, limit=12, text_limit=120)),
        "detail_tools": [dict(tool) for tool in detail_tools if isinstance(tool, Mapping)][:6],
        "note": _text(note, 320),
    }
    if blocking_reason:
        payload["blocking_reason"] = _text(blocking_reason, 260)
    items.append(payload)


def _action(action_type: str,
            title: str,
            *,
            tool: str = "",
            args: Mapping[str, object] | None = None,
            reason: str = "",
            source: str = "",
            priority: int = 50) -> dict:
    payload = {
        "type": _text(action_type, 60),
        "title": _text(title, 180),
        "reason": _text(reason, 260),
        "source": _text(source, 100),
        "priority": int(priority),
    }
    if tool:
        payload["tool"] = _text(tool, 120)
    if args is not None:
        payload["args"] = dict(args)
    return payload


def _contract_actions(contracts: Sequence[Mapping[str, object]], *,
                      used: Sequence[str],
                      limit: int = 12) -> tuple[dict, ...]:
    actions: list[dict] = []
    for index, contract in enumerate(_mapping_items(contracts), start=1):
        completeness = _text(contract.get("completeness"), 40).lower()
        if completeness not in INCOMPLETE_COMPLETENESS and completeness not in ERROR_COMPLETENESS:
            continue
        scope = _text(contract.get("scope_summary") or f"contract-{index}", 180)
        aggregation = contract.get("aggregation_tool")
        detail = contract.get("detail_tool")
        if isinstance(aggregation, Mapping) and aggregation.get("name"):
            actions.append(_action(
                "tool_call",
                f"聚合截断结果：{scope}",
                tool=str(aggregation.get("name")),
                args=aggregation.get("args") if isinstance(aggregation.get("args"), Mapping) else {},
                reason=f"当前工具结果为 {completeness}，需要先聚合再下统计结论。",
                source=f"tool_result_contract-{index}",
                priority=10,
            ))
        if isinstance(detail, Mapping) and detail.get("name"):
            actions.append(_action(
                "tool_call",
                f"读取原始详情：{scope}",
                tool=str(detail.get("name")),
                args=detail.get("args") if isinstance(detail.get("args"), Mapping) else {},
                reason=f"当前工具结果为 {completeness}，高风险结论前需要回拉 detail。",
                source=f"tool_result_contract-{index}",
                priority=20,
            ))
        for tool in contract.get("recommended_next_tools") or []:
            if _text(tool, 120):
                title = f"改用替代工具：{scope}" if completeness in ERROR_COMPLETENESS else f"继续取证：{scope}"
                reason = (
                    f"上一轮工具失败，runtime 推荐改用 {tool} 继续安全取证。"
                    if completeness in ERROR_COMPLETENESS
                    else f"工具协议推荐下一步；当前完整性为 {completeness}。"
                )
                actions.append(_action(
                    "tool_call",
                    title,
                    tool=str(tool),
                    reason=reason,
                    source=f"tool_result_contract-{index}",
                    priority=15 if completeness in ERROR_COMPLETENESS else 30,
                ))
    deduped: list[dict] = []
    seen = set()
    used_set = set(used or [])
    for action in sorted(actions, key=lambda item: item.get("priority", 50)):
        key = (action.get("type"), action.get("tool"), str(action.get("args") or {}), action.get("title"))
        if key in seen:
            continue
        seen.add(key)
        if action.get("tool") and action.get("tool") in used_set and not action.get("args"):
            continue
        deduped.append(action)
        if len(deduped) >= limit:
            break
    return tuple(deduped)


def _playbook_ids(playbook_plan: Mapping[str, object]) -> set[str]:
    return {
        _text(playbook.get("id"), 120)
        for playbook in _mapping_items(playbook_plan.get("selected_playbooks") or [])
        if _text(playbook.get("id"), 120)
    }


def _seeded_calls(playbook_plan: Mapping[str, object]) -> tuple[Mapping[str, object], ...]:
    return tuple(_mapping_items(playbook_plan.get("seeded_tool_calls") or []))


def _seed_by_tool(playbook_plan: Mapping[str, object], tool_names: Sequence[str]) -> Mapping[str, object] | None:
    wanted = set(tool_names or [])
    for seed in _seeded_calls(playbook_plan):
        if _text(seed.get("name") or seed.get("tool"), 120) in wanted:
            return seed
    return None


def _seed_targets(playbook_plan: Mapping[str, object]) -> tuple[str, ...]:
    values: list[object] = []
    for seed in _seeded_calls(playbook_plan):
        args = seed.get("args") if isinstance(seed.get("args"), Mapping) else {}
        for key in ("queries", "refdes_list"):
            raw = args.get(key)
            if isinstance(raw, list):
                values.extend(raw)
            elif raw:
                values.append(raw)
    return _dedupe(values, limit=24, text_limit=120)


def _connection_review_stage_status(*,
                                    evidence_types: set[str],
                                    used: Sequence[str],
                                    stage_evidence: set[str],
                                    stage_tools: set[str],
                                    fallback_complete: bool = False) -> str:
    if fallback_complete or evidence_types & stage_evidence:
        return "completed"
    if set(used or []) & stage_tools:
        return "in_progress"
    return "pending"


def _append_connection_review_items(items: list[dict],
                                    *,
                                    playbook_plan: Mapping[str, object],
                                    evidence_nodes: Sequence[Mapping[str, object]],
                                    evidence_types: set[str],
                                    evidence_ids: Sequence[object],
                                    used: Sequence[str]) -> None:
    if CONNECTION_REVIEW_PLAYBOOK_ID not in _playbook_ids(playbook_plan):
        return

    targets = _seed_targets(playbook_plan)
    schematic_complete = bool(evidence_types & CONNECTION_REVIEW_SCHEMATIC_EVIDENCE)
    identity_complete = bool(evidence_types & CONNECTION_REVIEW_IDENTITY_EVIDENCE)
    locator_complete = bool(evidence_types & CONNECTION_REVIEW_DATASHEET_LOCATOR_EVIDENCE)
    gap_present = "datasheet_gap" in evidence_types
    detail_complete = _has_detail_datasheet_node(evidence_nodes) or gap_present

    _append_item(
        items,
        item_id="connection-review-targets",
        title="解读用户连接反查目标",
        status="completed" if targets else "in_progress",
        source="connection_review_phase",
        note=(
            f"已从用户问题提取目标：{', '.join(targets[:8])}。"
            if targets else
            "需要先定位用户问题中的位号、网络、rail、接口或时序关键词。"
        ),
        evidence_ids=evidence_ids if targets else (),
        recommended_tools=("batch_query_report_entities", "summarize_llm_topology_netlist"),
    )
    _append_item(
        items,
        item_id="connection-review-schematic-evidence",
        title="读取原理图/网表/拓扑连接 evidence",
        status=_connection_review_stage_status(
            evidence_types=evidence_types,
            used=used,
            stage_evidence=CONNECTION_REVIEW_SCHEMATIC_EVIDENCE,
            stage_tools={"batch_query_llm_topology_netlist", "query_llm_topology_netlist", "get_llm_topology_edge", "get_llm_topology_node", "trace_project_source", "search_project_text", "batch_query_report_entities"},
        ),
        source="connection_review_phase",
        note="连接判断必须先有 topology edge/node、pin-net、report entity 或 source trace evidence。",
        evidence_ids=evidence_ids if schematic_complete else (),
        recommended_tools=("batch_query_llm_topology_netlist", "get_llm_topology_edge", "get_llm_topology_node", "trace_project_source", "search_project_text"),
    )
    _append_item(
        items,
        item_id="connection-review-identity",
        title="确认相关元件身份和 pin-net 上下文",
        status=_connection_review_stage_status(
            evidence_types=evidence_types,
            used=used,
            stage_evidence=CONNECTION_REVIEW_IDENTITY_EVIDENCE,
            stage_tools={"batch_get_component_identity_cards", "get_component_identity_card", "search_component_identity_cards"},
        ),
        source="connection_review_phase",
        note="按 refdes 确认 HQ、型号、power nets、interface nets 后再套用 datasheet 条件。",
        evidence_ids=evidence_ids if identity_complete else (),
        recommended_tools=("batch_get_component_identity_cards", "get_component_identity_card"),
    )
    _append_item(
        items,
        item_id="connection-review-datasheet-locator",
        title="匹配 MinerU-backed datasheet 候选",
        status=_connection_review_stage_status(
            evidence_types=evidence_types,
            used=used,
            stage_evidence=CONNECTION_REVIEW_DATASHEET_LOCATOR_EVIDENCE,
            stage_tools={"list_datasheet_sources", "batch_match_component_datasheets", "match_component_datasheets", "search_datasheets", "batch_search_datasheet_chunks"},
        ),
        source="connection_review_phase",
        note="先确认本地 MinerU 索引和 refdes 对应 datasheet 候选；无命中应保留 gap。",
        evidence_ids=evidence_ids if locator_complete else (),
        recommended_tools=("list_datasheet_sources", "batch_match_component_datasheets", "match_component_datasheets"),
    )
    _append_item(
        items,
        item_id="connection-review-datasheet-detail",
        title="读取 datasheet detail/parameter 原文",
        status=_connection_review_stage_status(
            evidence_types=evidence_types,
            used=used,
            stage_evidence=set(),
            stage_tools={"search_datasheet_parameters", "get_datasheet_parameter", "batch_search_datasheet_chunks", "get_datasheet_chunk", "get_datasheet_excerpt", "get_datasheet_page_excerpt"},
            fallback_complete=detail_complete,
        ),
        source="connection_review_phase",
        note=(
            "当前已有 datasheet gap，可在结论中明确说明缺口，不猜参数。"
            if gap_present and not _has_detail_datasheet_node(evidence_nodes) else
            "定量/电气事实必须来自 get_datasheet_chunk/get_datasheet_parameter 等 detail evidence。"
        ),
        evidence_ids=evidence_ids if detail_complete else (),
        recommended_tools=("search_datasheet_parameters", "get_datasheet_parameter", "batch_search_datasheet_chunks", "get_datasheet_chunk"),
        detail_tools=_detail_tools_from_nodes(evidence_nodes),
    )
    _append_item(
        items,
        item_id="connection-review-backcheck",
        title="用 datasheet fact 反查连接风险",
        status="in_progress" if schematic_complete and detail_complete else "pending",
        source="connection_review_phase",
        note="最终回答应逐项给出 pass-like observation、evidence-backed risk 或 evidence gap。",
        evidence_ids=evidence_ids if schematic_complete and detail_complete else (),
    )


def _connection_review_actions(*,
                               playbook_plan: Mapping[str, object],
                               evidence_nodes: Sequence[Mapping[str, object]],
                               evidence_types: set[str],
                               used: Sequence[str]) -> tuple[dict, ...]:
    if CONNECTION_REVIEW_PLAYBOOK_ID not in _playbook_ids(playbook_plan):
        return ()

    actions: list[dict] = []

    def add_seed(tool_names: Sequence[str], title: str, reason: str, priority: int) -> None:
        seed = _seed_by_tool(playbook_plan, tool_names)
        if not isinstance(seed, Mapping):
            return
        tool = _text(seed.get("name") or seed.get("tool"), 120)
        args = seed.get("args") if isinstance(seed.get("args"), Mapping) else {}
        if not tool or tool in set(used or []):
            return
        actions.append(_action(
            "tool_call",
            title,
            tool=tool,
            args=args,
            reason=reason,
            source="connection_review_phase",
            priority=priority,
        ))

    if not (evidence_types & CONNECTION_REVIEW_SCHEMATIC_EVIDENCE):
        add_seed(
            ("batch_query_llm_topology_netlist", "batch_query_report_entities"),
            "连接反查先读取原理图/网表 evidence",
            "连接反查不能先看 datasheet 下结论；需要先拿 topology/pin-net/source evidence。",
            9,
        )
    if not (evidence_types & CONNECTION_REVIEW_IDENTITY_EVIDENCE):
        add_seed(
            ("batch_get_component_identity_cards",),
            "连接反查确认相关元件身份",
            "需要先确认 refdes 对应 HQ、型号、pin-net、power/interface nets，再匹配 datasheet。",
            10,
        )
    if not (evidence_types & CONNECTION_REVIEW_DATASHEET_LOCATOR_EVIDENCE):
        add_seed(
            ("batch_match_component_datasheets", "list_datasheet_sources"),
            "连接反查匹配 MinerU-backed datasheet 候选",
            "需要确认本地 datasheet 索引和 refdes/PDF 匹配；无命中也要形成 gap evidence。",
            11,
        )
    if not _has_detail_datasheet_node(evidence_nodes) and "datasheet_gap" not in evidence_types:
        detail_action_added = False
        for detail in _detail_tools_from_nodes(evidence_nodes, limit=3):
            actions.append(_action(
                "tool_call",
                "连接反查打开 datasheet detail 原文",
                tool=str(detail.get("name") or ""),
                args=detail.get("args") if isinstance(detail.get("args"), Mapping) else {},
                reason="当前只有 datasheet locator/search evidence；反查连接风险前必须打开 detail chunk/parameter 原文。",
                source="connection_review_phase",
                priority=12,
            ))
            detail_action_added = True
        if not detail_action_added:
            add_seed(
                ("search_datasheet_parameters", "batch_search_datasheet_chunks"),
                "连接反查检索 datasheet 参数和章节",
                "缺少 datasheet detail/gap；应先按接口、电源、reset、clock 或 strap 关键词检索 MinerU-backed evidence。",
                13,
            )
    if not (evidence_types & {"source_trace", "file_excerpt"}):
        add_seed(
            ("trace_project_source", "search_project_text"),
            "连接反查保留原始文件追溯证据",
            "高风险连接结论最好能回到 PSTX/Cadence line-number excerpt；无法定位时再说明缺口。",
            18,
        )
    return tuple(actions[:8])


def build_task_ledger(*,
                      goal: object,
                      capability_plan: Sequence[Mapping[str, object]] = (),
                      playbook_plan: Mapping[str, object] | None = None,
                      observations: Sequence[Mapping[str, object]] = (),
                      tool_result_contracts: Sequence[Mapping[str, object]] = (),
                      project_context: Mapping[str, object] | None = None,
                      max_items: int = 18) -> dict:
    """Build a compact, deterministic task ledger for model planning and trace replay."""

    playbook_plan = playbook_plan or {}
    project_context = project_context or {}
    observation_items = _mapping_items(observations)
    observed_count = len(observation_items)
    evidence_ids = _evidence_ids_from_observations(observation_items)
    evidence_nodes = _evidence_nodes_from_observations(observation_items)
    evidence_types = _present_evidence_types(evidence_nodes)
    used = _used_tools(observation_items)
    session_memory = project_context.get("session_memory_summary")
    session_memory = session_memory if isinstance(session_memory, Mapping) else {}
    items: list[dict] = []

    plans = _mapping_items(capability_plan)
    if not plans:
        plans = [{"id": "quick_scan", "title": "快速证据收集", "description": "先收集核心证据。"}]
    for index, plan in enumerate(plans[:8], start=1):
        status = "completed" if observed_count and index == 1 else ("in_progress" if index == 1 else "pending")
        _append_item(
            items,
            item_id=f"capability-{_text(plan.get('id') or index, 80)}",
            title=plan.get("title") or plan.get("id") or f"能力 {index}",
            status=status,
            source="capability_plan",
            note="已有工具观察，继续补齐细节和引用。" if status == "completed" else plan.get("description") or "",
            evidence_ids=evidence_ids if status == "completed" else (),
        )

    selected_playbooks = _mapping_items(playbook_plan.get("selected_playbooks") or [])
    recommended_first = list(playbook_plan.get("recommended_first_tools") or [])
    for index, playbook in enumerate(selected_playbooks[:8], start=1):
        tools = list(playbook.get("preferred_batch_tools") or []) + list(playbook.get("preferred_tools") or [])
        tool_hit = any(tool in used for tool in tools)
        status = "completed" if tool_hit and evidence_ids else ("in_progress" if index == 1 else "pending")
        _append_item(
            items,
            item_id=f"playbook-{_text(playbook.get('id') or index, 80)}",
            title=playbook.get("title") or playbook.get("id") or f"取证路线 {index}",
            status=status,
            source="playbook_plan",
            note="路线已有证据，下一步聚合/回拉详情后再总结。" if status == "completed" else "按 playbook 推荐路线取证。",
            evidence_ids=evidence_ids if tool_hit else (),
            recommended_tools=[tool for tool in recommended_first if tool in tools] or tools,
        )

    _append_connection_review_items(
        items,
        playbook_plan=playbook_plan,
        evidence_nodes=evidence_nodes,
        evidence_types=evidence_types,
        evidence_ids=evidence_ids,
        used=used,
    )

    for index, contract in enumerate(_mapping_items(tool_result_contracts)[-10:], start=1):
        completeness = _text(contract.get("completeness"), 40).lower()
        if completeness not in INCOMPLETE_COMPLETENESS and completeness not in ERROR_COMPLETENESS:
            continue
        detail_tools = []
        if isinstance(contract.get("aggregation_tool"), Mapping):
            detail_tools.append(contract["aggregation_tool"])
        if isinstance(contract.get("detail_tool"), Mapping):
            detail_tools.append(contract["detail_tool"])
        if completeness in ERROR_COMPLETENESS:
            _append_item(
                items,
                item_id=f"contract-{index}-{completeness}",
                title=f"修正失败工具调用：{contract.get('scope_summary') or completeness}",
                status="pending",
                source="tool_error_contract",
                note="上一次工具调用失败，下一步应换用推荐白名单工具或修正参数，不要重复同一失败调用。",
                recommended_tools=contract.get("recommended_next_tools") or [],
                detail_tools=detail_tools,
            )
        else:
            _append_item(
                items,
                item_id=f"contract-{index}-{completeness}",
                title=f"补齐不完整工具结果：{contract.get('scope_summary') or completeness}",
                status="in_progress",
                source="tool_result_contract",
                note=f"当前结果完整性为 {completeness}，不能只用 preview/摘要下最终统计结论。",
                recommended_tools=contract.get("recommended_next_tools") or [],
                detail_tools=detail_tools,
            )

    for index, question in enumerate(_mapping_items(project_context.get("pending_questions") or [])[:8], start=1):
        qid = _text(question.get("question_id") or question.get("id") or f"q-{index}", 80)
        _append_item(
            items,
            item_id=f"clarification-{qid}",
            title=question.get("question") or question.get("prompt") or "等待用户补充信息",
            status="blocked",
            source="clarification",
            note="需要用户补充后才能继续当前任务。",
            evidence_ids=question.get("related_evidence_ids") or [],
            blocking_reason="; ".join(_dedupe(question.get("missing_fields") or [], limit=8, text_limit=80)),
        )

    for index, title in enumerate(_dedupe(session_memory.get("open_items") or [], limit=8, text_limit=260), start=1):
        _append_item(
            items,
            item_id=f"session-memory-open-{index}",
            title=title,
            status="pending",
            source="session_memory",
            note="来自项目级滚动记忆的未完成任务；需结合当前问题和 evidence 继续确认。",
            evidence_ids=session_memory.get("evidence_ids") or [],
        )

    next_actions: list[dict] = []
    used_set = set(used)
    seeded_names: set[str] = set()
    for seed in _mapping_items(playbook_plan.get("seeded_tool_calls") or []):
        tool_name = _text(seed.get("name") or seed.get("tool"), 120)
        args = seed.get("args") if isinstance(seed.get("args"), Mapping) else {}
        if not tool_name or tool_name in used_set:
            continue
        seeded_names.add(tool_name)
        next_actions.append(_action(
            "tool_call",
            f"按 playbook 带参种子取证：{tool_name}",
            tool=tool_name,
            args=args,
            reason=seed.get("reason") or "本地 playbook 已从用户问题提取实体并生成工具参数。",
            source=seed.get("source") or "playbook_seed",
            priority=8,
        ))
    for tool in recommended_first:
        text = _text(tool, 120)
        if text and text not in used_set and text not in seeded_names:
            next_actions.append(_action(
                "tool_call",
                f"按 playbook 首选工具取证：{text}",
                tool=text,
                reason="本地 planner/playbook 判断该工具是当前问题的优先取证入口。",
                source="playbook_plan",
                priority=15,
            ))
    next_actions.extend(_connection_review_actions(
        playbook_plan=playbook_plan,
        evidence_nodes=evidence_nodes,
        evidence_types=evidence_types,
        used=used,
    ))
    next_actions.extend(_contract_actions(tool_result_contracts, used=used))
    if project_context.get("pending_questions"):
        next_actions.append(_action(
            "ask_user",
            "等待用户补充阻塞信息",
            reason="存在 pending_questions，继续推理前需要用户补充。",
            source="clarification",
            priority=5,
        ))
    for index, action_text in enumerate(_dedupe(session_memory.get("next_actions") or [], limit=8, text_limit=260), start=1):
        next_actions.append(_action(
            "review_memory_next_action",
            f"复用项目记忆下一步：{action_text}",
            reason="上一轮项目级滚动记忆留下的建议，模型应结合当前任务判断是否需要转成白名单工具调用或追问。",
            source="session_memory",
            priority=60 + index,
        ))

    deduped_actions: list[dict] = []
    seen_actions = set()
    for action in sorted(next_actions, key=lambda item: item.get("priority", 50)):
        key = (action.get("type"), action.get("tool"), str(action.get("args") or {}), action.get("title"))
        if key in seen_actions:
            continue
        seen_actions.add(key)
        deduped_actions.append(action)
        if len(deduped_actions) >= 12:
            break

    capped_items = items[:max_items]
    progress = _progress(capped_items)
    return {
        "version": TASK_LEDGER_VERSION,
        "protocol_version": PROTOCOL_VERSION,
        "goal": _text(goal, 600),
        "items": capped_items,
        "progress": progress,
        "next_actions": deduped_actions,
        "evidence_ids": list(evidence_ids),
        "used_tools": list(used),
        "notes": (
            "task_ledger 是本地确定性任务账本；模型应优先处理 blocked/in_progress 与 next_actions，"
            "最终结论必须引用 evidence。"
        ),
    }


def todo_list_from_task_ledger(task_ledger: Mapping[str, object]) -> AgentTodoList:
    """Build the legacy TodoList view from the richer task ledger."""

    items: list[AgentTodoItem] = []
    for index, item in enumerate(_mapping_items(task_ledger.get("items") or [])[:8], start=1):
        items.append(AgentTodoItem(
            id=f"todo-{index}",
            title=_text(item.get("title") or item.get("id") or f"任务 {index}", 160),
            status=_text(item.get("status") or "pending", 40),
            evidence_ids=tuple(_dedupe(item.get("evidence_ids") or [], limit=12, text_limit=120)),
            note=_text(item.get("note") or item.get("source") or "", 300),
        ))
    return AgentTodoList(goal=_text(task_ledger.get("goal"), 500), items=tuple(items))
