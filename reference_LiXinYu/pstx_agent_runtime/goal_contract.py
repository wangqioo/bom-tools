# -*- coding: utf-8 -*-
"""Evidence goal contracts for Codex/Claude-Code style agent execution.

The playbook tells the model which route to try. This module turns that route
into a deterministic evidence checklist so the local runtime can notice when a
final answer is trying to stop before the necessary evidence types exist.
"""

from __future__ import annotations

import json
import re
from collections.abc import Mapping, Sequence


EVIDENCE_GOAL_CONTRACT_VERSION = "agent-evidence-goal-contract/v1"
CONNECTION_REVIEW_PLAYBOOK_ID = "schematic_datasheet_connection_review"
CONNECTION_REVIEW_DETAIL_TOOLS = {
    "get_datasheet_chunk",
    "get_datasheet_excerpt",
    "get_datasheet_page_excerpt",
    "get_datasheet_parameter",
}
CONNECTION_REVIEW_PHASES = (
    {
        "id": "schematic_connection",
        "title": "原理图/网表/拓扑连接 evidence",
        "evidence_types": ("llm_topology_edge", "llm_topology_node", "llm_topology_summary", "source_trace", "file_excerpt", "component", "net", "table_row"),
        "seed_tools": ("batch_query_llm_topology_netlist", "batch_query_report_entities", "trace_project_source", "search_project_text"),
    },
    {
        "id": "component_identity",
        "title": "元件身份和 pin-net 上下文",
        "evidence_types": ("component_identity",),
        "seed_tools": ("batch_get_component_identity_cards", "get_component_identity_card"),
    },
    {
        "id": "datasheet_locator",
        "title": "MinerU-backed datasheet 匹配或缺口",
        "evidence_types": ("datasheet_match", "datasheet_document", "datasheet_gap"),
        "seed_tools": ("batch_match_component_datasheets", "match_component_datasheets", "list_datasheet_sources"),
    },
    {
        "id": "datasheet_detail",
        "title": "datasheet detail/parameter 原文或明确 gap",
        "evidence_types": ("datasheet_chunk", "datasheet_excerpt", "datasheet_parameter", "datasheet_gap"),
        "seed_tools": ("search_datasheet_parameters", "batch_search_datasheet_chunks"),
    },
)


def _text(value: object, limit: int = 240) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _mapping_items(items: Sequence[object] | object) -> list[Mapping[str, object]]:
    return [item for item in items or [] if isinstance(item, Mapping)]


def _dedupe(items: Sequence[object], *, limit: int = 80, text_limit: int = 160) -> tuple[str, ...]:
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


def _target_key(value: object) -> str:
    return _text(value, 160).upper()


def _is_high_confidence_target(value: object) -> bool:
    text = _text(value, 160)
    upper = text.upper()
    if not text:
        return False
    if re.fullmatch(r"HQ[0-9A-Z]{3,}", upper):
        return True
    if re.fullmatch(r"(?:P?[RUCL]\d+[A-Z]?\d*|P?C\d+[A-Z]?\d*|PU\d+[A-Z]?\d*|XU\d+[A-Z]?\d*|U\d+[A-Z]?\d*|J\d+[A-Z]?\d*|CN\d+[A-Z]?\d*)", upper):
        return True
    # Datasheet/spec identifiers often mix letters and digits, for example LCMXO3LF.
    return len(upper) >= 5 and any(ch.isdigit() for ch in upper) and any(ch.isalpha() for ch in upper)


_TARGET_BATCH_ARG_KEYS = {
    "batch_query_report_entities": "queries",
    "batch_query_compare_diff": "queries",
    "batch_query_llm_topology_netlist": "queries",
    "batch_query_chip_topology": "queries",
    "batch_expand_topology_review_tasks": "task_ids",
    "batch_search_feishu_cache_rows": "queries",
    "batch_search_datasheet_chunks": "queries",
    "batch_get_component_identity_cards": "refdes_list",
    "batch_match_component_datasheets": "refdes_list",
}


_TARGET_SINGLE_ARG_KEYS = {
    "query_report_entity": "keyword",
    "query_compare_diff": "query",
    "search_feishu_cache_rows": "query",
    "search_project_text": "query",
    "search_datasheet_chunks": "query",
    "get_component_identity_card": "refdes",
    "match_component_datasheets": "refdes",
}


def _target_repair_arg_key(tool: str, arg_key: str) -> str:
    if tool in _TARGET_BATCH_ARG_KEYS:
        return _TARGET_BATCH_ARG_KEYS[tool]
    return arg_key


def _target_tool_priority(value: object, tool: str) -> int:
    upper = _target_key(value)
    is_hq = bool(re.fullmatch(r"HQ[0-9A-Z]{3,}", upper))
    is_refdes = bool(re.fullmatch(r"(?:P?[RUCL]\d+[A-Z]?\d*|P?C\d+[A-Z]?\d*|PU\d+[A-Z]?\d*|XU\d+[A-Z]?\d*|U\d+[A-Z]?\d*|J\d+[A-Z]?\d*|CN\d+[A-Z]?\d*)", upper))
    if is_refdes:
        order = {
            "batch_get_component_identity_cards": 10,
            "get_component_identity_card": 12,
            "batch_query_report_entities": 20,
            "query_report_entity": 22,
            "batch_query_compare_diff": 24,
            "query_compare_diff": 26,
            "batch_query_llm_topology_netlist": 28,
            "batch_query_chip_topology": 30,
            "batch_expand_topology_review_tasks": 32,
            "search_project_text": 34,
            "batch_search_feishu_cache_rows": 70,
            "batch_search_datasheet_chunks": 80,
        }
        return order.get(tool, 90)
    if is_hq:
        order = {
            "batch_search_feishu_cache_rows": 10,
            "search_feishu_cache_rows": 12,
            "batch_search_datasheet_chunks": 20,
            "search_datasheet_chunks": 22,
            "batch_query_report_entities": 30,
            "batch_query_compare_diff": 32,
            "search_project_text": 34,
        }
        return order.get(tool, 90)
    order = {
        "batch_search_datasheet_chunks": 10,
        "search_datasheet_chunks": 12,
        "batch_query_compare_diff": 20,
        "batch_query_report_entities": 22,
        "batch_search_feishu_cache_rows": 30,
        "search_project_text": 35,
    }
    return order.get(tool, 90)


def _targets_from_seeded_calls(playbook_plan: Mapping[str, object]) -> list[dict]:
    candidates: dict[str, dict] = {}
    for seed in _mapping_items(playbook_plan.get("seeded_tool_calls") or []):
        tool = _text(seed.get("name") or seed.get("tool"), 120)
        args = seed.get("args") if isinstance(seed.get("args"), Mapping) else {}
        if tool in _TARGET_BATCH_ARG_KEYS:
            arg_key = _TARGET_BATCH_ARG_KEYS[tool]
            raw_values = args.get(arg_key)
            values = raw_values if isinstance(raw_values, list) else [raw_values]
        elif tool in _TARGET_SINGLE_ARG_KEYS:
            arg_key = _TARGET_SINGLE_ARG_KEYS[tool]
            values = [args.get(arg_key)]
        else:
            continue
        for value in values:
            target = _text(value, 160)
            normalized = _target_key(target)
            if not _is_high_confidence_target(target):
                continue
            priority = _target_tool_priority(target, tool)
            existing = candidates.get(normalized)
            if existing and int(existing.get("_priority") or 999) <= priority:
                continue
            candidates[normalized] = {
                "value": target,
                "normalized": normalized,
                "source_tool": tool,
                "repair_tool": tool,
                "repair_arg_key": _target_repair_arg_key(tool, arg_key),
                "source": "playbook_seed",
                "_priority": priority,
            }
            if len(candidates) >= 40:
                break
    return [
        {key: value for key, value in item.items() if key != "_priority"}
        for item in sorted(candidates.values(), key=lambda item: (int(item.get("_priority") or 999), str(item.get("normalized") or "")))
    ]


def _evidence_nodes_from_observations(observations: Sequence[Mapping[str, object]]) -> list[Mapping[str, object]]:
    nodes: list[Mapping[str, object]] = []
    for observation in _mapping_items(observations):
        nodes.extend(_mapping_items(observation.get("evidence_nodes") or []))
        layers = observation.get("evidence_layers")
        cards = layers.get("evidence_card_layer") if isinstance(layers, Mapping) else []
        for card in _mapping_items(cards or []):
            nodes.append(card)
    return nodes


def _present_evidence_types(evidence_nodes: Sequence[Mapping[str, object]]) -> tuple[str, ...]:
    return _dedupe([node.get("type") for node in _mapping_items(evidence_nodes)], limit=80, text_limit=120)


def _selected_playbook_ids(playbook_plan: Mapping[str, object]) -> set[str]:
    return {
        _text(playbook.get("id"), 120)
        for playbook in _mapping_items(playbook_plan.get("selected_playbooks") or [])
        if _text(playbook.get("id"), 120)
    }


def _node_source_tool(node: Mapping[str, object]) -> str:
    source = node.get("source") if isinstance(node.get("source"), Mapping) else {}
    return _text(source.get("tool"), 120)


def _is_datasheet_detail_node(node: Mapping[str, object]) -> bool:
    evidence_type = _text(node.get("type"), 120)
    return (
        evidence_type in {"datasheet_chunk", "datasheet_excerpt", "datasheet_parameter"}
        and _node_source_tool(node) in CONNECTION_REVIEW_DETAIL_TOOLS
    )


def _detail_actions_from_nodes(evidence_nodes: Sequence[Mapping[str, object]], *, limit: int = 4) -> list[dict]:
    actions: list[dict] = []
    seen = set()
    for node in _mapping_items(evidence_nodes):
        detail = node.get("detail_tool") if isinstance(node.get("detail_tool"), Mapping) else {}
        tool = _text(detail.get("name"), 120)
        args = detail.get("args") if isinstance(detail.get("args"), Mapping) else {}
        if not tool or not args:
            continue
        key = (tool, str(dict(args)))
        if key in seen:
            continue
        seen.add(key)
        actions.append({
            "type": "tool_call",
            "tool": tool,
            "args": dict(args),
            "title": f"打开 datasheet detail 原文：{_text(node.get('title') or node.get('id'), 140)}",
            "reason": "连接反查阶段发现当前只有 datasheet locator/search evidence；反查前需要读取 detail chunk/parameter 原文。",
            "source": "missing_connection_review_phase",
            "priority": 12,
        })
        if len(actions) >= limit:
            break
    return actions


def _seed_for_tools(playbook_plan: Mapping[str, object], tool_names: Sequence[str]) -> Mapping[str, object] | None:
    wanted = set(tool_names or [])
    for seed in _mapping_items(playbook_plan.get("seeded_tool_calls") or []):
        tool = _text(seed.get("name") or seed.get("tool"), 120)
        if tool in wanted:
            return seed
    return None


def _connection_review_phase_contract(playbook_plan: Mapping[str, object],
                                      evidence_nodes: Sequence[Mapping[str, object]],
                                      present_types: set[str],
                                      *,
                                      max_actions: int = 8) -> dict:
    if CONNECTION_REVIEW_PLAYBOOK_ID not in _selected_playbook_ids(playbook_plan):
        return {
            "status": "not_required",
            "phases": [],
            "missing_phases": [],
            "repair_actions": [],
        }

    has_detail_node = any(_is_datasheet_detail_node(node) for node in _mapping_items(evidence_nodes))
    has_datasheet_gap = "datasheet_gap" in present_types
    phases: list[dict] = []
    actions: list[dict] = []

    for phase in CONNECTION_REVIEW_PHASES:
        phase_id = str(phase["id"])
        evidence_types = set(str(item) for item in phase.get("evidence_types") or [])
        if phase_id == "datasheet_detail":
            covered = sorted(present_types & evidence_types)
            satisfied = has_detail_node or has_datasheet_gap
        else:
            covered = sorted(present_types & evidence_types)
            satisfied = bool(covered)
        status = "satisfied" if satisfied else "missing"
        phases.append({
            "id": phase_id,
            "title": _text(phase.get("title"), 180),
            "required_evidence": sorted(evidence_types),
            "covered_evidence": covered,
            "status": status,
        })
        if satisfied:
            continue
        if phase_id == "datasheet_detail":
            actions.extend(_detail_actions_from_nodes(evidence_nodes, limit=max_actions - len(actions)))
            if len(actions) >= max_actions:
                continue
        seed = _seed_for_tools(playbook_plan, phase.get("seed_tools") or [])
        if isinstance(seed, Mapping):
            tool = _text(seed.get("name") or seed.get("tool"), 120)
            args = seed.get("args") if isinstance(seed.get("args"), Mapping) else {}
            if tool:
                actions.append({
                    "type": "tool_call",
                    "tool": tool,
                    "args": dict(args),
                    "title": f"补齐连接反查阶段：{phase.get('title')}",
                    "reason": "连接 × datasheet 反查需要按阶段补齐连接、身份、datasheet locator/detail evidence；缺失阶段不能直接下 pass/fail 结论。",
                    "source": "missing_connection_review_phase",
                    "priority": 14 + len(actions),
                })
        if len(actions) >= max_actions:
            break

    missing = [dict(item) for item in phases if item.get("status") != "satisfied"]
    if not missing:
        status = "satisfied"
    elif len(missing) == len(phases):
        status = "missing"
    else:
        status = "partial"
    return {
        "status": status,
        "phases": phases,
        "missing_phases": missing,
        "repair_actions": actions[:max_actions],
    }


def _node_search_blob(node: Mapping[str, object]) -> str:
    try:
        return json.dumps(node, ensure_ascii=False, sort_keys=True, default=str).upper()
    except (TypeError, ValueError):
        return str(node).upper()


def _contains_target_token(text: object, normalized: object) -> bool:
    target = _target_key(normalized)
    blob = str(text or "").upper()
    if not target or not blob:
        return False
    if re.fullmatch(r"[A-Z0-9_]+", target):
        return re.search(rf"(?<![A-Z0-9_]){re.escape(target)}(?![A-Z0-9_])", blob) is not None
    return target in blob


def _target_coverage(required_targets: Sequence[Mapping[str, object]],
                     evidence_nodes: Sequence[Mapping[str, object]]) -> dict:
    blobs = [_node_search_blob(node) for node in _mapping_items(evidence_nodes)]
    covered: list[dict] = []
    missing: list[dict] = []
    for target in _mapping_items(required_targets):
        normalized = _target_key(target.get("normalized") or target.get("value"))
        item = {
            "value": _text(target.get("value"), 160),
            "normalized": normalized,
            "repair_tool": _text(target.get("repair_tool"), 120),
            "repair_arg_key": _text(target.get("repair_arg_key"), 80),
            "source": _text(target.get("source"), 80),
        }
        if normalized and any(_contains_target_token(blob, normalized) for blob in blobs):
            covered.append(item)
        else:
            missing.append(item)
    if not required_targets:
        status = "not_required"
    elif not missing:
        status = "satisfied"
    elif len(missing) == len(required_targets):
        status = "missing"
    else:
        status = "partial"
    return {
        "status": status,
        "covered_targets": covered,
        "missing_targets": missing,
    }


def _seed_actions(playbook_plan: Mapping[str, object], *, limit: int = 8) -> list[dict]:
    actions: list[dict] = []
    for seed in _mapping_items(playbook_plan.get("seeded_tool_calls") or []):
        name = _text(seed.get("name") or seed.get("tool"), 120)
        if not name:
            continue
        args = seed.get("args") if isinstance(seed.get("args"), Mapping) else {}
        actions.append({
            "type": "tool_call",
            "tool": name,
            "args": dict(args),
            "title": f"补齐证据目标：{name}",
            "reason": _text(seed.get("reason") or "证据目标契约发现缺少当前问题所需 evidence，优先执行 playbook 带参工具。", 360),
            "source": "missing_evidence_goal",
            "priority": 12,
        })
        if len(actions) >= limit:
            break
    if actions:
        return actions
    for name in list(playbook_plan.get("recommended_first_tools") or [])[:limit]:
        tool = _text(name, 120)
        if not tool:
            continue
        actions.append({
            "type": "tool_call",
            "tool": tool,
            "title": f"补齐证据目标：{tool}",
            "reason": "证据目标契约发现缺少当前问题所需 evidence，建议先沿 playbook 首选工具取证。",
            "source": "missing_evidence_goal",
            "priority": 24,
        })
    return actions


def _target_repair_actions(missing_targets: Sequence[Mapping[str, object]], *, limit: int = 8) -> list[dict]:
    grouped: dict[tuple[str, str], list[str]] = {}
    for target in _mapping_items(missing_targets):
        tool = _text(target.get("repair_tool"), 120)
        arg_key = _text(target.get("repair_arg_key"), 80)
        value = _text(target.get("value"), 160)
        if not tool or not arg_key or not value:
            continue
        grouped.setdefault((tool, arg_key), []).append(value)
    actions: list[dict] = []
    for (tool, arg_key), values in grouped.items():
        deduped = list(_dedupe(values, limit=20, text_limit=160))
        if not deduped:
            continue
        if arg_key in {"queries", "refdes_list"}:
            args = {arg_key: deduped}
            if arg_key == "queries":
                args["limit_per_query"] = 10
        else:
            args = {arg_key: deduped[0]}
        actions.append({
            "type": "tool_call",
            "tool": tool,
            "args": args,
            "title": f"补齐目标对象取证：{', '.join(deduped[:3])}",
            "reason": "问题目标覆盖契约发现用户提到的对象/料号/型号尚未出现在 evidence 中，应先补齐对应目标证据。",
            "source": "missing_target_coverage",
            "priority": 10,
        })
        if len(actions) >= limit:
            break
    return actions


def build_evidence_goal_contract(*,
                                 playbook_plan: Mapping[str, object] | None = None,
                                 evidence_nodes: Sequence[Mapping[str, object]] | None = None,
                                 observations: Sequence[Mapping[str, object]] | None = None,
                                 max_goals: int = 24) -> dict:
    """Build a compact checklist of evidence types implied by selected playbooks."""

    playbook_plan = playbook_plan or {}
    all_nodes: list[Mapping[str, object]] = []
    all_nodes.extend(_mapping_items(evidence_nodes or []))
    all_nodes.extend(_evidence_nodes_from_observations(_mapping_items(observations or [])))
    present_types = set(_present_evidence_types(all_nodes))
    selected_playbooks = _mapping_items(playbook_plan.get("selected_playbooks") or [])
    goal_items: list[dict] = []
    required_all: list[object] = []
    for playbook in selected_playbooks:
        required = list(_dedupe(playbook.get("required_evidence") or [], limit=max_goals, text_limit=120))
        if not required:
            continue
        covered = [item for item in required if item in present_types]
        missing = [item for item in required if item not in present_types]
        if not missing:
            status = "satisfied"
        elif covered:
            status = "partial"
        else:
            status = "missing"
        goal_items.append({
            "playbook_id": _text(playbook.get("id"), 120),
            "title": _text(playbook.get("title") or playbook.get("id"), 180),
            "required_evidence": required,
            "covered_evidence": covered,
            "missing_evidence": missing,
            "status": status,
        })
        required_all.extend(required)

    required_types = list(_dedupe(required_all, limit=max_goals, text_limit=120))
    missing_types = [item for item in required_types if item not in present_types]
    required_targets = _targets_from_seeded_calls(playbook_plan)
    target_coverage = _target_coverage(required_targets, all_nodes)
    target_repair_actions = _target_repair_actions(target_coverage["missing_targets"])
    connection_review = _connection_review_phase_contract(
        playbook_plan,
        all_nodes,
        present_types,
    )
    connection_repair_actions = list(connection_review.get("repair_actions") or [])
    if not required_types:
        status = "not_required"
    elif not missing_types:
        status = "satisfied"
    elif len(missing_types) == len(required_types):
        status = "missing"
    else:
        status = "partial"
    repair_actions = _seed_actions(playbook_plan) if status in {"missing", "partial"} else []
    repair_actions.extend(target_repair_actions)
    repair_actions.extend(connection_repair_actions)
    return {
        "version": EVIDENCE_GOAL_CONTRACT_VERSION,
        "status": status,
        "selected_playbook_count": len(selected_playbooks),
        "required_evidence_types": required_types,
        "present_evidence_types": list(_dedupe(sorted(present_types), limit=80, text_limit=120)),
        "missing_evidence_types": missing_types,
        "target_status": target_coverage["status"],
        "required_targets": required_targets[:max_goals],
        "covered_targets": target_coverage["covered_targets"][:max_goals],
        "missing_targets": target_coverage["missing_targets"][:max_goals],
        "connection_review_phase_status": connection_review["status"],
        "connection_review_phases": connection_review["phases"][:max_goals],
        "missing_connection_review_phases": connection_review["missing_phases"][:max_goals],
        "goal_items": goal_items[:max_goals],
        "recommended_next_tools": list(_dedupe([action.get("tool") for action in repair_actions], limit=12, text_limit=120)),
        "repair_actions": repair_actions[:8],
        "connection_review_repair_actions": connection_repair_actions[:8],
        "notes": (
            "证据目标契约用于防止模型在关键 evidence 类型缺失时过早 final_answer；"
            "它不会扩大权限，只会建议当前 profile 白名单内的只读工具继续取证。"
        ),
    }
