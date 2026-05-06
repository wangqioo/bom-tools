# -*- coding: utf-8 -*-
"""Deterministic safe evidence prefetch for PSTX agent runtime.

The planner/playbook layer can produce fully-parameterized seeded tool calls.
This module selects a tiny safe subset so the runtime can gather initial
observations before the first model step, similar to a coding agent reading
the most relevant files before proposing a plan.
"""

from __future__ import annotations

import json
from collections.abc import Mapping, Sequence
from typing import Any


PREFETCH_PLAN_VERSION = "agent-prefetch-plan/v1"
GOAL_PREFETCH_PLAN_VERSION = "agent-goal-prefetch-plan/v1"
_PLACEHOLDER_MARKERS = ("<", ">", "待补充", "TODO", "todo", "示例")
_SAFE_GOAL_PREFETCH_TOOLS = {
    "list_report_tables": 10,
    "summarize_schematic_page_count": 12,
    "summarize_dfmea_readiness": 20,
    "summarize_llm_topology_netlist": 21,
    "summarize_chip_topology": 22,
    "summarize_topology_review_tasks": 23,
    "list_feishu_cache_libraries": 30,
    "list_datasheet_documents": 35,
    "list_datasheet_sources": 36,
    "list_agent_ref_sources": 40,
    "list_review_checklist_sources": 45,
    "list_document_search_sources": 48,
    "summarize_compare_risks": 50,
    "list_compare_sections": 52,
}
_TOOL_PRIORITY = {
    "summarize_schematic_page_count": 10,
    "list_datasheet_sources": 17,
    "summarize_llm_topology_netlist": 18,
    "summarize_chip_topology": 20,
    "summarize_topology_review_tasks": 21,
    "batch_query_llm_topology_netlist": 23,
    "batch_get_component_identity_cards": 24,
    "batch_match_component_datasheets": 26,
    "batch_query_chip_topology": 25,
    "batch_expand_topology_review_tasks": 27,
    "resolve_compare_page_range": 30,
    "compare_cadence_page_semantics": 35,
    "batch_query_compare_diff": 40,
    "batch_get_compare_rows": 45,
    "trace_project_source": 45,
    "search_project_text": 46,
    "batch_query_report_entities": 50,
    "summarize_table_column_values": 55,
    "search_documents": 60,
    "batch_search_documents": 65,
    "batch_search_feishu_cache_rows": 70,
    "search_datasheet_parameters": 90,
    "batch_search_datasheet_chunks": 120,
}


def _stable_json(value: Any) -> str:
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    except TypeError:
        return repr(value)


def _has_placeholder(value: Any) -> bool:
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return False
        return any(marker in text for marker in _PLACEHOLDER_MARKERS)
    if isinstance(value, Mapping):
        return any(_has_placeholder(item) for item in value.values())
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray)):
        return any(_has_placeholder(item) for item in value)
    return False


def select_seeded_prefetch_tool_calls(playbook_plan: Mapping[str, object] | None,
                                      *,
                                      allowed_tools: set[str] | Sequence[str],
                                      max_calls: int = 2,
                                      remaining_tool_calls: int | None = None,
                                      enabled: bool = True) -> dict:
    """Select safe playbook seeded tool calls for runtime prefetch.

    Selection is intentionally conservative:
    - only already-parameterized seeded calls are considered;
    - tool name must be in the current profile whitelist;
    - args must be a JSON-like mapping and must not contain placeholders;
    - calls are deduplicated and capped by the remaining tool budget.
    """

    plan = dict(playbook_plan or {})
    raw_calls = plan.get("seeded_tool_calls") or []
    allowed = set(allowed_tools or [])
    skipped: list[dict] = []
    selected: list[dict] = []
    seen: set[str] = set()
    budget = max(0, int(max_calls or 0))
    if remaining_tool_calls is not None:
        budget = min(budget, max(0, int(remaining_tool_calls or 0)))

    if not enabled:
        return {
            "version": PREFETCH_PLAN_VERSION,
            "enabled": False,
            "candidate_count": len(raw_calls) if isinstance(raw_calls, list) else 0,
            "selected_count": 0,
            "tool_calls": [],
            "skipped": [{"reason": "prefetch disabled for this profile"}],
        }

    if budget <= 0:
        return {
            "version": PREFETCH_PLAN_VERSION,
            "enabled": True,
            "candidate_count": len(raw_calls) if isinstance(raw_calls, list) else 0,
            "selected_count": 0,
            "tool_calls": [],
            "skipped": [{"reason": "no remaining tool-call budget"}],
        }

    if not isinstance(raw_calls, list):
        raw_calls = list(raw_calls) if isinstance(raw_calls, tuple) else []

    indexed_calls = list(enumerate(raw_calls))
    indexed_calls.sort(key=lambda pair: _TOOL_PRIORITY.get(
        str(pair[1].get("name") or pair[1].get("tool") or "") if isinstance(pair[1], Mapping) else "",
        100,
    ))

    for index, item in indexed_calls:
        if len(selected) >= budget:
            skipped.append({"index": index, "reason": "prefetch limit reached"})
            continue
        if not isinstance(item, Mapping):
            skipped.append({"index": index, "reason": "seed is not an object"})
            continue
        name = str(item.get("name") or item.get("tool") or "").strip()
        args = item.get("args") if isinstance(item.get("args"), Mapping) else {}
        if not name:
            skipped.append({"index": index, "reason": "missing tool name"})
            continue
        if name not in allowed:
            skipped.append({"index": index, "name": name, "reason": "tool not allowed by active profile"})
            continue
        if not isinstance(args, Mapping):
            skipped.append({"index": index, "name": name, "reason": "args must be an object"})
            continue
        args = dict(args)
        if _has_placeholder(args):
            skipped.append({"index": index, "name": name, "reason": "args contain placeholder text"})
            continue
        signature = f"{name}:{_stable_json(args)}"
        if signature in seen:
            skipped.append({"index": index, "name": name, "reason": "duplicate seeded call"})
            continue
        seen.add(signature)
        selected.append({
            "name": name,
            "args": args,
            "reason": str(item.get("reason") or "runtime playbook prefetch"),
            "source": str(item.get("source") or "runtime_prefetch"),
        })

    return {
        "version": PREFETCH_PLAN_VERSION,
        "enabled": True,
        "candidate_count": len(raw_calls),
        "selected_count": len(selected),
        "tool_calls": selected,
        "skipped": skipped[:20],
    }


def select_goal_prefetch_tool_calls(playbook_plan: Mapping[str, object] | None,
                                    *,
                                    allowed_tools: set[str] | Sequence[str],
                                    max_calls: int = 1,
                                    remaining_tool_calls: int | None = None,
                                    previous_tool_signatures: Sequence[str] | None = None,
                                    enabled: bool = True) -> dict:
    """Select a safe no-arg evidence-goal prefetch when no concrete seed exists.

    This is deliberately narrower than seeded prefetch. It only runs tools that
    are useful "overview/read-index" calls and do not require model-inferred
    arguments, giving the model real project context before it drafts a plan.
    """

    plan = dict(playbook_plan or {})
    recommended = [
        str(item or "").strip()
        for item in plan.get("recommended_first_tools") or []
        if str(item or "").strip()
    ]
    allowed = set(allowed_tools or [])
    previous = set(str(item) for item in previous_tool_signatures or [] if str(item))
    skipped: list[dict] = []
    selected: list[dict] = []
    budget = max(0, int(max_calls or 0))
    if remaining_tool_calls is not None:
        budget = min(budget, max(0, int(remaining_tool_calls or 0)))
    base = {
        "version": GOAL_PREFETCH_PLAN_VERSION,
        "enabled": bool(enabled),
        "candidate_count": len(recommended),
        "selected_count": 0,
        "tool_calls": [],
        "skipped": [],
    }
    if not enabled:
        return {**base, "skipped": [{"reason": "goal prefetch disabled"}]}
    if budget <= 0:
        return {**base, "skipped": [{"reason": "no remaining tool-call budget"}]}
    candidates = sorted(
        recommended,
        key=lambda name: _SAFE_GOAL_PREFETCH_TOOLS.get(name, 999),
    )
    seen: set[str] = set()
    for index, name in enumerate(candidates):
        if len(selected) >= budget:
            skipped.append({"index": index, "name": name, "reason": "goal prefetch limit reached"})
            continue
        if name in seen:
            skipped.append({"index": index, "name": name, "reason": "duplicate recommended tool"})
            continue
        seen.add(name)
        if name not in allowed:
            skipped.append({"index": index, "name": name, "reason": "tool not allowed by active profile"})
            continue
        if name not in _SAFE_GOAL_PREFETCH_TOOLS:
            skipped.append({"index": index, "name": name, "reason": "tool requires task-specific args"})
            continue
        signature = f"{name}:{_stable_json({})}"
        alt_signature = f"{name}::{_stable_json({})}"
        if signature in previous or alt_signature in previous:
            skipped.append({"index": index, "name": name, "reason": "goal prefetch already executed"})
            continue
        selected.append({
            "name": name,
            "args": {},
            "reason": "本地证据目标预取：先执行安全总览工具，让模型在真实项目轮廓上继续规划。",
            "source": "runtime_goal_prefetch",
        })
    return {
        **base,
        "selected_count": len(selected),
        "tool_calls": selected,
        "skipped": skipped[:20],
    }


def _first_mapping(items: Any) -> dict:
    if isinstance(items, Sequence) and not isinstance(items, (str, bytes, bytearray)):
        for item in items:
            if isinstance(item, Mapping):
                return dict(item)
    return {}


def _first_datasheet_match(result: Mapping[str, object]) -> dict:
    direct = _first_mapping(result.get("matches"))
    if direct:
        return direct
    for item in result.get("items") or []:
        if isinstance(item, Mapping) and str(item.get("status") or "") == "found":
            found = _first_mapping(item.get("matches"))
            if found:
                return found
    return {}


def _first_document_match(result: Mapping[str, object]) -> dict:
    direct = _first_mapping(result.get("matches"))
    if direct:
        return direct
    for item in result.get("items") or []:
        if isinstance(item, Mapping) and str(item.get("status") or "") == "found":
            found = _first_mapping(item.get("matches"))
            if found:
                return found
    return {}


def _first_compare_match(result: Mapping[str, object]) -> dict:
    direct = _first_mapping(result.get("matches"))
    if direct:
        return direct
    for item in result.get("items") or []:
        if isinstance(item, Mapping) and str(item.get("status") or "") == "found":
            found = _first_mapping(item.get("matches"))
            if found:
                return found
    return {}


def _first_feishu_row(result: Mapping[str, object]) -> dict:
    direct = _first_mapping(result.get("rows"))
    if direct:
        return direct
    for item in result.get("items") or []:
        if isinstance(item, Mapping) and str(item.get("status") or "") == "found":
            found = _first_mapping(item.get("rows"))
            if found:
                return found
    return {}


def _first_topology_edge_id(result: Mapping[str, object]) -> str:
    direct = _first_mapping(result.get("edges"))
    if direct.get("edge_id"):
        return str(direct.get("edge_id") or "")
    for item in result.get("items") or []:
        if not isinstance(item, Mapping):
            continue
        if isinstance(item.get("edge"), Mapping) and item["edge"].get("edge_id"):
            return str(item["edge"].get("edge_id") or "")
        for nested in item.get("items") or []:
            if isinstance(nested, Mapping) and isinstance(nested.get("edge"), Mapping) and nested["edge"].get("edge_id"):
                return str(nested["edge"].get("edge_id") or "")
    return ""


def _first_topology_review_task_id(result: Mapping[str, object]) -> str:
    direct = _first_mapping(result.get("tasks"))
    if direct.get("task_id"):
        return str(direct.get("task_id") or "")
    direct = _first_mapping(result.get("review_tasks"))
    if direct.get("task_id"):
        return str(direct.get("task_id") or "")
    evidence_cards = result.get("evidence_cards") if isinstance(result.get("evidence_cards"), Mapping) else {}
    direct = _first_mapping(evidence_cards.get("review_tasks") if isinstance(evidence_cards, Mapping) else [])
    locator = direct.get("locator") if isinstance(direct.get("locator"), Mapping) else {}
    if locator.get("task_id"):
        return str(locator.get("task_id") or "")
    return str(direct.get("id") or "")


def _candidate_detail_tool(tool_name: str, result: Mapping[str, object]) -> dict:
    if tool_name in {"search_datasheet_chunks", "batch_search_datasheet_chunks", "match_component_datasheets", "batch_match_component_datasheets"}:
        match = _first_datasheet_match(result)
        if match.get("doc_id") and match.get("chunk_id"):
            return {
                "name": "get_datasheet_chunk",
                "args": {"doc_id": int(match["doc_id"]), "chunk_id": str(match["chunk_id"]), "max_chars": 4000},
                "reason": "预取命中 datasheet chunk 后，自动读取首个片段原文以支撑定量/规格结论。",
            }
    if tool_name == "search_datasheet_parameters":
        parameter = _first_mapping(result.get("parameters"))
        if parameter.get("parameter_id"):
            return {
                "name": "get_datasheet_parameter",
                "args": {"parameter_id": int(parameter["parameter_id"]), "max_chars": 2400},
                "reason": "预取命中 datasheet 参数卡后，自动读取首个参数详情以支撑定量结论。",
            }
    if tool_name in {"search_documents", "batch_search_documents"}:
        match = _first_document_match(result)
        if match.get("doc_id"):
            return {
                "name": "get_document_excerpt",
                "args": {
                    "doc_id": str(match["doc_id"]),
                    "char_start": int(match.get("char_start") or 0),
                    "before_chars": 800,
                    "after_chars": 1600,
                    "max_chars": 5000,
                },
                "reason": "预取命中文档关键词后，自动读取首个命中段落上下文。",
            }
    if tool_name in {"query_compare_diff", "batch_query_compare_diff", "batch_get_compare_rows"}:
        match = _first_compare_match(result)
        if match.get("section_id") and match.get("row_index") is not None:
            return {
                "name": "get_compare_row",
                "args": {"section_id": str(match["section_id"]), "row_index": int(match["row_index"])},
                "reason": "预取命中对比差异后，自动读取首条差异详情。",
            }
    if tool_name == "summarize_topology_review_tasks":
        task_id = _first_topology_review_task_id(result)
        if task_id:
            return {
                "name": "get_topology_review_task",
                "args": {"task_id": task_id},
                "reason": "预取生成拓扑 review 队列后，自动读取首个高优先级任务详情。",
            }
    if tool_name in {"summarize_llm_topology_netlist", "query_llm_topology_netlist", "batch_query_llm_topology_netlist"}:
        task_id = _first_topology_review_task_id(result)
        if task_id:
            return {
                "name": "get_topology_review_task",
                "args": {"task_id": task_id},
                "reason": "预取生成 LLM 拓扑网表后，优先读取首个 review task 详情。",
            }
        edge_id = _first_topology_edge_id(result)
        if edge_id:
            return {
                "name": "get_llm_topology_edge",
                "args": {"edge_id": edge_id},
                "reason": "预取生成 LLM 拓扑网表后，自动读取首条芯片间连接详情。",
            }
    if tool_name in {"summarize_chip_topology", "query_chip_topology", "batch_query_chip_topology"}:
        edge_id = _first_topology_edge_id(result)
        if edge_id:
            return {
                "name": "get_chip_topology_edge",
                "args": {"edge_id": edge_id},
                "reason": "预取生成芯片级拓扑后，自动读取首条芯片间连接详情。",
            }
    if tool_name in {"search_feishu_cache_rows", "batch_search_feishu_cache_rows"}:
        row = _first_feishu_row(result)
        row_id = row.get("id") or row.get("row_id")
        if row_id:
            return {
                "name": "get_feishu_cache_row",
                "args": {"row_id": int(row_id)},
                "reason": "预取命中飞书缓存后，自动读取首条物料完整详情。",
            }
    return {}


def select_prefetch_followup_tool_calls(raw_observations: Sequence[Mapping[str, object]] | None,
                                        *,
                                        allowed_tools: set[str] | Sequence[str],
                                        max_calls: int = 1,
                                        remaining_tool_calls: int | None = None,
                                        previous_tool_signatures: Sequence[str] | None = None,
                                        enabled: bool = True) -> dict:
    """Select detail tools that should follow safe prefetch observations.

    This is the "open the most relevant hit" layer: after a search/batch
    prefetch, the runtime can read one concrete row/chunk/excerpt before the
    model is asked to reason over the evidence.
    """

    raw_list = [dict(item) for item in raw_observations or [] if isinstance(item, Mapping)]
    allowed = set(allowed_tools or [])
    previous = set(str(item) for item in previous_tool_signatures or [] if str(item))
    selected: list[dict] = []
    skipped: list[dict] = []
    budget = max(0, int(max_calls or 0))
    if remaining_tool_calls is not None:
        budget = min(budget, max(0, int(remaining_tool_calls or 0)))
    if not enabled:
        return {
            "version": PREFETCH_PLAN_VERSION,
            "enabled": False,
            "candidate_count": len(raw_list),
            "selected_count": 0,
            "tool_calls": [],
            "skipped": [{"reason": "follow-up disabled"}],
        }
    if budget <= 0:
        return {
            "version": PREFETCH_PLAN_VERSION,
            "enabled": True,
            "candidate_count": len(raw_list),
            "selected_count": 0,
            "tool_calls": [],
            "skipped": [{"reason": "no remaining tool-call budget"}],
        }
    for index, observation in enumerate(raw_list):
        if len(selected) >= budget:
            skipped.append({"index": index, "reason": "follow-up limit reached"})
            continue
        tool = str(observation.get("tool") or "").strip()
        result = observation.get("raw_result")
        if not tool or not isinstance(result, Mapping):
            skipped.append({"index": index, "tool": tool, "reason": "missing raw tool result"})
            continue
        candidate = _candidate_detail_tool(tool, result)
        name = str(candidate.get("name") or "").strip()
        args = candidate.get("args") if isinstance(candidate.get("args"), Mapping) else {}
        if not name:
            skipped.append({"index": index, "tool": tool, "reason": "no supported detail candidate"})
            continue
        if name not in allowed:
            skipped.append({"index": index, "tool": tool, "name": name, "reason": "detail tool not allowed"})
            continue
        if _has_placeholder(args):
            skipped.append({"index": index, "tool": tool, "name": name, "reason": "detail args contain placeholder text"})
            continue
        signature = f"{name}::{_stable_json(dict(args))}"
        if signature in previous:
            skipped.append({"index": index, "tool": tool, "name": name, "reason": "detail call already executed"})
            continue
        selected.append({
            "name": name,
            "args": dict(args),
            "reason": str(candidate.get("reason") or "runtime prefetch detail follow-up"),
            "source": "runtime_prefetch_followup",
        })
    return {
        "version": PREFETCH_PLAN_VERSION,
        "enabled": True,
        "candidate_count": len(raw_list),
        "selected_count": len(selected),
        "tool_calls": selected,
        "skipped": skipped[:20],
    }
