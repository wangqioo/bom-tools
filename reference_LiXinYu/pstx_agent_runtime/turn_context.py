# -*- coding: utf-8 -*-
"""Codex-inspired turn context and tool dispatch summaries for harness runs."""

from __future__ import annotations

import json
from collections import Counter
from collections.abc import Iterable, Mapping, Sequence
from typing import Any


TURN_CONTEXT_SCHEMA_VERSION = "pstx-harness-turn-context.v1"
TOOL_DISPATCH_TRACE_SCHEMA_VERSION = "pstx-tool-dispatch-trace.v1"
TOOL_DISPATCH_SUMMARY_SCHEMA_VERSION = "pstx-tool-dispatch-summary.v1"


def _text(value: object, limit: int = 240) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _stable_json_chars(value: object) -> int:
    try:
        return len(json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")))
    except (TypeError, ValueError):
        return len(str(value))


def _ids(values: Iterable[object] | None, *, limit: int = 20) -> list[str]:
    result: list[str] = []
    for value in values or []:
        item = _text(value, 160)
        if not item or item in result:
            continue
        result.append(item)
        if len(result) >= limit:
            break
    return result


def _mapping(value: object) -> dict:
    return dict(value) if isinstance(value, Mapping) else {}


def _selected_skill_ids(selected_skills: Mapping[str, object] | None) -> list[str]:
    skills = _mapping(selected_skills)
    selected = []
    for item in skills.get("selected") or skills.get("skills") or []:
        if isinstance(item, Mapping):
            selected.append(item.get("id") or item.get("name"))
        else:
            selected.append(item)
    return _ids(selected, limit=20)


def _playbook_ids(playbook_plan: Mapping[str, object] | None) -> list[str]:
    plan = _mapping(playbook_plan)
    values: list[object] = []
    values.extend(plan.get("playbook_ids") or [])
    values.extend(plan.get("selected_playbooks") or [])
    if plan.get("playbook_id"):
        values.append(plan.get("playbook_id"))
    if plan.get("id"):
        values.append(plan.get("id"))
    return _ids(values, limit=16)


def _tool_boundary(tool_list: Sequence[Mapping[str, object]] | None,
                   allowed_tools: Iterable[object] | None) -> dict:
    allowed = set(_ids(allowed_tools, limit=1000))
    readonly_false: list[str] = []
    file_access: list[str] = []
    mutating: list[str] = []
    parallel_capable: list[str] = []
    approval_scopes: Counter[str] = Counter()
    evidence_kinds: Counter[str] = Counter()
    for tool in tool_list or []:
        if not isinstance(tool, Mapping):
            continue
        name = _text(tool.get("name"), 160)
        if not name or (allowed and name not in allowed):
            continue
        if tool.get("readonly") is False:
            readonly_false.append(name)
        if tool.get("file_access"):
            file_access.append(name)
        if tool.get("mutating"):
            mutating.append(name)
        if tool.get("supports_parallel"):
            parallel_capable.append(name)
        scope = _text(tool.get("approval_scope") or ("read_project_file" if tool.get("file_access") else "none"), 80)
        kind = _text(tool.get("evidence_kind") or tool.get("target") or "general", 80)
        approval_scopes[scope or "none"] += 1
        evidence_kinds[kind or "general"] += 1
    return {
        "allowed_tool_count": len(allowed),
        "allowed_tools": sorted(allowed)[:80],
        "allowed_tools_truncated": len(allowed) > 80,
        "readonly": not readonly_false,
        "non_readonly_tools": sorted(readonly_false)[:20],
        "file_access_tools": sorted(file_access)[:40],
        "file_access_tool_count": len(file_access),
        "mutating": bool(mutating),
        "mutating_tools": sorted(mutating)[:20],
        "parallel_capable_tools": sorted(parallel_capable)[:40],
        "parallel_capable_tool_count": len(parallel_capable),
        "approval_scopes": dict(sorted(approval_scopes.items())),
        "evidence_kinds": dict(sorted(evidence_kinds.items())),
    }


def compact_dispatch_args(args: Mapping[str, object] | None, *, debug: bool) -> dict:
    payload = dict(args or {})
    compact = {
        "arg_keys": sorted(str(key) for key in payload.keys())[:40],
        "arg_count": len(payload),
        "args_json_chars": _stable_json_chars(payload),
    }
    if debug:
        compact["args"] = payload
    return compact


def build_tool_dispatch_event(*,
                              event_index: int,
                              tool: object,
                              args: Mapping[str, object] | None,
                              status: str,
                              call_id: object = "",
                              reason: object = "",
                              profile_label: object = "",
                              capability_profiles: Sequence[object] = (),
                              signature: object = "",
                              batch: bool = False,
                              duplicate: bool = False,
                              allowed: bool = True,
                              debug: bool = False,
                              call_index: object = None,
                              evidence_ids: Iterable[object] | None = None,
                              contract: Mapping[str, object] | None = None,
                              tool_metadata: Mapping[str, object] | None = None,
                              preflight_status: object = "",
                              error: object = "",
                              duration_ms: object = None,
                              raw_result_json_chars: object = None) -> dict:
    event = {
        "schema_version": TOOL_DISPATCH_TRACE_SCHEMA_VERSION,
        "event_index": int(event_index or 0),
        "call_id": _text(call_id, 120),
        "status": _text(status, 60),
        "tool": _text(tool, 160),
        "allowed": bool(allowed),
        "duplicate": bool(duplicate),
        "batch": bool(batch),
        "reason": _text(reason, 360),
        "profile": _text(profile_label, 120),
        "capability_profiles": _ids(capability_profiles, limit=20),
        "signature": _text(signature, 260),
        **compact_dispatch_args(args, debug=debug),
    }
    if call_index is not None:
        event["call_index"] = call_index
    ids = _ids(evidence_ids, limit=24)
    if ids:
        event["evidence_node_ids"] = ids
    contract_map = _mapping(contract)
    if contract_map:
        event["contract"] = {
            "completeness": _text(contract_map.get("completeness"), 80),
            "recommended_next_tools": _ids(contract_map.get("recommended_next_tools") or (), limit=10),
            "detail_tool": _mapping(contract_map.get("detail_tool")),
            "aggregation_tool": _mapping(contract_map.get("aggregation_tool")),
        }
    metadata = _mapping(tool_metadata)
    if metadata:
        event["tool_boundary"] = {
            "readonly": metadata.get("readonly") is not False,
            "file_access": bool(metadata.get("file_access")),
            "mutating": bool(metadata.get("mutating")),
            "approval_scope": _text(metadata.get("approval_scope") or "none", 80),
            "evidence_kind": _text(metadata.get("evidence_kind") or "general", 80),
        }
    if error:
        event["error"] = _text(error, 500)
    if preflight_status:
        event["preflight_status"] = _text(preflight_status, 80)
    if duration_ms is not None:
        try:
            event["duration_ms"] = round(float(duration_ms), 3)
        except (TypeError, ValueError):
            event["duration_ms"] = 0.0
    if raw_result_json_chars is not None:
        try:
            event["raw_result_json_chars"] = int(raw_result_json_chars)
        except (TypeError, ValueError):
            event["raw_result_json_chars"] = _stable_json_chars(raw_result_json_chars)
    return event


def summarize_tool_dispatch_trace(trace: Sequence[Mapping[str, object]] | None) -> dict:
    events = [item for item in trace or [] if isinstance(item, Mapping)]
    status_counts = Counter(_text(item.get("status"), 60) or "unknown" for item in events)
    tool_counts = Counter(_text(item.get("tool"), 160) or "unknown" for item in events)
    preflight_counts = Counter(_text(item.get("preflight_status"), 60) or "unknown" for item in events)
    durations = []
    for item in events:
        try:
            durations.append(max(0.0, float(item.get("duration_ms") or 0)))
        except (TypeError, ValueError):
            durations.append(0.0)
    slowest_index = max(range(len(events)), key=lambda idx: durations[idx], default=None)
    slowest_event = events[slowest_index] if slowest_index is not None else {}
    blocked_statuses = {"blocked", "duplicate", "failed", "limit"}
    return {
        "schema_version": TOOL_DISPATCH_SUMMARY_SCHEMA_VERSION,
        "event_count": len(events),
        "status_counts": dict(status_counts),
        "tool_counts": dict(tool_counts),
        "preflight_status_counts": dict(preflight_counts),
        "completed_count": status_counts.get("completed", 0),
        "blocked_count": sum(status_counts.get(status, 0) for status in blocked_statuses),
        "duplicate_count": status_counts.get("duplicate", 0),
        "failed_count": status_counts.get("failed", 0),
        "limit_count": status_counts.get("limit", 0),
        "preflight_failed_count": preflight_counts.get("failed", 0),
        "file_access_call_count": sum(
            1
            for item in events
            if isinstance(item.get("tool_boundary"), Mapping) and item["tool_boundary"].get("file_access")
        ),
        "duration_ms_total": round(sum(durations), 3),
        "duration_ms_max": round(max(durations, default=0.0), 3),
        "slowest_tool": _text(slowest_event.get("tool"), 160) if slowest_event else "",
        "slowest_call_id": _text(slowest_event.get("call_id"), 120) if slowest_event else "",
        "unique_tool_count": len(tool_counts),
    }


def build_harness_turn_context_snapshot(*,
                                        agent_run_id: object = "",
                                        mode: object = "",
                                        profile: object = "",
                                        capability_profiles: Sequence[object] = (),
                                        model_provider: object = "",
                                        model_mode: object = "",
                                        guidance_summary: Mapping[str, object] | None = None,
                                        selected_skills: Mapping[str, object] | None = None,
                                        playbook_plan: Mapping[str, object] | None = None,
                                        allowed_tools: Iterable[object] | None = None,
                                        tool_list: Sequence[Mapping[str, object]] | None = None,
                                        context_budget: Mapping[str, object] | None = None,
                                        runtime_state: Mapping[str, object] | None = None,
                                        limits: Mapping[str, object] | None = None,
                                        safeguards: Sequence[object] = (),
                                        source: object = "pstx-harness") -> dict:
    guidance = _mapping(guidance_summary)
    skills = _mapping(selected_skills)
    budget = _mapping(context_budget)
    runtime = _mapping(runtime_state)
    ledger = _mapping(runtime.get("task_ledger"))
    goal_contract = _mapping(runtime.get("evidence_goal_contract"))
    memory = _mapping(runtime.get("memory_summary"))
    tool_boundary = _tool_boundary(tool_list, allowed_tools)
    return {
        "schema_version": TURN_CONTEXT_SCHEMA_VERSION,
        "source": _text(source, 120),
        "agent_run_id": _text(agent_run_id, 160),
        "mode": _text(mode, 120),
        "profile": _text(profile, 120),
        "capability_profiles": _ids(capability_profiles, limit=20),
        "model": {
            "provider": _text(model_provider, 120),
            "mode": _text(model_mode, 120),
        },
        "guidance": {
            "source_count": int(guidance.get("source_count") or 0),
            "sources": _ids(guidance.get("sources") or guidance.get("source_paths") or (), limit=12),
            "truncated": bool(guidance.get("truncated")),
        },
        "skills": {
            "selected_count": int(skills.get("selected_count") or len(_selected_skill_ids(skills))),
            "selected_ids": _selected_skill_ids(skills),
        },
        "playbooks": {
            "ids": _playbook_ids(playbook_plan),
            "recommended_first_tools": _ids((_mapping(playbook_plan)).get("recommended_first_tools") or (), limit=16),
            "planner_warning_count": len((_mapping(playbook_plan)).get("planner_warnings") or []),
        },
        "tool_boundary": tool_boundary,
        "context_budget": {
            "truncated": bool(budget.get("truncated")),
            "model_observation_json_chars": int(budget.get("model_observation_json_chars") or 0),
            "source_observation_count": int(budget.get("source_observation_count") or 0),
            "model_observation_count": int(budget.get("model_observation_count") or 0),
        },
        "runtime_state": {
            "protocol_version": _text(runtime.get("protocol_version"), 120),
            "evidence_id_count": int(runtime.get("evidence_id_count") or 0),
            "memory_fact_count": len(memory.get("facts") or []),
            "task_ledger_progress": _mapping(ledger.get("progress")),
            "evidence_goal_status": _text(goal_contract.get("status"), 80),
            "missing_evidence_goal_count": len(goal_contract.get("missing_evidence_types") or []),
        },
        "limits": dict(limits or {}),
        "safeguards": [_text(item, 260) for item in safeguards or [] if _text(item, 260)][:12],
    }
