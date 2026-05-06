# -*- coding: utf-8 -*-
"""Quality-gate repair planning for PSTX agent runtime."""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence


SAFE_NO_ARG_TOOLS = {
    "list_report_tables",
    "summarize_schematic_page_count",
    "list_feishu_cache_libraries",
    "list_datasheet_sources",
    "list_agent_ref_sources",
    "list_review_checklist_sources",
    "list_compare_sections",
    "summarize_compare_risks",
}

REPAIR_TOOL_CALL_SOURCES = {
    "answer_missing_target_coverage",
    "answer_target_citation_missing",
    "citation_detail_required",
    "missing_connection_review_phase",
    "missing_target_coverage",
    "open_next_actions",
    "quantitative_claim_detail_required",
    "missing_evidence_goal",
}

AUTO_EXECUTE_REPAIR_SOURCES = {
    "answer_missing_target_coverage",
    "answer_target_citation_missing",
    "citation_detail_required",
    "missing_connection_review_phase",
    "missing_target_coverage",
    "quantitative_claim_detail_required",
}

OPEN_NEXT_ACTION_REASON_IDS = {
    "answer_missing_target_coverage",
    "answer_target_citation_missing",
    "citation_detail_required",
    "empty_answer",
    "low_effort_answer",
    "missing_connection_review_phase",
    "missing_evidence_goal",
    "missing_target_coverage",
    "quantitative_claim_detail_required",
}


def _text(value: object, limit: int = 240) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _has_placeholder(value: object) -> bool:
    if isinstance(value, Mapping):
        return any(_has_placeholder(item) for item in value.values())
    if isinstance(value, list):
        return any(_has_placeholder(item) for item in value)
    text = str(value or "").strip()
    return bool(text) and (
        text.startswith("<") and text.endswith(">")
        or "<需要" in text
        or "TODO" in text.upper()
    )


def _quality_reason_ids(quality_gate: Mapping[str, object] | None) -> set[str]:
    reasons = (quality_gate or {}).get("reasons") or []
    return {
        str(item.get("id") or "").strip()
        for item in reasons
        if isinstance(item, Mapping) and str(item.get("id") or "").strip()
    }


def build_quality_repair_tool_calls(quality_gate: Mapping[str, object] | None,
                                    *,
                                    allowed_tools: Iterable[str],
                                    max_calls: int = 2) -> dict:
    """Select executable local tool calls from final-answer repair actions.

    The helper is intentionally conservative: it only emits tool calls already
    allowed by the current profile, skips placeholder args, and avoids tools
    that normally require arguments unless they are explicitly safe no-arg
    status/list tools.
    """

    allowed = {str(item) for item in allowed_tools or [] if str(item)}
    actions = [
        dict(item)
        for item in (quality_gate or {}).get("repair_actions") or []
        if isinstance(item, Mapping)
    ]
    budget = max(0, int(max_calls or 0))
    if budget <= 0:
        return {
            "version": "quality-repair-plan/v1",
            "candidate_action_count": len([item for item in actions if str(item.get("type") or "") == "tool_call"]),
            "selected_tool_call_count": 0,
            "tool_calls": [],
            "skipped_actions": [{"reason": "max_calls_exhausted"}],
        }
    selected: list[dict] = []
    skipped: list[dict] = []
    seen = set()
    reason_ids = _quality_reason_ids(quality_gate)
    allow_open_next_actions = bool(reason_ids & OPEN_NEXT_ACTION_REASON_IDS)
    for action in sorted(actions, key=lambda item: int(item.get("priority") or 50)):
        if str(action.get("type") or "") != "tool_call":
            continue
        source = _text(action.get("source"), 120)
        is_incomplete_result = source.startswith("incomplete_tool_result")
        is_open_next_action = source in REPAIR_TOOL_CALL_SOURCES
        if not (is_incomplete_result or is_open_next_action):
            skipped.append({"tool": _text(action.get("tool"), 120), "title": _text(action.get("title")), "reason": "non_evidence_repair"})
            continue
        if is_open_next_action and not allow_open_next_actions:
            skipped.append({"tool": _text(action.get("tool"), 120), "title": _text(action.get("title")), "reason": "open_next_not_needed"})
            continue
        tool = _text(action.get("tool"), 120)
        args = action.get("args") if isinstance(action.get("args"), Mapping) else {}
        key = (tool, str(dict(args)))
        if not tool:
            skipped.append({"title": _text(action.get("title")), "reason": "missing_tool"})
            continue
        if tool not in allowed:
            skipped.append({"tool": tool, "title": _text(action.get("title")), "reason": "not_allowed"})
            continue
        if key in seen:
            skipped.append({"tool": tool, "title": _text(action.get("title")), "reason": "duplicate"})
            continue
        if not args and tool not in SAFE_NO_ARG_TOOLS:
            skipped.append({"tool": tool, "title": _text(action.get("title")), "reason": "missing_args"})
            continue
        if _has_placeholder(args):
            skipped.append({"tool": tool, "title": _text(action.get("title")), "reason": "placeholder_args"})
            continue
        seen.add(key)
        selected.append({
            "name": tool,
            "args": dict(args),
            "reason": _text(action.get("reason") or action.get("title") or "quality repair", 360),
            "source": source,
        })
        if len(selected) >= budget:
            break
    return {
        "version": "quality-repair-plan/v1",
        "candidate_action_count": len([item for item in actions if str(item.get("type") or "") == "tool_call"]),
        "selected_tool_call_count": len(selected),
        "tool_calls": selected,
        "skipped_actions": skipped[:12],
    }


def filter_auto_quality_repair_tool_calls(repair_plan: Mapping[str, object] | None,
                                          quality_gate: Mapping[str, object] | None = None,
                                          provided_citation_count: int = 0) -> dict:
    """Return the subset of repair calls safe enough for the live agent loop.

    The full repair plan is useful for trace/replay, but automatically executing
    every warning can make deterministic providers loop or chase broad playbook
    gaps after an otherwise usable answer. Auto execution is therefore limited
    to precise repairs: detail evidence, quantitative spec grounding, and
    concrete target coverage.
    """

    plan = dict(repair_plan or {})
    valid_citation_count = int((quality_gate or {}).get("valid_citation_count") or 0)
    provided_citation_count = int(provided_citation_count or 0)
    goal_contract = (quality_gate or {}).get("evidence_goal_contract")
    if not isinstance(goal_contract, Mapping):
        goal_contract = {}
    required_target_count = len([item for item in goal_contract.get("required_targets") or [] if isinstance(item, Mapping)])
    selected = []
    skipped = list(plan.get("skipped_actions") or [])
    for call in plan.get("tool_calls") or []:
        if not isinstance(call, Mapping):
            continue
        source = _text(call.get("source"), 120)
        if source.startswith("incomplete_tool_result") and valid_citation_count > 0 and provided_citation_count > 0:
            skipped.append({
                "tool": _text(call.get("name") or call.get("tool"), 120),
                "reason": "incomplete_result_has_valid_citation",
                "source": source,
            })
        elif (
            source == "missing_target_coverage"
            and valid_citation_count > 0
            and provided_citation_count > 0
            and required_target_count <= 1
        ):
            skipped.append({
                "tool": _text(call.get("name") or call.get("tool"), 120),
                "reason": "single_target_has_valid_citation",
                "source": source,
            })
        elif source.startswith("incomplete_tool_result") or source in AUTO_EXECUTE_REPAIR_SOURCES:
            selected.append(dict(call))
        else:
            skipped.append({
                "tool": _text(call.get("name") or call.get("tool"), 120),
                "reason": "trace_only_repair",
                "source": source,
            })
    plan["tool_calls"] = selected
    plan["selected_tool_call_count"] = len(selected)
    plan["auto_filtered"] = True
    plan["skipped_actions"] = skipped[:12]
    return plan
