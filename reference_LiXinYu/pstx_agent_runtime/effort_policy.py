# -*- coding: utf-8 -*-
"""Effort policy helpers for preventing premature agent abandonment."""

from __future__ import annotations

from typing import Mapping, Sequence

from .finalizer import build_perseverance_retry_note, is_low_effort_answer


EFFORT_POLICY_VERSION = "pstx-effort-policy/v1"


def _recommended_tools(playbook_plan: Mapping[str, object] | None,
                       tool_result_contracts: Sequence[Mapping[str, object]] | None,
                       task_ledger: Mapping[str, object] | None) -> list[str]:
    result: list[str] = []
    if isinstance(playbook_plan, Mapping):
        for name in playbook_plan.get("recommended_first_tools") or []:
            text = str(name or "").strip()
            if text and text not in result:
                result.append(text)
    for contract in tool_result_contracts or []:
        if not isinstance(contract, Mapping):
            continue
        for name in contract.get("recommended_next_tools") or []:
            text = str(name or "").strip()
            if text and text not in result:
                result.append(text)
        for key in ("detail_tool", "aggregation_tool"):
            tool = contract.get(key)
            if isinstance(tool, Mapping):
                text = str(tool.get("name") or "").strip()
                if text and text not in result:
                    result.append(text)
    if isinstance(task_ledger, Mapping):
        for action in task_ledger.get("next_actions") or []:
            if not isinstance(action, Mapping):
                continue
            text = str(action.get("tool") or "").strip()
            if text and text not in result:
                result.append(text)
    return result[:12]


def build_effort_policy_state(*,
                              step_type: str = "",
                              answer: object = "",
                              tool_call_count: int = 0,
                              max_tool_calls: int = 0,
                              playbook_plan: Mapping[str, object] | None = None,
                              tool_result_contracts: Sequence[Mapping[str, object]] | None = None,
                              task_ledger: Mapping[str, object] | None = None,
                              evidence_node_count: int = 0,
                              citation_count: int = 0,
                              allow_needs_user_input: bool = True,
                              retry_limit: int = 2,
                              retry_count: int = 0) -> dict:
    retry_note = build_perseverance_retry_note(
        step_type=step_type,
        answer=answer,
        tool_call_count=tool_call_count,
        max_tool_calls=max_tool_calls,
        playbook_plan=playbook_plan,
        tool_result_contracts=tool_result_contracts,
        task_ledger=task_ledger,
        evidence_node_count=evidence_node_count,
        citation_count=citation_count,
        allow_needs_user_input=allow_needs_user_input,
    )
    recommended = _recommended_tools(playbook_plan, tool_result_contracts, task_ledger)
    return {
        "version": EFFORT_POLICY_VERSION,
        "mode": "try_safe_tools_then_ask",
        "step_type": str(step_type or ""),
        "low_effort_answer": is_low_effort_answer(answer) if step_type == "final_answer" else False,
        "recommended_tools": recommended,
        "has_safe_next_step": bool(recommended),
        "retry_limit": max(0, int(retry_limit or 0)),
        "retry_count": max(0, int(retry_count or 0)),
        "retry_available": bool(retry_note) and max(0, int(retry_count or 0)) < max(0, int(retry_limit or 0)),
        "retry_note": retry_note,
        "rules": [
            "不得在未尝试安全只读工具前直接说信息不足。",
            "preview/partial/truncated 结果不能直接支撑全量统计结论。",
            "仍有 detail/aggregation/next action 时优先继续取证。",
            "确实缺少外部信息时返回 waiting_for_user 并列出具体问题。",
        ],
    }


def build_effort_policy_retry_note(**kwargs) -> str:
    return str(build_effort_policy_state(**kwargs).get("retry_note") or "")
