# -*- coding: utf-8 -*-
"""Runtime state helpers for PSTX agent memory and todo summaries."""

from __future__ import annotations

from typing import Mapping, Sequence

from .protocol import AgentTodoItem, AgentTodoList, MemorySummary, PROTOCOL_VERSION
from .goal_contract import build_evidence_goal_contract
from .task_ledger import build_task_ledger, todo_list_from_task_ledger


def _text(value: object, limit: int = 220) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _dedupe(items: Sequence[object], *, limit: int, text_limit: int = 220) -> tuple[str, ...]:
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


def _evidence_ids_from_observations(observations: Sequence[Mapping[str, object]], *, limit: int = 80) -> tuple[str, ...]:
    ids: list[str] = []
    for observation in observations or []:
        if not isinstance(observation, Mapping):
            continue
        for evidence_id in observation.get("evidence_node_ids", []) or []:
            ids.append(evidence_id)
        for node in observation.get("evidence_nodes", []) or []:
            if isinstance(node, Mapping):
                ids.append(node.get("id", ""))
    return _dedupe(ids, limit=limit, text_limit=120)


def _facts_from_observations(observations: Sequence[Mapping[str, object]], *, limit: int = 16) -> tuple[str, ...]:
    facts: list[str] = []
    for observation in observations or []:
        if not isinstance(observation, Mapping):
            continue
        tool = _text(observation.get("tool", ""), 80)
        summary = _text(observation.get("summary", ""), 220)
        if tool or summary:
            facts.append(f"{tool}: {summary}" if tool and summary else tool or summary)
    return _dedupe(facts[-limit:], limit=limit, text_limit=260)


def _answers_from_context(project_context: Mapping[str, object]) -> tuple[str, ...]:
    facts: list[str] = []
    for item in list(project_context.get("answers") or [])[-12:]:
        if not isinstance(item, Mapping):
            continue
        question_id = _text(item.get("question_id", ""), 80)
        answer = _text(item.get("answer", ""), 260)
        if answer:
            facts.append(f"{question_id}: {answer}" if question_id else answer)
    return _dedupe(facts, limit=12, text_limit=320)


def _open_questions_from_context(project_context: Mapping[str, object]) -> tuple[str, ...]:
    questions: list[str] = []
    for item in list(project_context.get("pending_questions") or [])[:12]:
        if not isinstance(item, Mapping):
            continue
        question = _text(item.get("question", ""), 260)
        fields = ",".join(_text(field, 60) for field in list(item.get("missing_fields") or [])[:6])
        if question and fields:
            questions.append(f"{question}（缺失：{fields}）")
        elif question:
            questions.append(question)
    return _dedupe(questions, limit=12, text_limit=320)


def _project_session_memory(project_context: Mapping[str, object]) -> Mapping[str, object]:
    memory = project_context.get("session_memory_summary")
    return memory if isinstance(memory, Mapping) else {}


def _todo_from_capabilities(goal: str,
                            capability_plan: Sequence[Mapping[str, object]],
                            observations: Sequence[Mapping[str, object]]) -> AgentTodoList:
    observed_count = len([item for item in observations or [] if isinstance(item, Mapping)])
    items: list[AgentTodoItem] = []
    plans = [item for item in capability_plan or [] if isinstance(item, Mapping)]
    if not plans:
        plans = [{"id": "quick_scan", "title": "快速证据收集"}]
    for index, plan in enumerate(plans[:8], start=1):
        title = _text(plan.get("title") or plan.get("id") or f"任务 {index}", 160)
        status = "completed" if observed_count and index == 1 else ("in_progress" if index == 1 else "pending")
        note = "已有工具观察，继续按需读取细节。" if status == "completed" else _text(plan.get("description", ""), 220)
        items.append(AgentTodoItem(
            id=f"todo-{index}",
            title=title,
            status=status,
            note=note,
        ))
    return AgentTodoList(goal=_text(goal, 500), items=tuple(items))


def build_runtime_state(*,
                        goal: object,
                        capability_plan: Sequence[Mapping[str, object]] = (),
                        playbook_plan: Mapping[str, object] | None = None,
                        observations: Sequence[Mapping[str, object]] = (),
                        tool_result_contracts: Sequence[Mapping[str, object]] = (),
                        project_context: Mapping[str, object] | None = None,
                        truncated: bool = False) -> dict:
    """Build a compact local runtime state card for model prompts and traces."""

    project_context = project_context or {}
    observation_items = [item for item in observations or [] if isinstance(item, Mapping)]
    evidence_ids = _evidence_ids_from_observations(observation_items)
    observation_facts = _facts_from_observations(observation_items)
    context_facts = _answers_from_context(project_context)
    project_memory = _project_session_memory(project_context)
    memory = MemorySummary.from_parts(
        goal=goal,
        facts=(
            list(project_memory.get("facts") or [])
            + list(observation_facts)
            + [f"用户补充：{item}" for item in context_facts]
        ),
        decisions=list(project_memory.get("decisions") or []),
        open_questions=list(project_memory.get("open_questions") or []) + list(_open_questions_from_context(project_context)),
        evidence_ids=list(project_memory.get("evidence_ids") or []) + list(evidence_ids),
    )
    task_ledger = build_task_ledger(
        goal=goal,
        capability_plan=capability_plan,
        playbook_plan=playbook_plan,
        observations=observation_items,
        tool_result_contracts=tool_result_contracts,
        project_context=project_context,
    )
    evidence_goal_contract = build_evidence_goal_contract(
        playbook_plan=playbook_plan,
        observations=observation_items,
    )
    todo = todo_list_from_task_ledger(task_ledger) or _todo_from_capabilities(str(goal or ""), capability_plan, observation_items)
    return {
        "protocol_version": PROTOCOL_VERSION,
        "todo_list": todo.to_dict(),
        "task_ledger": task_ledger,
        "evidence_goal_contract": evidence_goal_contract,
        "memory_summary": memory.to_dict(),
        "observation_count": len(observation_items),
        "evidence_id_count": len(evidence_ids),
        "context_answer_count": len(project_context.get("answers") or []),
        "pending_question_count": len(project_context.get("pending_questions") or []),
        "truncated": bool(truncated),
        "notes": (
            "runtime_state 为本地压缩任务记忆；完整表格、文件和 evidence 细节需通过工具读取。"
            if truncated else
            "runtime_state 在当前预算内。"
        ),
    }
