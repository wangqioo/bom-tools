# -*- coding: utf-8 -*-
"""Lightweight per-run session state for PSTX agent prompts and traces."""

from __future__ import annotations

from dataclasses import dataclass, field
import json
from collections.abc import Mapping, Sequence

from .protocol import PROTOCOL_VERSION


PROJECT_SESSION_MEMORY_VERSION = "agent-project-session-memory/v1"


def _text(value: object, limit: int = 500) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _dedupe_text(items: Sequence[object], *, limit: int, text_limit: int = 180) -> tuple[str, ...]:
    output: list[str] = []
    seen = set()
    for item in items or []:
        text = _text(item, text_limit)
        key = text.lower()
        if not text or key in seen:
            continue
        output.append(text)
        seen.add(key)
        if len(output) >= limit:
            break
    return tuple(output)


def _mapping_items(items: Sequence[object]) -> list[Mapping[str, object]]:
    return [item for item in items or [] if isinstance(item, Mapping)]


def _args_preview(args: object, *, limit: int = 180) -> str:
    if not isinstance(args, Mapping) or not args:
        return ""
    try:
        text = json.dumps(dict(args), ensure_ascii=False, sort_keys=True, default=str)
    except (TypeError, ValueError):
        text = str(dict(args))
    return _text(text, limit)


def _strings_from_mapping_items(items: Sequence[Mapping[str, object]],
                                keys: Sequence[str],
                                *,
                                limit: int,
                                text_limit: int = 240) -> tuple[str, ...]:
    values: list[str] = []
    for item in items:
        for key in keys:
            if item.get(key):
                values.append(_text(item.get(key), text_limit))
                break
        if len(values) >= limit:
            break
    return _dedupe_text(values, limit=limit, text_limit=text_limit)


def _recent_evidence_ids(runtime_state: Mapping[str, object],
                         observations: Sequence[Mapping[str, object]],
                         *,
                         limit: int = 40) -> tuple[str, ...]:
    ids: list[object] = []
    memory = runtime_state.get("memory_summary") if isinstance(runtime_state.get("memory_summary"), Mapping) else {}
    ids.extend(memory.get("evidence_ids") or [])
    for observation in observations or []:
        if not isinstance(observation, Mapping):
            continue
        ids.extend(observation.get("evidence_node_ids") or [])
        for node in observation.get("evidence_nodes") or []:
            if isinstance(node, Mapping):
                ids.append(node.get("id"))
    return _dedupe_text(ids, limit=limit, text_limit=120)


def _pending_questions(project_context: Mapping[str, object],
                       runtime_state: Mapping[str, object],
                       *,
                       limit: int = 12) -> tuple[dict, ...]:
    questions: list[dict] = []
    for item in list(project_context.get("pending_questions") or [])[:limit]:
        if not isinstance(item, Mapping):
            continue
        question = _text(item.get("question") or item.get("prompt") or item.get("title"), 360)
        if not question:
            continue
        questions.append({
            "question_id": _text(item.get("question_id") or item.get("id") or f"pending-{len(questions) + 1}", 120),
            "question": question,
            "missing_fields": list(_dedupe_text(item.get("missing_fields") or [], limit=12, text_limit=80)),
            "related_evidence_ids": list(_dedupe_text(item.get("related_evidence_ids") or [], limit=12, text_limit=120)),
        })
    if questions:
        return tuple(questions[:limit])
    memory = runtime_state.get("memory_summary") if isinstance(runtime_state.get("memory_summary"), Mapping) else {}
    for index, question in enumerate(_dedupe_text(memory.get("open_questions") or [], limit=limit, text_limit=360), start=1):
        questions.append({
            "question_id": f"runtime-open-question-{index}",
            "question": question,
            "missing_fields": [],
            "related_evidence_ids": [],
        })
    return tuple(questions[:limit])


def merge_project_session_memory(project_context: Mapping[str, object] | None,
                                 result: Mapping[str, object] | None,
                                 *,
                                 max_facts: int = 24,
                                 max_open_questions: int = 16,
                                 max_open_items: int = 16,
                                 max_next_actions: int = 16,
                                 max_evidence_ids: int = 80) -> dict:
    """Merge one agent result into the run-scoped rolling project memory.

    This memory is a compact navigation aid for the next prompt. It deliberately
    stores summaries and evidence ids, not raw observations or full table/PDF/file
    content.
    """

    project_context = project_context or {}
    result = result or {}
    previous = project_context.get("session_memory_summary")
    previous = previous if isinstance(previous, Mapping) else {}
    runtime_state = result.get("runtime_state") if isinstance(result.get("runtime_state"), Mapping) else {}
    runtime_memory = runtime_state.get("memory_summary") if isinstance(runtime_state.get("memory_summary"), Mapping) else {}
    task_ledger = runtime_state.get("task_ledger") if isinstance(runtime_state.get("task_ledger"), Mapping) else {}
    continuation_pack = result.get("continuation_pack") if isinstance(result.get("continuation_pack"), Mapping) else {}
    quality_gate = result.get("final_answer_quality_gate") if isinstance(result.get("final_answer_quality_gate"), Mapping) else {}
    agent_run_id = _text(result.get("agent_run_id"), 120)

    facts: list[object] = []
    facts.extend(previous.get("facts") or [])
    facts.extend(runtime_memory.get("facts") or [])
    answer = _text(result.get("answer"), 260)
    if answer:
        facts.append(f"最近回答：{answer}")
    if continuation_pack.get("continuation_brief"):
        facts.append(f"交接摘要：{continuation_pack.get('continuation_brief')}")

    decisions: list[object] = []
    decisions.extend(previous.get("decisions") or [])
    decisions.extend(runtime_memory.get("decisions") or [])
    status = _text(result.get("status"), 60)
    stopped = _text((result.get("model_metadata") or {}).get("stopped_reason") if isinstance(result.get("model_metadata"), Mapping) else "", 100)
    if status or stopped:
        decisions.append(f"最近运行状态：{status or 'unknown'}；停止原因：{stopped or 'unknown'}")
    if quality_gate.get("status"):
        decisions.append(f"质量门禁：{quality_gate.get('status')}，score={quality_gate.get('score', 0)}")

    open_questions: list[object] = []
    open_questions.extend(previous.get("open_questions") or [])
    open_questions.extend(runtime_memory.get("open_questions") or [])
    needs_user_input = result.get("needs_user_input") if isinstance(result.get("needs_user_input"), Mapping) else {}
    for item in _mapping_items(needs_user_input.get("questions") or []):
        question = _text(item.get("question"), 320)
        fields = ",".join(_text(field, 60) for field in list(item.get("missing_fields") or [])[:8])
        if question and fields:
            open_questions.append(f"{question}（缺失：{fields}）")
        elif question:
            open_questions.append(question)
    open_questions.extend(continuation_pack.get("pending_questions") or [])

    open_items: list[object] = []
    open_items.extend(previous.get("open_items") or [])
    for item in _mapping_items(task_ledger.get("items") or []):
        if str(item.get("status") or "") in {"pending", "in_progress", "blocked"}:
            title = _text(item.get("title") or item.get("id"), 220)
            note = _text(item.get("note") or item.get("blocking_reason") or item.get("source"), 220)
            if title and note:
                open_items.append(f"{title}：{note}")
            elif title:
                open_items.append(title)
    open_items.extend(_strings_from_mapping_items(_mapping_items(continuation_pack.get("open_ledger_items") or []), ("title", "id"), limit=max_open_items))

    next_actions: list[object] = []
    next_actions.extend(previous.get("next_actions") or [])
    for action in _mapping_items(task_ledger.get("next_actions") or []):
        title = _text(action.get("title") or action.get("tool") or action.get("type"), 180)
        tool = _text(action.get("tool"), 120)
        reason = _text(action.get("reason"), 220)
        args = _args_preview(action.get("args"))
        if title and tool:
            arg_text = f"，args={args}" if args else ""
            next_actions.append(f"{title}（tool={tool}{arg_text}）：{reason}")
        elif title:
            next_actions.append(title)
    for item in _mapping_items(continuation_pack.get("suggested_tool_calls") or []):
        name = _text(item.get("tool") or item.get("name"), 120)
        title = _text(item.get("title") or name, 180)
        args = _args_preview(item.get("args"))
        if name and args:
            next_actions.append(f"{title}（tool={name}，args={args}）")
        elif name:
            next_actions.append(f"{title}（tool={name}）")

    evidence_ids: list[object] = []
    evidence_ids.extend(previous.get("evidence_ids") or [])
    evidence_ids.extend(runtime_memory.get("evidence_ids") or [])
    evidence_ids.extend(task_ledger.get("evidence_ids") or [])
    evidence_ids.extend(continuation_pack.get("evidence_ids") or [])
    for node in _mapping_items(result.get("final_evidence") or []):
        evidence_ids.append(node.get("id"))

    source_runs = list(previous.get("source_agent_run_ids") or [])
    if agent_run_id:
        source_runs.append(agent_run_id)

    goal = (
        _text(continuation_pack.get("goal"), 500)
        or _text(runtime_memory.get("goal"), 500)
        or _text(previous.get("goal"), 500)
        or _text((result.get("request") or {}).get("question") if isinstance(result.get("request"), Mapping) else "", 500)
    )

    return {
        "version": PROJECT_SESSION_MEMORY_VERSION,
        "protocol_version": PROTOCOL_VERSION,
        "goal": goal,
        "facts": list(_dedupe_text(facts, limit=max_facts, text_limit=320)),
        "decisions": list(_dedupe_text(decisions, limit=16, text_limit=260)),
        "open_questions": list(_dedupe_text(open_questions, limit=max_open_questions, text_limit=320)),
        "open_items": list(_dedupe_text(open_items, limit=max_open_items, text_limit=320)),
        "next_actions": list(_dedupe_text(next_actions, limit=max_next_actions, text_limit=320)),
        "evidence_ids": list(_dedupe_text(evidence_ids, limit=max_evidence_ids, text_limit=120)),
        "source_agent_run_ids": list(_dedupe_text(source_runs, limit=24, text_limit=120)),
        "updated_from_agent_run_id": agent_run_id,
        "status": status,
        "quality_status": _text(quality_gate.get("status"), 80),
        "notes": "项目级滚动记忆只保存摘要和 evidence id；原始证据仍需通过 detail/aggregation 工具读取。",
    }


@dataclass(frozen=True)
class AgentSessionState:
    agent_run_id: str
    goal: str
    todo_list: dict = field(default_factory=dict)
    task_ledger: dict = field(default_factory=dict)
    memory_summary: dict = field(default_factory=dict)
    recent_evidence_ids: tuple[str, ...] = ()
    pending_questions: tuple[dict, ...] = ()
    context_answer_count: int = 0
    protocol_version: str = PROTOCOL_VERSION

    def to_dict(self) -> dict:
        return {
            "protocol_version": self.protocol_version,
            "agent_run_id": self.agent_run_id,
            "goal": self.goal,
            "todo_list": dict(self.todo_list or {}),
            "task_ledger": dict(self.task_ledger or {}),
            "memory_summary": dict(self.memory_summary or {}),
            "recent_evidence_ids": list(self.recent_evidence_ids),
            "pending_questions": [dict(item) for item in self.pending_questions],
            "context_answer_count": int(self.context_answer_count or 0),
            "notes": "session_state 是本轮 agent 的轻量会话记忆；第一版仅进程内传递，不写原始项目。",
        }


def build_agent_session_state(*,
                              agent_run_id: object,
                              goal: object,
                              runtime_state: Mapping[str, object] | None = None,
                              project_context: Mapping[str, object] | None = None,
                              observations: Sequence[Mapping[str, object]] = ()) -> dict:
    """Build a compact session state card from runtime memory and user context."""

    runtime_state = runtime_state or {}
    project_context = project_context or {}
    todo = runtime_state.get("todo_list") if isinstance(runtime_state.get("todo_list"), Mapping) else {}
    task_ledger = runtime_state.get("task_ledger") if isinstance(runtime_state.get("task_ledger"), Mapping) else {}
    memory = runtime_state.get("memory_summary") if isinstance(runtime_state.get("memory_summary"), Mapping) else {}
    session = AgentSessionState(
        agent_run_id=_text(agent_run_id, 120),
        goal=_text(goal, 600),
        todo_list=dict(todo),
        task_ledger=dict(task_ledger),
        memory_summary=dict(memory),
        recent_evidence_ids=_recent_evidence_ids(runtime_state, observations),
        pending_questions=_pending_questions(project_context, runtime_state),
        context_answer_count=len(project_context.get("answers") or []),
    )
    return session.to_dict()
