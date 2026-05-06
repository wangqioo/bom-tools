# -*- coding: utf-8 -*-
"""Generic long-task dispatch protocol helpers for durable agent runs."""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from typing import Mapping, Sequence

from .protocol import AgentProtocolError


TASK_DISPATCH_SCHEMA_VERSION = "pstx-agent-task-dispatch.v1"
DEFAULT_MAX_DISPATCH_TASKS = 6
TASK_PROFILE_FALLBACK = "auto"


def _text(value: object, limit: int = 1000) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "..."


def _list_of_text(value: object, *, limit: int = 12, item_limit: int = 120) -> tuple[str, ...]:
    if value is None:
        return ()
    source = value if isinstance(value, list) else [value]
    items: list[str] = []
    for item in source:
        text = _text(item, item_limit)
        if text and text not in items:
            items.append(text)
        if len(items) >= limit:
            break
    return tuple(items)


def _int_range(value: object, *, minimum: int = 0, maximum: int = 200) -> int:
    try:
        number = int(value or 0)
    except Exception:
        return 0
    if number <= 0:
        return 0
    return max(minimum, min(maximum, number))


def _raw_task_items(raw: object) -> tuple[list[object], str]:
    if isinstance(raw, Mapping) and isinstance(raw.get("dispatch_tasks"), list):
        return list(raw.get("dispatch_tasks") or []), _text(raw.get("reason") or raw.get("summary") or "", 800)
    if isinstance(raw, Mapping) and isinstance(raw.get("task_dispatch"), Mapping):
        payload = raw.get("task_dispatch") or {}
        tasks = payload.get("tasks") if isinstance(payload, Mapping) else None
        if isinstance(tasks, list):
            return list(tasks), _text(payload.get("reason") or raw.get("reason") or "", 800)
    if isinstance(raw, Mapping) and isinstance(raw.get("long_task_dispatch"), Mapping):
        payload = raw.get("long_task_dispatch") or {}
        tasks = payload.get("tasks") if isinstance(payload, Mapping) else None
        if isinstance(tasks, list):
            return list(tasks), _text(payload.get("reason") or raw.get("reason") or "", 800)
    if isinstance(raw, list):
        return list(raw), ""
    raise AgentProtocolError("dispatch_tasks 必须是对象数组。")


@dataclass(frozen=True)
class DispatchTask:
    """A model-requested long task that may become a durable child run."""

    task_id: str
    title: str
    question: str
    profile: str = TASK_PROFILE_FALLBACK
    reason: str = ""
    priority: str = "normal"
    depends_on: tuple[str, ...] = ()
    expected_outputs: tuple[str, ...] = ()
    scope: dict = field(default_factory=dict)
    max_steps: int = 0
    max_tool_calls: int = 0

    @classmethod
    def from_mapping(cls, value: Mapping[str, object], *, index: int) -> "DispatchTask":
        if not isinstance(value, Mapping):
            raise AgentProtocolError("dispatch task 必须是对象。")
        task_id = _text(value.get("task_id") or value.get("id") or f"task-{index}", 80)
        title = _text(value.get("title") or value.get("name") or value.get("summary") or "", 180)
        question = _text(value.get("question") or value.get("prompt") or value.get("instruction") or "", 1200)
        if not question:
            raise AgentProtocolError(f"{task_id} 缺少 question。")
        if not title:
            title = question[:120]
        raw_scope = value.get("scope") if isinstance(value.get("scope"), Mapping) else {}
        return cls(
            task_id=task_id,
            title=title,
            question=question,
            profile=_text(value.get("profile") or value.get("kind") or TASK_PROFILE_FALLBACK, 80) or TASK_PROFILE_FALLBACK,
            reason=_text(value.get("reason") or "", 600),
            priority=_text(value.get("priority") or "normal", 40) or "normal",
            depends_on=_list_of_text(value.get("depends_on"), limit=12, item_limit=80),
            expected_outputs=_list_of_text(value.get("expected_outputs") or value.get("outputs"), limit=12, item_limit=160),
            scope={str(key)[:80]: _text(item, 500) for key, item in dict(raw_scope).items()},
            max_steps=_int_range(value.get("max_steps"), minimum=1, maximum=100),
            max_tool_calls=_int_range(value.get("max_tool_calls"), minimum=1, maximum=200),
        )

    def to_dict(self) -> dict:
        payload = {
            "task_id": self.task_id,
            "title": self.title,
            "profile": self.profile,
            "question": self.question,
            "reason": self.reason,
            "priority": self.priority,
            "depends_on": list(self.depends_on),
            "expected_outputs": list(self.expected_outputs),
            "scope": dict(self.scope),
        }
        if self.max_steps:
            payload["max_steps"] = self.max_steps
        if self.max_tool_calls:
            payload["max_tool_calls"] = self.max_tool_calls
        return payload


def normalize_dispatch_tasks(raw: object, *, max_tasks: int = DEFAULT_MAX_DISPATCH_TASKS) -> dict:
    """Normalize model-emitted long-task dispatch payloads."""

    if max_tasks < 1:
        raise AgentProtocolError("max_tasks 必须大于 0。")
    raw_items, reason = _raw_task_items(raw)
    if not raw_items:
        raise AgentProtocolError("dispatch_tasks 至少需要一个任务。")
    if len(raw_items) > max_tasks:
        raise AgentProtocolError(f"一次最多允许分发 {max_tasks} 个任务。")
    tasks: list[dict] = []
    seen: set[str] = set()
    for index, item in enumerate(raw_items, start=1):
        task = DispatchTask.from_mapping(item, index=index).to_dict()
        original_id = str(task.get("task_id") or f"task-{index}")
        task_id = original_id
        while task_id in seen:
            task_id = f"{original_id}-{uuid.uuid4().hex[:4]}"
        task["task_id"] = task_id
        seen.add(task_id)
        tasks.append(task)
    return {
        "schema_version": TASK_DISPATCH_SCHEMA_VERSION,
        "reason": reason,
        "task_count": len(tasks),
        "tasks": tasks,
    }


def compact_dispatch_records(records: Sequence[Mapping[str, object]], *, limit: int = 20) -> list[dict]:
    """Return a status-friendly compact list of child dispatch records."""

    compact: list[dict] = []
    for record in list(records or [])[: max(1, int(limit or 20))]:
        if not isinstance(record, Mapping):
            continue
        compact.append({
            "task_id": _text(record.get("task_id") or "", 80),
            "title": _text(record.get("title") or "", 180),
            "profile": _text(record.get("profile") or "", 80),
            "agent_run_id": _text(record.get("agent_run_id") or "", 120),
            "status": _text(record.get("status") or "", 40),
            "status_url": _text(record.get("status_url") or "", 240),
            "question": _text(record.get("question") or "", 300),
        })
    return compact
