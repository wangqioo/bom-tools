# -*- coding: utf-8 -*-
"""Protocol primitives for the PSTX agent runtime.

The report agent loop in ``pstx_harness.report_agent`` remains the production
executor. This module is intentionally small and side-effect free so we can migrate toward a
Codex/Claude-Code style runtime incrementally: explicit todo state, memory
summary, batch tool-call envelopes, observation bundles, and final answer
contracts.
"""

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass, field
from typing import Iterable, Mapping, Optional, Sequence


PROTOCOL_VERSION = "pstx-agent-runtime/v1"
TODO_STATUSES = {"pending", "in_progress", "completed", "blocked"}


class AgentProtocolError(ValueError):
    """Raised when an agent protocol payload fails local validation."""


def _text(value: object, limit: int = 1000) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _json_chars(value: object) -> int:
    try:
        return len(json.dumps(value, ensure_ascii=False, sort_keys=True))
    except (TypeError, ValueError):
        return len(str(value))


def _extract_balanced_json(text: str) -> Optional[dict]:
    """Extract the first balanced JSON object from a model response."""

    content = str(text or "").strip()
    fence_start = content.find("```")
    if fence_start >= 0:
        content = content.replace("```json", "```").replace("```JSON", "```")
        parts = content.split("```")
        if len(parts) >= 3:
            content = parts[1].strip()
    start = content.find("{")
    if start < 0:
        return None
    depth = 0
    in_string = False
    escape = False
    for index in range(start, len(content)):
        char = content[index]
        if in_string:
            if escape:
                escape = False
            elif char == "\\":
                escape = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                try:
                    parsed = json.loads(content[start:index + 1])
                except json.JSONDecodeError:
                    return None
                return parsed if isinstance(parsed, dict) else None
    return None


@dataclass(frozen=True)
class AgentTodoItem:
    """A single planned work item visible to the local harness."""

    id: str
    title: str
    status: str = "pending"
    evidence_ids: tuple[str, ...] = ()
    note: str = ""

    def __post_init__(self) -> None:
        if not self.id:
            raise AgentProtocolError("todo item id 不能为空。")
        if not self.title:
            raise AgentProtocolError("todo item title 不能为空。")
        if self.status not in TODO_STATUSES:
            raise AgentProtocolError(f"未知 todo 状态：{self.status}")

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "title": self.title,
            "status": self.status,
            "evidence_ids": list(self.evidence_ids),
            "note": self.note,
        }


@dataclass(frozen=True)
class AgentTodoList:
    """Small todo list that can be carried across multi-turn agent work."""

    goal: str
    items: tuple[AgentTodoItem, ...] = ()

    @classmethod
    def from_titles(cls, goal: str, titles: Sequence[object]) -> "AgentTodoList":
        items = tuple(
            AgentTodoItem(id=f"todo-{index}", title=_text(title, 160))
            for index, title in enumerate(titles or [], start=1)
            if _text(title, 160)
        )
        return cls(goal=_text(goal, 500), items=items)

    def mark(self, item_id: str, status: str, *, note: str = "", evidence_ids: Sequence[object] = ()) -> "AgentTodoList":
        updated = []
        found = False
        for item in self.items:
            if item.id != item_id:
                updated.append(item)
                continue
            found = True
            updated.append(AgentTodoItem(
                id=item.id,
                title=item.title,
                status=status,
                evidence_ids=tuple(_text(value, 120) for value in (evidence_ids or item.evidence_ids) if _text(value, 120)),
                note=_text(note or item.note, 300),
            ))
        if not found:
            raise AgentProtocolError(f"未找到 todo item：{item_id}")
        return AgentTodoList(goal=self.goal, items=tuple(updated))

    def to_dict(self) -> dict:
        return {
            "goal": self.goal,
            "items": [item.to_dict() for item in self.items],
            "open_count": sum(1 for item in self.items if item.status in {"pending", "in_progress", "blocked"}),
        }


@dataclass(frozen=True)
class ToolBatchCall:
    """Validated tool-call envelope used by the planned batch runtime."""

    name: str
    args: dict = field(default_factory=dict)
    reason: str = ""
    call_id: str = ""

    @classmethod
    def from_mapping(cls,
                     value: Mapping[str, object],
                     *,
                     allowed_tools: Optional[Iterable[str]] = None,
                     call_id: str = "") -> "ToolBatchCall":
        if not isinstance(value, Mapping):
            raise AgentProtocolError("tool call 必须是对象。")
        name = _text(value.get("name") or value.get("tool"), 120)
        if not name:
            raise AgentProtocolError("tool call 缺少 name。")
        allowed = set(allowed_tools or [])
        if allowed and name not in allowed:
            raise AgentProtocolError(f"工具不在白名单中：{name}")
        raw_args = value.get("args", {})
        if raw_args is None:
            raw_args = {}
        if not isinstance(raw_args, dict):
            raise AgentProtocolError(f"{name}.args 必须是 JSON 对象。")
        return cls(
            name=name,
            args=dict(raw_args),
            reason=_text(value.get("reason") or "", 500),
            call_id=_text(value.get("call_id") or call_id or f"tc-{uuid.uuid4().hex[:10]}", 80),
        )

    def to_dict(self) -> dict:
        return {
            "call_id": self.call_id,
            "name": self.name,
            "args": self.args,
            "reason": self.reason,
        }


def normalize_tool_batch(raw: object,
                         *,
                         allowed_tools: Optional[Iterable[str]] = None,
                         max_calls: int = 8) -> list[dict]:
    """Normalize future ``tool_batch_call`` payloads into validated call dicts."""

    if max_calls < 1:
        raise AgentProtocolError("max_calls 必须大于 0。")
    if isinstance(raw, Mapping) and isinstance(raw.get("tool_call"), Mapping):
        raw_items = [raw["tool_call"]]
    elif isinstance(raw, Mapping) and isinstance(raw.get("tool_batch_call"), list):
        raw_items = list(raw["tool_batch_call"])
    elif isinstance(raw, list):
        raw_items = raw
    else:
        raise AgentProtocolError("工具调用必须是 tool_call、tool_batch_call 或对象数组。")
    if len(raw_items) > max_calls:
        raise AgentProtocolError(f"一次最多允许 {max_calls} 个工具调用。")
    return [
        ToolBatchCall.from_mapping(item, allowed_tools=allowed_tools, call_id=f"tc-{index}").to_dict()
        for index, item in enumerate(raw_items, start=1)
    ]


@dataclass(frozen=True)
class AgentModelStep:
    """A normalized single model step emitted by the local runtime protocol."""

    type: str
    raw: dict = field(default_factory=dict)
    tool_calls: tuple[dict, ...] = ()
    dispatch_tasks: tuple[dict, ...] = ()
    task_dispatch: dict = field(default_factory=dict)
    final_answer: str = ""
    needs_user_input: dict = field(default_factory=dict)

    def to_legacy_dict(self) -> dict:
        if self.type == "tool_call":
            return {
                "type": "tool_call",
                "tool_call": self.tool_calls[0] if self.tool_calls else {},
                "tool_calls": list(self.tool_calls),
                "raw": self.raw,
            }
        if self.type == "tool_batch_call":
            return {
                "type": "tool_batch_call",
                "tool_calls": list(self.tool_calls),
                "raw": self.raw,
            }
        if self.type == "needs_user_input":
            return {
                "type": "needs_user_input",
                "needs_user_input": self.needs_user_input,
                "raw": self.raw,
            }
        if self.type == "dispatch_tasks":
            return {
                "type": "dispatch_tasks",
                "dispatch_tasks": list(self.dispatch_tasks),
                "task_dispatch": dict(self.task_dispatch),
                "raw": self.raw,
            }
        if self.type == "final_answer":
            return {
                "type": "final_answer",
                "final_answer": self.final_answer,
                "raw": self.raw,
            }
        return {"type": self.type, "raw": self.raw}


def parse_agent_model_step(answer: str,
                           *,
                           allowed_tools: Optional[Iterable[str]] = None,
                           max_batch_calls: int = 6,
                           max_dispatch_tasks: int = 6,
                           allow_batch_tools: bool = True,
                           allow_needs_user_input: bool = True,
                           allow_task_dispatch: bool = True) -> Optional[AgentModelStep]:
    """Parse model JSON into the canonical PSTX agent step contract."""

    parsed = _extract_balanced_json(answer)
    if not parsed:
        return None
    if allow_task_dispatch and (
        isinstance(parsed.get("dispatch_tasks"), list)
        or isinstance(parsed.get("task_dispatch"), Mapping)
        or isinstance(parsed.get("long_task_dispatch"), Mapping)
    ):
        from .task_dispatch import normalize_dispatch_tasks

        dispatch = normalize_dispatch_tasks(parsed, max_tasks=max_dispatch_tasks)
        return AgentModelStep(
            type="dispatch_tasks",
            raw=parsed,
            dispatch_tasks=tuple(dispatch.get("tasks") or []),
            task_dispatch=dispatch,
        )
    if isinstance(parsed.get("tool_call"), Mapping):
        calls = normalize_tool_batch(
            {"tool_call": parsed["tool_call"]},
            allowed_tools=allowed_tools,
            max_calls=1,
        )
        return AgentModelStep(type="tool_call", raw=parsed, tool_calls=tuple(calls))
    if isinstance(parsed.get("tool_batch_call"), list):
        if not allow_batch_tools:
            raise AgentProtocolError("当前 runtime 不允许 tool_batch_call。")
        calls = normalize_tool_batch(
            parsed.get("tool_batch_call") or [],
            allowed_tools=allowed_tools,
            max_calls=max_batch_calls,
        )
        return AgentModelStep(type="tool_batch_call", raw=parsed, tool_calls=tuple(calls))
    if allow_needs_user_input and isinstance(parsed.get("needs_user_input"), dict):
        return AgentModelStep(
            type="needs_user_input",
            raw=parsed,
            needs_user_input=dict(parsed["needs_user_input"]),
        )
    if isinstance(parsed.get("final_answer"), str):
        return AgentModelStep(
            type="final_answer",
            raw=parsed,
            final_answer=str(parsed["final_answer"]),
        )
    return None


@dataclass(frozen=True)
class ObservationBundle:
    """Compressed observation envelope for model context."""

    id: str
    summary: str
    evidence_ids: tuple[str, ...] = ()
    observations: tuple[dict, ...] = ()
    truncated: bool = False
    omitted_count: int = 0

    @classmethod
    def from_observations(cls,
                          observations: Sequence[dict],
                          *,
                          bundle_id: str = "obs-bundle",
                          max_items: int = 6,
                          max_chars: int = 12000) -> "ObservationBundle":
        source = [item for item in observations or [] if isinstance(item, dict)]
        compact = source[-max_items:]
        truncated = len(source) > len(compact) or _json_chars(compact) > max_chars
        while compact and _json_chars(compact) > max_chars:
            compact = compact[1:]
            truncated = True
        evidence_ids: list[str] = []
        for item in source:
            for evidence_id in item.get("evidence_node_ids", []) or []:
                text = _text(evidence_id, 120)
                if text and text not in evidence_ids:
                    evidence_ids.append(text)
        return cls(
            id=bundle_id,
            summary=f"{len(source)} 个观察，发送 {len(compact)} 个，证据 {len(evidence_ids)} 条。",
            evidence_ids=tuple(evidence_ids[:80]),
            observations=tuple(compact),
            truncated=truncated,
            omitted_count=max(0, len(source) - len(compact)),
        )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "summary": self.summary,
            "evidence_ids": list(self.evidence_ids),
            "observations": list(self.observations),
            "truncated": self.truncated,
            "omitted_count": self.omitted_count,
        }


@dataclass(frozen=True)
class MemorySummary:
    """Rolling memory summary for long agent tasks."""

    goal: str
    facts: tuple[str, ...] = ()
    decisions: tuple[str, ...] = ()
    open_questions: tuple[str, ...] = ()
    evidence_ids: tuple[str, ...] = ()

    @classmethod
    def from_parts(cls,
                   *,
                   goal: object,
                   facts: Sequence[object] = (),
                   decisions: Sequence[object] = (),
                   open_questions: Sequence[object] = (),
                   evidence_ids: Sequence[object] = ()) -> "MemorySummary":
        return cls(
            goal=_text(goal, 500),
            facts=tuple(_text(item, 220) for item in facts if _text(item, 220))[:24],
            decisions=tuple(_text(item, 220) for item in decisions if _text(item, 220))[:24],
            open_questions=tuple(_text(item, 220) for item in open_questions if _text(item, 220))[:24],
            evidence_ids=tuple(_text(item, 120) for item in evidence_ids if _text(item, 120))[:80],
        )

    def to_dict(self) -> dict:
        return {
            "goal": self.goal,
            "facts": list(self.facts),
            "decisions": list(self.decisions),
            "open_questions": list(self.open_questions),
            "evidence_ids": list(self.evidence_ids),
        }


def build_agent_protocol_brief(*,
                               allow_batch_tools: bool = False,
                               allow_task_dispatch: bool = False) -> str:
    """Return the compact protocol note injected into model prompts."""

    batch_line = (
        "生产 loop 支持 tool_batch_call 数组，但每个工具仍由本地按白名单、schema 和调用上限逐个校验执行。"
        if allow_batch_tools
        else "当前每轮只允许一个 tool_call；批处理协议对象已在本地定义，迁移前不得让模型一次请求多个工具。"
    )
    dispatch_line = (
        "异步 durable run 可选支持 dispatch_tasks，把互相独立的长耗时分支交给本地后台创建 child runs；模型只声明任务，不直接执行。"
        if allow_task_dispatch
        else "长任务分发协议由本地 runtime 定义，只有启用异步 durable dispatch 时才会创建 child runs。"
    )
    return (
        f"Runtime protocol: {PROTOCOL_VERSION}。\n"
        "长期任务按 Plan/TodoList -> ToolCall/ObservationBundle -> MemorySummary -> FinalAnswer/NeedsUserInput 推进。\n"
        "ObservationBundle 只给模型压缩证据卡和摘要；完整数据必须通过 detail 工具按需读取。\n"
        f"{batch_line}\n"
        f"{dispatch_line}"
        "\n最终回答可选包含 scratch_files 数组，用于声明本地 runtime 代写临时文本文件："
        '{"filename":"notes.md","content":"...","content_type":"text/markdown"}；'
        "这些文件只会进入 agent_workspace scratch 临时区，不授权模型直接写文件。"
    )
