# -*- coding: utf-8 -*-
"""In-memory trace/replay store for PSTX agent runs.

The production Web API still returns the full historical agent payload for
backward compatibility. This module adds a runtime-owned trace envelope around
that payload so future UI, diagnostics, and memory layers can replay runs
without scraping Web globals.
"""

from __future__ import annotations

from collections import OrderedDict
from collections.abc import Iterator, Mapping
from typing import Any

from .turn_context import summarize_tool_dispatch_trace


DEFAULT_MAX_AGENT_TRACE_ITEMS = 50


def _text(value: object, limit: int = 500) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _list_count(value: object) -> int:
    return len(value) if isinstance(value, list) else 0


def _mapping_copy(value: object) -> dict:
    return dict(value) if isinstance(value, Mapping) else {}


def _compact_value(value: object, *, depth: int = 0, text_limit: int = 220) -> object:
    if depth >= 3:
        return _text(value, text_limit)
    if isinstance(value, Mapping):
        return {
            str(key)[:80]: _compact_value(item, depth=depth + 1, text_limit=min(text_limit, 180))
            for key, item in list(value.items())[:12]
        }
    if isinstance(value, list):
        items = [_compact_value(item, depth=depth + 1, text_limit=min(text_limit, 180)) for item in value[:12]]
        if len(value) > 12:
            items.append({"__truncated__": True, "__remaining__": len(value) - 12})
        return items
    return _text(value, text_limit)


def _ids(value: object, *, limit: int = 12) -> list[str]:
    result: list[str] = []
    for item in value or []:
        text = _text(item, 120)
        if text and text not in result:
            result.append(text)
        if len(result) >= limit:
            break
    return result


def _status(ok: object = None, *, default: str = "ok") -> str:
    if ok is False:
        return "error"
    if ok is True:
        return "ok"
    return default


def _event(index: int,
           phase: str,
           event_type: str,
           title: str,
           *,
           summary: object = "",
           status: str = "info",
           tool: object = "",
           evidence_ids: object = (),
           metadata: Mapping[str, object] | None = None) -> dict:
    payload = {
        "index": index,
        "phase": _text(phase, 60),
        "type": _text(event_type, 80),
        "title": _text(title, 180),
        "summary": _text(summary, 500),
        "status": _text(status, 40),
    }
    if tool:
        payload["tool"] = _text(tool, 120)
    ids = _ids(evidence_ids)
    if ids:
        payload["evidence_ids"] = ids
    if metadata:
        payload["metadata"] = {str(key)[:80]: _text(value, 240) for key, value in list(metadata.items())[:12]}
    return payload


def build_execution_journal(payload: Mapping[str, object] | None, *, max_events: int = 120) -> list[dict]:
    """Build a compact Codex-style run journal from an agent payload."""

    source = dict(payload or {})
    events: list[dict] = []

    def add(phase: str, event_type: str, title: str, **kwargs) -> None:
        if len(events) >= max(1, int(max_events or 120)):
            return
        events.append(_event(len(events) + 1, phase, event_type, title, **kwargs))

    add(
        "start",
        "run_started",
        "Agent run started",
        summary=source.get("profile") or source.get("mode") or source.get("agent_run_id") or "",
        status="info",
        metadata={"agent_run_id": source.get("agent_run_id", ""), "mode": source.get("mode", "")},
    )
    for step in source.get("agent_steps") or []:
        if not isinstance(step, Mapping):
            continue
        step_type = _text(step.get("type"), 80)
        phase = "model"
        if "tool" in step_type:
            phase = "tool"
        elif step_type in {"final_answer", "needs_user_input"}:
            phase = "final"
        elif "error" in step_type or step.get("ok") is False:
            phase = "error"
        add(
            phase,
            step_type or "agent_step",
            step_type or "Agent step",
            summary=step.get("summary") or step.get("error") or "",
            status=_status(step.get("ok"), default="info"),
            tool=step.get("tool", ""),
            metadata={"step_index": step.get("index", ""), "provider": step.get("provider", "")},
        )
    for call in source.get("tool_calls") or []:
        if not isinstance(call, Mapping):
            continue
        add(
            "tool",
            "tool_call",
            f"Tool call: {call.get('tool') or ''}",
            summary=call.get("reason") or call.get("error") or "",
            status=_status(call.get("ok"), default="ok"),
            tool=call.get("tool", ""),
            evidence_ids=call.get("evidence_node_ids") or (),
            metadata={"call_index": call.get("index", ""), "batch": call.get("batch", False)},
        )
    for dispatch in source.get("tool_dispatch_trace") or []:
        if not isinstance(dispatch, Mapping):
            continue
        status = _text(dispatch.get("status"), 40)
        add(
            "dispatch",
            "tool_dispatch",
            f"Tool dispatch: {dispatch.get('tool') or ''}",
            summary=dispatch.get("reason") or dispatch.get("error") or status,
            status="ok" if status == "completed" else ("warn" if status in {"duplicate", "limit"} else "error"),
            tool=dispatch.get("tool", ""),
            evidence_ids=dispatch.get("evidence_node_ids") or (),
            metadata={
                "event_index": dispatch.get("event_index", ""),
                "status": status,
                "allowed": dispatch.get("allowed", ""),
                "duplicate": dispatch.get("duplicate", ""),
            },
        )
    quality = _mapping_copy(source.get("final_answer_quality_gate"))
    if quality:
        add(
            "quality",
            "final_answer_quality_gate",
            f"Final answer quality: {quality.get('status') or ''}",
            summary="；".join(str(item.get("message") or item.get("id") or "") for item in (quality.get("reasons") or [])[:4] if isinstance(item, Mapping)),
            status="ok" if quality.get("status") == "pass" else ("error" if quality.get("status") == "fail" else "warn"),
            metadata={"score": quality.get("score", ""), "repair_action_count": quality.get("repair_action_count", 0)},
        )
    ledger = _mapping_copy((_mapping_copy(source.get("runtime_state")).get("task_ledger")))
    progress = _mapping_copy(ledger.get("progress"))
    if progress:
        add(
            "ledger",
            "task_ledger",
            "Task ledger progress",
            summary=f"completed={progress.get('completed', 0)} open={progress.get('open', 0)} blocked={progress.get('blocked', 0)}",
            status="warn" if int(progress.get("blocked") or 0) else "info",
            metadata=progress,
        )
    add(
        "finish",
        "run_finished",
        "Agent run finished",
        summary=source.get("answer") or "",
        status="ok" if source.get("ok", True) else "error",
        metadata={"status": source.get("status", ""), "stopped_reason": (_mapping_copy(source.get("model_metadata")).get("stopped_reason") or "")},
    )
    return events


def build_journal_summary(journal: Sequence[Mapping[str, object]] | None) -> dict:
    events = [item for item in journal or [] if isinstance(item, Mapping)]
    phases: dict[str, int] = {}
    statuses: dict[str, int] = {}
    for event in events:
        phase = _text(event.get("phase"), 60) or "unknown"
        status = _text(event.get("status"), 40) or "info"
        phases[phase] = phases.get(phase, 0) + 1
        statuses[status] = statuses.get(status, 0) + 1
    return {
        "version": "agent-run-journal/v1",
        "event_count": len(events),
        "phase_counts": phases,
        "status_counts": statuses,
        "tool_event_count": phases.get("tool", 0),
        "warning_count": statuses.get("warn", 0),
        "error_count": statuses.get("error", 0),
    }


def _open_ledger_items(runtime_state: Mapping[str, object], *, limit: int = 8) -> list[dict]:
    ledger = _mapping_copy(runtime_state.get("task_ledger"))
    result: list[dict] = []
    for item in ledger.get("items") or []:
        if not isinstance(item, Mapping):
            continue
        status = _text(item.get("status"), 40)
        if status == "completed":
            continue
        result.append({
            "id": _text(item.get("id"), 100),
            "title": _text(item.get("title"), 180),
            "status": status or "pending",
            "recommended_tools": _ids(item.get("recommended_tools") or (), limit=8),
            "detail_tools": [
                {str(key)[:80]: _text(value, 180) for key, value in list(tool.items())[:8]}
                for tool in (item.get("detail_tools") or [])
                if isinstance(tool, Mapping)
            ][:4],
            "note": _text(item.get("note"), 260),
        })
        if len(result) >= limit:
            break
    return result


def _quality_tool_suggestions(quality_gate: Mapping[str, object], *, limit: int = 8) -> list[dict]:
    result: list[dict] = []
    for action in quality_gate.get("repair_actions") or []:
        if not isinstance(action, Mapping) or _text(action.get("type"), 60) != "tool_call":
            continue
        tool = _text(action.get("tool"), 120)
        if not tool:
            continue
        payload = {
            "tool": tool,
            "title": _text(action.get("title"), 180),
            "reason": _text(action.get("reason"), 260),
            "source": _text(action.get("source"), 120),
        }
        if isinstance(action.get("args"), Mapping):
            payload["args"] = _compact_value(action.get("args"), depth=1)
        result.append(payload)
        if len(result) >= limit:
            break
    return result


def _task_ledger_tool_suggestions(runtime_state: Mapping[str, object], *, limit: int = 8) -> list[dict]:
    ledger = _mapping_copy(runtime_state.get("task_ledger"))
    result: list[dict] = []
    seen = set()
    for action in ledger.get("next_actions") or []:
        if not isinstance(action, Mapping):
            continue
        if _text(action.get("type"), 60) not in {"", "tool_call"}:
            continue
        tool = _text(action.get("tool"), 120)
        if not tool:
            continue
        args = action.get("args") if isinstance(action.get("args"), Mapping) else {}
        marker = (tool, str(args))
        if marker in seen:
            continue
        seen.add(marker)
        payload = {
            "tool": tool,
            "title": _text(action.get("title"), 180),
            "reason": _text(action.get("reason"), 260),
            "source": _text(action.get("source") or "task_ledger", 120),
        }
        if args:
            payload["args"] = _compact_value(args, depth=1)
        result.append(payload)
        if len(result) >= limit:
            break
    return result


def _pending_questions_from_payload(source: Mapping[str, object], session_state: Mapping[str, object]) -> list[dict]:
    needs = _mapping_copy(source.get("needs_user_input"))
    questions = needs.get("questions") if isinstance(needs.get("questions"), list) else session_state.get("pending_questions")
    result: list[dict] = []
    for item in questions or []:
        if not isinstance(item, Mapping):
            continue
        result.append({
            "question_id": _text(item.get("question_id") or item.get("id"), 120),
            "question": _text(item.get("question") or item.get("prompt") or item.get("title"), 360),
            "missing_fields": _ids(item.get("missing_fields") or (), limit=12),
            "related_evidence_ids": _ids(item.get("related_evidence_ids") or (), limit=12),
        })
        if len(result) >= 12:
            break
    return result


def _evidence_ids_from_payload(source: Mapping[str, object], session_state: Mapping[str, object]) -> list[str]:
    ids: list[object] = []
    ids.extend(session_state.get("recent_evidence_ids") or [])
    ids.extend(item.get("id") for item in source.get("final_evidence") or [] if isinstance(item, Mapping))
    ids.extend(item.get("id") for item in source.get("citations") or [] if isinstance(item, Mapping) and item.get("valid") is not False)
    return _ids(ids, limit=40)


def build_continuation_pack(payload: Mapping[str, object] | None) -> dict:
    """Build a compact handoff card for the next agent turn."""

    source = dict(payload or {})
    metadata = _mapping_copy(source.get("model_metadata"))
    runtime_state = _mapping_copy(source.get("runtime_state"))
    session_state = _mapping_copy(source.get("session_state"))
    quality_gate = _mapping_copy(source.get("final_answer_quality_gate"))
    status = _text(source.get("status") or ("completed" if source.get("ok", True) else "failed"), 80)
    stopped_reason = _text(metadata.get("stopped_reason"), 100)
    pending_questions = _pending_questions_from_payload(source, session_state)
    open_items = _open_ledger_items(runtime_state)
    quality_status = _text(quality_gate.get("status"), 40)
    if status == "waiting_for_user" or pending_questions:
        next_intent = "await_user_input"
    elif quality_status in {"warn", "fail"} or open_items:
        next_intent = "continue_evidence_gathering"
    elif source.get("ok") is False:
        next_intent = "manual_review"
    else:
        next_intent = "completed"
    suggested_tools = _quality_tool_suggestions(quality_gate)
    suggested_markers = {
        (item.get("tool"), str(item.get("args") or {}))
        for item in suggested_tools
        if isinstance(item, Mapping)
    }
    for item in _task_ledger_tool_suggestions(runtime_state):
        marker = (item.get("tool"), str(item.get("args") or {}))
        if marker in suggested_markers:
            continue
        suggested_markers.add(marker)
        suggested_tools.append(item)
        if len(suggested_tools) >= 8:
            break
    if not suggested_tools:
        for item in open_items:
            for tool in item.get("detail_tools") or []:
                if isinstance(tool, Mapping) and tool.get("name"):
                    payload = {
                        "tool": _text(tool.get("name"), 120),
                        "title": _text(item.get("title"), 180),
                        "reason": _text(item.get("note"), 260),
                        "source": "task_ledger",
                    }
                    if isinstance(tool.get("args"), Mapping):
                        payload["args"] = _compact_value(tool.get("args"), depth=1)
                    suggested_tools.append(payload)
            for tool in item.get("recommended_tools") or []:
                suggested_tools.append({
                    "tool": _text(tool, 120),
                    "title": _text(item.get("title"), 180),
                    "reason": _text(item.get("note"), 260),
                    "source": "task_ledger",
                })
            if len(suggested_tools) >= 8:
                break
    goal = _text((source.get("request") or {}).get("question") if isinstance(source.get("request"), Mapping) else "", 500)
    if not goal:
        goal = _text(session_state.get("goal") or source.get("profile") or source.get("mode"), 500)
    brief_parts = [
        f"目标：{goal}" if goal else "",
        f"状态：{status}",
        f"停止原因：{stopped_reason}" if stopped_reason else "",
        f"续跑意图：{next_intent}",
        f"有效证据 {len(_evidence_ids_from_payload(source, session_state))} 个",
        f"未完成项 {len(open_items)} 个" if open_items else "",
        f"待用户补充 {len(pending_questions)} 项" if pending_questions else "",
    ]
    return {
        "version": "agent-continuation-pack/v1",
        "agent_run_id": _text(source.get("agent_run_id"), 120),
        "mode": _text(source.get("mode"), 120),
        "profile": _text(source.get("profile"), 120),
        "goal": goal,
        "status": status,
        "stopped_reason": stopped_reason,
        "next_intent": next_intent,
        "answer_preview": _text(source.get("answer"), 900),
        "evidence_ids": _evidence_ids_from_payload(source, session_state),
        "pending_questions": pending_questions,
        "open_ledger_items": open_items,
        "suggested_tool_calls": suggested_tools[:8],
        "quality_status": quality_status,
        "quality_score": quality_gate.get("score", 0),
        "context_budget": _mapping_copy(source.get("context_budget")),
        "journal_summary": _mapping_copy(source.get("journal_summary")),
        "continuation_brief": "；".join(part for part in brief_parts if part),
        "model_rules": [
            "续跑时优先使用本 pack 的 evidence_ids、open_ledger_items 和 suggested_tool_calls。",
            "pack 是压缩交接摘要，不代表完整证据；高风险结论必须通过 detail_tool 或原始 trace 取证。",
            "如果 next_intent=await_user_input，应先处理 pending_questions。",
        ],
    }


def build_trace_envelope(payload: Mapping[str, object] | None,
                         *,
                         agent_run_id: object = None) -> dict:
    """Build a compact replay envelope around a full agent run payload."""

    source = dict(payload or {})
    run_id = _text(agent_run_id or source.get("agent_run_id"), 120)
    if run_id:
        source["agent_run_id"] = run_id
    ok_value = source.get("ok")
    ok = bool(ok_value) if ok_value is not None else True
    status = _text(source.get("status") or ("completed" if ok else "failed"), 80)
    answer = _text(source.get("answer") or source.get("final_answer") or "", 1000)
    trace_summary = _mapping_copy(source.get("trace_summary"))
    runtime_state = _mapping_copy(source.get("runtime_state"))
    context_budget = _mapping_copy(source.get("context_budget"))
    request = _mapping_copy(source.get("request"))
    dispatch_trace = source.get("tool_dispatch_trace") if isinstance(source.get("tool_dispatch_trace"), list) else []
    dispatch_summary = (
        source.get("tool_dispatch_summary")
        if isinstance(source.get("tool_dispatch_summary"), Mapping)
        else summarize_tool_dispatch_trace(dispatch_trace)
    )
    source["tool_dispatch_summary"] = dispatch_summary
    journal = source.get("execution_journal")
    if not isinstance(journal, list):
        journal = build_execution_journal(source)
        source["execution_journal"] = journal
    journal_summary = source.get("journal_summary") if isinstance(source.get("journal_summary"), Mapping) else build_journal_summary(journal)
    source["journal_summary"] = journal_summary
    continuation_pack = source.get("continuation_pack") if isinstance(source.get("continuation_pack"), Mapping) else build_continuation_pack(source)
    source["continuation_pack"] = continuation_pack

    return {
        "agent_run_id": run_id,
        "mode": _text(source.get("mode"), 120),
        "profile": _text(source.get("profile"), 120),
        "status": status,
        "ok": ok,
        "answer_preview": answer,
        "trace_summary": trace_summary,
        "runtime_state": runtime_state,
        "context_budget": context_budget,
        "turn_context_snapshot": _mapping_copy(source.get("turn_context_snapshot")),
        "tool_dispatch_summary": dispatch_summary,
        "tool_call_count": _list_count(source.get("tool_calls")),
        "observation_count": _list_count(source.get("observations")),
        "agent_step_count": _list_count(source.get("agent_steps")),
        "evidence_node_count": _list_count(source.get("final_evidence")),
        "citation_count": _list_count(source.get("citations")),
        "proposed_action_count": _list_count(source.get("proposed_actions")),
        "subagent_count": _list_count(source.get("subagents")),
        "execution_journal": journal,
        "journal_summary": journal_summary,
        "continuation_pack": continuation_pack,
        "started_at": _text(source.get("started_at"), 80),
        "finished_at": _text(source.get("finished_at"), 80),
        "request": request,
        "payload": source,
    }


class AgentTraceStore:
    """Small LRU store that preserves legacy payload access and new envelopes."""

    def __init__(self, max_items: int = DEFAULT_MAX_AGENT_TRACE_ITEMS):
        self.max_items = max(1, int(max_items or DEFAULT_MAX_AGENT_TRACE_ITEMS))
        self._items: "OrderedDict[str, dict]" = OrderedDict()

    def remember(self,
                 payload: Mapping[str, object] | None,
                 *,
                 agent_run_id: object = None) -> dict | None:
        envelope = build_trace_envelope(payload, agent_run_id=agent_run_id)
        run_id = str(envelope.get("agent_run_id") or "").strip()
        if not run_id:
            return None
        self._items[run_id] = envelope
        self._items.move_to_end(run_id)
        while len(self._items) > self.max_items:
            self._items.popitem(last=False)
        return envelope

    def get_envelope(self, agent_run_id: object, default: Any = None) -> dict | Any:
        run_id = str(agent_run_id or "").strip()
        if not run_id or run_id not in self._items:
            return default
        return self._items[run_id]

    def get(self, agent_run_id: object, default: Any = None) -> dict | Any:
        envelope = self.get_envelope(agent_run_id)
        if not envelope:
            return default
        return envelope.get("payload") or default

    def clear(self) -> None:
        self._items.clear()

    def keys(self) -> list[str]:
        return list(self._items.keys())

    def values(self) -> list[dict]:
        return [dict(item.get("payload") or {}) for item in self._items.values()]

    def items(self) -> list[tuple[str, dict]]:
        return [(key, dict(item.get("payload") or {})) for key, item in self._items.items()]

    def envelope_items(self) -> list[tuple[str, dict]]:
        return list(self._items.items())

    def __contains__(self, agent_run_id: object) -> bool:
        return str(agent_run_id or "").strip() in self._items

    def __len__(self) -> int:
        return len(self._items)

    def __iter__(self) -> Iterator[str]:
        return iter(self._items)

    def __getitem__(self, agent_run_id: object) -> dict:
        envelope = self.get_envelope(agent_run_id)
        if not envelope:
            raise KeyError(agent_run_id)
        return dict(envelope.get("payload") or {})
