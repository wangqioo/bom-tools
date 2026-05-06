# -*- coding: utf-8 -*-
"""Step-level checkpoint reporter for durable PSTX agent runs."""

from __future__ import annotations

from typing import Mapping
import time

from .workspace import append_workspace_log


CHECKPOINT_REPORTER_VERSION = "pstx-agent-checkpoint-reporter/v1"


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S")


def _compact(value: object, *, depth: int = 3, list_limit: int = 40, text_limit: int = 4000) -> object:
    if depth <= 0:
        text = "" if value is None else str(value)
        return text if len(text) <= text_limit else text[: max(0, text_limit - 1)] + "…"
    if isinstance(value, Mapping):
        return {
            str(key)[:80]: _compact(item, depth=depth - 1, list_limit=list_limit, text_limit=max(400, text_limit // 2))
            for key, item in list(value.items())[:64]
        }
    if isinstance(value, list):
        return [
            _compact(item, depth=depth - 1, list_limit=list_limit, text_limit=max(400, text_limit // 2))
            for item in value[:list_limit]
        ]
    text = "" if value is None else str(value)
    return text if len(text) <= text_limit else text[: max(0, text_limit - 1)] + "…"


def _evidence_ids_from_payload(payload: Mapping[str, object]) -> list[str]:
    explicit = [str(item) for item in payload.get("evidence_ids") or [] if str(item or "").strip()]
    if explicit:
        return explicit[-300:]
    ids: list[str] = []
    for item in payload.get("evidence_nodes") or payload.get("final_evidence") or []:
        if isinstance(item, Mapping) and item.get("id"):
            ids.append(str(item.get("id")))
    return ids[-300:]


def _progress_from_payload(payload: Mapping[str, object]) -> dict:
    progress = dict(payload.get("progress") or {}) if isinstance(payload.get("progress"), Mapping) else {}
    if "step_index" not in progress and payload.get("step_index") is not None:
        progress["step_index"] = int(payload.get("step_index") or 0)
    if "max_steps" not in progress and payload.get("max_steps") is not None:
        progress["max_steps"] = int(payload.get("max_steps") or 0)
    if "tool_call_count" not in progress:
        progress["tool_call_count"] = len(payload.get("tool_calls") or [])
    if "max_tool_calls" not in progress and payload.get("max_tool_calls") is not None:
        progress["max_tool_calls"] = int(payload.get("max_tool_calls") or 0)
    if "evidence_count" not in progress:
        progress["evidence_count"] = len(_evidence_ids_from_payload(payload))
    return progress


class AgentCheckpointReporter:
    """Bridge agent loop events into durable run JSON + workspace JSONL logs."""

    def __init__(self, store, agent_run_id: object, *, scope_id: object = "", kind: str = ""):
        self.store = store
        self.agent_run_id = str(agent_run_id or "")
        self.scope_id = str(scope_id or "")
        self.kind = str(kind or "")

    def __call__(self, payload: Mapping[str, object] | None = None, **kwargs) -> dict:
        data = dict(payload or {})
        data.update(kwargs)
        return self.emit(data)

    def emit(self, payload: Mapping[str, object] | None = None, **kwargs) -> dict:
        data = dict(payload or {})
        data.update(kwargs)
        phase = str(data.get("phase") or data.get("current_phase") or "checkpoint")
        heartbeat_at = _now()
        progress = _progress_from_payload(data)
        checkpoint = {
            "version": CHECKPOINT_REPORTER_VERSION,
            "phase": phase,
            "heartbeat_at": heartbeat_at,
            "progress": progress,
            "summary": _compact(data.get("summary") or data.get("message") or "", depth=1, text_limit=1200),
        }
        partial_trace = {
            "phase": phase,
            "progress": progress,
            "steps": _compact(data.get("agent_steps") or data.get("steps") or [], depth=3, list_limit=80),
            "tool_calls": _compact(data.get("tool_calls") or [], depth=3, list_limit=120),
            "partial_observations": _compact(data.get("partial_observations") or data.get("observations") or [], depth=3, list_limit=40),
            "evidence_ids": _evidence_ids_from_payload(data),
            "selected_skills": _compact(data.get("selected_skills") or {}, depth=3),
            "playbook_plan": _compact(data.get("playbook_plan") or {}, depth=3),
            "task_ledger": _compact(data.get("task_ledger") or {}, depth=3),
            "retry_reasons": _compact(data.get("retry_reasons") or [], depth=2),
        }
        patch = {
            "status": "running" if phase not in {"waiting_for_user", "completed", "failed", "cancelled", "incomplete"} else phase,
            "current_phase": phase,
            "heartbeat_at": heartbeat_at,
            "progress": progress,
            "checkpoint": checkpoint,
            "partial_trace": partial_trace,
            "partial_observations": partial_trace["partial_observations"],
            "evidence_ids": partial_trace["evidence_ids"],
        }
        for key in ("agent_steps", "steps"):
            if isinstance(data.get(key), list):
                patch["steps"] = data[key]
                break
        if isinstance(data.get("tool_calls"), list):
            patch["tool_calls"] = data["tool_calls"]
        if isinstance(data.get("selected_skills"), Mapping):
            patch["selected_skills"] = data["selected_skills"]
        if isinstance(data.get("task_ledger"), Mapping):
            patch["task_ledger"] = data["task_ledger"]
            patch["next_actions"] = data["task_ledger"].get("next_actions") or []
        if isinstance(data.get("continuation_pack"), Mapping):
            patch["continuation_pack"] = data["continuation_pack"]
        if data.get("retry_reasons") is not None:
            patch["retry_reasons"] = list(data.get("retry_reasons") or [])
        if data.get("error"):
            patch["error"] = str(data.get("error"))
        record = self.store.update_record(self.agent_run_id, **patch)
        scope_id = self.scope_id or (record or {}).get("scope_id") or "scope"
        append_workspace_log(scope_id, self.agent_run_id, {
            "version": CHECKPOINT_REPORTER_VERSION,
            "agent_run_id": self.agent_run_id,
            "kind": self.kind or (record or {}).get("kind", ""),
            "phase": phase,
            "heartbeat_at": heartbeat_at,
            "progress": progress,
            "summary": checkpoint.get("summary", ""),
            "error": str(data.get("error") or ""),
        }, root=self.store.root)
        return record

    def cancel_requested(self) -> bool:
        record = self.store.read_record(self.agent_run_id)
        return bool(record.get("cancel_requested")) if record else False


__all__ = ["CHECKPOINT_REPORTER_VERSION", "AgentCheckpointReporter"]
