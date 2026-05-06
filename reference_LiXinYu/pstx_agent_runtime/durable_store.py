# -*- coding: utf-8 -*-
"""File-backed durable run store for PSTX Harness Agent jobs."""

from __future__ import annotations

from pathlib import Path
import json
import os
import threading
import time
import uuid
from typing import Mapping

from .workspace import (
    AGENT_WORKSPACE_VERSION,
    ensure_scope_workspace,
    list_workspace_artifacts,
    safe_workspace_id,
    write_workspace_draft,
    write_task_markdown,
    write_workspace_artifact,
    agent_workspace_root,
)
from .task_dispatch import compact_dispatch_records


DURABLE_RUN_STORE_VERSION = "pstx-agent-durable-runs/v1"
RUN_STATUSES = {"queued", "running", "waiting_for_user", "completed", "failed", "cancelled", "incomplete"}
DEFAULT_HEARTBEAT_TIMEOUT_SECONDS = 900


def new_agent_run_id(prefix: str = "ar") -> str:
    prefix = safe_workspace_id(prefix, "ar")[:12].lower()
    return f"{prefix}_{uuid.uuid4().hex[:16]}"


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S")


def _parse_time(value: object) -> float:
    text = str(value or "").strip()
    if not text:
        return 0.0
    try:
        return time.mktime(time.strptime(text[:19], "%Y-%m-%dT%H:%M:%S"))
    except Exception:
        return 0.0


def _compact(value: object, limit: int = 4000) -> object:
    if isinstance(value, Mapping):
        return {str(key)[:80]: _compact(item, limit=limit // 2) for key, item in list(value.items())[:32]}
    if isinstance(value, list):
        return [_compact(item, limit=limit // 2) for item in value[:40]]
    text = "" if value is None else str(value)
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _task_markdown(record: Mapping[str, object]) -> str:
    result = record.get("result") if isinstance(record.get("result"), Mapping) else {}
    request = record.get("request") if isinstance(record.get("request"), Mapping) else {}
    lines = [
        "# PSTX Agent Workspace Task",
        "",
        f"- store_version: `{DURABLE_RUN_STORE_VERSION}`",
        f"- agent_run_id: `{record.get('agent_run_id', '')}`",
        f"- kind: `{record.get('kind', '')}`",
        f"- status: `{record.get('status', '')}`",
        f"- scope_id: `{record.get('scope_id', '')}`",
        f"- parent_agent_run_id: `{record.get('parent_agent_run_id', '')}`",
        f"- child_agent_run_count: `{len(record.get('child_agent_run_ids') or [])}`",
        f"- updated_at: `{record.get('updated_at', '')}`",
        "",
        "## Request",
        "",
        f"- profile: `{request.get('profile', '')}`",
        f"- question: {_compact(request.get('question', ''), 1200)}",
        "",
        "## Latest Answer",
        "",
        str(_compact(result.get("answer", ""), 1800) or "-"),
        "",
        "## Next",
        "",
        str(_compact((result.get("continuation_pack") or {}).get("continuation_brief", ""), 1200) if isinstance(result.get("continuation_pack"), Mapping) else "-"),
    ]
    return "\n".join(lines).strip() + "\n"


def _ledger_markdown(record: Mapping[str, object]) -> str:
    request = record.get("request") if isinstance(record.get("request"), Mapping) else {}
    ledger = record.get("task_ledger") if isinstance(record.get("task_ledger"), Mapping) else {}
    checkpoint = record.get("checkpoint") if isinstance(record.get("checkpoint"), Mapping) else {}
    lines = [
        "# Agent Task Ledger",
        "",
        f"- agent_run_id: `{record.get('agent_run_id', '')}`",
        f"- status: `{record.get('status', '')}`",
        f"- current_phase: `{record.get('current_phase') or checkpoint.get('phase') or ''}`",
        f"- heartbeat_at: `{record.get('heartbeat_at', '')}`",
        f"- question: {_compact(request.get('question', ''), 1200)}",
        "",
        "## Ledger",
        "",
        "```json",
        json.dumps(_compact(ledger, 6000), ensure_ascii=False, indent=2, default=str),
        "```",
        "",
        "## Next Actions",
        "",
    ]
    for item in record.get("next_actions") or (ledger.get("next_actions") if isinstance(ledger, Mapping) else []) or []:
        lines.append(f"- {_compact(item, 500)}")
    if lines[-1] == "## Next Actions":
        lines.append("- 暂无。")
    return "\n".join(lines).strip() + "\n"


def _review_draft_markdown(record: Mapping[str, object], result: Mapping[str, object]) -> str:
    answer = str(result.get("answer") or "").strip()
    citations = result.get("citations") if isinstance(result.get("citations"), list) else []
    actions = result.get("proposed_actions") if isinstance(result.get("proposed_actions"), list) else []
    lines = [
        "# Agent Review Draft",
        "",
        f"- agent_run_id: `{record.get('agent_run_id', '')}`",
        f"- status: `{record.get('status', '')}`",
        "",
        "## Answer",
        "",
        answer or "暂无最终回答。",
        "",
        "## Citations",
        "",
    ]
    for item in citations[:50]:
        lines.append(f"- {_compact(item, 600)}")
    if not citations:
        lines.append("- 暂无。")
    lines.extend(["", "## Proposed Actions", ""])
    for item in actions[:50]:
        lines.append(f"- {_compact(item, 600)}")
    if not actions:
        lines.append("- 暂无。")
    return "\n".join(lines).strip() + "\n"


class AgentDurableRunStore:
    """Atomic JSON store under agent_workspace for background agent runs."""

    def __init__(self, *, root: str | Path | None = None):
        self.root = Path(root).expanduser().resolve() if root else agent_workspace_root()
        self._lock = threading.RLock()

    def create_run(self,
                   *,
                   scope_id: object,
                   kind: str,
                   request: Mapping[str, object],
                   initial_status: str = "queued",
                   agent_run_id: str = "",
                   parent_agent_run_id: object = "",
                   root_agent_run_id: object = "",
                   dispatch_task: Mapping[str, object] | None = None,
                   dispatch_group_id: object = "") -> dict:
        status = initial_status if initial_status in RUN_STATUSES else "queued"
        run_id = safe_workspace_id(agent_run_id, "") or new_agent_run_id("agent")
        workspace = ensure_scope_workspace(scope_id, root=self.root)
        parent_id = safe_workspace_id(parent_agent_run_id, "") if parent_agent_run_id else ""
        root_id = safe_workspace_id(root_agent_run_id, "") if root_agent_run_id else (parent_id or run_id)
        record = {
            "version": DURABLE_RUN_STORE_VERSION,
            "workspace_version": AGENT_WORKSPACE_VERSION,
            "agent_run_id": run_id,
            "scope_id": workspace["scope_id"],
            "kind": str(kind or "agent"),
            "parent_agent_run_id": parent_id,
            "root_agent_run_id": root_id,
            "child_agent_run_ids": [],
            "dispatch_group_id": safe_workspace_id(dispatch_group_id, "") if dispatch_group_id else "",
            "dispatch_task": dict(dispatch_task or {}),
            "dispatch_tasks": [],
            "task_dispatch_summary": {},
            "status": status,
            "request": dict(request or {}),
            "checkpoint": {},
            "current_phase": status,
            "heartbeat_at": _now(),
            "progress": {
                "step_index": 0,
                "max_steps": int(request.get("max_steps") or 0) if isinstance(request, Mapping) else 0,
                "tool_call_count": 0,
                "max_tool_calls": int(request.get("max_tool_calls") or 0) if isinstance(request, Mapping) else 0,
                "evidence_count": 0,
            },
            "steps": [],
            "tool_calls": [],
            "partial_observations": [],
            "evidence_ids": [],
            "retry_reasons": [],
            "selected_skills": {},
            "task_ledger": {},
            "continuation_pack": {},
            "partial_trace": {},
            "next_actions": [],
            "result": {},
            "error": "",
            "cancel_requested": False,
            "created_at": _now(),
            "updated_at": _now(),
            "workspace": workspace,
            "artifacts": [],
        }
        self.write_record(record)
        return record

    def _path_for(self, scope_id: object, agent_run_id: object) -> Path:
        workspace = ensure_scope_workspace(scope_id, root=self.root)
        return Path(workspace["runs_dir"]) / f"{safe_workspace_id(agent_run_id, 'run')}.json"

    def _find_path(self, agent_run_id: object, scope_id: object = "") -> Path | None:
        run_id = safe_workspace_id(agent_run_id, "")
        if not run_id:
            return None
        if scope_id:
            path = self._path_for(scope_id, run_id)
            return path if path.is_file() else None
        for path in self.root.glob(f"*/runs/{run_id}.json"):
            if path.is_file():
                return path
        return None

    def read_record(self, agent_run_id: object, *, scope_id: object = "") -> dict:
        with self._lock:
            path = self._find_path(agent_run_id, scope_id=scope_id)
            if not path:
                return {}
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
            except Exception as exc:
                return {
                    "version": DURABLE_RUN_STORE_VERSION,
                    "agent_run_id": safe_workspace_id(agent_run_id, "run"),
                    "status": "failed",
                    "error": f"durable run JSON unreadable: {exc}",
                    "path": str(path),
                }
            if isinstance(payload, dict):
                payload.setdefault("path", str(path))
                return payload
            return {"status": "failed", "error": "durable run JSON is not an object", "path": str(path)}

    def write_record(self, record: Mapping[str, object]) -> dict:
        with self._lock:
            payload = dict(record or {})
            scope_id = payload.get("scope_id") or "scope"
            run_id = payload.get("agent_run_id") or new_agent_run_id("agent")
            payload["agent_run_id"] = safe_workspace_id(run_id, "run")
            payload["scope_id"] = safe_workspace_id(scope_id, "scope")
            payload["updated_at"] = _now()
            path = self._path_for(payload["scope_id"], payload["agent_run_id"])
            tmp = path.with_suffix(".json.tmp")
            tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True, default=str), encoding="utf-8")
            tmp.replace(path)
            payload["path"] = str(path)
            return payload

    def update_record(self, agent_run_id: object, *, scope_id: object = "", **patch) -> dict:
        with self._lock:
            record = self.read_record(agent_run_id, scope_id=scope_id)
            if not record:
                return {}
            record.update(patch)
            if "status" in record and record["status"] not in RUN_STATUSES:
                record["status"] = "failed"
            return self.write_record(record)

    def append_child_runs(self,
                          agent_run_id: object,
                          child_runs: list[Mapping[str, object]],
                          *,
                          scope_id: object = "",
                          task_dispatch_summary: Mapping[str, object] | None = None) -> dict:
        """Attach durable child run metadata to a parent run."""

        record = self.read_record(agent_run_id, scope_id=scope_id)
        if not record:
            return {}
        existing_ids = [str(item or "") for item in (record.get("child_agent_run_ids") or []) if str(item or "")]
        dispatch_tasks = list(record.get("dispatch_tasks") or [])
        for child in child_runs or []:
            if not isinstance(child, Mapping):
                continue
            child_id = safe_workspace_id(child.get("agent_run_id"), "")
            if child_id and child_id not in existing_ids:
                existing_ids.append(child_id)
            dispatch_tasks.append(dict(child))
        summary = dict(record.get("task_dispatch_summary") if isinstance(record.get("task_dispatch_summary"), Mapping) else {})
        if task_dispatch_summary:
            summary.update(dict(task_dispatch_summary))
        summary["child_count"] = len(existing_ids)
        return self.update_record(
            agent_run_id,
            scope_id=scope_id,
            child_agent_run_ids=existing_ids,
            dispatch_tasks=compact_dispatch_records(dispatch_tasks, limit=50),
            task_dispatch_summary=summary,
        )

    def mark_incomplete(self, agent_run_id: object, *, reason: object = "") -> dict:
        record = self.read_record(agent_run_id)
        if not record:
            return {}
        if str(record.get("status") or "") not in {"queued", "running", "failed"}:
            return record
        return self.update_record(
            agent_run_id,
            status="incomplete",
            current_phase="incomplete",
            error=str(reason or record.get("error") or "agent run became incomplete"),
            checkpoint={
                **(record.get("checkpoint") if isinstance(record.get("checkpoint"), Mapping) else {}),
                "phase": "incomplete",
                "reason": str(reason or ""),
                "heartbeat_at": record.get("heartbeat_at", ""),
            },
        )

    def mark_cancel_requested(self, agent_run_id: object) -> dict:
        record = self.read_record(agent_run_id)
        if not record:
            return {}
        status = str(record.get("status") or "")
        next_status = "cancelled" if status == "queued" else status
        phase = "cancelled" if next_status == "cancelled" else "cancel_requested"
        return self.update_record(
            agent_run_id,
            cancel_requested=True,
            status=next_status,
            current_phase=phase,
            checkpoint={"phase": phase, "heartbeat_at": _now()},
        )

    def finish_record(self, agent_run_id: object, result: Mapping[str, object]) -> dict:
        record = self.read_record(agent_run_id)
        if not record:
            return {}
        result_payload = dict(result or {})
        status = "completed"
        if result_payload.get("status") == "waiting_for_user":
            status = "waiting_for_user"
        elif result_payload.get("ok") is False:
            status = "failed"
        if record.get("cancel_requested"):
            status = "cancelled"
        artifacts = list(record.get("artifacts") or [])
        artifacts.append(write_workspace_artifact(
            record.get("scope_id"),
            agent_run_id,
            "result.json",
            result_payload,
            root=self.root,
            content_type="application/json",
        ))
        trace_payload = {
            "agent_run_id": agent_run_id,
            "trace_summary": result_payload.get("trace_summary") or {},
            "agent_steps": result_payload.get("agent_steps") or [],
            "tool_calls": result_payload.get("tool_calls") or [],
            "citations": result_payload.get("citations") or [],
            "final_evidence": result_payload.get("final_evidence") or [],
            "execution_journal": result_payload.get("execution_journal") or [],
            "journal_summary": result_payload.get("journal_summary") or {},
            "model_metadata": result_payload.get("model_metadata") or {},
            "continuation_pack": result_payload.get("continuation_pack") or {},
            "task_dispatch_summary": result_payload.get("task_dispatch_summary") or {},
            "dispatched_tasks": result_payload.get("dispatched_tasks") or [],
        }
        artifacts.append(write_workspace_artifact(
            record.get("scope_id"),
            agent_run_id,
            "trace.json",
            trace_payload,
            root=self.root,
            content_type="application/json",
        ))
        artifacts.append(write_workspace_artifact(
            record.get("scope_id"),
            agent_run_id,
            "evidence_cards.json",
            result_payload.get("final_evidence") or [],
            root=self.root,
            content_type="application/json",
        ))
        if result_payload.get("answer"):
            artifacts.append(write_workspace_artifact(
                record.get("scope_id"),
                agent_run_id,
                "answer.md",
                str(result_payload.get("answer") or ""),
                root=self.root,
                content_type="text/markdown",
            ))
        if result_payload.get("task_dispatch_summary") or result_payload.get("dispatched_tasks"):
            artifacts.append(write_workspace_artifact(
                record.get("scope_id"),
                agent_run_id,
                "task_dispatch.json",
                {
                    "task_dispatch_summary": result_payload.get("task_dispatch_summary") or {},
                    "dispatched_tasks": result_payload.get("dispatched_tasks") or [],
                },
                root=self.root,
                content_type="application/json",
            ))
        child_ids = list(record.get("child_agent_run_ids") or [])
        for item in result_payload.get("dispatched_tasks") or []:
            if not isinstance(item, Mapping):
                continue
            child_id = safe_workspace_id(item.get("agent_run_id"), "")
            if child_id and child_id not in child_ids:
                child_ids.append(child_id)
        artifacts.append(write_workspace_artifact(
            record.get("scope_id"),
            agent_run_id,
            "task_ledger.md",
            _ledger_markdown({**record, "status": status, "task_ledger": (result_payload.get("runtime_state") or {}).get("task_ledger") if isinstance(result_payload.get("runtime_state"), Mapping) else record.get("task_ledger") or {}}),
            root=self.root,
            content_type="text/markdown",
        ))
        artifacts.append(write_workspace_draft(
            record.get("scope_id"),
            agent_run_id,
            "review_draft.md",
            _review_draft_markdown({**record, "status": status}, result_payload),
            root=self.root,
            content_type="text/markdown",
        ))
        updated = self.update_record(
            agent_run_id,
            status=status,
            current_phase=status,
            heartbeat_at=_now(),
            result=result_payload,
            continuation_pack=result_payload.get("continuation_pack") or {},
            selected_skills=result_payload.get("selected_skills") or {},
            child_agent_run_ids=child_ids,
            dispatch_tasks=compact_dispatch_records(result_payload.get("dispatched_tasks") or record.get("dispatch_tasks") or [], limit=50),
            task_dispatch_summary=result_payload.get("task_dispatch_summary") or record.get("task_dispatch_summary") or {},
            task_ledger=(result_payload.get("runtime_state") or {}).get("task_ledger") if isinstance(result_payload.get("runtime_state"), Mapping) else {},
            retry_reasons=result_payload.get("retry_reasons") or [],
            steps=result_payload.get("agent_steps") or [],
            tool_calls=result_payload.get("tool_calls") or [],
            evidence_ids=[str((item or {}).get("id") or "") for item in (result_payload.get("final_evidence") or []) if isinstance(item, Mapping) and (item or {}).get("id")],
            partial_trace=trace_payload,
            progress={
                "step_index": len(result_payload.get("agent_steps") or []),
                "max_steps": int(((result_payload.get("limits") or {}).get("max_steps") if isinstance(result_payload.get("limits"), Mapping) else 0) or 0),
                "tool_call_count": len(result_payload.get("tool_calls") or []),
                "max_tool_calls": int(((result_payload.get("limits") or {}).get("max_tool_calls") if isinstance(result_payload.get("limits"), Mapping) else 0) or 0),
                "evidence_count": len(result_payload.get("final_evidence") or []),
            },
            next_actions=((result_payload.get("runtime_state") or {}).get("task_ledger") or {}).get("next_actions", []) if isinstance(result_payload.get("runtime_state"), Mapping) and isinstance((result_payload.get("runtime_state") or {}).get("task_ledger"), Mapping) else [],
            artifacts=artifacts,
            error="",
        )
        write_task_markdown(updated.get("scope_id"), _task_markdown(updated), root=self.root)
        return updated

    def fail_record(self, agent_run_id: object, error: object) -> dict:
        record = self.read_record(agent_run_id)
        if not record:
            return {}
        status = "cancelled" if record.get("cancel_requested") else "failed"
        return self.update_record(
            agent_run_id,
            status=status,
            current_phase=status,
            heartbeat_at=_now(),
            error=str(error or "unknown error"),
            checkpoint={
                **(record.get("checkpoint") if isinstance(record.get("checkpoint"), Mapping) else {}),
                "phase": status,
                "error": str(error or "unknown error"),
            },
        )

    def list_artifacts(self, agent_run_id: object) -> dict:
        record = self.read_record(agent_run_id)
        if not record:
            return {}
        return list_workspace_artifacts(record.get("scope_id"), agent_run_id, root=self.root)

    def public_status(self, agent_run_id: object) -> dict:
        record = self.read_record(agent_run_id)
        if not record:
            return {"ok": False, "error": f"未找到 agent_run_id：{agent_run_id}"}
        record = self._maybe_mark_stale_incomplete(record)
        result = record.get("result") if isinstance(record.get("result"), Mapping) else {}
        status = str(record.get("status") or "")
        checkpoint = record.get("checkpoint") if isinstance(record.get("checkpoint"), Mapping) else {}
        current_phase = record.get("current_phase") or checkpoint.get("phase") or status
        partial_trace = record.get("partial_trace") if isinstance(record.get("partial_trace"), Mapping) else {}
        return {
            "ok": True,
            "version": DURABLE_RUN_STORE_VERSION,
            "agent_run_id": record.get("agent_run_id", ""),
            "scope_id": record.get("scope_id", ""),
            "kind": record.get("kind", ""),
            "status": status,
            "current_phase": current_phase,
            "heartbeat_at": record.get("heartbeat_at", ""),
            "progress": record.get("progress") or {},
            "can_continue": status in {"waiting_for_user", "failed", "incomplete"} and not bool(record.get("cancel_requested")),
            "can_cancel": status in {"queued", "running"},
            "created_at": record.get("created_at", ""),
            "updated_at": record.get("updated_at", ""),
            "request": record.get("request") or {},
            "checkpoint": record.get("checkpoint") or {},
            "continuation_pack": record.get("continuation_pack") or {},
            "selected_skills": record.get("selected_skills") or {},
            "retry_reasons": record.get("retry_reasons") or [],
            "error": record.get("error", ""),
            "cancel_requested": bool(record.get("cancel_requested")),
            "artifact_count": len(record.get("artifacts") or []),
            "partial_trace": partial_trace,
            "next_actions": record.get("next_actions") or [],
            "parent_agent_run_id": record.get("parent_agent_run_id", ""),
            "root_agent_run_id": record.get("root_agent_run_id", ""),
            "child_agent_run_ids": record.get("child_agent_run_ids") or [],
            "dispatch_group_id": record.get("dispatch_group_id", ""),
            "dispatch_task": record.get("dispatch_task") or {},
            "dispatch_tasks": record.get("dispatch_tasks") or [],
            "task_dispatch_summary": record.get("task_dispatch_summary") or {},
            "result_available": bool(result),
            "agent_run": result,
            "trace": result.get("trace_summary") if isinstance(result, Mapping) else partial_trace,
            "workspace": record.get("workspace") or {},
        }

    def _maybe_mark_stale_incomplete(self, record: Mapping[str, object]) -> dict:
        if str(record.get("status") or "") != "running":
            return dict(record)
        try:
            timeout = int(os.environ.get("PSTX_AGENT_HEARTBEAT_TIMEOUT_SECONDS") or DEFAULT_HEARTBEAT_TIMEOUT_SECONDS)
        except Exception:
            timeout = DEFAULT_HEARTBEAT_TIMEOUT_SECONDS
        if timeout <= 0:
            return dict(record)
        heartbeat_at = _parse_time(record.get("heartbeat_at") or record.get("updated_at"))
        if heartbeat_at and time.time() - heartbeat_at > timeout:
            updated = self.mark_incomplete(record.get("agent_run_id"), reason=f"heartbeat stale for more than {timeout}s")
            return updated or dict(record)
        return dict(record)
