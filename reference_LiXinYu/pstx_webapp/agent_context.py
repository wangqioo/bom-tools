"""Project-scoped in-memory Agent context helpers for Web routes."""

from __future__ import annotations

import time
from typing import Tuple

from pstx_agent_runtime import (
    build_project_evidence_memory,
    compact_project_evidence_memory,
    merge_project_session_memory,
)
from pstx_webapp.json_utils import compact_value
from pstx_webapp.state import AGENT_CONTEXT_CACHE, MAX_RUNS


def new_agent_context() -> dict:
    return {
        "answers": [],
        "pending_questions": [],
        "recent_agent_runs": [],
        "recent_evidence_ids": [],
        "evidence_memory_cards": [],
        "session_memory_summary": {},
        "active_continuation_pack": {},
        "latest_continuation_pack": {},
        "updated_at": "",
    }


def get_agent_context(run_id: str) -> dict:
    context = AGENT_CONTEXT_CACHE.get(run_id)
    if context is None:
        context = new_agent_context()
        AGENT_CONTEXT_CACHE[run_id] = context
    AGENT_CONTEXT_CACHE.move_to_end(run_id)
    while len(AGENT_CONTEXT_CACHE) > MAX_RUNS:
        AGENT_CONTEXT_CACHE.popitem(last=False)
    return context


def agent_context_public(run_id: str, context: dict) -> dict:
    return {
        "run_id": run_id,
        "answer_count": len(context.get("answers") or []),
        "answers": list(context.get("answers") or [])[-24:],
        "pending_questions": list(context.get("pending_questions") or [])[:24],
        "recent_agent_runs": list(context.get("recent_agent_runs") or [])[-12:],
        "recent_evidence_ids": list(context.get("recent_evidence_ids") or [])[-48:],
        "evidence_memory_card_count": len(context.get("evidence_memory_cards") or []),
        "evidence_memory_cards": compact_project_evidence_memory(context.get("evidence_memory_cards") or [], limit=24),
        "session_memory_summary": dict(context.get("session_memory_summary") or {}),
        "active_continuation_pack": dict(context.get("active_continuation_pack") or {}),
        "latest_continuation_pack": dict(context.get("latest_continuation_pack") or {}),
        "updated_at": context.get("updated_at", ""),
        "storage": "memory",
    }


def append_agent_context_answers(context: dict, answers: Tuple[dict, ...], *, source_agent_run_id: str = "") -> None:
    if not answers:
        return
    now = time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())
    existing = list(context.get("answers") or [])
    answered_ids = set()
    for item in answers:
        question_id = str(item.get("question_id") or "").strip()
        answer = str(item.get("answer") or "").strip()
        if not question_id or not answer:
            continue
        answered_ids.add(question_id)
        existing.append({
            "question_id": question_id,
            "answer": answer[:4000],
            "applies_to": dict(item.get("applies_to") or {}),
            "source_agent_run_id": source_agent_run_id,
            "created_at": now,
        })
    context["answers"] = existing[-100:]
    context["pending_questions"] = [
        item for item in list(context.get("pending_questions") or [])
        if str(item.get("question_id") or "") not in answered_ids
    ]
    context["updated_at"] = now


def update_agent_context_after_run(run_id: str, context: dict, result: dict) -> None:
    now = time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())
    agent_run_id = str(result.get("agent_run_id") or "")
    if agent_run_id:
        recent_runs = list(context.get("recent_agent_runs") or [])
        recent_runs.append({
            "agent_run_id": agent_run_id,
            "profile": result.get("profile", ""),
            "status": result.get("status", ""),
            "stopped_reason": (result.get("model_metadata") or {}).get("stopped_reason", ""),
            "question": compact_value((result.get("request") or {}).get("question", ""), 300),
            "created_at": now,
        })
        context["recent_agent_runs"] = recent_runs[-20:]
    evidence_ids = list(context.get("recent_evidence_ids") or [])
    for evidence_id in [str(item.get("id") or "") for item in list(result.get("final_evidence") or [])[:24]]:
        if evidence_id and evidence_id not in evidence_ids:
            evidence_ids.append(evidence_id)
    context["recent_evidence_ids"] = evidence_ids[-80:]
    needs_user_input = result.get("needs_user_input") if isinstance(result.get("needs_user_input"), dict) else {}
    if result.get("status") == "waiting_for_user":
        context["pending_questions"] = list(needs_user_input.get("questions") or [])
    elif result.get("status") == "completed":
        context["pending_questions"] = []
    if isinstance(result.get("continuation_pack"), dict):
        context["latest_continuation_pack"] = dict(result.get("continuation_pack") or {})
    context["evidence_memory_cards"] = build_project_evidence_memory(context, result)
    context["session_memory_summary"] = merge_project_session_memory(context, result)
    context["updated_at"] = now
