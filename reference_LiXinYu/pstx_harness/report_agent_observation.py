# -*- coding: utf-8 -*-
"""Observation and trace payload helpers for the report harness agent."""

from __future__ import annotations

from typing import List, Optional, Sequence

from pstx_agent_runtime import (
    build_context_budget_summary,
    compact_project_evidence_memory,
    fit_items_to_json_budget,
    json_char_count as runtime_json_char_count,
)
from pstx_harness.report_agent_config import (
    HARNESS_AGENT_MODEL_JSON_BUDGET,
    HARNESS_AGENT_MODEL_NODE_LIMIT,
    HARNESS_AGENT_MODEL_OBSERVATION_LIMIT,
    HARNESS_AGENT_MODEL_TEXT_LIMIT,
)


def summarize_observation(tool_name: str, result: dict) -> dict:
    return {
        "tool": tool_name,
        "ok": True,
        "id": result.get("id", tool_name),
        "title": result.get("title", tool_name),
        "summary": result.get("summary", ""),
        "keys": sorted(str(key) for key in result.keys())[:20],
    }


def preview(value, limit: int = 500):
    if isinstance(value, dict):
        return {str(key): preview(item, 180) for key, item in list(value.items())[:20]}
    if isinstance(value, list):
        return [preview(item, 180) for item in value[:12]]
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def json_char_count(value) -> int:
    return runtime_json_char_count(value)


def model_preview(value, *, depth: int = 0, text_limit: int = HARNESS_AGENT_MODEL_TEXT_LIMIT):
    if depth >= 4:
        return preview(value, 220)
    if isinstance(value, dict):
        priority_keys = [
            "id", "title", "summary", "severity", "issue_count", "target", "table_id", "count",
            "metrics", "notes", "sample_rows", "rows", "query_result", "path", "truncated",
            "tool_result_contract", "evidence_layers", "completeness", "recommended_next_tools",
        ]
        keys = []
        for key in priority_keys:
            if key in value and key not in keys:
                keys.append(key)
        for key in value.keys():
            if key not in keys:
                keys.append(key)
            if len(keys) >= 16:
                break
        return {str(key): model_preview(value.get(key), depth=depth + 1, text_limit=240) for key in keys}
    if isinstance(value, list):
        limit = 8 if depth <= 1 else 5
        items = [model_preview(item, depth=depth + 1, text_limit=240) for item in value[:limit]]
        if len(value) > limit:
            items.append({"omitted_count": len(value) - limit})
        return items
    return preview(value, text_limit)


def compact_project_context_for_model(context: dict) -> dict:
    if not isinstance(context, dict):
        return {}
    answers = []
    for item in list(context.get("answers") or [])[-16:]:
        if not isinstance(item, dict):
            continue
        answers.append({
            "question_id": preview(item.get("question_id", ""), 120),
            "answer": preview(item.get("answer", ""), 500),
            "applies_to": model_preview(item.get("applies_to") or {}, depth=1, text_limit=160),
            "source_agent_run_id": preview(item.get("source_agent_run_id", ""), 80),
        })
    pending_questions = []
    for item in list(context.get("pending_questions") or [])[:12]:
        if not isinstance(item, dict):
            continue
        pending_questions.append({
            "question_id": preview(item.get("question_id", ""), 120),
            "question": preview(item.get("question", ""), 300),
            "missing_fields": list(item.get("missing_fields") or [])[:8],
            "applies_to": model_preview(item.get("applies_to") or {}, depth=1, text_limit=160),
        })

    def _compact_pack(pack: dict) -> dict:
        if not isinstance(pack, dict):
            return {}
        return {
            "version": preview(pack.get("version", ""), 80),
            "agent_run_id": preview(pack.get("agent_run_id", ""), 100),
            "profile": preview(pack.get("profile", ""), 80),
            "status": preview(pack.get("status", ""), 80),
            "stopped_reason": preview(pack.get("stopped_reason", ""), 120),
            "next_intent": preview(pack.get("next_intent", ""), 100),
            "goal": preview(pack.get("goal", ""), 420),
            "continuation_brief": preview(pack.get("continuation_brief", ""), 700),
            "evidence_ids": list(pack.get("evidence_ids") or [])[:24],
            "pending_questions": model_preview(pack.get("pending_questions") or [], depth=1, text_limit=220),
            "open_ledger_items": model_preview(pack.get("open_ledger_items") or [], depth=1, text_limit=220),
            "suggested_tool_calls": model_preview(pack.get("suggested_tool_calls") or [], depth=1, text_limit=220),
            "quality_status": preview(pack.get("quality_status", ""), 80),
            "quality_score": pack.get("quality_score", 0),
        }

    def _compact_session_memory(memory: dict) -> dict:
        if not isinstance(memory, dict):
            return {}
        return {
            "version": preview(memory.get("version", ""), 80),
            "goal": preview(memory.get("goal", ""), 420),
            "facts": model_preview(memory.get("facts") or [], depth=1, text_limit=220),
            "decisions": model_preview(memory.get("decisions") or [], depth=1, text_limit=220),
            "open_questions": model_preview(memory.get("open_questions") or [], depth=1, text_limit=220),
            "open_items": model_preview(memory.get("open_items") or [], depth=1, text_limit=220),
            "next_actions": model_preview(memory.get("next_actions") or [], depth=1, text_limit=220),
            "evidence_ids": list(memory.get("evidence_ids") or [])[:32],
            "source_agent_run_ids": list(memory.get("source_agent_run_ids") or [])[-8:],
            "updated_from_agent_run_id": preview(memory.get("updated_from_agent_run_id", ""), 100),
            "quality_status": preview(memory.get("quality_status", ""), 80),
        }

    return {
        "answer_count": len(context.get("answers") or []),
        "answers": answers,
        "pending_questions": pending_questions,
        "recent_agent_runs": list(context.get("recent_agent_runs") or [])[-8:],
        "recent_evidence_ids": list(context.get("recent_evidence_ids") or [])[-24:],
        "evidence_memory_cards": compact_project_evidence_memory(context.get("evidence_memory_cards") or [], limit=16),
        "session_memory_summary": _compact_session_memory(context.get("session_memory_summary") or {}),
        "active_continuation_pack": _compact_pack(context.get("active_continuation_pack") or {}),
        "latest_continuation_pack": _compact_pack(context.get("latest_continuation_pack") or {}),
    }


def compact_evidence_node_for_model(node: dict) -> dict:
    compact = {
        "id": str(node.get("id") or ""),
        "type": str(node.get("type") or ""),
        "title": preview(node.get("title") or "", 120),
        "summary": preview(node.get("summary") or "", 180),
        "locator": model_preview(node.get("locator") or {}, depth=1, text_limit=160),
    }
    missing_fields = list(node.get("missing_fields") or [])[:12]
    if missing_fields:
        compact["missing_fields"] = missing_fields
    detail_tool = node.get("detail_tool")
    if isinstance(detail_tool, dict):
        compact["detail_tool"] = model_preview(detail_tool, depth=1, text_limit=160)
    source = node.get("source")
    if isinstance(source, dict):
        compact["source"] = model_preview(source, depth=1, text_limit=120)
    return compact


def model_observation(tool_name: str, result: dict, nodes: Sequence[dict]) -> dict:
    summary = summarize_observation(tool_name, result)
    compact_nodes = [compact_evidence_node_for_model(node) for node in list(nodes)[:HARNESS_AGENT_MODEL_NODE_LIMIT]]
    observation = {
        **summary,
        "result_preview": model_preview(result),
        "result_json_chars": json_char_count(result),
        "evidence_node_ids": [str(node.get("id") or "") for node in nodes if node.get("id")],
        "evidence_nodes": compact_nodes,
        "truncated_for_model": True,
    }
    if len(nodes) > HARNESS_AGENT_MODEL_NODE_LIMIT:
        observation["omitted_evidence_node_count"] = len(nodes) - HARNESS_AGENT_MODEL_NODE_LIMIT
    return observation


def summarize_omitted_observations(observations: Sequence[dict]) -> dict:
    evidence_ids = []
    evidence_nodes = []
    summaries = []
    for item in observations:
        summaries.append({
            "tool": item.get("tool", ""),
            "title": item.get("title", ""),
            "summary": preview(item.get("summary", ""), 180),
            "tool_result_contract": model_preview(item.get("tool_result_contract") or {}, depth=1, text_limit=160),
        })
        for evidence_id in item.get("evidence_node_ids", []) or []:
            if evidence_id and evidence_id not in evidence_ids:
                evidence_ids.append(str(evidence_id))
        for node in item.get("evidence_nodes", []) or []:
            node_id = str(node.get("id") or "")
            if node_id and all(existing.get("id") != node_id for existing in evidence_nodes):
                evidence_nodes.append(compact_evidence_node_for_model(node))
    return {
        "tool": "harness_context_summary",
        "ok": True,
        "id": "omitted-observations",
        "title": "已压缩的早期观察",
        "summary": f"前面 {len(observations)} 个观察已压缩，仅保留摘要和证据 id。",
        "omitted_observation_count": len(observations),
        "omitted_observation_summaries": summaries[:8],
        "evidence_node_ids": evidence_ids[:24],
        "evidence_nodes": evidence_nodes[:HARNESS_AGENT_MODEL_NODE_LIMIT],
        "truncated_for_model": True,
    }


def fit_observations_to_budget(observations: List[dict]) -> List[dict]:
    return fit_items_to_json_budget(
        observations,
        json_budget_chars=HARNESS_AGENT_MODEL_JSON_BUDGET,
        compact_item=lambda item: {
            "tool": item.get("tool", ""),
            "ok": item.get("ok", True),
            "id": item.get("id", ""),
            "title": item.get("title", ""),
            "summary": preview(item.get("summary", ""), 220),
            "evidence_node_ids": list(item.get("evidence_node_ids") or [])[:16],
            "tool_result_contract": model_preview(item.get("tool_result_contract") or {}, depth=1, text_limit=160),
            "evidence_layers": model_preview(item.get("evidence_layers") or {}, depth=1, text_limit=180),
            "evidence_nodes": [
                compact_evidence_node_for_model(node)
                for node in list(item.get("evidence_nodes") or [])[:4]
                if isinstance(node, dict)
            ],
            "result_preview_omitted": True,
            "truncated_for_model": True,
        },
        fallback_limit=max(1, HARNESS_AGENT_MODEL_OBSERVATION_LIMIT // 2),
    )


def context_budget_summary(source_observations: Sequence[dict], model_observations: Sequence[dict]) -> dict:
    return build_context_budget_summary(
        source_observations,
        model_observations,
        json_budget_chars=HARNESS_AGENT_MODEL_JSON_BUDGET,
        truncated_note="observations 已压缩；如需完整行/对象，请继续调用 detail 工具。",
        ok_note="observations 在当前预算内。",
        include_observation_bundle=True,
        bundle_id="harness-observation-bundle",
    )


def observations_for_model_context(observations: Sequence[dict]) -> List[dict]:
    items = list(observations or [])
    if len(items) > HARNESS_AGENT_MODEL_OBSERVATION_LIMIT:
        omitted = items[:-HARNESS_AGENT_MODEL_OBSERVATION_LIMIT]
        items = [summarize_omitted_observations(omitted)] + items[-HARNESS_AGENT_MODEL_OBSERVATION_LIMIT:]
    return fit_observations_to_budget(items)


def public_tool_result(result: dict, *, debug: bool) -> dict:
    if debug:
        return result
    public = dict(result)
    if "content" in public:
        content = str(public.pop("content") or "")
        public["content_preview"] = content[:500]
        public["content_hidden"] = len(content) > 500
    if "rows" in public and isinstance(public["rows"], list):
        public["rows"] = public["rows"][:5]
    if "cards" in public and isinstance(public["cards"], list):
        public["cards"] = public["cards"][:5]
    if "ready_cards" in public and isinstance(public["ready_cards"], list):
        public["ready_cards"] = public["ready_cards"][:5]
    if "needs_context_cards" in public and isinstance(public["needs_context_cards"], list):
        public["needs_context_cards"] = public["needs_context_cards"][:5]
    if "matches" in public and isinstance(public["matches"], list):
        public["matches"] = public["matches"][:5]
    if "values" in public and isinstance(public["values"], list):
        public["values"] = public["values"][:80]
        public["values_truncated"] = len(result.get("values") or []) > len(public["values"])
    if "top_values" in public and isinstance(public["top_values"], list):
        public["top_values"] = public["top_values"][:20]
    if "items" in public and isinstance(public["items"], list):
        compact_items = []
        for item in public["items"][:8]:
            if not isinstance(item, dict):
                compact_items.append(item)
                continue
            compact = dict(item)
            if "rows" in compact and isinstance(compact["rows"], list):
                compact["rows"] = compact["rows"][:3]
            if "matches" in compact and isinstance(compact["matches"], list):
                compact["matches"] = compact["matches"][:3]
            if "card" in compact and isinstance(compact["card"], dict):
                compact["card"] = model_preview(compact["card"], depth=1, text_limit=180)
            compact_items.append(compact)
        public["items"] = compact_items
        public["items_truncated"] = len(result.get("items") or []) > len(compact_items)
    if "matched_cards" in public and isinstance(public["matched_cards"], list):
        public["matched_cards"] = public["matched_cards"][:5]
    if "gap_cards" in public and isinstance(public["gap_cards"], list):
        public["gap_cards"] = public["gap_cards"][:5]
    return public


def step_payload(index: int,
                 kind: str,
                 *,
                 provider: str = "",
                 raw_model_output: str = "",
                 tool_name: str = "",
                 args: Optional[dict] = None,
                 ok: bool = True,
                 error: str = "",
                 summary: str = "",
                 debug: bool = False) -> dict:
    payload = {
        "index": index,
        "type": kind,
        "ok": ok,
        "provider": provider,
        "tool": tool_name,
        "summary": summary,
    }
    if error:
        payload["error"] = error
    if debug:
        payload["raw_model_output"] = raw_model_output[:4000]
        payload["args"] = args or {}
    return payload
