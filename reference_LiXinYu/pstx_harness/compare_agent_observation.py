# -*- coding: utf-8 -*-
"""Observation compaction and public-result shaping for the compare agent."""

from __future__ import annotations

from typing import List, Optional, Sequence

from pstx_agent_runtime import (
    build_context_budget_summary,
    json_char_count as runtime_json_char_count,
)
from pstx_harness.compare_agent_config import (
    COMPARE_AGENT_MODEL_JSON_BUDGET,
    COMPARE_AGENT_MODEL_NODE_LIMIT,
    COMPARE_AGENT_MODEL_OBSERVATION_LIMIT,
)


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


def compact_for_model(value, depth: int = 0):
    if depth >= 4:
        return preview(value, 220)
    if isinstance(value, dict):
        keys = []
        priority = [
            "id", "title", "summary", "section_id", "section_title", "row_index", "row_number",
            "total_rows", "diff_totals", "risk_items", "matches", "rows", "files", "path",
            "side", "truncated", "content", "tool_result_contract", "completeness",
            "recommended_next_tools", "evidence_layers",
        ]
        for key in priority:
            if key in value and key not in keys:
                keys.append(key)
        for key in value.keys():
            if key not in keys:
                keys.append(key)
            if len(keys) >= 18:
                break
        return {str(key): compact_for_model(value.get(key), depth + 1) for key in keys}
    if isinstance(value, list):
        return [compact_for_model(item, depth + 1) for item in value[:12]]
    return preview(value, 900)


def observations_for_model_context(observations: List[dict]) -> List[dict]:
    source_count = len(observations or [])
    omitted_count = max(0, source_count - COMPARE_AGENT_MODEL_OBSERVATION_LIMIT)
    compact = []
    for observation in observations[-COMPARE_AGENT_MODEL_OBSERVATION_LIMIT:]:
        nodes = list(observation.get("evidence_nodes") or [])[:COMPARE_AGENT_MODEL_NODE_LIMIT]
        item = {
            "tool": observation.get("tool"),
            "ok": observation.get("ok", True),
            "error": preview(observation.get("error") or "", 300),
            "summary": observation.get("summary"),
            "tool_result_contract": compact_for_model(observation.get("tool_result_contract") or {}),
            "evidence_layers": compact_for_model(observation.get("evidence_layers") or {}),
            "evidence_nodes": compact_for_model(nodes),
            "result_preview": compact_for_model(observation.get("result") or {}),
            "truncated_for_model": True,
        }
        compact.append(item)
    if omitted_count:
        compact.insert(0, {
            "tool": "compare_context_summary",
            "summary": f"前面 {omitted_count} 个观察已压缩，仅保留最近观察和证据摘要。",
            "omitted_observation_count": omitted_count,
            "truncated_for_model": True,
        })
    while compact and json_char_count(compact) > COMPARE_AGENT_MODEL_JSON_BUDGET:
        compact.pop(0)
    if not compact and observations:
        last_observation = observations[-1]
        compact.append({
            "tool": last_observation.get("tool"),
            "ok": last_observation.get("ok", True),
            "summary": preview(last_observation.get("summary") or "最后一个观察结果过大，已仅保留摘要和证据索引。", 500),
            "tool_result_contract": compact_for_model(last_observation.get("tool_result_contract") or {}),
            "evidence_layers": compact_for_model(last_observation.get("evidence_layers") or {}),
            "evidence_nodes": compact_for_model(list(last_observation.get("evidence_nodes") or [])[:8]),
            "result_preview": {
                "truncated_for_model": True,
                "detail_hint": "该观察超过 compare 模型上下文预算，请优先调用 detail/aggregation 工具读取原始证据。",
            },
            "truncated_for_model": True,
        })
    return compact


def context_budget_summary(source_observations: Sequence[dict],
                           model_observations: Sequence[dict]) -> dict:
    return build_context_budget_summary(
        source_observations,
        model_observations,
        json_budget_chars=COMPARE_AGENT_MODEL_JSON_BUDGET,
        truncated_note="compare observations 已压缩；如需完整差异行/文件片段，请继续调用 compare detail 工具。",
        ok_note="compare observations 在当前预算内。",
    )


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
    if "matches" in public and isinstance(public["matches"], list):
        public["matches"] = public["matches"][:5]
    if "items" in public and isinstance(public["items"], list):
        compact_items = []
        for item in public["items"][:8]:
            if not isinstance(item, dict):
                compact_items.append(item)
                continue
            compact = dict(item)
            if "matches" in compact and isinstance(compact["matches"], list):
                compact["matches"] = compact["matches"][:3]
            if "row" in compact and isinstance(compact["row"], dict):
                compact["row"] = compact_for_model(compact["row"], depth=1)
            if "object" in compact and isinstance(compact["object"], dict):
                compact["object"] = compact_for_model(compact["object"], depth=1)
            compact_items.append(compact)
        public["items"] = compact_items
        public["items_truncated"] = len(result.get("items") or []) > len(compact_items)
    if "page_results" in public and isinstance(public["page_results"], list):
        compact_pages = []
        for page in public["page_results"][:12]:
            if not isinstance(page, dict):
                continue
            compact_pages.append({
                "page": page.get("page"),
                "status": page.get("status"),
                "diff_count": page.get("diff_count"),
                "returned_diff_count": page.get("returned_diff_count"),
                "omitted_diff_count": page.get("omitted_diff_count"),
                "left_digest": page.get("left_digest"),
                "right_digest": page.get("right_digest"),
                "diffs": list(page.get("diffs") or [])[:3],
            })
        public["page_results"] = compact_pages
        public["page_results_truncated"] = len(result.get("page_results") or []) > len(compact_pages)
    return public


def summarize_observation(tool_name: str, result: dict) -> dict:
    return {
        "tool": tool_name,
        "ok": True,
        "id": result.get("id", tool_name),
        "title": result.get("title", tool_name),
        "summary": result.get("summary", ""),
        "keys": sorted(str(key) for key in result.keys())[:20],
    }


def step_payload(index: int,
                 step_type: str,
                 *,
                 provider: str = "",
                 raw_model_output: str = "",
                 tool_name: str = "",
                 args: Optional[dict] = None,
                 summary: str = "",
                 ok: bool = True,
                 error: str = "",
                 debug: bool = False) -> dict:
    payload = {
        "index": index,
        "type": step_type,
        "provider": provider,
        "tool": tool_name,
        "summary": summary,
        "ok": ok,
        "error": error,
    }
    if debug:
        payload["args"] = args or {}
        payload["raw_model_output"] = raw_model_output[:5000]
    return payload
