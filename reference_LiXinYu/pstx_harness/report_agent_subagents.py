# -*- coding: utf-8 -*-
"""Focused subagent orchestration for the report harness agent."""

from __future__ import annotations

import copy
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, List, Optional, Tuple

from pstx_agent_runtime import (
    build_subagent_question,
    compact_subagent_result,
    plan_subagents,
    summarize_subagent_results,
)
from pstx_harness.model import MockHarnessModelProvider
from pstx_harness.report_agent_config import (
    HARNESS_AGENT_MAX_STEPS,
    HARNESS_AGENT_MAX_SUBAGENTS,
    HARNESS_AGENT_MAX_TOOL_CALLS,
    HARNESS_AGENT_PROFILES,
    HarnessAgentRequest,
)
from pstx_harness.report_tools import HarnessToolRegistry


def subagent_provider_parallel_safe(provider) -> bool:
    if callable(getattr(provider, "clone_for_subagent", None)):
        return True
    if isinstance(provider, MockHarnessModelProvider):
        return True
    try:
        from pstx_harness.model import AsterHarnessModelProvider

        return isinstance(provider, AsterHarnessModelProvider)
    except Exception:
        return False


def fresh_subagent_provider(provider):
    clone_fn = getattr(provider, "clone_for_subagent", None)
    if callable(clone_fn):
        try:
            return clone_fn()
        except Exception:
            pass
    if isinstance(provider, MockHarnessModelProvider):
        return MockHarnessModelProvider()
    try:
        from pstx_harness.model import AsterHarnessModelProvider

        if isinstance(provider, AsterHarnessModelProvider):
            return AsterHarnessModelProvider(ask_model=getattr(provider, "_ask_model", None))
    except Exception:
        pass
    return provider


def subagent_question(parent_question: str, profile: str) -> str:
    definitions = plan_subagents(
        [profile],
        HARNESS_AGENT_PROFILES,
        max_subagents=1,
        disallowed_profiles=("full_review",),
    ).get("definitions", [])
    definition = definitions[0] if definitions else {"profile": profile, "title": profile, "prompt": f"请按 {profile} 聚焦审查。"}
    return build_subagent_question(parent_question, definition)


def run_subagents_parallel(report: dict,
                           bundle: dict,
                           parent_request: HarnessAgentRequest,
                           model_provider,
                           registry: HarnessToolRegistry,
                           *,
                           run_agent: Callable[..., dict],
                           project_context: Optional[dict] = None) -> Tuple[List[dict], dict]:
    plan = plan_subagents(
        parent_request.subagent_profiles,
        HARNESS_AGENT_PROFILES,
        max_subagents=parent_request.max_subagents,
        disallowed_profiles=("full_review",),
        parent_profile=parent_request.profile,
    )
    definitions = list(plan.get("definitions") or [])
    profiles = [str(item.get("profile") or "") for item in definitions]
    if not parent_request.enable_subagents or not definitions:
        return [], {
            "schema_version": plan.get("schema_version", "pstx-agent-subagents.v1"),
            "enabled": bool(parent_request.enable_subagents),
            "planned_count": 0,
            "completed_count": 0,
            "failed_count": 0,
            "profiles": [],
            "skipped": plan.get("skipped") or [],
            "warnings": plan.get("warnings") or [],
        }

    provider_parallel_safe = subagent_provider_parallel_safe(model_provider)

    def failed_payload(definition: dict, exc: Exception) -> dict:
        profile = str(definition.get("profile") or "")
        return {
            "schema_version": plan.get("schema_version", "pstx-agent-subagents.v1"),
            "profile": profile,
            "title": definition.get("title", profile),
            "ok": False,
            "status": "failed",
            "agent_run_id": "",
            "answer": f"Subagent 执行失败：{exc}",
            "trace_summary": {"stopped_reason": "subagent_error"},
            "citation_count": 0,
            "evidence_node_count": 0,
            "proposed_action_count": 0,
            "citations": [],
            "proposed_actions": [],
            "model_metadata": {"ok": False, "error": str(exc), "error_type": exc.__class__.__name__},
            "definition": definition,
            "isolation": definition.get("isolation", "fresh_context"),
            "provider_parallel_safe": provider_parallel_safe,
        }

    def worker(definition: dict) -> dict:
        profile = str(definition.get("profile") or "")
        started_at = time.time()
        child_request = HarnessAgentRequest(
            profile=profile,
            question=build_subagent_question(parent_request.question, definition),
            max_steps=min(parent_request.max_steps, int(definition.get("max_steps") or 8), HARNESS_AGENT_MAX_STEPS),
            max_tool_calls=min(parent_request.max_tool_calls, int(definition.get("max_tool_calls") or 14), HARNESS_AGENT_MAX_TOOL_CALLS),
            max_rows_per_table=parent_request.max_rows_per_table,
            debug=parent_request.debug,
            enable_subagents=False,
            subagent_profiles=(),
            max_subagents=0,
        )
        child_context = copy.deepcopy(project_context or {})
        child_context["subagent_definition"] = definition
        child_context["subagent_parent_profile"] = parent_request.profile
        child_context["subagent_parent_question"] = parent_request.question
        result = run_agent(
            report,
            bundle,
            child_request,
            model_provider=fresh_subagent_provider(model_provider),
            registry=registry,
            project_context=child_context,
        )
        row = compact_subagent_result(result, definition=definition)
        row["elapsed_ms"] = int((time.time() - started_at) * 1000)
        row["provider_parallel_safe"] = provider_parallel_safe
        return row

    started_at = time.time()
    results: List[dict] = []
    if provider_parallel_safe:
        max_workers = max(1, min(len(definitions), parent_request.max_subagents or 1, HARNESS_AGENT_MAX_SUBAGENTS))
    else:
        max_workers = 1
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(worker, definition): definition for definition in definitions}
        for future in as_completed(future_map):
            definition = future_map[future]
            try:
                results.append(future.result())
            except Exception as exc:
                results.append(failed_payload(definition, exc))
    order = {profile: index for index, profile in enumerate(profiles)}
    results.sort(key=lambda item: order.get(str(item.get("profile")), 999))
    summary = summarize_subagent_results(
        results,
        plan=plan,
        max_workers=max_workers,
        elapsed_ms=int((time.time() - started_at) * 1000),
        provider_parallel_safe=provider_parallel_safe,
    )
    return results, summary
