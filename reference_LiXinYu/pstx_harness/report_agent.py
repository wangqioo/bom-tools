# -*- coding: utf-8 -*-
"""Controlled multi-step agent loop for the local PSTX harness."""

from __future__ import annotations

from collections.abc import Mapping
import time
import uuid
from typing import Callable, List, Optional

from pstx_agent_runtime import (
    build_agent_session_state,
    build_continuation_pack,
    build_evidence_goal_contract,
    build_final_answer_quality_gate,
    build_agentic_envelope,
    build_execution_journal,
    build_journal_summary,
    build_harness_turn_context_snapshot,
    persist_agentic_task_memory,
    build_quality_repair_tool_calls,
    build_runtime_state,
    build_tool_error_observation,
    update_agentic_effort,
    dedupe_profile_ids,
    execute_runtime_tool_calls,
    filter_auto_quality_repair_tool_calls,
    recommended_tools_for_recovery,
    select_goal_prefetch_tool_calls,
    select_project_memory_prefetch_tool_calls,
    select_prefetch_followup_tool_calls,
    select_seeded_prefetch_tool_calls,
    is_recoverable_tool_error,
    merge_runtime_tool_execution,
    summarize_tool_dispatch_trace,
    write_workspace_scratch_files,
)
from pstx_harness.model import MockHarnessModelProvider
from pstx_harness.report_agent_config import (
    HARNESS_AGENT_MAX_TOOL_BATCH_CALLS,
    HarnessAgentRequest,
    list_harness_agent_profiles,
    profile_config as _profile_config,
)
from pstx_harness.report_agent_finalization import (
    normalize_citations as _normalize_citations,
    normalize_needs_user_input as _normalize_needs_user_input,
    normalize_proposed_actions as _normalize_proposed_actions,
    status_from_metadata as _status_from_metadata,
)
from pstx_harness.report_agent_evidence import (
    evidence_nodes_from_tool_result as _evidence_nodes_from_tool_result,
)
from pstx_harness.report_agent_protocol import (
    extract_balanced_json as _extract_balanced_json,
    parse_model_step as _parse_model_step_impl,
)
from pstx_harness.report_agent_planning import (
    allowed_tool_names as _allowed_tool_names,
    build_agent_prompt as _build_agent_prompt_impl,
    capability_plan as _capability_plan,
    memory_planning_context as _memory_planning_context,
    playbook_plan as _playbook_plan,
    selected_profile_ids as _selected_profile_ids,
)
from pstx_harness.report_agent_observation import (
    compact_project_context_for_model as _compact_project_context_for_model,
    context_budget_summary as _context_budget_summary,
    json_char_count as _json_char_count,
    model_observation as _model_observation,
    observations_for_model_context as _observations_for_model_context,
    public_tool_result as _public_tool_result,
    summarize_observation as _summarize_observation,
    step_payload as _step_payload,
)
from pstx_harness.report_agent_subagents import run_subagents_parallel as _run_subagents_parallel
from pstx_harness.report_tools import (
    HarnessToolContext,
    HarnessToolRegistry,
    build_default_harness_registry,
)


def _build_agent_prompt(request: HarnessAgentRequest,
                        report: dict,
                        registry: HarnessToolRegistry,
                        observations: List[dict],
                        *,
                        playbook_plan: Optional[dict] = None,
                        project_context: Optional[dict] = None,
                        planning_context: str = "",
                        context_budget: Optional[dict] = None,
                        runtime_state: Optional[dict] = None,
                        session_state: Optional[dict] = None,
                        agentic_context: Optional[dict] = None,
                        retry_note: str = "") -> str:
    return _build_agent_prompt_impl(
        request,
        report,
        registry,
        observations,
        playbook_plan_payload=playbook_plan,
        project_context=project_context,
        planning_context=planning_context,
        context_budget=context_budget,
        runtime_state=runtime_state,
        session_state=session_state,
        agentic_context=agentic_context,
        retry_note=retry_note,
        compact_project_context=_compact_project_context_for_model,
    )


def _parse_model_step(answer: str) -> Optional[dict]:
    return _parse_model_step_impl(answer, max_batch_calls=HARNESS_AGENT_MAX_TOOL_BATCH_CALLS)


def _agent_memory_run_id(report: dict, bundle: dict, project_context: dict) -> str:
    for value in (
        project_context.get("run_id"),
        project_context.get("memory_run_id"),
        report.get("run_id"),
        bundle.get("run_id"),
        (bundle.get("_cli_bundle_cache") or {}).get("path") if isinstance(bundle.get("_cli_bundle_cache"), dict) else "",
    ):
        text = str(value or "").strip()
        if text:
            return text
    return ""


def run_harness_agent(report: dict,
                      bundle: dict,
                      request: HarnessAgentRequest,
                      model_provider=None,
                      registry: Optional[HarnessToolRegistry] = None,
                      project_context: Optional[dict] = None,
                      checkpoint_callback: Optional[Callable[[dict], None]] = None,
                      should_cancel: Optional[Callable[[], bool]] = None,
                      dispatch_callback: Optional[Callable[[dict], object]] = None,
                      resume_context: Optional[dict] = None) -> dict:
    request.validate()
    started_at = time.time()
    agent_run_id = uuid.uuid4().hex[:12]
    registry = registry or build_default_harness_registry()
    provider = model_provider or MockHarnessModelProvider()
    project_context = dict(project_context or {})
    context = HarnessToolContext(report=report, bundle=bundle, request=request, project_context=project_context)
    planning_context = _memory_planning_context(project_context)
    selected_profile_ids = _selected_profile_ids(request, planning_context=planning_context)
    capability_plan_items = _capability_plan(request, planning_context=planning_context)
    allowed_tools = set(_allowed_tool_names(request, registry, planning_context=planning_context))
    playbook_plan = _playbook_plan(request, registry, planning_context=planning_context)
    memory_run_id = _agent_memory_run_id(report, bundle, project_context)
    agentic_context = build_agentic_envelope(
        run_id=memory_run_id or agent_run_id,
        question=f"{request.question or ''} {planning_context or ''}",
        capability_profiles=selected_profile_ids,
        playbook_plan=playbook_plan,
        root=None,
        include_skill_body=True,
    )
    observations_for_model: List[dict] = []
    public_observations: List[dict] = []
    raw_observations: List[dict] = []
    tool_calls: List[dict] = []
    tool_signatures: List[str] = []
    tool_dispatch_trace: List[dict] = []
    tool_result_contracts: List[dict] = []
    agent_steps: List[dict] = []
    evidence_nodes: List[dict] = []
    citations: List[dict] = []
    proposed_actions: List[dict] = []
    scratch_files: dict = {}
    needs_user_input: dict = {}
    final_answer_quality_gate: dict = {}
    task_dispatch_summary: dict = {}
    dispatched_tasks: List[dict] = []
    context_budget: dict = _context_budget_summary([], [])
    runtime_state: dict = build_runtime_state(
        goal=request.question or _profile_config(request.profile).get("default_question", ""),
        capability_plan=capability_plan_items,
        playbook_plan=playbook_plan,
        observations=[],
        tool_result_contracts=tool_result_contracts,
        project_context=project_context,
    )
    session_state: dict = build_agent_session_state(
        agent_run_id=agent_run_id,
        goal=request.question or _profile_config(request.profile).get("default_question", ""),
        runtime_state=runtime_state,
        project_context=project_context,
    )
    metadata = {
        "provider": provider.__class__.__name__,
        "ok": True,
        "stopped_reason": "",
        "profile": request.profile,
        "capability_profiles": selected_profile_ids,
        "planning_context_chars": len(planning_context),
        "agentic_kernel_version": agentic_context.get("version", ""),
        "selected_skill_count": (agentic_context.get("selected_skills") or {}).get("selected_count", 0),
        "guidance_source_count": (agentic_context.get("guidance_summary") or {}).get("source_count", 0),
        "task_memory_found": bool((agentic_context.get("task_memory_summary") or {}).get("found")),
    }
    answer = ""

    resume_context = dict(resume_context or {})

    def _checkpoint(phase: str, **extra) -> None:
        if not checkpoint_callback:
            return
        try:
            checkpoint_callback({
                "phase": phase,
                "step_index": int(extra.pop("step_index", len(agent_steps)) or 0),
                "max_steps": request.max_steps,
                "max_tool_calls": request.max_tool_calls,
                "tool_calls": tool_calls,
                "agent_steps": agent_steps,
                "partial_observations": public_observations[-30:],
                "evidence_ids": [str((item or {}).get("id") or "") for item in evidence_nodes if isinstance(item, dict) and (item or {}).get("id")],
                "selected_skills": agentic_context.get("selected_skills", {}),
                "playbook_plan": playbook_plan,
                "task_ledger": runtime_state.get("task_ledger", {}) if isinstance(runtime_state, dict) else {},
                "continuation_pack": resume_context.get("continuation_pack") or {},
                "retry_reasons": list(metadata.get("retry_reasons") or metadata.get("perseverance_retry_notes") or []),
                **extra,
            })
        except Exception:
            # Checkpointing must never change a synchronous agent answer.
            return

    def _is_cancelled() -> bool:
        try:
            return bool(should_cancel and should_cancel())
        except Exception:
            return False

    _checkpoint("planning", summary="已完成 profile、skill、playbook 与工具白名单规划。")

    memory_prefetch_plan = select_project_memory_prefetch_tool_calls(
        request.question,
        project_context,
        allowed_tools=allowed_tools,
        max_calls=1,
        remaining_tool_calls=request.max_tool_calls - len(tool_calls) - 1,
        enabled=request.profile == "auto",
    )
    metadata["memory_prefetch_plan"] = memory_prefetch_plan
    if memory_prefetch_plan.get("tool_calls"):
        _checkpoint("prefetch", summary="正在预取项目 evidence memory。")
        memory_execution = execute_runtime_tool_calls(
            tool_call_items=list(memory_prefetch_plan["tool_calls"]),
            is_batch_call=False,
            registry=registry,
            context=context,
            allowed_tools=allowed_tools,
            existing_tool_call_count=len(tool_calls),
            max_tool_calls=request.max_tool_calls,
            previous_tool_calls=tool_calls,
            previous_tool_signatures=tool_signatures,
            debug=request.debug,
            profile_label=request.profile,
            capability_profiles=selected_profile_ids,
            rejection_prefix="证据记忆预取工具调用被本地 harness 拒绝",
            empty_message="本地 evidence memory 没有生成可执行预取工具。",
            make_evidence_nodes=lambda name, result, index, args: _evidence_nodes_from_tool_result(
                name,
                result,
                call_index=index,
                args=args,
            ),
            summarize_observation=_summarize_observation,
            make_model_observation=lambda name, result, nodes, _observation: _model_observation(name, result, nodes),
            make_public_result=lambda result, debug: _public_tool_result(result, debug=debug),
        )
        merge_runtime_tool_execution(
            execution=memory_execution,
            tool_calls=tool_calls,
            tool_signatures=tool_signatures,
            tool_dispatch_trace=tool_dispatch_trace,
            tool_result_contracts=tool_result_contracts,
            observations_for_model=observations_for_model,
            public_observations=public_observations,
            raw_observations=raw_observations,
            evidence_nodes=evidence_nodes,
            metadata=metadata,
            metadata_prefix="memory_prefetch",
        )
        agent_steps.append(_step_payload(
            0,
            "runtime_memory_prefetch",
            tool_name=memory_execution["tool_name"],
            args=memory_execution["args"],
            ok=bool(memory_execution["ok"]),
            error=memory_execution["error"],
            summary=("本地 evidence memory 预取证据：" + memory_execution["summary"])[:500],
            debug=request.debug,
        ))
        _checkpoint("prefetch", summary="项目 evidence memory 预取完成。")

    connection_datasheet_prefetch = (
        "connection_datasheet_review" in selected_profile_ids
        or any(
            isinstance(item, Mapping) and item.get("id") == "schematic_datasheet_connection_review"
            for item in playbook_plan.get("selected_playbooks", [])
        )
    )
    seeded_prefetch_enabled = request.profile == "auto" or request.profile == "connection_datasheet_review"
    seeded_prefetch_max_calls = 4 if connection_datasheet_prefetch and request.max_tool_calls >= 8 else (2 if request.max_tool_calls >= 6 else 1)
    prefetch_plan = select_seeded_prefetch_tool_calls(
        playbook_plan,
        allowed_tools=allowed_tools,
        max_calls=seeded_prefetch_max_calls,
        remaining_tool_calls=request.max_tool_calls - len(tool_calls) - 1,
        enabled=seeded_prefetch_enabled,
    )
    metadata["prefetch_plan"] = prefetch_plan
    if prefetch_plan.get("tool_calls"):
        _checkpoint("prefetch", summary="正在执行 playbook 预取证据。")
        execution = execute_runtime_tool_calls(
            tool_call_items=list(prefetch_plan["tool_calls"]),
            is_batch_call=len(prefetch_plan["tool_calls"]) > 1,
            registry=registry,
            context=context,
            allowed_tools=allowed_tools,
            existing_tool_call_count=len(tool_calls),
            max_tool_calls=request.max_tool_calls,
            previous_tool_calls=tool_calls,
            previous_tool_signatures=tool_signatures,
            debug=request.debug,
            profile_label=request.profile,
            capability_profiles=selected_profile_ids,
            rejection_prefix="预取工具调用被本地 harness 拒绝",
            empty_message="本地 playbook 没有生成可执行预取工具。",
            make_evidence_nodes=lambda name, result, index, args: _evidence_nodes_from_tool_result(
                name,
                result,
                call_index=index,
                args=args,
            ),
            summarize_observation=_summarize_observation,
            make_model_observation=lambda name, result, nodes, _observation: _model_observation(name, result, nodes),
            make_public_result=lambda result, debug: _public_tool_result(result, debug=debug),
        )
        merge_runtime_tool_execution(
            execution=execution,
            tool_calls=tool_calls,
            tool_signatures=tool_signatures,
            tool_dispatch_trace=tool_dispatch_trace,
            tool_result_contracts=tool_result_contracts,
            observations_for_model=observations_for_model,
            public_observations=public_observations,
            raw_observations=raw_observations,
            evidence_nodes=evidence_nodes,
            metadata=metadata,
            metadata_prefix="prefetch",
        )
        agent_steps.append(_step_payload(
            0,
            "runtime_prefetch",
            tool_name=execution["tool_name"],
            args=execution["args"],
            ok=bool(execution["ok"]),
            error=execution["error"],
            summary=("本地 playbook 预取证据：" + execution["summary"])[:500],
            debug=request.debug,
        ))
        _checkpoint("prefetch", summary="playbook 预取证据完成。")
        followup_plan = select_prefetch_followup_tool_calls(
            execution.get("raw_observations") or [],
            allowed_tools=allowed_tools,
            max_calls=2 if connection_datasheet_prefetch and request.max_tool_calls >= 8 else 1,
            remaining_tool_calls=request.max_tool_calls - len(tool_calls) - 1,
            previous_tool_signatures=tool_signatures,
            enabled=bool(execution["ok"]),
        )
        metadata["prefetch_followup_plan"] = followup_plan
        if followup_plan.get("tool_calls"):
            _checkpoint("prefetch", summary="正在读取预取证据详情。")
            followup_execution = execute_runtime_tool_calls(
                tool_call_items=list(followup_plan["tool_calls"]),
                is_batch_call=False,
                registry=registry,
                context=context,
                allowed_tools=allowed_tools,
                existing_tool_call_count=len(tool_calls),
                max_tool_calls=request.max_tool_calls,
                previous_tool_calls=tool_calls,
                previous_tool_signatures=tool_signatures,
                debug=request.debug,
                profile_label=request.profile,
                capability_profiles=selected_profile_ids,
                rejection_prefix="预取详情工具调用被本地 harness 拒绝",
                empty_message="本地 playbook 没有生成可执行预取详情工具。",
                make_evidence_nodes=lambda name, result, index, args: _evidence_nodes_from_tool_result(
                    name,
                    result,
                    call_index=index,
                    args=args,
                ),
                summarize_observation=_summarize_observation,
                make_model_observation=lambda name, result, nodes, _observation: _model_observation(name, result, nodes),
                make_public_result=lambda result, debug: _public_tool_result(result, debug=debug),
            )
            merge_runtime_tool_execution(
                execution=followup_execution,
                tool_calls=tool_calls,
                tool_signatures=tool_signatures,
                tool_dispatch_trace=tool_dispatch_trace,
                tool_result_contracts=tool_result_contracts,
                observations_for_model=observations_for_model,
                public_observations=public_observations,
                raw_observations=raw_observations,
                evidence_nodes=evidence_nodes,
                metadata=metadata,
                metadata_prefix="prefetch_followup",
                record_observation_count=False,
            )
            agent_steps.append(_step_payload(
                0,
                "runtime_prefetch_followup",
                tool_name=followup_execution["tool_name"],
                args=followup_execution["args"],
                ok=bool(followup_execution["ok"]),
                error=followup_execution["error"],
                summary=("本地 playbook 预取详情：" + followup_execution["summary"])[:500],
                debug=request.debug,
            ))
            _checkpoint("prefetch", summary="预取证据详情读取完成。")

    goal_prefetch_plan = select_goal_prefetch_tool_calls(
        playbook_plan,
        allowed_tools=allowed_tools,
        max_calls=1,
        remaining_tool_calls=request.max_tool_calls - len(tool_calls) - 1,
        previous_tool_signatures=tool_signatures,
        enabled=seeded_prefetch_enabled and not observations_for_model,
    )
    metadata["goal_prefetch_plan"] = goal_prefetch_plan
    if goal_prefetch_plan.get("tool_calls"):
        _checkpoint("prefetch", summary="正在按 evidence goal 预取证据。")
        goal_execution = execute_runtime_tool_calls(
            tool_call_items=list(goal_prefetch_plan["tool_calls"]),
            is_batch_call=False,
            registry=registry,
            context=context,
            allowed_tools=allowed_tools,
            existing_tool_call_count=len(tool_calls),
            max_tool_calls=request.max_tool_calls,
            previous_tool_calls=tool_calls,
            previous_tool_signatures=tool_signatures,
            debug=request.debug,
            profile_label=request.profile,
            capability_profiles=selected_profile_ids,
            rejection_prefix="证据目标预取工具调用被本地 harness 拒绝",
            empty_message="本地 evidence goal 没有生成可执行预取工具。",
            make_evidence_nodes=lambda name, result, index, args: _evidence_nodes_from_tool_result(
                name,
                result,
                call_index=index,
                args=args,
            ),
            summarize_observation=_summarize_observation,
            make_model_observation=lambda name, result, nodes, _observation: _model_observation(name, result, nodes),
            make_public_result=lambda result, debug: _public_tool_result(result, debug=debug),
        )
        merge_runtime_tool_execution(
            execution=goal_execution,
            tool_calls=tool_calls,
            tool_signatures=tool_signatures,
            tool_dispatch_trace=tool_dispatch_trace,
            tool_result_contracts=tool_result_contracts,
            observations_for_model=observations_for_model,
            public_observations=public_observations,
            raw_observations=raw_observations,
            evidence_nodes=evidence_nodes,
            metadata=metadata,
            metadata_prefix="goal_prefetch",
        )
        agent_steps.append(_step_payload(
            0,
            "runtime_goal_prefetch",
            tool_name=goal_execution["tool_name"],
            args=goal_execution["args"],
            ok=bool(goal_execution["ok"]),
            error=goal_execution["error"],
            summary=("本地 evidence goal 预取证据：" + goal_execution["summary"])[:500],
            debug=request.debug,
        ))
        _checkpoint("prefetch", summary="evidence goal 预取完成。")

    for step_index in range(1, request.max_steps + 1):
        if _is_cancelled():
            metadata.update({"ok": False, "stopped_reason": "cancelled"})
            answer = "Agent 任务已取消。"
            agent_steps.append(_step_payload(step_index, "cancelled", ok=False, error=answer, debug=request.debug))
            _checkpoint("cancelled", step_index=step_index, summary=answer)
            break
        model_observations = _observations_for_model_context(observations_for_model)
        context_budget = _context_budget_summary(observations_for_model, model_observations)
        runtime_state = build_runtime_state(
            goal=request.question or _profile_config(request.profile).get("default_question", ""),
            capability_plan=capability_plan_items,
            playbook_plan=playbook_plan,
            observations=model_observations,
            tool_result_contracts=tool_result_contracts,
            project_context=project_context,
            truncated=bool(context_budget.get("truncated")),
        )
        session_state = build_agent_session_state(
            agent_run_id=agent_run_id,
            goal=request.question or _profile_config(request.profile).get("default_question", ""),
            runtime_state=runtime_state,
            project_context=project_context,
            observations=model_observations,
        )
        agentic_context = update_agentic_effort(
            agentic_context,
            step_type="planning",
            answer="",
            tool_call_count=len(tool_calls),
            max_tool_calls=request.max_tool_calls,
            playbook_plan=playbook_plan,
            tool_result_contracts=tool_result_contracts,
            task_ledger=runtime_state.get("task_ledger", {}),
            evidence_node_count=len(evidence_nodes),
            citation_count=len(citations),
            allow_needs_user_input=True,
            retry_count=int(metadata.get("perseverance_retry_count") or 0),
        )
        prompt = _build_agent_prompt(
            request,
            report,
            registry,
            model_observations,
            playbook_plan=playbook_plan,
            project_context=project_context,
            planning_context=planning_context,
            context_budget=context_budget,
            runtime_state=runtime_state,
            session_state=session_state,
            agentic_context=agentic_context,
        )
        try:
            _checkpoint("model_call", step_index=step_index, summary=f"第 {step_index} 轮模型推理中。")
            response = provider.generate_agent_step(prompt, inputs={
                "question": request.question,
                "observations": model_observations,
                "context_budget": context_budget,
                "observation_bundle": context_budget.get("observation_bundle", {}),
                "runtime_state": runtime_state,
                "session_state": session_state,
                "agentic_context": agentic_context,
                "guidance_summary": agentic_context.get("guidance_summary", {}),
                "selected_skills": agentic_context.get("selected_skills", {}),
                "task_memory_summary": agentic_context.get("task_memory_summary", {}),
                "effort_policy": agentic_context.get("effort_policy", {}),
                "task_ledger": runtime_state.get("task_ledger", {}),
                "playbook_plan": playbook_plan,
                "tool_count": len(tool_calls),
                "step_index": step_index,
                "agent_run_id": agent_run_id,
                "agent_profile": request.profile,
                "capability_profiles": selected_profile_ids,
                "planning_context": planning_context,
                "project_context": _compact_project_context_for_model(project_context),
                "context_answers": list(request.context_answers),
                "continue_agent_run_id": request.continue_agent_run_id,
            })
            _checkpoint("model_call", step_index=step_index, summary=f"第 {step_index} 轮模型返回，正在解析协议。")
        except Exception as exc:
            metadata.update({"ok": False, "stopped_reason": "model_error", "error": str(exc)})
            answer = f"模型 provider 调用失败：{exc}"
            agent_steps.append(_step_payload(
                step_index,
                "model_error",
                ok=False,
                error=str(exc),
                debug=request.debug,
            ))
            _checkpoint("failed", step_index=step_index, summary=answer, error=str(exc))
            break
        metadata.update({
            "provider": response.provider,
            "mode": response.mode,
            **dict(response.metadata or {}),
        })
        metadata["last_prompt_chars"] = len(prompt)
        metadata["last_model_observation_json_chars"] = _json_char_count(model_observations)
        metadata["last_context_budget"] = context_budget
        metadata["last_runtime_state"] = runtime_state
        metadata["last_session_state"] = session_state
        metadata["input_truncated"] = bool(metadata.get("input_truncated")) or bool(context_budget.get("truncated"))
        metadata["max_model_observation_json_chars"] = max(
            int(metadata.get("max_model_observation_json_chars") or 0),
            int(context_budget.get("model_observation_json_chars") or 0),
        )
        parsed = _parse_model_step(response.answer)
        raw_answer = response.answer
        if parsed is None:
            _checkpoint("quality_repair_start", step_index=step_index, summary="模型输出不是合法 JSON，正在请求协议修复。")
            retry_prompt = _build_agent_prompt(
                request,
                report,
                registry,
                model_observations,
                playbook_plan=playbook_plan,
                project_context=project_context,
                planning_context=planning_context,
                context_budget=context_budget,
                runtime_state=runtime_state,
                session_state=session_state,
                agentic_context=agentic_context,
                retry_note="必须输出包含 tool_call、tool_batch_call、dispatch_tasks、needs_user_input 或 final_answer 的 JSON 对象。",
            )
            response = provider.generate_agent_step(retry_prompt, inputs={
                "question": request.question,
                "observations": model_observations,
                "context_budget": context_budget,
                "observation_bundle": context_budget.get("observation_bundle", {}),
                "runtime_state": runtime_state,
                "session_state": session_state,
                "agentic_context": agentic_context,
                "guidance_summary": agentic_context.get("guidance_summary", {}),
                "selected_skills": agentic_context.get("selected_skills", {}),
                "task_memory_summary": agentic_context.get("task_memory_summary", {}),
                "effort_policy": agentic_context.get("effort_policy", {}),
                "task_ledger": runtime_state.get("task_ledger", {}),
                "playbook_plan": playbook_plan,
                "tool_count": len(tool_calls),
                "step_index": step_index,
                "retry": True,
                "agent_run_id": agent_run_id,
                "agent_profile": request.profile,
                "capability_profiles": selected_profile_ids,
                "planning_context": planning_context,
                "project_context": _compact_project_context_for_model(project_context),
                "context_answers": list(request.context_answers),
                "continue_agent_run_id": request.continue_agent_run_id,
            })
            raw_answer = response.answer
            parsed = _parse_model_step(response.answer)

        if parsed is None:
            metadata.update({"ok": False, "stopped_reason": "invalid_model_json"})
            answer = "模型未返回合法 JSON，已停止 agent loop；请基于现有报告继续人工确认。"
            agent_steps.append(_step_payload(
                step_index,
                "model_error",
                provider=metadata.get("provider", ""),
                raw_model_output=raw_answer,
                ok=False,
                error="模型未返回合法 JSON。",
                debug=request.debug,
            ))
            _checkpoint("failed", step_index=step_index, summary=answer, error="invalid_model_json")
            break

        if parsed["type"] == "protocol_error":
            metadata.update({"ok": False, "stopped_reason": "protocol_error"})
            answer = f"模型输出不符合 harness 通讯协议：{parsed.get('error')}"
            agent_steps.append(_step_payload(
                step_index,
                "model_error",
                provider=response.provider,
                raw_model_output=raw_answer,
                ok=False,
                error=answer,
                debug=request.debug,
            ))
            _checkpoint("failed", step_index=step_index, summary=answer, error="protocol_error")
            break

        retry_note = ""
        if parsed["type"] in {"needs_user_input", "final_answer"}:
            agentic_context = update_agentic_effort(
                agentic_context,
                step_type=str(parsed["type"]),
                answer=parsed.get("final_answer") if parsed["type"] == "final_answer" else "",
                tool_call_count=len(tool_calls),
                max_tool_calls=request.max_tool_calls,
                playbook_plan=playbook_plan,
                tool_result_contracts=tool_result_contracts,
                task_ledger=runtime_state.get("task_ledger", {}),
                evidence_node_count=len(evidence_nodes),
                citation_count=len(((parsed.get("raw") if isinstance(parsed.get("raw"), dict) else {}) or {}).get("citations") or []) if parsed["type"] == "final_answer" else 0,
                allow_needs_user_input=True,
                retry_count=int(metadata.get("perseverance_retry_count") or 0),
            )
            retry_note = str((agentic_context.get("effort_policy") or {}).get("retry_note") or "")
        if retry_note and int(metadata.get("perseverance_retry_count") or 0) < 2:
            metadata["perseverance_retry_count"] = int(metadata.get("perseverance_retry_count") or 0) + 1
            notes = list(metadata.get("perseverance_retry_notes") or [])
            notes.append(retry_note)
            metadata["perseverance_retry_notes"] = notes[:4]
            metadata["retry_reasons"] = notes[:4]
            retry_prompt = _build_agent_prompt(
                request,
                report,
                registry,
                model_observations,
                playbook_plan=playbook_plan,
                project_context=project_context,
                planning_context=planning_context,
                context_budget=context_budget,
                runtime_state=runtime_state,
                session_state=session_state,
                agentic_context=agentic_context,
                retry_note=retry_note,
            )
            response = provider.generate_agent_step(retry_prompt, inputs={
                "question": request.question,
                "observations": model_observations,
                "context_budget": context_budget,
                "observation_bundle": context_budget.get("observation_bundle", {}),
                "runtime_state": runtime_state,
                "session_state": session_state,
                "agentic_context": agentic_context,
                "guidance_summary": agentic_context.get("guidance_summary", {}),
                "selected_skills": agentic_context.get("selected_skills", {}),
                "task_memory_summary": agentic_context.get("task_memory_summary", {}),
                "effort_policy": agentic_context.get("effort_policy", {}),
                "task_ledger": runtime_state.get("task_ledger", {}),
                "playbook_plan": playbook_plan,
                "tool_count": len(tool_calls),
                "step_index": step_index,
                "retry": True,
                "perseverance_retry": True,
                "perseverance_retry_note": retry_note,
                "agent_run_id": agent_run_id,
                "agent_profile": request.profile,
                "capability_profiles": selected_profile_ids,
                "planning_context": planning_context,
                "project_context": _compact_project_context_for_model(project_context),
                "context_answers": list(request.context_answers),
                "continue_agent_run_id": request.continue_agent_run_id,
            })
            _checkpoint("quality_repair_continue", step_index=step_index, summary="低努力回答被拒绝，已完成继续取证重试。")
            metadata.update({
                "provider": response.provider,
                "mode": response.mode,
                **dict(response.metadata or {}),
            })
            raw_answer = response.answer
            parsed = _parse_model_step(response.answer)
            if parsed is None:
                metadata.update({"ok": False, "stopped_reason": "invalid_model_json"})
                answer = "模型在继续取证重试后仍未返回合法 JSON，已停止 agent loop；请基于现有报告继续人工确认。"
                agent_steps.append(_step_payload(
                    step_index,
                    "model_error",
                    provider=metadata.get("provider", ""),
                    raw_model_output=raw_answer,
                    ok=False,
                    error="模型在继续取证重试后未返回合法 JSON。",
                    summary=retry_note[:500],
                    debug=request.debug,
                ))
                _checkpoint("failed", step_index=step_index, summary=answer, error="invalid_model_json")
                break
            if parsed["type"] == "protocol_error":
                metadata.update({"ok": False, "stopped_reason": "protocol_error"})
                answer = f"模型输出不符合 harness 通讯协议：{parsed.get('error')}"
                agent_steps.append(_step_payload(
                    step_index,
                    "model_error",
                    provider=response.provider,
                    raw_model_output=raw_answer,
                    ok=False,
                    error=answer,
                    summary=retry_note[:500],
                    debug=request.debug,
                ))
                _checkpoint("failed", step_index=step_index, summary=answer, error="protocol_error")
                break

        if parsed["type"] == "dispatch_tasks":
            raw_dispatch = parsed.get("task_dispatch") if isinstance(parsed.get("task_dispatch"), dict) else {}
            requested_tasks = list(parsed.get("dispatch_tasks") or raw_dispatch.get("tasks") or [])
            dispatch_request = {
                "schema_version": raw_dispatch.get("schema_version") or "pstx-agent-task-dispatch.v1",
                "source": "report",
                "parent_agent_run_id": agent_run_id,
                "profile": request.profile,
                "reason": raw_dispatch.get("reason") or "",
                "task_count": len(requested_tasks),
                "tasks": requested_tasks,
                "context": {
                    "run_id": project_context.get("run_id") or report.get("run_id") or bundle.get("run_id") or "",
                    "project_name": report.get("project_name") or bundle.get("project_name") or "",
                },
            }
            metadata["stopped_reason"] = "task_dispatched"
            metadata["task_dispatch_requested_count"] = len(requested_tasks)
            if dispatch_callback:
                _checkpoint("task_dispatch", step_index=step_index, summary=f"正在分发 {len(requested_tasks)} 个后台子任务。")
                try:
                    dispatch_result = dispatch_callback(dispatch_request)
                except Exception as exc:
                    metadata.update({"ok": False, "stopped_reason": "task_dispatch_error", "error": str(exc)})
                    answer = f"长任务分发失败：{exc}"
                    agent_steps.append(_step_payload(
                        step_index,
                        "task_dispatch",
                        provider=response.provider,
                        raw_model_output=raw_answer,
                        args={"task_count": len(requested_tasks)},
                        ok=False,
                        error=answer,
                        debug=request.debug,
                    ))
                    _checkpoint("failed", step_index=step_index, summary=answer, error="task_dispatch_error")
                    break
                if isinstance(dispatch_result, Mapping):
                    task_dispatch_summary = dict(dispatch_result.get("task_dispatch_summary") or dispatch_result.get("summary") or {})
                    dispatched_tasks = [dict(item) for item in (dispatch_result.get("dispatched_tasks") or dispatch_result.get("children") or []) if isinstance(item, Mapping)]
                elif isinstance(dispatch_result, list):
                    dispatched_tasks = [dict(item) for item in dispatch_result if isinstance(item, Mapping)]
                    task_dispatch_summary = {}
                else:
                    dispatched_tasks = []
                    task_dispatch_summary = {}
                task_dispatch_summary = {
                    "schema_version": raw_dispatch.get("schema_version") or "pstx-agent-task-dispatch.v1",
                    "available": True,
                    "task_count": len(requested_tasks),
                    "dispatched_count": len(dispatched_tasks),
                    **task_dispatch_summary,
                }
                answer = f"已分发 {len(dispatched_tasks)} 个后台子任务；父任务先记录分发摘要，子任务会继续在 durable run 中取证。"
                _checkpoint(
                    "task_dispatch",
                    step_index=step_index,
                    summary=answer,
                    task_dispatch_summary=task_dispatch_summary,
                    dispatched_tasks=dispatched_tasks,
                )
            else:
                dispatched_tasks = [
                    {**dict(item), "status": "not_dispatched", "agent_run_id": "", "status_url": ""}
                    for item in requested_tasks
                    if isinstance(item, Mapping)
                ]
                task_dispatch_summary = {
                    "schema_version": raw_dispatch.get("schema_version") or "pstx-agent-task-dispatch.v1",
                    "available": False,
                    "task_count": len(requested_tasks),
                    "dispatched_count": 0,
                    "reason": "当前执行器未配置后台 dispatch callback。",
                }
                answer = "已生成长任务分发计划；当前执行器未启用后台分发，因此没有创建 child runs。"
                _checkpoint("task_dispatch", step_index=step_index, summary=answer, task_dispatch_summary=task_dispatch_summary)
            metadata["task_dispatch_available"] = bool(task_dispatch_summary.get("available"))
            metadata["task_dispatch_dispatched_count"] = int(task_dispatch_summary.get("dispatched_count") or 0)
            agent_steps.append(_step_payload(
                step_index,
                "task_dispatch",
                provider=response.provider,
                raw_model_output=raw_answer,
                args={"task_count": len(requested_tasks), "available": bool(task_dispatch_summary.get("available"))},
                summary=answer[:500],
                debug=request.debug,
            ))
            break

        if parsed["type"] == "needs_user_input":
            raw = parsed.get("raw") if isinstance(parsed.get("raw"), dict) else {}
            needs_user_input = _normalize_needs_user_input(raw, evidence_nodes)
            metadata["stopped_reason"] = "needs_user_input"
            answer = needs_user_input.get("reason") or "需要用户补充信息后继续。"
            agent_steps.append(_step_payload(
                step_index,
                "needs_user_input",
                provider=response.provider,
                raw_model_output=raw_answer,
                summary=answer[:500],
                debug=request.debug,
            ))
            _checkpoint("waiting_for_user", step_index=step_index, summary=answer)
            break

        if parsed["type"] == "final_answer":
            answer = str(parsed.get("final_answer") or "").strip()
            metadata["stopped_reason"] = "final_answer"
            raw = parsed.get("raw") if isinstance(parsed.get("raw"), dict) else {}
            if isinstance(raw.get("scratch_files"), list):
                scratch_scope_id = (
                    project_context.get("agent_workspace_scope_id")
                    or project_context.get("run_id")
                    or memory_run_id
                    or agent_run_id
                )
                scratch_run_id = project_context.get("agent_workspace_agent_run_id") or agent_run_id
                scratch_files = write_workspace_scratch_files(
                    scratch_scope_id,
                    scratch_run_id,
                    raw.get("scratch_files") or [],
                )
                metadata["scratch_file_count"] = int(scratch_files.get("file_count") or 0)
                metadata["scratch_file_warnings"] = list(scratch_files.get("warnings") or [])
            citations, citation_meta = _normalize_citations(raw, evidence_nodes)
            proposed_actions = _normalize_proposed_actions(raw)
            evidence_goal_contract = build_evidence_goal_contract(
                playbook_plan=playbook_plan,
                evidence_nodes=evidence_nodes,
            )
            final_answer_quality_gate = build_final_answer_quality_gate(
                answer=answer,
                citations=citations,
                proposed_actions=proposed_actions,
                evidence_nodes=evidence_nodes,
                tool_result_contracts=tool_result_contracts,
                task_ledger=runtime_state.get("task_ledger", {}),
                evidence_goal_contract=evidence_goal_contract,
            )
            repair_plan = build_quality_repair_tool_calls(
                final_answer_quality_gate,
                allowed_tools=allowed_tools,
                max_calls=min(2, request.max_tool_calls - len(tool_calls)),
            )
            metadata["last_quality_repair_plan"] = repair_plan
            auto_repair_plan = filter_auto_quality_repair_tool_calls(
                repair_plan,
                final_answer_quality_gate,
                provided_citation_count=len(raw.get("citations") or []),
            )
            metadata["last_auto_quality_repair_plan"] = auto_repair_plan
            if (
                auto_repair_plan.get("tool_calls")
                and int(metadata.get("quality_repair_attempt_count") or 0) < 1
                and len(tool_calls) < request.max_tool_calls
            ):
                _checkpoint("quality_repair_start", step_index=step_index, summary="最终回答缺证据，正在自动补证据。")
                metadata["quality_repair_attempt_count"] = int(metadata.get("quality_repair_attempt_count") or 0) + 1
                execution = execute_runtime_tool_calls(
                    tool_call_items=auto_repair_plan["tool_calls"],
                    is_batch_call=len(auto_repair_plan["tool_calls"]) > 1,
                    registry=registry,
                    context=context,
                    allowed_tools=allowed_tools,
                    existing_tool_call_count=len(tool_calls),
                    max_tool_calls=request.max_tool_calls,
                    previous_tool_calls=tool_calls,
                    previous_tool_signatures=tool_signatures,
                    debug=request.debug,
                    profile_label=request.profile,
                    capability_profiles=selected_profile_ids,
                    rejection_prefix="质量门禁修复工具调用被本地 harness 拒绝",
                    empty_message="质量门禁没有生成可执行修复工具调用。",
                    make_evidence_nodes=lambda name, result, index, args: _evidence_nodes_from_tool_result(
                        name,
                        result,
                        call_index=index,
                        args=args,
                    ),
                    summarize_observation=_summarize_observation,
                    make_model_observation=lambda name, result, nodes, _observation: _model_observation(name, result, nodes),
                    make_public_result=lambda result, debug: _public_tool_result(result, debug=debug),
                )
                repair_counts = merge_runtime_tool_execution(
                    execution=execution,
                    tool_calls=tool_calls,
                    tool_signatures=tool_signatures,
                    tool_dispatch_trace=tool_dispatch_trace,
                    tool_result_contracts=tool_result_contracts,
                    observations_for_model=observations_for_model,
                    public_observations=public_observations,
                    raw_observations=raw_observations,
                    evidence_nodes=evidence_nodes,
                )
                metadata["quality_repair_tool_count"] = int(metadata.get("quality_repair_tool_count") or 0) + repair_counts["tool_count"]
                agent_steps.append(_step_payload(
                    step_index,
                    "quality_repair_tool_call",
                    provider=response.provider,
                    raw_model_output=raw_answer,
                    tool_name=execution["tool_name"],
                    args=execution["args"],
                    ok=bool(execution["ok"]),
                    error=execution["error"],
                    summary=("质量门禁触发补证据：" + execution["summary"])[:500],
                    debug=request.debug,
                ))
                _checkpoint("quality_repair_tool_call", step_index=step_index, summary="质量门禁补证据工具调用完成。")
                if execution["ok"]:
                    answer = ""
                    citations = []
                    proposed_actions = []
                    final_answer_quality_gate = {}
                    metadata["stopped_reason"] = "quality_repair_continue"
                    _checkpoint("quality_repair_continue", step_index=step_index, summary="质量门禁补证据成功，继续模型审查。")
                    continue
                metadata["quality_repair_failed"] = True
                metadata["quality_repair_error"] = execution["error"]
            metadata.update(citation_meta)
            metadata["final_answer_quality_gate"] = final_answer_quality_gate
            agent_steps.append(_step_payload(
                step_index,
                "final_answer",
                provider=response.provider,
                raw_model_output=raw_answer,
                summary=answer[:500],
                debug=request.debug,
            ))
            _checkpoint("finalizing", step_index=step_index, summary="最终回答已生成，正在写入 trace 与 artifact。")
            break

        tool_call_items = list(parsed.get("tool_calls") or [])
        if not tool_call_items and isinstance(parsed.get("tool_call"), dict):
            tool_call_items = [parsed["tool_call"]]
        is_batch_call = parsed.get("type") == "tool_batch_call"
        if not tool_call_items:
            metadata.update({"ok": False, "stopped_reason": "invalid_model_json"})
            answer = "模型未返回可执行的工具调用，已停止 agent loop。"
            agent_steps.append(_step_payload(
                step_index,
                "model_error",
                provider=response.provider,
                raw_model_output=raw_answer,
                ok=False,
                error=answer,
                debug=request.debug,
            ))
            break
        if len(tool_calls) + len(tool_call_items) > request.max_tool_calls:
            metadata["stopped_reason"] = "max_tool_calls"
            answer = "已达到最大工具调用次数，需人工基于现有观察继续确认。"
            agent_steps.append(_step_payload(
                step_index,
                "limit",
                provider=response.provider,
                raw_model_output=raw_answer,
                tool_name="tool_batch_call" if is_batch_call else str(tool_call_items[0].get("name") or ""),
                args={"requested_tool_count": len(tool_call_items)} if is_batch_call else (tool_call_items[0].get("args") if isinstance(tool_call_items[0].get("args"), dict) else {}),
                ok=False,
                error=answer,
                debug=request.debug,
            ))
            _checkpoint("failed", step_index=step_index, summary=answer, error="max_tool_calls")
            break

        _checkpoint("batch_tool_call" if is_batch_call else "tool_call", step_index=step_index, summary="正在执行模型选择的取证工具。")
        execution = execute_runtime_tool_calls(
            tool_call_items=tool_call_items,
            is_batch_call=is_batch_call,
            registry=registry,
            context=context,
            allowed_tools=allowed_tools,
            existing_tool_call_count=len(tool_calls),
            max_tool_calls=request.max_tool_calls,
            previous_tool_calls=tool_calls,
            previous_tool_signatures=tool_signatures,
            debug=request.debug,
            profile_label=request.profile,
            capability_profiles=selected_profile_ids,
            rejection_prefix="工具调用被本地 harness 拒绝",
            empty_message="模型未返回可执行的工具调用，已停止 agent loop。",
            make_evidence_nodes=lambda name, result, index, args: _evidence_nodes_from_tool_result(
                name,
                result,
                call_index=index,
                args=args,
            ),
            summarize_observation=_summarize_observation,
            make_model_observation=lambda name, result, nodes, _observation: _model_observation(name, result, nodes),
            make_public_result=lambda result, debug: _public_tool_result(result, debug=debug),
        )
        merge_runtime_tool_execution(
            execution=execution,
            tool_calls=tool_calls,
            tool_signatures=tool_signatures,
            tool_dispatch_trace=tool_dispatch_trace,
            tool_result_contracts=tool_result_contracts,
            observations_for_model=observations_for_model,
            public_observations=public_observations,
            raw_observations=raw_observations,
            evidence_nodes=evidence_nodes,
        )
        _checkpoint("batch_tool_call" if is_batch_call else "tool_call", step_index=step_index, summary="取证工具调用完成。")
        if not execution["ok"]:
            can_recover_tool_error = (
                execution.get("stopped_reason") == "tool_error"
                and is_recoverable_tool_error(execution.get("error") or execution.get("answer"))
                and int(metadata.get("tool_error_recovery_count") or 0) < 1
                and step_index < request.max_steps
                and len(tool_calls) < request.max_tool_calls
            )
            if can_recover_tool_error:
                metadata["tool_error_recovery_count"] = int(metadata.get("tool_error_recovery_count") or 0) + 1
                recovery = build_tool_error_observation(
                    execution=execution,
                    call_index=len(tool_calls),
                    debug=request.debug,
                    recommended_next_tools=recommended_tools_for_recovery(
                        playbook_plan,
                        allowed_tools=allowed_tools,
                        failed_tool=execution.get("tool_name"),
                    ),
                    summarize_observation=_summarize_observation,
                    make_model_observation=lambda name, result, nodes, _observation: _model_observation(name, result, nodes),
                    make_public_result=lambda result, debug: _public_tool_result(result, debug=debug),
                )
                observations_for_model.append(recovery["model_observation"])
                public_observations.append(recovery["public_observation"])
                raw_observations.append(recovery["raw_observation"])
                tool_result_contracts.append(recovery["contract"])
                agent_steps.append(_step_payload(
                    step_index,
                    "tool_error_recovery",
                    provider=response.provider,
                    raw_model_output=raw_answer,
                    tool_name=execution["tool_name"],
                    args=execution["args"],
                    ok=False,
                    error=execution["error"],
                    summary=("工具失败已作为 observation 反馈给下一轮：" + execution["summary"])[:500],
                    debug=request.debug,
                ))
                metadata["stopped_reason"] = "tool_error_recovery"
                _checkpoint("quality_repair_continue", step_index=step_index, summary="工具失败已转为可恢复 observation。")
                continue
            metadata.update({"ok": False, "stopped_reason": execution["stopped_reason"]})
            answer = execution["answer"]
            agent_steps.append(_step_payload(
                step_index,
                execution["step_type"],
                provider=response.provider,
                raw_model_output=raw_answer,
                tool_name=execution["tool_name"],
                args=execution["args"],
                ok=False,
                error=execution["error"],
                summary=execution["summary"],
                debug=request.debug,
            ))
            _checkpoint("failed", step_index=step_index, summary=answer, error=execution["error"])
            break
        agent_steps.append(_step_payload(
            step_index,
            execution["step_type"],
            provider=response.provider,
            raw_model_output=raw_answer,
            tool_name=execution["tool_name"],
            args=execution["args"],
            summary=execution["summary"],
            debug=request.debug,
        ))
    else:
        metadata["stopped_reason"] = "max_steps"
        answer = "已达到最大 agent 轮数，需人工基于现有观察继续确认。"

    if not answer:
        metadata["stopped_reason"] = metadata.get("stopped_reason") or "empty_answer"
        answer = "Agent loop 未生成最终回答，请基于已收集的观察结果人工确认。"

    if metadata.get("stopped_reason") == "task_dispatched":
        subagents, subagent_summary = [], {"enabled": False, "skipped_reason": "task_dispatched"}
    else:
        subagents, subagent_summary = _run_subagents_parallel(
            report,
            bundle,
            request,
            provider,
            registry,
            run_agent=run_harness_agent,
            project_context=project_context,
        )
    metadata["subagents_enabled"] = bool(request.enable_subagents)
    metadata["subagent_count"] = len(subagents)

    finished_at = time.time()
    elapsed_ms = int((finished_at - started_at) * 1000)
    tool_dispatch_summary = summarize_tool_dispatch_trace(tool_dispatch_trace)
    trace_summary = {
        "agent_run_id": agent_run_id,
        "profile": request.profile,
        "capability_profiles": selected_profile_ids,
        "steps": len(agent_steps),
        "tool_call_count": len(tool_calls),
        "tool_dispatch_event_count": tool_dispatch_summary.get("event_count", 0),
        "tool_dispatch_blocked_count": tool_dispatch_summary.get("blocked_count", 0),
        "tool_signature_count": len(tool_signatures),
        "tool_result_contract_count": len(tool_result_contracts),
        "observation_count": len(public_observations),
        "raw_observation_count": len(raw_observations),
        "evidence_node_count": len(evidence_nodes),
        "citation_count": len(citations),
        "subagent_count": len(subagents),
        "task_dispatch_count": len(dispatched_tasks),
        "scratch_file_count": int(scratch_files.get("file_count") or 0) if isinstance(scratch_files, dict) else 0,
        "task_dispatch_available": bool(task_dispatch_summary.get("available")),
        "subagent_failed_count": int(subagent_summary.get("failed_count") or 0),
        "input_truncated": bool(metadata.get("input_truncated")),
        "last_context_budget": context_budget,
        "runtime_memory_fact_count": len((runtime_state.get("memory_summary") or {}).get("facts") or []),
        "runtime_evidence_id_count": int(runtime_state.get("evidence_id_count") or 0),
        "task_ledger_open_count": int(((runtime_state.get("task_ledger") or {}).get("progress") or {}).get("open") or 0),
        "task_ledger_next_action_count": len((runtime_state.get("task_ledger") or {}).get("next_actions") or []),
        "final_quality_status": final_answer_quality_gate.get("status", ""),
        "final_quality_score": final_answer_quality_gate.get("score", 0) if final_answer_quality_gate else 0,
        "quality_repair_attempt_count": int(metadata.get("quality_repair_attempt_count") or 0),
        "quality_repair_tool_count": int(metadata.get("quality_repair_tool_count") or 0),
        "tool_error_recovery_count": int(metadata.get("tool_error_recovery_count") or 0),
        "selected_skill_count": (agentic_context.get("selected_skills") or {}).get("selected_count", 0),
        "guidance_source_count": (agentic_context.get("guidance_summary") or {}).get("source_count", 0),
        "task_memory_found": bool((agentic_context.get("task_memory_summary") or {}).get("found")),
        "session_recent_evidence_count": len(session_state.get("recent_evidence_ids") or []),
        "session_pending_question_count": len(session_state.get("pending_questions") or []),
        "evidence_goal_status": (runtime_state.get("evidence_goal_contract") or {}).get("status", ""),
        "missing_evidence_goal_count": len((runtime_state.get("evidence_goal_contract") or {}).get("missing_evidence_types") or []),
        "stopped_reason": metadata.get("stopped_reason", ""),
        "elapsed_ms": elapsed_ms,
        "planning_context_chars": len(planning_context),
    }
    status = _status_from_metadata(metadata)
    limits = {
        "max_steps": request.max_steps,
        "max_tool_calls": request.max_tool_calls,
        "max_rows_per_table": request.max_rows_per_table,
        "enable_subagents": request.enable_subagents,
        "subagent_profiles": list(request.subagent_profiles),
        "max_subagents": request.max_subagents,
        "task_dispatch": True,
    }
    safeguards = [
        "Agent loop 只执行本地白名单只读工具。",
        "Subagent 默认关闭；开启后也只运行 focused profile 的本地只读工具。",
        "长任务分发只在异步 durable run 中创建 child runs；子任务仍只执行本地白名单只读工具。",
        "Agent 临时文件只能由 final_answer.scratch_files 声明，并由本地 runtime 写入 agent_workspace scratch 临时区。",
        "Aster 只返回 JSON 决策，不直接执行工具或访问文件。",
        "项目文件读取被限制在当前 project_root 的 packaged、sch_1、module_order(.dat)、page.map 范围内。",
    ]
    turn_context_snapshot = build_harness_turn_context_snapshot(
        agent_run_id=agent_run_id,
        mode="local-agent-harness",
        profile=request.profile,
        capability_profiles=selected_profile_ids,
        model_provider=metadata.get("provider", ""),
        model_mode=getattr(provider, "mode", ""),
        guidance_summary=agentic_context.get("guidance_summary", {}),
        selected_skills=agentic_context.get("selected_skills", {}),
        playbook_plan=playbook_plan,
        allowed_tools=allowed_tools,
        tool_list=registry.list_tools(),
        context_budget=context_budget,
        runtime_state=runtime_state,
        limits=limits,
        safeguards=safeguards,
        source="openai-codex-inspired",
    )

    payload = {
        "ok": bool(metadata.get("ok", True)),
        "status": status,
        "mode": "local-agent-harness",
        "agent_run_id": agent_run_id,
        "profile": request.profile,
        "capability_plan": capability_plan_items,
        "planning_context": planning_context,
        "playbook_plan": playbook_plan,
        "guidance_summary": agentic_context.get("guidance_summary", {}),
        "selected_skills": agentic_context.get("selected_skills", {}),
        "effort_policy": agentic_context.get("effort_policy", {}),
        "task_memory_summary": agentic_context.get("task_memory_summary", {}),
        "retry_reasons": list(metadata.get("retry_reasons") or metadata.get("perseverance_retry_notes") or []),
        "tool_result_contracts": tool_result_contracts,
        "turn_context_snapshot": turn_context_snapshot,
        "tool_dispatch_trace": tool_dispatch_trace,
        "tool_dispatch_summary": tool_dispatch_summary,
        "planner_warnings": list(playbook_plan.get("planner_warnings") or []),
        "answer": answer[:2000],
        "trace_summary": trace_summary,
        "agent_steps": agent_steps,
        "tool_calls": tool_calls,
        "observations": public_observations,
        "raw_observations": raw_observations,
        "final_evidence": evidence_nodes,
        "citations": citations,
        "proposed_actions": proposed_actions,
        "scratch_files": scratch_files,
        "needs_user_input": needs_user_input,
        "task_dispatch_summary": task_dispatch_summary,
        "dispatched_tasks": dispatched_tasks,
        "final_answer_quality_gate": final_answer_quality_gate,
        "project_context_summary": _compact_project_context_for_model(project_context),
        "context_budget": context_budget,
        "runtime_state": runtime_state,
        "session_state": session_state,
        "evidence_goal_contract": runtime_state.get("evidence_goal_contract", {}),
        "subagents": subagents,
        "subagent_summary": subagent_summary,
        "model_metadata": metadata,
        "limits": limits,
        "started_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(started_at)),
        "finished_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(finished_at)),
        "elapsed_ms": elapsed_ms,
        "safeguards": safeguards,
    }
    if memory_run_id:
        try:
            payload["task_memory_summary"] = persist_agentic_task_memory(memory_run_id, payload)
        except Exception as exc:  # pragma: no cover - defensive runtime persistence.
            payload["task_memory_summary"] = {
                "found": False,
                "error": str(exc),
                "path": (agentic_context.get("task_memory_summary") or {}).get("path", ""),
            }
    payload["execution_journal"] = build_execution_journal(payload)
    payload["journal_summary"] = build_journal_summary(payload["execution_journal"])
    payload["continuation_pack"] = build_continuation_pack(payload)
    _checkpoint(
        "completed" if payload.get("ok") and payload.get("status") != "waiting_for_user" else str(payload.get("status") or "completed"),
        step_index=len(agent_steps),
        summary="Agent run 已完成并生成最终 payload。",
        continuation_pack=payload.get("continuation_pack") or {},
    )
    return payload
