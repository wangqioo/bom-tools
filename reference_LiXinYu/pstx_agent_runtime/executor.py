# -*- coding: utf-8 -*-
"""Controlled runtime tool executor shared by PSTX agents."""

from __future__ import annotations

import json
import time
from typing import Callable, Iterable, Mapping, Sequence

from .compression import build_evidence_layers, json_char_count
from .playbook import build_tool_result_contract
from .turn_context import build_tool_dispatch_event


def _stable_json(value: object) -> str:
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    except (TypeError, ValueError):
        return str(value)


def _as_args(tool_call: Mapping[str, object]) -> dict:
    args = tool_call.get("args") if isinstance(tool_call, Mapping) else {}
    return dict(args) if isinstance(args, dict) else {}


def _tool_name(tool_call: Mapping[str, object]) -> str:
    return str(tool_call.get("name") or tool_call.get("tool") or "").strip()


def _reason(tool_call: Mapping[str, object]) -> str:
    return str(tool_call.get("reason") or "").strip()


def _tool_attr(tool: object, name: str, default: object = None) -> object:
    if isinstance(tool, Mapping):
        return tool.get(name, default)
    return getattr(tool, name, default)


def _call_or_value(value: object, default: object = "") -> object:
    if callable(value):
        try:
            return value()
        except TypeError:
            return default
    return value


def _tool_boundary_metadata(tool: object) -> dict:
    file_access = bool(_tool_attr(tool, "file_access", False))
    approval_scope = _call_or_value(_tool_attr(tool, "normalized_approval_scope", None), "")
    evidence_kind = _call_or_value(_tool_attr(tool, "normalized_evidence_kind", None), "")
    return {
        "readonly": _tool_attr(tool, "readonly", True) is not False,
        "file_access": file_access,
        "mutating": bool(_tool_attr(tool, "mutating", False)),
        "supports_parallel": bool(_tool_attr(tool, "supports_parallel", False)),
        "approval_scope": str(approval_scope or _tool_attr(tool, "approval_scope", "") or ("read_project_file" if file_access else "none")),
        "evidence_kind": str(evidence_kind or _tool_attr(tool, "evidence_kind", "") or _tool_attr(tool, "target", "") or "general"),
    }


def _assert_tool_boundary(tool_name: str, metadata: Mapping[str, object]) -> None:
    if metadata.get("readonly") is not True:
        raise ValueError(f"工具 {tool_name} 不是只读工具，已拒绝执行。")
    if metadata.get("mutating"):
        raise ValueError(f"工具 {tool_name} 声明为 mutating，已拒绝执行。")
    if metadata.get("approval_scope") not in {"none", "read_project_file"}:
        raise ValueError(f"工具 {tool_name} 的 approval_scope 不在允许范围内。")


def tool_call_signature(tool_name: object, args: Mapping[str, object] | None = None) -> str:
    """Build a deterministic signature for duplicate read-only tool-call detection."""

    return f"{str(tool_name or '').strip()}::{_stable_json(dict(args or {}))}"


def is_recoverable_tool_error(error: object) -> bool:
    """Return whether a tool failure can be converted into an observation."""

    text = str(error or "")
    non_recoverable_markers = (
        "不允许调用工具",
        "项目根目录之外",
        "不允许读取",
        "路径越权",
        "权限",
        "secret",
        "token",
        "Authorization",
    )
    return not any(marker in text for marker in non_recoverable_markers)


def recommended_tools_for_recovery(playbook_plan: Mapping[str, object] | None,
                                   *,
                                   allowed_tools: Iterable[str],
                                   failed_tool: object = "",
                                   limit: int = 4) -> list[str]:
    """Return safe playbook tools that can be tried after a recoverable error."""

    if not isinstance(playbook_plan, Mapping):
        return []
    allowed = {str(item) for item in allowed_tools or [] if str(item)}
    failed = str(failed_tool or "").strip()
    result: list[str] = []
    for item in playbook_plan.get("recommended_first_tools") or []:
        name = str(item or "").strip()
        if not name or name == failed:
            continue
        if allowed and name not in allowed:
            continue
        if name in result:
            continue
        result.append(name)
        if len(result) >= max(0, int(limit or 0)):
            break
    return result


def build_tool_error_observation(*,
                                 execution: Mapping[str, object],
                                 call_index: int,
                                 debug: bool,
                                 recommended_next_tools: Sequence[object] = (),
                                 summarize_observation: Callable[[str, dict], dict] | None = None,
                                 make_model_observation: Callable[[str, dict, list, dict], dict] | None = None,
                                 make_public_result: Callable[[dict, bool], dict] | None = None) -> dict:
    """Build synthetic observations for a recoverable tool error."""

    tool_name = str(execution.get("tool_name") or "tool_error")
    error = str(execution.get("error") or execution.get("answer") or "工具调用失败。")
    result = {
        "ok": False,
        "id": f"tool-error-{call_index}",
        "title": f"工具调用失败：{tool_name}",
        "summary": f"工具 {tool_name} 调用失败：{error}",
        "tool": tool_name,
        "error": error,
    }
    if debug:
        args = execution.get("args")
        result["args"] = dict(args) if isinstance(args, Mapping) else {}

    contract = {
        "completeness": "error",
        "recommended_next_tools": [str(item) for item in recommended_next_tools or [] if str(item)],
        "scope_summary": result["summary"][:500],
    }
    summarize_observation = summarize_observation or (lambda name, payload: {
        "tool": name,
        "ok": False,
        "id": payload.get("id", name),
        "title": payload.get("title", name),
        "summary": payload.get("summary", ""),
    })
    make_model_observation = make_model_observation or (
        lambda _name, payload, nodes, observation: {**observation, "result": payload, "evidence_nodes": nodes}
    )
    make_public_result = make_public_result or (lambda payload, _debug: dict(payload or {}))

    observation = dict(summarize_observation(tool_name, result) or {})
    observation.update({
        "ok": False,
        "error": error,
        "tool_result_contract": contract,
    })
    layers = build_evidence_layers(
        tool_name=tool_name,
        result=result,
        evidence_nodes=[],
        observation=observation,
        tool_result_contract=contract,
        include_raw_preview=False,
    )
    model_observation = dict(make_model_observation(tool_name, result, [], observation) or {})
    model_observation.update({
        "ok": False,
        "error": error,
        "tool_result_contract": contract,
    })
    model_observation.setdefault("evidence_layers", layers)

    public_result = dict(make_public_result(result, debug) or {})
    public_result.setdefault("tool_result_contract", contract)
    public_layers = build_evidence_layers(
        tool_name=tool_name,
        result=result,
        evidence_nodes=[],
        observation=observation,
        tool_result_contract=contract,
        include_raw_preview=True,
        raw_preview=public_result,
    )
    public_observation = {
        **observation,
        "result": public_result,
        "evidence_node_ids": [],
        "evidence_nodes": [],
        "evidence_layers": public_layers,
    }
    raw_observation = {
        "tool": tool_name,
        "call_index": call_index,
        "summary": result["summary"][:500],
        "evidence_node_ids": [],
        "evidence_layers": build_evidence_layers(
            tool_name=tool_name,
            result=result,
            evidence_nodes=[],
            observation=observation,
            tool_result_contract=contract,
            include_raw_preview=True,
            raw_preview=result,
        ),
        "raw_result_json_chars": json_char_count(result),
        "raw_result": result,
    }
    return {
        "model_observation": model_observation,
        "public_observation": public_observation,
        "raw_observation": raw_observation,
        "contract": {"tool": tool_name, "call_index": call_index, **contract},
    }


def _plan_label(profile_label: str, capability_profiles: Sequence[object]) -> str:
    profiles = [str(item) for item in capability_profiles or [] if str(item)]
    return ",".join(profiles) or profile_label


def _elapsed_ms(start: float) -> float:
    return max(0.0, (time.perf_counter() - start) * 1000.0)


def execute_runtime_tool_calls(*,
                               tool_call_items: Sequence[Mapping[str, object]],
                               is_batch_call: bool,
                               registry,
                               context,
                               allowed_tools: Iterable[str],
                               existing_tool_call_count: int,
                               max_tool_calls: int,
                               debug: bool,
                               profile_label: str,
                               capability_profiles: Sequence[object] = (),
                               previous_tool_calls: Sequence[Mapping[str, object]] = (),
                               previous_tool_signatures: Iterable[str] = (),
                               rejection_prefix: str = "工具调用被本地 harness 拒绝",
                               empty_message: str = "模型未返回可执行的工具调用。",
                               limit_message: str = "已达到最大工具调用次数，需人工基于现有观察继续确认。",
                               make_evidence_nodes: Callable[[str, dict, int, dict], list] | None = None,
                               summarize_observation: Callable[[str, dict], dict] | None = None,
                               make_model_observation: Callable[[str, dict, list, dict], dict] | None = None,
                               make_public_result: Callable[[dict, bool], dict] | None = None,
                               make_tool_result_contract: Callable[[str, dict], dict] | None = None) -> dict:
    """Execute one model tool step under local whitelist/schema control."""

    items = [dict(item) for item in tool_call_items or [] if isinstance(item, Mapping)]
    step_type = "tool_batch_call" if is_batch_call else "tool_call"
    if not items:
        return {
            "ok": False,
            "stopped_reason": "invalid_model_json",
            "answer": empty_message,
            "step_type": "model_error",
            "tool_name": "",
            "args": {},
            "error": empty_message,
            "summary": "",
            "tool_calls": [],
            "observations_for_model": [],
            "public_observations": [],
            "raw_observations": [],
            "evidence_nodes": [],
            "tool_result_contracts": [],
            "tool_signatures": [],
            "tool_dispatch_trace": [],
        }
    if existing_tool_call_count + len(items) > max_tool_calls:
        first = items[0]
        name = "tool_batch_call" if is_batch_call else _tool_name(first)
        args = {"requested_tool_count": len(items)} if is_batch_call else _as_args(first)
        return {
            "ok": False,
            "stopped_reason": "max_tool_calls",
            "answer": limit_message,
            "step_type": "limit",
            "tool_name": name,
            "args": args,
            "error": limit_message,
            "summary": "",
            "tool_calls": [],
            "observations_for_model": [],
            "public_observations": [],
            "raw_observations": [],
            "evidence_nodes": [],
            "tool_result_contracts": [],
            "tool_signatures": [],
            "tool_dispatch_trace": [build_tool_dispatch_event(
                event_index=1,
                tool=name,
                args=args,
                status="limit",
                reason=limit_message,
                profile_label=profile_label,
                capability_profiles=capability_profiles,
                batch=is_batch_call,
                allowed=True,
                debug=debug,
                error=limit_message,
            )],
        }

    allowed = set(allowed_tools or [])
    all_tool_calls: list[dict] = []
    observations_for_model: list[dict] = []
    public_observations: list[dict] = []
    raw_observations: list[dict] = []
    evidence_nodes: list[dict] = []
    tool_result_contracts: list[dict] = []
    tool_signatures: list[str] = []
    tool_dispatch_trace: list[dict] = []
    summaries: list[str] = []
    previous_signatures = {
        str(call.get("signature") or tool_call_signature(call.get("tool") or call.get("name"), call.get("args") if isinstance(call.get("args"), Mapping) else {}))
        for call in previous_tool_calls or []
        if isinstance(call, Mapping) and (call.get("tool") or call.get("name"))
    }
    previous_signatures.update(str(item) for item in previous_tool_signatures or [] if str(item))
    current_signatures: set[str] = set()

    make_evidence_nodes = make_evidence_nodes or (lambda _name, _result, _index, _args: [])
    summarize_observation = summarize_observation or (lambda name, result: {
        "tool": name,
        "ok": True,
        "id": result.get("id", name) if isinstance(result, dict) else name,
        "title": result.get("title", name) if isinstance(result, dict) else name,
        "summary": result.get("summary", "") if isinstance(result, dict) else "",
    })
    make_model_observation = make_model_observation or (
        lambda _name, result, nodes, observation: {**observation, "result": result, "evidence_nodes": nodes}
    )
    make_public_result = make_public_result or (lambda result, _debug: dict(result or {}))
    make_tool_result_contract = make_tool_result_contract or (
        lambda name, result: build_tool_result_contract(name, result).to_dict()
    )

    for item in items:
        name = _tool_name(item)
        args = _as_args(item)
        reason = _reason(item)
        signature = tool_call_signature(name, args)
        repeated_signature = signature in previous_signatures or signature in current_signatures
        current_signatures.add(signature)
        tool_metadata: dict = {}
        call_index = existing_tool_call_count + len(all_tool_calls) + 1
        call_id = str(item.get("call_id") or f"tool-call-{call_index}")
        call_started = time.perf_counter()
        preflight_status = "not_started"
        try:
            tool = registry.get(name)
            tool_metadata = _tool_boundary_metadata(tool)
            if name not in allowed:
                preflight_status = "blocked_by_profile"
                plan = _plan_label(profile_label, capability_profiles)
                raise ValueError(f"profile {profile_label}（capability plan: {plan}）不允许调用工具：{name}")
            _assert_tool_boundary(name, tool_metadata)
            preflight_status = "passed"
            result = dict(registry.run(name, context, args=args) or {})
        except Exception as exc:
            error = str(exc)
            if preflight_status == "not_started":
                preflight_status = "tool_lookup_failed" if not tool_metadata else "failed"
            allowed_by_profile = name in allowed
            failed_trace = tool_dispatch_trace + [build_tool_dispatch_event(
                event_index=len(tool_dispatch_trace) + 1,
                tool=name,
                args=args,
                status="failed" if allowed_by_profile else "blocked",
                call_id=call_id,
                reason=reason,
                profile_label=profile_label,
                capability_profiles=capability_profiles,
                signature=signature,
                batch=is_batch_call,
                allowed=allowed_by_profile,
                debug=debug,
                call_index=call_index,
                tool_metadata=tool_metadata,
                preflight_status=preflight_status,
                error=error,
                duration_ms=_elapsed_ms(call_started),
            )]
            return {
                "ok": False,
                "stopped_reason": "tool_error",
                "answer": f"{rejection_prefix}：{error}",
                "step_type": step_type,
                "tool_name": name,
                "args": args,
                "error": error,
                "summary": reason,
                "tool_calls": all_tool_calls + [{
                    "tool": name,
                    "args": args if debug else {},
                    "ok": False,
                    "error": error,
                    "reason": reason,
                    "batch": is_batch_call,
                }],
                "observations_for_model": observations_for_model,
                "public_observations": public_observations,
                "raw_observations": raw_observations,
                "evidence_nodes": evidence_nodes,
                "tool_result_contracts": tool_result_contracts,
                "tool_signatures": tool_signatures,
                "tool_dispatch_trace": failed_trace,
            }

        nodes = list(make_evidence_nodes(name, result, call_index, args) or [])
        observation = dict(summarize_observation(name, result) or {})
        contract = dict(make_tool_result_contract(name, result) or {})
        if contract:
            observation["tool_result_contract"] = contract
        model_layers = build_evidence_layers(
            tool_name=name,
            result=result,
            evidence_nodes=nodes,
            observation=observation,
            tool_result_contract=contract,
            include_raw_preview=False,
        )
        model_observation = dict(make_model_observation(name, result, nodes, observation) or {})
        if contract:
            model_observation["tool_result_contract"] = contract
        model_observation.setdefault("evidence_layers", model_layers)
        public_result = dict(make_public_result(result, debug) or {})
        if contract:
            public_result.setdefault("tool_result_contract", contract)
        public_layers = build_evidence_layers(
            tool_name=name,
            result=result,
            evidence_nodes=nodes,
            observation=observation,
            tool_result_contract=contract,
            include_raw_preview=True,
            raw_preview=public_result,
        )
        evidence_ids = [str(node.get("id")) for node in nodes if isinstance(node, Mapping) and node.get("id")]

        evidence_nodes.extend(nodes)
        if contract:
            tool_result_contracts.append({
                "tool": name,
                "call_index": call_index,
                **contract,
            })
        observations_for_model.append(model_observation)
        public_observations.append({
            **observation,
            "result": public_result,
            "evidence_node_ids": evidence_ids,
            "evidence_nodes": nodes,
            "evidence_layers": public_layers,
        })
        raw_observations.append({
            "tool": name,
            "call_index": call_index,
            "summary": str(observation.get("summary") or "")[:500],
            "evidence_node_ids": evidence_ids,
            "evidence_layers": build_evidence_layers(
                tool_name=name,
                result=result,
                evidence_nodes=nodes,
                observation=observation,
                tool_result_contract=contract,
                include_raw_preview=True,
                raw_preview=result,
            ),
            "raw_result_json_chars": json_char_count(result),
            "raw_result": result,
        })
        summaries.append(reason or str(observation.get("summary") or "") or name)
        all_tool_calls.append({
            "index": call_index,
            "tool": name,
            "args": args if debug else {},
            "ok": True,
            "reason": reason,
            "batch": is_batch_call,
            "duplicate": repeated_signature,
            "batch_step_index": None,
            "evidence_node_ids": evidence_ids,
            "tool_result_contract": contract,
        })
        tool_signatures.append(signature)
        tool_dispatch_trace.append(build_tool_dispatch_event(
            event_index=len(tool_dispatch_trace) + 1,
            tool=name,
            args=args,
            status="completed",
            call_id=call_id,
            reason=reason or str(observation.get("summary") or ""),
            profile_label=profile_label,
            capability_profiles=capability_profiles,
            signature=signature,
            batch=is_batch_call,
            duplicate=repeated_signature,
            allowed=True,
            debug=debug,
            call_index=call_index,
            evidence_ids=evidence_ids,
            contract=contract,
            tool_metadata=tool_metadata,
            preflight_status=preflight_status,
            duration_ms=_elapsed_ms(call_started),
            raw_result_json_chars=json_char_count(result),
        ))

    return {
        "ok": True,
        "stopped_reason": "",
        "answer": "",
        "step_type": step_type,
        "tool_name": ",".join(_tool_name(item) for item in items) if is_batch_call else _tool_name(items[0]),
        "args": {"tool_count": len(items)} if is_batch_call else _as_args(items[0]),
        "error": "",
        "summary": "；".join(item for item in summaries if item)[:500] or "批量工具执行完成。",
        "tool_calls": all_tool_calls,
        "observations_for_model": observations_for_model,
        "public_observations": public_observations,
        "raw_observations": raw_observations,
        "evidence_nodes": evidence_nodes,
        "tool_result_contracts": tool_result_contracts,
        "tool_signatures": tool_signatures,
        "tool_dispatch_trace": tool_dispatch_trace,
    }


def merge_runtime_tool_execution(*,
                                 execution: Mapping[str, object],
                                 tool_calls: list,
                                 tool_signatures: list,
                                 tool_dispatch_trace: list,
                                 tool_result_contracts: list,
                                 observations_for_model: list,
                                 public_observations: list,
                                 raw_observations: list,
                                 evidence_nodes: list,
                                 metadata: dict | None = None,
                                 metadata_prefix: str = "",
                                 record_observation_count: bool = True) -> dict:
    """Merge an executor result into mutable agent-loop state lists.

    Report and Compare agents both maintain the same runtime state shape after
    each local tool execution. Keeping this append logic in the runtime layer
    prevents future prefetch/repair paths from forgetting one of the companion
    lists such as dispatch trace, contracts, or raw observations.
    """

    payload = dict(execution or {})
    new_tool_calls = list(payload.get("tool_calls") or [])
    new_tool_signatures = list(payload.get("tool_signatures") or [])
    new_dispatch_trace = list(payload.get("tool_dispatch_trace") or [])
    new_contracts = list(payload.get("tool_result_contracts") or [])
    new_model_observations = list(payload.get("observations_for_model") or [])
    new_public_observations = list(payload.get("public_observations") or [])
    new_raw_observations = list(payload.get("raw_observations") or [])
    new_evidence_nodes = list(payload.get("evidence_nodes") or [])

    tool_calls.extend(new_tool_calls)
    tool_signatures.extend(new_tool_signatures)
    tool_dispatch_trace.extend(new_dispatch_trace)
    tool_result_contracts.extend(new_contracts)
    observations_for_model.extend(new_model_observations)
    public_observations.extend(new_public_observations)
    raw_observations.extend(new_raw_observations)
    evidence_nodes.extend(new_evidence_nodes)

    counts = {
        "tool_count": len(new_tool_calls),
        "tool_signature_count": len(new_tool_signatures),
        "dispatch_trace_count": len(new_dispatch_trace),
        "tool_result_contract_count": len(new_contracts),
        "observation_count": len(new_model_observations),
        "public_observation_count": len(new_public_observations),
        "raw_observation_count": len(new_raw_observations),
        "evidence_node_count": len(new_evidence_nodes),
    }
    if metadata is not None and metadata_prefix:
        prefix = str(metadata_prefix).strip().rstrip("_")
        metadata[f"{prefix}_ok"] = bool(payload.get("ok"))
        metadata[f"{prefix}_tool_count"] = counts["tool_count"]
        if record_observation_count:
            metadata[f"{prefix}_observation_count"] = counts["observation_count"]
        if not payload.get("ok"):
            metadata[f"{prefix}_error"] = payload.get("error") or payload.get("answer") or ""
    return counts
