# -*- coding: utf-8 -*-
"""Generic subagent definitions, planning, and result compaction."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Mapping, Sequence


SUBAGENT_SCHEMA_VERSION = "pstx-agent-subagents.v1"
DEFAULT_SUBAGENT_ISOLATION = "fresh_context"


def _text(value: object, limit: int = 1000) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 3)] + "..."


def _list_of_text(value: object, *, limit: int = 40, item_limit: int = 120) -> list[str]:
    if value is None:
        return []
    source = value if isinstance(value, (list, tuple, set)) else [value]
    items: list[str] = []
    for item in source:
        text = _text(item, item_limit)
        if text and text not in items:
            items.append(text)
        if len(items) >= limit:
            break
    return items


def _positive_int(value: object, *, default: int = 0, maximum: int = 200) -> int:
    try:
        number = int(value if value is not None else default)
    except Exception:
        number = default
    if number <= 0:
        return 0
    return min(number, maximum)


@dataclass(frozen=True)
class SubagentDefinition:
    """A reusable focused-agent definition derived from a profile catalog."""

    id: str
    title: str
    profile: str
    description: str = ""
    prompt: str = ""
    tools: tuple[str, ...] = ()
    max_steps: int = 0
    max_tool_calls: int = 0
    model: str = ""
    background: bool = False
    isolation: str = DEFAULT_SUBAGENT_ISOLATION
    metadata: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "schema_version": SUBAGENT_SCHEMA_VERSION,
            "id": self.id,
            "title": self.title,
            "profile": self.profile,
            "description": self.description,
            "prompt": self.prompt,
            "tools": list(self.tools),
            "max_steps": self.max_steps,
            "max_tool_calls": self.max_tool_calls,
            "model": self.model,
            "background": self.background,
            "isolation": self.isolation,
            "metadata": dict(self.metadata),
        }


def build_subagent_definition(profile_id: object,
                              profile_config: Mapping[str, object],
                              *,
                              allowed_tools: Sequence[object] | None = None,
                              max_steps: int = 0,
                              max_tool_calls: int = 0,
                              background: bool = False,
                              isolation: str = DEFAULT_SUBAGENT_ISOLATION) -> dict:
    """Build a stable subagent definition from a generic profile mapping."""

    profile = _text(profile_id, 80)
    config = dict(profile_config or {})
    title = _text(config.get("title") or profile, 180)
    description = _text(config.get("description") or "", 500)
    prompt = _text(config.get("default_question") or description or title, 1200)
    configured_tools = _list_of_text(config.get("tools"), limit=120)
    tools = _list_of_text(allowed_tools, limit=120) if allowed_tools is not None else configured_tools
    definition = SubagentDefinition(
        id=f"subagent-{profile}",
        title=title,
        profile=profile,
        description=description,
        prompt=prompt,
        tools=tuple(tools),
        max_steps=_positive_int(max_steps or config.get("max_steps"), maximum=200),
        max_tool_calls=_positive_int(max_tool_calls or config.get("max_tool_calls"), maximum=400),
        model=_text(config.get("model") or "", 80),
        background=bool(background),
        isolation=_text(isolation or DEFAULT_SUBAGENT_ISOLATION, 80),
        metadata={
            "source": "profile_catalog",
            "profile": profile,
        },
    )
    return definition.to_dict()


def plan_subagents(requested_profiles: Sequence[object],
                   profile_catalog: Mapping[str, Mapping[str, object]],
                   *,
                   max_subagents: int = 4,
                   disallowed_profiles: Sequence[object] = (),
                   profile_allowed_tools: Mapping[str, Sequence[object]] | None = None,
                   parent_profile: object = "") -> dict:
    """Normalize requested profiles into executable subagent definitions."""

    catalog = dict(profile_catalog or {})
    disallowed = set(_list_of_text(disallowed_profiles, limit=40, item_limit=80))
    max_count = max(0, int(max_subagents or 0))
    profiles: list[str] = []
    skipped: list[dict] = []
    for raw_profile in requested_profiles or []:
        profile = _text(raw_profile, 80)
        if not profile:
            continue
        if profile in profiles:
            skipped.append({"profile": profile, "reason": "duplicate"})
            continue
        if profile in disallowed:
            skipped.append({"profile": profile, "reason": "disallowed"})
            continue
        if profile not in catalog:
            skipped.append({"profile": profile, "reason": "unknown_profile"})
            continue
        if max_count <= 0 or len(profiles) >= max_count:
            skipped.append({"profile": profile, "reason": "max_subagents"})
            continue
        profiles.append(profile)
    definitions = [
        build_subagent_definition(
            profile,
            catalog.get(profile) or {},
            allowed_tools=(profile_allowed_tools or {}).get(profile) if profile_allowed_tools else None,
        )
        for profile in profiles
    ]
    return {
        "schema_version": SUBAGENT_SCHEMA_VERSION,
        "parent_profile": _text(parent_profile, 80),
        "requested_count": len(list(requested_profiles or [])),
        "planned_count": len(definitions),
        "profiles": profiles,
        "definitions": definitions,
        "skipped": skipped,
        "warnings": [
            "no_subagents_planned"
        ] if requested_profiles and not definitions else [],
    }


def build_subagent_question(parent_question: object,
                            definition: Mapping[str, object],
                            *,
                            extra_instruction: object = "") -> str:
    """Build a focused child-agent question from a definition and parent goal."""

    title = _text(definition.get("title") if isinstance(definition, Mapping) else "", 180)
    prompt = _text(definition.get("prompt") if isinstance(definition, Mapping) else "", 1400)
    profile = _text(definition.get("profile") if isinstance(definition, Mapping) else "", 80)
    parent = _text(parent_question or "请按 profile 聚焦审查。", 1200)
    extra = _text(extra_instruction, 500)
    parts = [
        prompt or f"请按 {title or profile} 聚焦审查。",
        f"父任务问题：{parent}",
        "请只输出当前 subagent 定义范围内的证据、风险和人工复核建议。",
    ]
    if extra:
        parts.append(extra)
    return "\n".join(parts)[:2200]


def compact_subagent_result(result: Mapping[str, object],
                            *,
                            definition: Mapping[str, object] | None = None,
                            answer_limit: int = 1800,
                            max_citations: int = 8,
                            max_actions: int = 8) -> dict:
    """Compact a child run payload into the parent-facing subagent row."""

    payload = dict(result or {})
    definition_payload = dict(definition or payload.get("subagent_definition") or {})
    citations = payload.get("citations") if isinstance(payload.get("citations"), list) else []
    final_evidence = payload.get("final_evidence") if isinstance(payload.get("final_evidence"), list) else []
    proposed_actions = payload.get("proposed_actions") if isinstance(payload.get("proposed_actions"), list) else []
    return {
        "schema_version": SUBAGENT_SCHEMA_VERSION,
        "profile": _text(payload.get("profile") or definition_payload.get("profile") or "", 80),
        "title": _text(definition_payload.get("title") or payload.get("title") or payload.get("profile") or "", 180),
        "ok": bool(payload.get("ok")),
        "status": "completed" if payload.get("ok") else _text(payload.get("status") or "failed", 80),
        "agent_run_id": _text(payload.get("agent_run_id") or "", 120),
        "answer": _text(payload.get("answer") or "", answer_limit),
        "trace_summary": dict(payload.get("trace_summary") if isinstance(payload.get("trace_summary"), Mapping) else {}),
        "citation_count": len(citations),
        "evidence_node_count": len(final_evidence),
        "proposed_action_count": len(proposed_actions),
        "citations": list(citations[:max_citations]),
        "proposed_actions": list(proposed_actions[:max_actions]),
        "model_metadata": dict(payload.get("model_metadata") if isinstance(payload.get("model_metadata"), Mapping) else {}),
        "definition": definition_payload,
        "isolation": _text(definition_payload.get("isolation") or DEFAULT_SUBAGENT_ISOLATION, 80),
    }


def summarize_subagent_results(results: Sequence[Mapping[str, object]],
                               *,
                               plan: Mapping[str, object] | None = None,
                               max_workers: int = 1,
                               elapsed_ms: int = 0,
                               provider_parallel_safe: bool = True) -> dict:
    """Build a stable parent-facing subagent summary."""

    rows = [dict(item) for item in (results or []) if isinstance(item, Mapping)]
    planned_profiles = _list_of_text((plan or {}).get("profiles"), limit=80)
    failed_profiles = [str(item.get("profile") or "") for item in rows if not item.get("ok")]
    return {
        "schema_version": SUBAGENT_SCHEMA_VERSION,
        "enabled": bool((plan or {}).get("planned_count") or rows),
        "planned_count": int((plan or {}).get("planned_count") or len(planned_profiles) or len(rows)),
        "completed_count": len(rows),
        "failed_count": len(failed_profiles),
        "profiles": planned_profiles or [str(item.get("profile") or "") for item in rows],
        "max_workers": max(1, int(max_workers or 1)),
        "elapsed_ms": max(0, int(elapsed_ms or 0)),
        "total_evidence_node_count": sum(int(item.get("evidence_node_count") or 0) for item in rows),
        "total_proposed_action_count": sum(int(item.get("proposed_action_count") or 0) for item in rows),
        "degraded": bool(failed_profiles),
        "failed_profiles": failed_profiles,
        "provider_parallel_safe": bool(provider_parallel_safe),
        "skipped": list((plan or {}).get("skipped") or []),
        "warnings": list((plan or {}).get("warnings") or []),
    }
