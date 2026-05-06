# -*- coding: utf-8 -*-
"""Capability planning helpers for PSTX agent runtimes."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping, Sequence


@dataclass(frozen=True)
class AgentCapabilityProfile:
    id: str
    title: str
    description: str
    tools: tuple[str, ...]
    default_question: str
    max_steps: int
    max_tool_calls: int
    subagent_profiles: tuple[str, ...] = ()

    @classmethod
    def from_mapping(cls, profile_id: str, value: Mapping[str, object]) -> "AgentCapabilityProfile":
        return cls(
            id=str(profile_id),
            title=str(value.get("title") or profile_id),
            description=str(value.get("description") or ""),
            tools=tuple(str(item) for item in value.get("tools", []) or []),
            default_question=str(value.get("default_question") or ""),
            max_steps=int(value.get("max_steps") or 1),
            max_tool_calls=int(value.get("max_tool_calls") or 0),
            subagent_profiles=tuple(str(item) for item in value.get("subagent_profiles", []) or []),
        )

    def to_public_dict(self, *, include_subagents: bool = False) -> dict:
        payload = {
            "id": self.id,
            "title": self.title,
            "description": self.description,
            "tools": list(self.tools),
            "default_question": self.default_question,
            "max_steps": self.max_steps,
            "max_tool_calls": self.max_tool_calls,
        }
        if include_subagents:
            payload["subagent_profiles"] = list(self.subagent_profiles)
        return payload


@dataclass(frozen=True)
class AgentCapabilityPlan:
    requested_profile: str
    capability_profiles: tuple[str, ...]
    allowed_tools: tuple[str, ...]
    plan_items: tuple[dict, ...]

    def to_dict(self) -> dict:
        return {
            "requested_profile": self.requested_profile,
            "capability_profiles": list(self.capability_profiles),
            "allowed_tools": list(self.allowed_tools),
            "plan_items": [dict(item) for item in self.plan_items],
        }


def profile_config(profiles: Mapping[str, Mapping[str, object]],
                   profile_id: str,
                   *,
                   default_profile: str) -> dict:
    return dict(profiles.get(profile_id) or profiles[default_profile])


def list_public_profiles(profiles: Mapping[str, Mapping[str, object]],
                         *,
                         include_subagents: bool = False) -> list[dict]:
    return [
        AgentCapabilityProfile.from_mapping(profile_id, config).to_public_dict(include_subagents=include_subagents)
        for profile_id, config in profiles.items()
    ]


def question_text(*,
                  profile_id: str,
                  question: str,
                  profiles: Mapping[str, Mapping[str, object]],
                  default_profile: str) -> str:
    config = profile_config(profiles, profile_id, default_profile=default_profile)
    return f"{question or ''} {config.get('default_question') or ''}".lower()


def dedupe_profile_ids(profile_ids: Sequence[str],
                       profiles: Mapping[str, Mapping[str, object]]) -> list[str]:
    result: list[str] = []
    for profile_id in profile_ids:
        text = str(profile_id)
        if text in profiles and text not in result and text != "auto":
            result.append(text)
    return result


def infer_capability_profiles(*,
                              requested_profile: str,
                              question: str,
                              profiles: Mapping[str, Mapping[str, object]],
                              default_profile: str,
                              rules: Sequence[tuple[str, Sequence[str]]],
                              quick_profile: str | None = None) -> list[str]:
    if requested_profile != "auto":
        return dedupe_profile_ids([requested_profile], profiles)
    text = question_text(
        profile_id=requested_profile,
        question=question,
        profiles=profiles,
        default_profile=default_profile,
    )
    upper = text.upper()
    selected: list[str] = []
    for profile_id, keywords in rules:
        if any(str(keyword).lower() in text or str(keyword).upper() in upper for keyword in keywords):
            selected.append(profile_id)
    if not selected:
        selected.append(default_profile)
    if quick_profile and len(selected) >= 3 and quick_profile not in selected:
        selected.insert(0, quick_profile)
    return dedupe_profile_ids(selected, profiles)


def build_capability_plan_items(profile_ids: Sequence[str],
                                profiles: Mapping[str, Mapping[str, object]]) -> list[dict]:
    items: list[dict] = []
    for profile_id in profile_ids:
        config = dict(profiles.get(profile_id) or {})
        items.append({
            "id": profile_id,
            "title": config.get("title", profile_id),
            "description": config.get("description", ""),
        })
    return items


def allowed_tool_names(*,
                       profile_ids: Sequence[str],
                       profiles: Mapping[str, Mapping[str, object]],
                       registry_tools: Sequence[Mapping[str, object]]) -> list[str]:
    all_names = [str(item.get("name") or "") for item in registry_tools if str(item.get("name") or "")]
    configured: list[str] = []
    for profile_id in profile_ids:
        profile_tools = [str(item) for item in (profiles.get(profile_id, {}).get("tools") or [])]
        if "*" in profile_tools:
            return all_names
        configured.extend(profile_tools)
    allowed = set(configured)
    return [name for name in all_names if name in allowed]


def filtered_tool_list(*,
                       profile_ids: Sequence[str],
                       profiles: Mapping[str, Mapping[str, object]],
                       registry_tools: Sequence[Mapping[str, object]]) -> list[dict]:
    allowed = set(allowed_tool_names(profile_ids=profile_ids, profiles=profiles, registry_tools=registry_tools))
    return [dict(tool) for tool in registry_tools if tool.get("name") in allowed]


def build_capability_plan(*,
                          requested_profile: str,
                          question: str,
                          profiles: Mapping[str, Mapping[str, object]],
                          default_profile: str,
                          rules: Sequence[tuple[str, Sequence[str]]],
                          registry_tools: Sequence[Mapping[str, object]],
                          quick_profile: str | None = None) -> AgentCapabilityPlan:
    profile_ids = infer_capability_profiles(
        requested_profile=requested_profile,
        question=question,
        profiles=profiles,
        default_profile=default_profile,
        rules=rules,
        quick_profile=quick_profile,
    )
    allowed = allowed_tool_names(profile_ids=profile_ids, profiles=profiles, registry_tools=registry_tools)
    return AgentCapabilityPlan(
        requested_profile=requested_profile,
        capability_profiles=tuple(profile_ids),
        allowed_tools=tuple(allowed),
        plan_items=tuple(build_capability_plan_items(profile_ids, profiles)),
    )
