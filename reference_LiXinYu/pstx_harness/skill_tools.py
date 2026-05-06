# -*- coding: utf-8 -*-
"""Shared Harness Skill card tools for report and compare agents.

These tools expose repository-owned `harness_skills/*/SKILL.md` cards as
read-only guidance. Skill cards never grant tool permissions; agent profiles
and the harness registry still decide which executable tools exist.
"""

from __future__ import annotations

from pathlib import Path

from pstx_agent_runtime.skill_registry import (
    SKILL_REGISTRY_VERSION,
    load_harness_skills,
    select_harness_skills,
)
from pstx_harness.tool_core import HarnessToolError


HARNESS_SKILL_TOOL_NAMES = (
    "list_harness_skills",
    "select_harness_skills",
    "get_harness_skill",
)

GUIDANCE_NOTE = (
    "Skill cards are guidance only; executable tools are still controlled by "
    "the Harness profile whitelist and registry."
)


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _limit(value, default: int, minimum: int, maximum: int) -> int:
    try:
        number = int(value if value is not None else default)
    except (TypeError, ValueError):
        number = default
    return max(minimum, min(number, maximum))


def _skill_payload(*, mode: str, skills: list[dict], available_count: int) -> dict:
    return {
        "schema_version": SKILL_REGISTRY_VERSION,
        "mode": mode,
        "available_count": available_count,
        "returned_count": len(skills),
        "skills": skills,
        "guidance_note": GUIDANCE_NOTE,
    }


def _list_harness_skills_tool(context, args: dict) -> dict:
    include_body = bool(args.get("include_body", False))
    max_body_chars = _limit(args.get("max_body_chars"), 1800, 200, 20000)
    limit = _limit(args.get("limit"), 24, 1, 200)
    skills = load_harness_skills(_repo_root())
    cards = [
        skill.card(include_body=include_body, max_body_chars=max_body_chars)
        for skill in skills[:limit]
    ]
    return {
        "id": "list_harness_skills",
        "title": "Harness Skill 清单",
        "target": "skill",
        "summary": f"当前仓库包含 {len(skills)} 个 Harness Skill，返回 {len(cards)} 个。",
        "harness_skills": _skill_payload(mode="list", skills=cards, available_count=len(skills)),
        "recommended_next_tools": ["select_harness_skills", "get_harness_skill"],
        "completeness": "partial" if len(skills) > len(cards) else "complete",
        "readonly": True,
    }


def _select_harness_skills_tool(context, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    capability_profiles = [str(item).strip() for item in (args.get("capability_profiles") or []) if str(item).strip()]
    playbooks = [str(item).strip() for item in (args.get("playbooks") or []) if str(item).strip()]
    tools = [str(item).strip() for item in (args.get("tools") or []) if str(item).strip()]
    include_body = bool(args.get("include_body", True))
    max_body_chars = _limit(args.get("max_body_chars"), 1800, 200, 20000)
    limit = _limit(args.get("limit"), 4, 1, 24)
    selected = select_harness_skills(
        question=query,
        capability_profiles=capability_profiles,
        playbook_plan={
            "selected_playbooks": [{"id": item} for item in playbooks],
            "recommended_first_tools": tools,
        },
        root=_repo_root(),
        max_selected=limit,
        include_body=include_body,
        max_body_chars=max_body_chars,
    )
    cards = list(selected.get("selected_skills") or [])
    return {
        "id": "select_harness_skills",
        "title": "选择 Harness Skill",
        "target": "skill",
        "summary": f"按 query/profile/playbook/tool 选择到 {len(cards)} 个 Harness Skill。",
        "query": query,
        "capability_profiles": capability_profiles,
        "playbooks": playbooks,
        "tools": tools,
        "harness_skills": _skill_payload(
            mode="select",
            skills=cards,
            available_count=int(selected.get("available_count") or 0),
        ),
        "skill_cards": list(selected.get("skill_cards") or []),
        "recommended_next_tools": ["get_harness_skill"] if cards else ["list_harness_skills"],
        "completeness": "complete" if cards else "missing",
        "readonly": True,
    }


def _get_harness_skill_tool(context, args: dict) -> dict:
    skill_id = str(args.get("skill_id") or "").strip()
    if not skill_id:
        raise HarnessToolError("get_harness_skill 需要 skill_id。")
    include_body = bool(args.get("include_body", True))
    max_body_chars = _limit(args.get("max_body_chars"), 4000, 200, 20000)
    skills = load_harness_skills(_repo_root())
    for skill in skills:
        if skill.id == skill_id:
            card = skill.card(include_body=include_body, max_body_chars=max_body_chars)
            return {
                "id": "get_harness_skill",
                "title": card.get("title") or skill_id,
                "target": "skill",
                "summary": f"读取 Harness Skill：{skill_id}。{GUIDANCE_NOTE}",
                "skill_id": skill_id,
                "harness_skills": _skill_payload(mode="single", skills=[card], available_count=len(skills)),
                "skill": card,
                "recommended_next_tools": list(card.get("allowed_tools") or [])[:12],
                "completeness": "preview" if card.get("body_truncated") else "complete",
                "readonly": True,
            }
    raise HarnessToolError(f"未知 Harness Skill：{skill_id}")
