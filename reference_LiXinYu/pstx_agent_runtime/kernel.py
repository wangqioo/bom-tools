# -*- coding: utf-8 -*-
"""Shared agentic envelope for PSTX report and compare harness loops."""

from __future__ import annotations

from pathlib import Path
from typing import Mapping, Sequence

from .effort_policy import build_effort_policy_state
from .guidance import compact_guidance_for_model, load_project_guidance
from .skill_registry import select_harness_skills
from .task_memory import read_task_memory, write_task_memory


KERNEL_VERSION = "pstx-agent-kernel/v2"


def build_agentic_envelope(*,
                           run_id: object,
                           question: object = "",
                           capability_profiles: Sequence[object] = (),
                           playbook_plan: Mapping[str, object] | None = None,
                           tool_result_contracts: Sequence[Mapping[str, object]] = (),
                           root: str | Path | None = None,
                           include_skill_body: bool = True) -> dict:
    guidance = compact_guidance_for_model(load_project_guidance(root))
    skills = select_harness_skills(
        question=question,
        capability_profiles=capability_profiles,
        playbook_plan=playbook_plan,
        tool_result_contracts=tool_result_contracts,
        root=root,
        include_body=include_skill_body,
    )
    memory = read_task_memory(run_id, root=root)
    return {
        "version": KERNEL_VERSION,
        "guidance_summary": guidance,
        "selected_skills": skills,
        "task_memory_summary": memory,
    }


def update_agentic_effort(envelope: Mapping[str, object] | None = None, **kwargs) -> dict:
    payload = dict(envelope or {})
    effort = build_effort_policy_state(**kwargs)
    payload["effort_policy"] = effort
    return payload


def persist_agentic_task_memory(run_id: object,
                                payload: Mapping[str, object],
                                *,
                                root: str | Path | None = None) -> dict:
    return write_task_memory(run_id, payload, root=root)
