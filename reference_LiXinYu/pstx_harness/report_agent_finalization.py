# -*- coding: utf-8 -*-
"""Final answer normalization helpers for the report harness agent."""

from __future__ import annotations

from typing import List, Sequence, Tuple

from pstx_agent_runtime import (
    normalize_citations as runtime_normalize_citations,
    normalize_needs_user_input as runtime_normalize_needs_user_input,
    normalize_proposed_actions as runtime_normalize_proposed_actions,
    status_from_stopped_reason as runtime_status_from_stopped_reason,
)


def normalize_proposed_actions(raw: dict) -> List[dict]:
    return runtime_normalize_proposed_actions(raw)


def normalize_needs_user_input(raw: dict, evidence_nodes: Sequence[dict]) -> dict:
    return runtime_normalize_needs_user_input(raw, evidence_nodes)


def normalize_citations(raw: dict, evidence_nodes: Sequence[dict]) -> Tuple[List[dict], dict]:
    return runtime_normalize_citations(raw, evidence_nodes, fallback_when_empty=True)


def status_from_metadata(metadata: dict) -> str:
    return runtime_status_from_stopped_reason(metadata.get("stopped_reason"))
