# -*- coding: utf-8 -*-
"""Protocol parsing helpers for the report harness agent."""

from __future__ import annotations

import json
from typing import Optional

from pstx_agent_runtime import (
    AgentProtocolError,
    parse_agent_model_step,
)


def extract_balanced_json(text: str) -> Optional[dict]:
    content = str(text or "").strip()
    fence_start = content.find("```")
    if fence_start >= 0:
        content = content.replace("```json", "```").replace("```JSON", "```")
        parts = content.split("```")
        if len(parts) >= 3:
            content = parts[1].strip()
    start = content.find("{")
    if start < 0:
        return None
    depth = 0
    in_string = False
    escape = False
    for index in range(start, len(content)):
        char = content[index]
        if in_string:
            if escape:
                escape = False
            elif char == "\\":
                escape = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                try:
                    parsed = json.loads(content[start:index + 1])
                except json.JSONDecodeError:
                    return None
                return parsed if isinstance(parsed, dict) else None
    return None


def parse_model_step(answer: str, *, max_batch_calls: int) -> Optional[dict]:
    try:
        step = parse_agent_model_step(
            answer,
            max_batch_calls=max_batch_calls,
            allow_batch_tools=True,
            allow_needs_user_input=True,
        )
    except AgentProtocolError as exc:
        return {"type": "protocol_error", "error": str(exc)}
    if not step:
        return None
    return step.to_legacy_dict()
