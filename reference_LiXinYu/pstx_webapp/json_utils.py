"""Small JSON/text helpers shared by Web view models and routes."""

from __future__ import annotations

import json
from typing import Any


def json_fingerprint(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)


def compact_value(value: Any, limit: int = 180) -> str:
    if isinstance(value, (dict, list, tuple)):
        text = json_fingerprint(value)
    else:
        text = str(value if value is not None else "")
    return text if len(text) <= limit else text[:limit - 1] + "…"
