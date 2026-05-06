# -*- coding: utf-8 -*-
"""Project-scoped evidence memory cards for PSTX agent continuity."""

from __future__ import annotations

import json
from collections.abc import Mapping, Sequence


PROJECT_EVIDENCE_MEMORY_VERSION = "agent-project-evidence-memory/v1"
MEMORY_PREFETCH_PLAN_VERSION = "agent-memory-prefetch-plan/v1"
_RECALL_CUES = (
    "继续",
    "接着",
    "刚才",
    "上次",
    "上一轮",
    "之前",
    "那个",
    "这些证据",
    "证据",
    "引用",
    "evidence",
    "memory",
)


def _text(value: object, limit: int = 500) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _compact_value(value: object, *, depth: int = 0, limit: int = 260) -> object:
    if depth >= 2:
        return _text(value, limit)
    if isinstance(value, Mapping):
        return {
            _text(key, 80): _compact_value(item, depth=depth + 1, limit=min(limit, 180))
            for key, item in list(value.items())[:10]
        }
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray)):
        items = [_compact_value(item, depth=depth + 1, limit=min(limit, 180)) for item in list(value)[:8]]
        if len(value) > 8:
            items.append({"omitted_count": len(value) - 8})
        return items
    return _text(value, limit)


def _compact_detail_tool(value: object, *, limit: int = 220) -> dict:
    """Compact a detail tool while preserving executable structured args."""

    if not isinstance(value, Mapping):
        return {}

    def compact_arg(item: object, *, depth: int = 0) -> object:
        if depth >= 3:
            return _text(item, min(limit, 180))
        if isinstance(item, Mapping):
            return {
                _text(key, 80): compact_arg(value, depth=depth + 1)
                for key, value in list(item.items())[:12]
            }
        if isinstance(item, Sequence) and not isinstance(item, (str, bytes, bytearray)):
            values = [compact_arg(value, depth=depth + 1) for value in list(item)[:12]]
            if len(item) > 12:
                values.append({"omitted_count": len(item) - 12})
            return values
        if isinstance(item, (int, float, bool)) or item is None:
            return item
        return _text(item, min(limit, 180))

    name = _text(value.get("name") or value.get("tool"), 120)
    args = value.get("args") if isinstance(value.get("args"), Mapping) else {}
    payload: dict[str, object] = {}
    if name:
        payload["name"] = name
    payload["args"] = compact_arg(args) if args else {}
    return payload


def _node_to_card(node: Mapping[str, object], *, result: Mapping[str, object]) -> dict:
    evidence_id = _text(node.get("id"), 160)
    if not evidence_id:
        return {}
    agent_run_id = _text(result.get("agent_run_id"), 120)
    memory_id = f"{agent_run_id}:{evidence_id}" if agent_run_id else evidence_id
    citations = [
        str(item.get("id") or "")
        for item in result.get("citations") or []
        if isinstance(item, Mapping) and item.get("valid") is not False
    ]
    card = {
        "version": PROJECT_EVIDENCE_MEMORY_VERSION,
        "id": evidence_id,
        "memory_id": memory_id,
        "type": _text(node.get("type"), 100),
        "title": _text(node.get("title"), 180),
        "summary": _text(node.get("summary"), 500),
        "source": _compact_value(node.get("source") or {}, depth=1, limit=220),
        "locator": _compact_value(node.get("locator") or {}, depth=1, limit=220),
        "detail_tool": _compact_detail_tool(node.get("detail_tool") or {}, limit=220),
        "missing_fields": _compact_value(node.get("missing_fields") or [], depth=1, limit=120),
        "agent_run_id": agent_run_id,
        "profile": _text(result.get("profile"), 80),
        "run_status": _text(result.get("status"), 80),
        "created_at": _text(result.get("finished_at") or result.get("started_at"), 80),
        "cited": evidence_id in citations,
    }
    preview = node.get("payload_preview")
    if preview:
        card["payload_preview"] = _compact_value(preview, depth=1, limit=220)
    return card


def _memory_key(card: Mapping[str, object]) -> str:
    memory_id = _text(card.get("memory_id"), 220)
    if memory_id:
        return memory_id
    evidence_id = _text(card.get("id"), 160)
    agent_run_id = _text(card.get("agent_run_id"), 120)
    return f"{agent_run_id}:{evidence_id}" if agent_run_id else evidence_id


def _iter_result_evidence_nodes(result: Mapping[str, object]):
    for node in result.get("final_evidence") or []:
        if isinstance(node, Mapping):
            yield node
    for observation in result.get("observations") or []:
        if not isinstance(observation, Mapping):
            continue
        for node in observation.get("evidence_nodes") or []:
            if isinstance(node, Mapping):
                yield node


def build_project_evidence_memory(project_context: Mapping[str, object] | None,
                                  result: Mapping[str, object] | None,
                                  *,
                                  max_cards: int = 120) -> list[dict]:
    """Merge result evidence into a compact project memory card list."""

    project_context = project_context or {}
    result = result or {}
    merged: dict[str, dict] = {}
    order: list[str] = []
    for card in project_context.get("evidence_memory_cards") or []:
        if not isinstance(card, Mapping):
            continue
        key = _memory_key(card)
        if not key:
            continue
        merged[key] = dict(card)
        order.append(key)
    for node in _iter_result_evidence_nodes(result):
        card = _node_to_card(node, result=result)
        key = _memory_key(card)
        if not key:
            continue
        if key not in merged:
            order.append(key)
        previous = merged.get(key) or {}
        first_seen = previous.get("first_seen_agent_run_id") or previous.get("agent_run_id") or card.get("agent_run_id")
        card["first_seen_agent_run_id"] = _text(first_seen, 120)
        merged[key] = card
    deduped = []
    seen = set()
    for key in order:
        if key in seen or key not in merged:
            continue
        seen.add(key)
        deduped.append(merged[key])
    return deduped[-max(1, int(max_cards or 120)):]


def compact_project_evidence_memory(cards: Sequence[Mapping[str, object]] | None,
                                    *,
                                    limit: int = 16) -> list[dict]:
    compact: list[dict] = []
    for card in list(cards or [])[-max(1, int(limit or 16)):]:
        if not isinstance(card, Mapping):
            continue
        compact.append({
            "id": _text(card.get("id"), 160),
            "type": _text(card.get("type"), 100),
            "title": _text(card.get("title"), 160),
            "summary": _text(card.get("summary"), 320),
            "locator": _compact_value(card.get("locator") or {}, depth=1, limit=160),
            "detail_tool": _compact_detail_tool(card.get("detail_tool") or {}, limit=160),
            "agent_run_id": _text(card.get("agent_run_id"), 120),
            "cited": bool(card.get("cited")),
        })
    return compact


def _search_blob(card: Mapping[str, object]) -> str:
    try:
        return json.dumps(card, ensure_ascii=False, sort_keys=True, default=str).lower()
    except (TypeError, ValueError):
        return str(card).lower()


def search_project_evidence_memory(cards: Sequence[Mapping[str, object]] | None,
                                   *,
                                   query: object = "",
                                   evidence_type: object = "",
                                   limit: int = 20,
                                   offset: int = 0) -> dict:
    query_text = _text(query, 200).lower()
    type_text = _text(evidence_type, 100).lower()
    matches: list[dict] = []
    for card in reversed([item for item in cards or [] if isinstance(item, Mapping)]):
        if type_text and _text(card.get("type"), 100).lower() != type_text:
            continue
        blob = _search_blob(card)
        if query_text and query_text not in blob:
            continue
        match = dict(card)
        match["match_reason"] = "type_filter" if type_text and not query_text else ("keyword" if query_text else "recent")
        matches.append(match)
    start = max(0, int(offset or 0))
    capped_limit = max(1, min(int(limit or 20), 100))
    rows = matches[start:start + capped_limit]
    return {
        "version": PROJECT_EVIDENCE_MEMORY_VERSION,
        "query": _text(query, 200),
        "evidence_type": _text(evidence_type, 100),
        "total_matches": len(matches),
        "offset": start,
        "limit": capped_limit,
        "has_more": start + capped_limit < len(matches),
        "cards": rows,
        "summary": f"项目证据记忆命中 {len(matches)} 条，返回 {len(rows)} 条。",
    }


def get_project_evidence_memory_card(cards: Sequence[Mapping[str, object]] | None,
                                     evidence_id: object) -> dict:
    target = _text(evidence_id, 160)
    for card in reversed([item for item in cards or [] if isinstance(item, Mapping)]):
        if _text(card.get("id"), 160) == target or _text(card.get("memory_id"), 220) == target:
            return {
                "version": PROJECT_EVIDENCE_MEMORY_VERSION,
                "found": True,
                "card": dict(card),
                "summary": f"已读取项目证据记忆：{target}",
            }
    return {
        "version": PROJECT_EVIDENCE_MEMORY_VERSION,
        "found": False,
        "card": {},
        "summary": f"未找到项目证据记忆：{target}",
    }


def _question_terms(question: object) -> list[str]:
    import re

    text = str(question or "")
    terms: list[str] = []
    for pattern in (
        r"\bHQ[0-9A-Z]{4,}\b",
        r"\b[A-Z]{1,4}\d+[A-Z0-9_]*\b",
        r"\bev-[A-Za-z0-9_.:-]+\b",
    ):
        for item in re.findall(pattern, text, flags=re.IGNORECASE):
            value = _text(item, 120)
            if value and value not in terms:
                terms.append(value)
    return terms[:6]


def _has_recall_cue(question: object) -> bool:
    text = str(question or "").lower()
    return any(cue.lower() in text for cue in _RECALL_CUES)


def select_project_memory_prefetch_tool_calls(question: object,
                                              project_context: Mapping[str, object] | None,
                                              *,
                                              allowed_tools: set[str] | Sequence[str],
                                              max_calls: int = 1,
                                              remaining_tool_calls: int | None = None,
                                              enabled: bool = True) -> dict:
    """Select a tiny project evidence-memory recall before the model step."""

    project_context = project_context or {}
    cards = [item for item in project_context.get("evidence_memory_cards") or [] if isinstance(item, Mapping)]
    allowed = set(allowed_tools or [])
    budget = max(0, int(max_calls or 0))
    if remaining_tool_calls is not None:
        budget = min(budget, max(0, int(remaining_tool_calls or 0)))
    base = {
        "version": MEMORY_PREFETCH_PLAN_VERSION,
        "enabled": bool(enabled),
        "card_count": len(cards),
        "selected_count": 0,
        "tool_calls": [],
        "skipped": [],
    }
    if not enabled:
        return {**base, "skipped": [{"reason": "memory prefetch disabled"}]}
    if not cards:
        return {**base, "skipped": [{"reason": "no evidence memory cards"}]}
    if budget <= 0:
        return {**base, "skipped": [{"reason": "no remaining tool-call budget"}]}

    question_text = str(question or "")
    terms = _question_terms(question_text)
    exact_id = ""
    lowered_question = question_text.lower()
    for card in reversed(cards):
        evidence_id = _text(card.get("id"), 160)
        if evidence_id and evidence_id.lower() in lowered_question:
            exact_id = evidence_id
            break
    if exact_id and "get_project_memory_evidence" in allowed:
        return {
            **base,
            "selected_count": 1,
            "tool_calls": [{
                "name": "get_project_memory_evidence",
                "args": {"evidence_id": exact_id},
                "reason": "用户问题直接提到历史 evidence id，本地 runtime 先读取该证据记忆卡。",
                "source": "runtime_memory_prefetch",
            }],
        }

    query = ""
    for term in terms:
        if search_project_evidence_memory(cards, query=term, limit=1).get("total_matches"):
            query = term
            break
    if not query and not _has_recall_cue(question_text):
        return {**base, "skipped": [{"reason": "question does not reference prior evidence memory"}]}
    if "list_project_memory_evidence" not in allowed:
        return {**base, "skipped": [{"reason": "memory search tool not allowed"}]}
    return {
        **base,
        "selected_count": 1,
        "tool_calls": [{
            "name": "list_project_memory_evidence",
            "args": {"query": query, "limit": 5},
            "reason": "用户问题疑似延续上一轮任务，本地 runtime 先召回项目证据记忆卡。",
            "source": "runtime_memory_prefetch",
        }],
    }
