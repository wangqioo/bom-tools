# -*- coding: utf-8 -*-
"""Observation compression and context budget helpers for PSTX agents."""

from __future__ import annotations

import json
from collections.abc import Mapping, Sequence
from typing import Callable

from .protocol import ObservationBundle


def json_char_count(value: object) -> int:
    try:
        return len(json.dumps(value, ensure_ascii=False, default=str))
    except (TypeError, ValueError):
        return len(str(value))


def evidence_ids_from_observations(observations: Sequence[Mapping[str, object]]) -> list[str]:
    evidence_ids: list[str] = []
    for item in observations or []:
        for evidence_id in item.get("evidence_node_ids", []) or []:
            text = str(evidence_id or "")
            if text and text not in evidence_ids:
                evidence_ids.append(text)
        for node in item.get("evidence_nodes", []) or []:
            if not isinstance(node, Mapping):
                continue
            text = str(node.get("id") or "")
            if text and text not in evidence_ids:
                evidence_ids.append(text)
    return evidence_ids


def _text(value: object, limit: int = 500) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _compact_value(value: object, *, depth: int = 0, text_limit: int = 500) -> object:
    if depth >= 4:
        return _text(value, min(text_limit, 220))
    if isinstance(value, Mapping):
        output = {}
        for index, (key, child) in enumerate(value.items()):
            if index >= 24:
                output["__truncated__"] = True
                output["__remaining__"] = len(value) - 24
                break
            output[str(key)] = _compact_value(child, depth=depth + 1, text_limit=min(text_limit, 260))
        return output
    if isinstance(value, (list, tuple)):
        items = [
            _compact_value(item, depth=depth + 1, text_limit=min(text_limit, 260))
            for item in list(value)[:12]
        ]
        if len(value) > 12:
            items.append({"__truncated__": True, "__remaining__": len(value) - 12})
        return items
    return _text(value, text_limit)


def _present(value: object) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value)
    if isinstance(value, Mapping):
        return bool(value)
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes)):
        return bool(value)
    return True


def _pick_first(mapping: Mapping[str, object], keys: Sequence[str]) -> object:
    for key in keys:
        if key in mapping and _present(mapping.get(key)):
            return mapping.get(key)
    return ""


def build_evidence_card(node: Mapping[str, object], *, include_payload_preview: bool = False) -> dict:
    """Build a compact, traceable evidence card from one evidence node."""

    source = node.get("source") if isinstance(node.get("source"), Mapping) else {}
    locator = node.get("locator") if isinstance(node.get("locator"), Mapping) else {}
    payload = node.get("payload_preview") if isinstance(node.get("payload_preview"), Mapping) else {}
    card = {
        "id": _text(node.get("id"), 160),
        "type": _text(node.get("type"), 120),
        "title": _text(node.get("title"), 180),
        "summary": _text(node.get("summary"), 360),
        "source": _compact_value(source, depth=1, text_limit=160),
        "locator": _compact_value(locator, depth=1, text_limit=220),
    }
    key_fields = {
        "table_id": ("table_id", "table", "表格"),
        "row_number": ("row_number", "row_index", "index", "行号"),
        "page": ("user_visible_page", "page", "真实页", "用户看到的真实页", "页面"),
        "refdes": ("refdes", "位号", "LOCATION"),
        "net": ("net", "network", "网络", "网络名"),
        "field_name": ("field", "field_name", "column", "变化字段", "字段"),
        "hq_no": ("hq_no", "HQ料号", "飞书HQ料号", "HQ编码"),
        "pi": ("pi", "PI"),
        "selection_order": ("selection_order", "选型顺序"),
        "match_reason": ("reason", "match_reason", "命中原因", "匹配原因"),
    }
    for field_name, aliases in key_fields.items():
        value = _pick_first(locator, aliases) or _pick_first(payload, aliases)
        if _present(value):
            card[field_name] = _text(value, 180)
    missing_fields = node.get("missing_fields")
    if isinstance(missing_fields, Sequence) and not isinstance(missing_fields, (str, bytes)):
        card["missing_fields"] = [_text(item, 120) for item in list(missing_fields)[:16] if _text(item, 120)]
    detail_tool = node.get("detail_tool")
    if isinstance(detail_tool, Mapping):
        card["detail_tool"] = _compact_value(detail_tool, depth=1, text_limit=220)
    if payload and include_payload_preview:
        card["payload_preview"] = _compact_value(payload, depth=1, text_limit=240)
    return {key: value for key, value in card.items() if _present(value)}


def _dedupe_tools(tools: Sequence[object]) -> list[dict]:
    output: list[dict] = []
    seen: set[str] = set()
    for item in tools or []:
        if not isinstance(item, Mapping):
            continue
        name = _text(item.get("name"), 120)
        args = item.get("args") if isinstance(item.get("args"), Mapping) else {}
        marker = json.dumps({"name": name, "args": dict(args)}, ensure_ascii=False, sort_keys=True, default=str)
        if not name or marker in seen:
            continue
        seen.add(marker)
        output.append({"name": name, "args": dict(args)})
    return output


def build_evidence_layers(*,
                          tool_name: str,
                          result: Mapping[str, object] | None,
                          evidence_nodes: Sequence[Mapping[str, object]] = (),
                          observation: Mapping[str, object] | None = None,
                          tool_result_contract: Mapping[str, object] | None = None,
                          include_raw_preview: bool = False,
                          raw_preview: object = None) -> dict:
    """Return the three-layer evidence envelope used by model, trace, and UI."""

    result = dict(result or {})
    observation = dict(observation or {})
    contract = dict(tool_result_contract or {})
    cards = [
        build_evidence_card(node, include_payload_preview=include_raw_preview)
        for node in evidence_nodes or []
        if isinstance(node, Mapping)
    ]
    evidence_ids = [card["id"] for card in cards if card.get("id")]
    detail_tools = _dedupe_tools(
        [card.get("detail_tool") for card in cards if isinstance(card.get("detail_tool"), Mapping)]
        + [contract.get("detail_tool") if isinstance(contract.get("detail_tool"), Mapping) else {}]
    )
    aggregation_tool = contract.get("aggregation_tool") if isinstance(contract.get("aggregation_tool"), Mapping) else {}
    raw_json_chars = json_char_count(result)
    summary_layer = {
        "tool": _text(tool_name, 120),
        "id": _text(observation.get("id") or result.get("id") or tool_name, 160),
        "title": _text(observation.get("title") or result.get("title") or tool_name, 180),
        "summary": _text(observation.get("summary") or result.get("summary") or "", 700),
        "completeness": _text(contract.get("completeness") or result.get("completeness") or "unknown", 80),
        "scope_summary": _text(contract.get("scope_summary") or "", 500),
        "evidence_count": len(cards),
        "evidence_ids": evidence_ids[:80],
        "recommended_next_tools": [
            _text(item, 120)
            for item in list(contract.get("recommended_next_tools") or [])[:12]
            if _text(item, 120)
        ],
        "detail_tools": detail_tools[:12],
    }
    if aggregation_tool:
        summary_layer["aggregation_tool"] = _compact_value(aggregation_tool, depth=1, text_limit=220)
    raw_layer = {
        "stored_in_trace": True,
        "available_to_model": False,
        "result_json_chars": raw_json_chars,
        "result_keys": sorted(str(key) for key in result.keys())[:80],
        "detail_tools": detail_tools[:12],
        "model_rule": "模型默认只读摘要层和证据卡层；高风险或不确定项必须通过 detail_tool/推荐工具二次取证。",
    }
    if include_raw_preview:
        raw_layer["preview"] = _compact_value(raw_preview if raw_preview is not None else result, depth=0, text_limit=500)
    else:
        raw_layer["preview_omitted_for_model"] = True
    return {
        "version": "three-layer-evidence/v1",
        "summary_layer": {key: value for key, value in summary_layer.items() if _present(value)},
        "evidence_card_layer": cards,
        "raw_layer": raw_layer,
    }


def observation_bundle_summary(observations: Sequence[Mapping[str, object]],
                               *,
                               bundle_id: str,
                               json_budget_chars: int) -> dict:
    bundle = ObservationBundle.from_observations(
        list(observations or []),
        bundle_id=bundle_id,
        max_items=max(1, len(observations or [])),
        max_chars=json_budget_chars,
    ).to_dict()
    bundle.pop("observations", None)
    bundle["model_observation_json_chars"] = json_char_count(observations or [])
    return bundle


def build_context_budget_summary(source_observations: Sequence[Mapping[str, object]],
                                 model_observations: Sequence[Mapping[str, object]],
                                 *,
                                 json_budget_chars: int,
                                 truncated_note: str,
                                 ok_note: str,
                                 include_observation_bundle: bool = False,
                                 bundle_id: str = "agent-observation-bundle") -> dict:
    sent_items = list(model_observations or [])
    source_count = len(source_observations or [])
    sent_count = len(sent_items)
    omitted_count = max(0, source_count - sent_count)
    summary_omitted_count = 0
    result_preview_omitted_count = 0
    omitted_evidence_node_count = 0
    raw_layer_available_count = 0
    detail_tool_count = 0
    for item in sent_items:
        summary_omitted_count += int(item.get("omitted_observation_count") or 0)
        omitted_evidence_node_count += int(item.get("omitted_evidence_node_count") or 0)
        if item.get("result_preview_omitted"):
            result_preview_omitted_count += 1
        layers = item.get("evidence_layers") if isinstance(item.get("evidence_layers"), Mapping) else {}
        raw_layer = layers.get("raw_layer") if isinstance(layers.get("raw_layer"), Mapping) else {}
        if raw_layer.get("stored_in_trace"):
            raw_layer_available_count += 1
        summary_layer = layers.get("summary_layer") if isinstance(layers.get("summary_layer"), Mapping) else {}
        detail_tool_count += len(summary_layer.get("detail_tools") or [])
    omitted_count = max(omitted_count, summary_omitted_count)
    evidence_ids = evidence_ids_from_observations(sent_items)
    json_chars = json_char_count(sent_items)
    truncated = (
        omitted_count > 0
        or result_preview_omitted_count > 0
        or omitted_evidence_node_count > 0
        or any(bool(item.get("truncated_for_model")) for item in sent_items)
        or json_chars >= json_budget_chars
    )
    payload = {
        "json_budget_chars": json_budget_chars,
        "model_observation_json_chars": json_chars,
        "source_observation_count": source_count,
        "sent_observation_count": sent_count,
        "omitted_observation_count": omitted_count,
        "result_preview_omitted_count": result_preview_omitted_count,
        "omitted_evidence_node_count": omitted_evidence_node_count,
        "sent_evidence_node_count": len(evidence_ids),
        "raw_layer_available_count": raw_layer_available_count,
        "detail_tool_count": detail_tool_count,
        "truncated": truncated,
        "notes": truncated_note if truncated else ok_note,
        "evidence_layer_policy": {
            "summary_layer": "默认发送给模型，用于快速规划和选择工具。",
            "evidence_card_layer": "发送 evidence id、来源、定位、关键字段和 detail_tool。",
            "raw_layer": "完整工具结果保留在本地 trace/store；模型需要细节时必须再次 tool-call。",
        },
    }
    if include_observation_bundle:
        payload["observation_bundle"] = observation_bundle_summary(
            sent_items,
            bundle_id=bundle_id,
            json_budget_chars=json_budget_chars,
        )
    return payload


def fit_items_to_json_budget(items: Sequence[dict],
                             *,
                             json_budget_chars: int,
                             compact_item: Callable[[dict], dict] | None = None,
                             fallback_limit: int = 1) -> list[dict]:
    output = [dict(item) for item in items or []]
    if json_char_count(output) <= json_budget_chars:
        return output
    if compact_item:
        output = [compact_item(item) for item in output]
        if json_char_count(output) <= json_budget_chars:
            return output
    while output and json_char_count(output) > json_budget_chars:
        output.pop(0)
    if output:
        return output
    return [compact_item(dict(item)) if compact_item else dict(item) for item in list(items or [])[-fallback_limit:]]
