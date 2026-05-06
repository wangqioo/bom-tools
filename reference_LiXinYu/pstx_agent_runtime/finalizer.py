# -*- coding: utf-8 -*-
"""Final answer, citation, and user-input normalization for PSTX agents."""

from __future__ import annotations

from dataclasses import dataclass, field
import json
import re
from collections.abc import Mapping, Sequence

EARLY_STOP_MARKERS = (
    "无法回答",
    "无法判断",
    "无法确认",
    "无法获取",
    "无法访问",
    "无法完成",
    "无法处理",
    "不能回答",
    "不能判断",
    "不能确认",
    "不能完成",
    "没有足够",
    "信息不足",
    "证据不足",
    "缺少必要",
    "无从判断",
    "不知道",
    "不清楚",
    "不确定",
    "需要人工",
    "人工确认",
    "请人工",
    "暂时无法",
    "I cannot",
    "I can't",
    "not enough information",
    "insufficient information",
)


def _preview(value: object, limit: int = 500) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[: max(0, limit - 1)] + "…"


def _preview_value(value: object, *, depth: int = 0, text_limit: int = 220):
    if depth >= 3:
        return _preview(value, text_limit)
    if isinstance(value, Mapping):
        return {str(key)[:80]: _preview_value(item, depth=depth + 1, text_limit=text_limit) for key, item in list(value.items())[:12]}
    if isinstance(value, list):
        return [_preview_value(item, depth=depth + 1, text_limit=text_limit) for item in value[:12]]
    return _preview(value, text_limit)


@dataclass(frozen=True)
class AgentFinalizationResult:
    answer: str = ""
    status: str = "completed"
    stopped_reason: str = ""
    citations: tuple[dict, ...] = ()
    proposed_actions: tuple[dict, ...] = ()
    invalid_citation_count: int = 0
    needs_user_input: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "answer": self.answer,
            "status": self.status,
            "stopped_reason": self.stopped_reason,
            "citations": [dict(item) for item in self.citations],
            "proposed_actions": [dict(item) for item in self.proposed_actions],
            "invalid_citation_count": self.invalid_citation_count,
            "needs_user_input": dict(self.needs_user_input or {}),
        }


def list_of_strings(value: object, *, limit: int = 16, text_limit: int = 120) -> list[str]:
    if value is None:
        return []
    items = value if isinstance(value, list) else [value]
    result: list[str] = []
    for item in items[:limit]:
        text = _preview(item, text_limit).strip()
        if text:
            result.append(text)
    return result


def normalize_needs_user_input(raw: Mapping[str, object],
                               evidence_nodes: Sequence[Mapping[str, object]]) -> dict:
    payload = raw.get("needs_user_input") if isinstance(raw.get("needs_user_input"), Mapping) else raw
    reason = _preview(payload.get("reason") or "当前证据不足，需要用户补充后再继续。", 360)
    top_missing_fields = list_of_strings(payload.get("missing_fields"), limit=16, text_limit=80)
    evidence_ids = {str(node.get("id")) for node in evidence_nodes if node.get("id")}
    related_ids = [
        {"id": evidence_id, "valid": evidence_id in evidence_ids}
        for evidence_id in list_of_strings(payload.get("related_evidence_ids"), limit=24, text_limit=140)
    ]
    questions = []
    raw_questions = payload.get("questions") if isinstance(payload.get("questions"), list) else []
    if not raw_questions:
        raw_questions = [{
            "question_id": "missing-context-1",
            "question": reason,
            "missing_fields": top_missing_fields,
        }]
    for index, item in enumerate(raw_questions[:12], start=1):
        if isinstance(item, Mapping):
            question_id = str(item.get("question_id") or item.get("id") or f"q-{index}").strip()[:120]
            question = _preview(item.get("question") or item.get("prompt") or item.get("title") or reason, 500)
            applies_to = item.get("applies_to") if isinstance(item.get("applies_to"), Mapping) else {}
            missing_fields = list_of_strings(item.get("missing_fields") or top_missing_fields, limit=12, text_limit=80)
            item_related_ids = list_of_strings(item.get("related_evidence_ids"), limit=12, text_limit=140)
        else:
            question_id = f"q-{index}"
            question = _preview(item, 500)
            applies_to = {}
            missing_fields = list(top_missing_fields)
            item_related_ids = []
        questions.append({
            "question_id": question_id,
            "question": question,
            "applies_to": _preview_value(applies_to, depth=1, text_limit=160),
            "missing_fields": missing_fields,
            "related_evidence_ids": item_related_ids,
            "required": True,
        })
    return {
        "reason": reason,
        "missing_fields": top_missing_fields,
        "related_evidence_ids": related_ids,
        "questions": questions,
    }


def citation_items(raw: Mapping[str, object]) -> list[dict]:
    items = raw.get("citations")
    if items is None:
        items = raw.get("evidence")
    result = []
    if isinstance(items, list):
        for item in items[:24]:
            if isinstance(item, Mapping):
                evidence_id = str(item.get("id") or item.get("evidence_id") or "").strip()
                if evidence_id:
                    result.append({"id": evidence_id, "note": str(item.get("note") or item.get("reason") or "")})
            else:
                evidence_id = str(item or "").strip()
                if evidence_id:
                    result.append({"id": evidence_id, "note": ""})
    return result


def normalize_citations(raw: Mapping[str, object],
                        evidence_nodes: Sequence[Mapping[str, object]],
                        *,
                        fallback_when_empty: bool = False,
                        fallback_note: str = "模型未给出有效证据引用，本地自动关联最近证据节点。") -> tuple[list[dict], dict]:
    by_id = {str(node.get("id")): node for node in evidence_nodes if node.get("id")}
    citations = []
    seen = set()
    invalid_count = 0
    for item in citation_items(raw):
        evidence_id = item["id"]
        if evidence_id in seen:
            continue
        seen.add(evidence_id)
        node = by_id.get(evidence_id)
        if node:
            citations.append({
                "id": evidence_id,
                "valid": True,
                "note": _preview(item.get("note", ""), 180),
                "title": node.get("title", ""),
                "type": node.get("type", ""),
                "locator": node.get("locator", {}),
                "source": node.get("source", {}),
            })
        else:
            invalid_count += 1
            citations.append({
                "id": evidence_id,
                "valid": False,
                "note": _preview(item.get("note", ""), 180),
                "title": "引用不存在",
                "type": "invalid",
                "locator": {},
                "source": {},
            })
    fallback_count = 0
    if fallback_when_empty and not any(item.get("valid") for item in citations) and evidence_nodes:
        for node in evidence_nodes[:3]:
            evidence_id = str(node.get("id") or "")
            if not evidence_id or evidence_id in seen:
                continue
            citations.append({
                "id": evidence_id,
                "valid": True,
                "fallback": True,
                "note": fallback_note,
                "title": node.get("title", ""),
                "type": node.get("type", ""),
                "locator": node.get("locator", {}),
                "source": node.get("source", {}),
            })
            fallback_count += 1
    return citations, {
        "citation_count": len(citations),
        "valid_citation_count": sum(1 for item in citations if item.get("valid")),
        "invalid_citation_count": invalid_count,
        "fallback_citation_count": fallback_count,
    }


def normalize_proposed_actions(raw: Mapping[str, object]) -> list[dict]:
    actions = raw.get("proposed_actions")
    if actions is None:
        actions = raw.get("actions")
    normalized = []
    if isinstance(actions, list):
        for index, item in enumerate(actions[:12], start=1):
            if isinstance(item, Mapping):
                normalized.append({
                    "id": str(item.get("id") or f"action-{index}"),
                    "title": _preview(item.get("title") or item.get("action") or item.get("summary") or f"建议 {index}", 160),
                    "reason": _preview(item.get("reason") or item.get("body") or "", 280),
                    "priority": _preview(item.get("priority") or "manual_review", 80),
                })
            else:
                normalized.append({
                    "id": f"action-{index}",
                    "title": _preview(item, 160),
                    "reason": "",
                    "priority": "manual_review",
                })
    return normalized


def status_from_stopped_reason(stopped_reason: object) -> str:
    stopped = str(stopped_reason or "")
    if stopped == "needs_user_input":
        return "waiting_for_user"
    if stopped == "tool_error":
        return "tool_error"
    if stopped in {"model_error", "invalid_model_json", "protocol_error"}:
        return "model_error"
    if stopped in {"max_tool_calls", "max_steps"}:
        return "limited"
    if stopped in {"empty_answer"}:
        return "incomplete"
    return "completed"


def _recommended_tools_from_playbook(playbook_plan: Mapping[str, object] | None) -> list[str]:
    if not isinstance(playbook_plan, Mapping):
        return []
    return list_of_strings(playbook_plan.get("recommended_first_tools"), limit=12, text_limit=100)


def _compact_args(args: Mapping[str, object] | None, *, limit: int = 220) -> str:
    if not isinstance(args, Mapping) or not args:
        return "{}"
    try:
        text = json.dumps(dict(args), ensure_ascii=False, sort_keys=True, default=str)
    except (TypeError, ValueError):
        text = str(dict(args))
    return _preview(text, limit)


def _seeded_tool_hints_from_playbook(playbook_plan: Mapping[str, object] | None) -> list[str]:
    if not isinstance(playbook_plan, Mapping):
        return []
    result: list[str] = []
    for item in playbook_plan.get("seeded_tool_calls") or []:
        if not isinstance(item, Mapping):
            continue
        name = str(item.get("name") or item.get("tool") or "").strip()
        if not name:
            continue
        hint = f"{name}({_compact_args(item.get('args') if isinstance(item.get('args'), Mapping) else {})})"
        if hint not in result:
            result.append(hint)
        if len(result) >= 4:
            break
    return result


def _seeded_tool_hints_from_task_ledger(task_ledger: Mapping[str, object] | None) -> list[str]:
    if not isinstance(task_ledger, Mapping):
        return []
    result: list[str] = []
    for action in task_ledger.get("next_actions") or []:
        if not isinstance(action, Mapping):
            continue
        if str(action.get("type") or "") not in {"", "tool_call"}:
            continue
        args = action.get("args") if isinstance(action.get("args"), Mapping) else {}
        if not args:
            continue
        name = str(action.get("tool") or "").strip()
        if not name:
            continue
        hint = f"{name}({_compact_args(args)})"
        if hint not in result:
            result.append(hint)
        if len(result) >= 4:
            break
    return result


def _seeded_retry_suffix(*,
                         playbook_plan: Mapping[str, object] | None,
                         task_ledger: Mapping[str, object] | None) -> str:
    hints = []
    for hint in [*_seeded_tool_hints_from_playbook(playbook_plan), *_seeded_tool_hints_from_task_ledger(task_ledger)]:
        if hint not in hints:
            hints.append(hint)
    if not hints:
        return ""
    return f" 可直接使用本地已生成的带参工具种子：{'; '.join(hints[:3])}。"


def _recommended_tools_from_contracts(tool_result_contracts: Sequence[Mapping[str, object]] | None) -> list[str]:
    result: list[str] = []
    for contract in list(tool_result_contracts or [])[-6:]:
        if not isinstance(contract, Mapping):
            continue
        for name in list_of_strings(contract.get("recommended_next_tools"), limit=8, text_limit=100):
            if name not in result:
                result.append(name)
        detail_tool = contract.get("detail_tool")
        if isinstance(detail_tool, Mapping):
            name = str(detail_tool.get("name") or "").strip()
            if name and name not in result:
                result.append(name)
        aggregation_tool = contract.get("aggregation_tool")
        if isinstance(aggregation_tool, Mapping):
            name = str(aggregation_tool.get("name") or "").strip()
            if name and name not in result:
                result.append(name)
    return result


def _recommended_tools_from_task_ledger(task_ledger: Mapping[str, object] | None) -> list[str]:
    return [
        action["tool"]
        for action in _recommended_tool_actions_from_task_ledger(task_ledger)
        if action.get("tool")
    ]


def _recommended_tool_actions_from_task_ledger(task_ledger: Mapping[str, object] | None) -> list[dict]:
    if not isinstance(task_ledger, Mapping):
        return []
    result: list[dict] = []
    for action in task_ledger.get("next_actions") or []:
        if not isinstance(action, Mapping):
            continue
        if str(action.get("type") or "") not in {"", "tool_call"}:
            continue
        name = str(action.get("tool") or "").strip()
        if not name or any(item.get("tool") == name and item.get("args") == action.get("args") for item in result):
            continue
        payload = {
            "tool": name,
            "reason": _preview(action.get("reason") or action.get("title") or "", 260),
        }
        if isinstance(action.get("args"), Mapping):
            payload["args"] = dict(action.get("args") or {})
        result.append(payload)
    return result


def _task_ledger_has_open_tool_work(task_ledger: Mapping[str, object] | None) -> bool:
    if not isinstance(task_ledger, Mapping):
        return False
    if _recommended_tools_from_task_ledger(task_ledger):
        return True
    progress = task_ledger.get("progress") if isinstance(task_ledger.get("progress"), Mapping) else {}
    return int(progress.get("in_progress") or 0) > 0


def _has_incomplete_contract(tool_result_contracts: Sequence[Mapping[str, object]] | None) -> bool:
    incomplete = {"preview", "partial", "truncated"}
    return any(
        str(contract.get("completeness") or "").lower() in incomplete
        for contract in tool_result_contracts or []
        if isinstance(contract, Mapping)
    )


def _repair_action(action_type: str,
                   title: str,
                   *,
                   severity: str = "warn",
                   tool: str = "",
                   args: Mapping[str, object] | None = None,
                   reason: str = "",
                   source: str = "",
                   priority: int = 50) -> dict:
    payload = {
        "type": _preview(action_type, 80),
        "title": _preview(title, 180),
        "severity": _preview(severity, 40),
        "reason": _preview(reason, 360),
        "source": _preview(source, 120),
        "priority": int(priority),
    }
    if tool:
        payload["tool"] = _preview(tool, 120)
    if args is not None:
        payload["args"] = dict(args)
    return payload


_DETAIL_REQUIRED_SOURCE_TOOLS = {
    "search_datasheet_chunks",
    "batch_search_datasheet_chunks",
}

_DATASHEET_DETAIL_SOURCE_TOOLS = {
    "get_datasheet_chunk",
    "get_datasheet_page_excerpt",
    "get_datasheet_excerpt",
}

_QUANTITATIVE_UNIT_PATTERN = re.compile(
    r"(?<![A-Z0-9_])"
    r"\d+(?:\.\d+)?\s*"
    r"(?:mV|V|kV|uA|µA|mA|A|nA|pF|nF|uF|µF|F|mΩ|Ω|ohm|Ohm|kΩ|MΩ|Hz|kHz|MHz|GHz|mW|W|°C|℃|%)"
    r"\b",
    flags=re.IGNORECASE,
)

_QUANTITATIVE_SPEC_KEYWORDS = (
    "absolute maximum",
    "recommended operating",
    "electrical characteristics",
    "operating condition",
    "工作电压",
    "推荐工作",
    "电气特性",
    "电气极限",
    "绝对最大",
    "耐压",
    "额定",
    "阈值",
    "电流",
    "电压",
    "温度",
)


def _has_quantitative_spec_claim(answer: object) -> bool:
    text = str(answer or "")
    if not text:
        return False
    if _QUANTITATIVE_UNIT_PATTERN.search(text):
        return True
    lowered = text.lower()
    has_spec_keyword = any(keyword.lower() in lowered for keyword in _QUANTITATIVE_SPEC_KEYWORDS)
    has_number = bool(re.search(r"\d+(?:\.\d+)?", text))
    return has_spec_keyword and has_number


def _node_by_id(evidence_nodes: Sequence[Mapping[str, object]] | None) -> dict[str, Mapping[str, object]]:
    return {
        str(node.get("id") or ""): node
        for node in evidence_nodes or []
        if isinstance(node, Mapping) and node.get("id")
    }


def _valid_citation_nodes(citations: Sequence[Mapping[str, object]] | None,
                          evidence_nodes: Sequence[Mapping[str, object]] | None) -> list[Mapping[str, object]]:
    by_id = _node_by_id(evidence_nodes)
    nodes: list[Mapping[str, object]] = []
    for citation in citations or []:
        if not isinstance(citation, Mapping) or not citation.get("valid"):
            continue
        node = by_id.get(str(citation.get("id") or ""))
        if isinstance(node, Mapping):
            nodes.append(node)
    return nodes


def _node_search_blob(node: Mapping[str, object]) -> str:
    try:
        return json.dumps(node, ensure_ascii=False, sort_keys=True, default=str).upper()
    except (TypeError, ValueError):
        return str(node).upper()


def _contains_target_token(text: object, normalized: object) -> bool:
    target = _preview(normalized, 160).upper()
    blob = str(text or "").upper()
    if not target or not blob:
        return False
    if re.fullmatch(r"[A-Z0-9_]+", target):
        return re.search(rf"(?<![A-Z0-9_]){re.escape(target)}(?![A-Z0-9_])", blob) is not None
    return target in blob


def _is_datasheet_detail_node(node: Mapping[str, object]) -> bool:
    if str(node.get("type") or "") not in {"datasheet_chunk", "datasheet_excerpt"}:
        return False
    source = node.get("source") if isinstance(node.get("source"), Mapping) else {}
    source_tool = str(source.get("tool") or "").strip()
    return source_tool in _DATASHEET_DETAIL_SOURCE_TOOLS


def _detail_actions_from_nodes(nodes: Sequence[Mapping[str, object]],
                               *,
                               source: str,
                               reason: str,
                               limit: int = 4,
                               priority: int = 18) -> list[dict]:
    actions: list[dict] = []
    seen = set()
    for node in nodes or []:
        if not isinstance(node, Mapping):
            continue
        detail = node.get("detail_tool") if isinstance(node.get("detail_tool"), Mapping) else {}
        tool = str(detail.get("name") or "").strip()
        args = detail.get("args") if isinstance(detail.get("args"), Mapping) else {}
        if not tool or not args:
            continue
        key = (tool, str(dict(args)))
        if key in seen:
            continue
        seen.add(key)
        actions.append(_repair_action(
            "tool_call",
            f"打开规格书原文详情：{node.get('title') or node.get('id')}",
            severity="warn",
            tool=tool,
            args=dict(args),
            reason=reason,
            source=source,
            priority=priority,
        ))
        if len(actions) >= max(0, int(limit or 0)):
            break
    return actions


def _target_repair_actions_from_targets(targets: Sequence[Mapping[str, object]],
                                        *,
                                        source: str,
                                        reason: str,
                                        title_prefix: str,
                                        priority: int = 16,
                                        limit: int = 4) -> list[dict]:
    grouped: dict[tuple[str, str], list[str]] = {}
    for target in targets or []:
        if not isinstance(target, Mapping):
            continue
        tool = str(target.get("repair_tool") or "").strip()
        arg_key = str(target.get("repair_arg_key") or "").strip()
        value = _preview(target.get("value"), 160)
        if not tool or not arg_key or not value:
            continue
        grouped.setdefault((tool, arg_key), []).append(value)

    actions: list[dict] = []
    for (tool, arg_key), values in grouped.items():
        unique_values = []
        for value in values:
            if value and value not in unique_values:
                unique_values.append(value)
            if len(unique_values) >= 20:
                break
        if not unique_values:
            continue
        if arg_key in {"queries", "refdes_list"}:
            args = {arg_key: unique_values}
            if arg_key == "queries":
                args["limit_per_query"] = 10
        else:
            args = {arg_key: unique_values[0]}
        actions.append(_repair_action(
            "tool_call",
            f"{title_prefix}：{', '.join(unique_values[:3])}",
            severity="warn",
            tool=tool,
            args=args,
            reason=reason,
            source=source,
            priority=priority,
        ))
        if len(actions) >= max(0, int(limit or 0)):
            break
    return actions


def _citation_detail_required_actions(citations: Sequence[Mapping[str, object]] | None,
                                      evidence_nodes: Sequence[Mapping[str, object]] | None,
                                      *,
                                      limit: int = 4) -> list[dict]:
    by_id = _node_by_id(evidence_nodes)
    candidate_nodes = []
    for citation in citations or []:
        if not isinstance(citation, Mapping) or not citation.get("valid"):
            continue
        node = by_id.get(str(citation.get("id") or ""))
        if not isinstance(node, Mapping):
            continue
        source = node.get("source") if isinstance(node.get("source"), Mapping) else {}
        detail = node.get("detail_tool") if isinstance(node.get("detail_tool"), Mapping) else {}
        source_tool = str(source.get("tool") or "").strip()
        evidence_type = str(node.get("type") or "")
        if not detail or _is_datasheet_detail_node(node):
            continue
        if source_tool in _DETAIL_REQUIRED_SOURCE_TOOLS or evidence_type in {"datasheet_chunk", "datasheet_excerpt"}:
            candidate_nodes.append(node)
    return _detail_actions_from_nodes(
        candidate_nodes,
        source="citation_detail_required",
        reason="最终回答引用了 datasheet 搜索命中摘要；规格书/定量类结论前应读取 detail chunk 原文。",
        limit=limit,
        priority=18,
    )


def _quantitative_claim_detail_actions(answer: object,
                                       citations: Sequence[Mapping[str, object]] | None,
                                       evidence_nodes: Sequence[Mapping[str, object]] | None,
                                       *,
                                       limit: int = 4) -> list[dict]:
    if not _has_quantitative_spec_claim(answer):
        return []
    cited_nodes = _valid_citation_nodes(citations, evidence_nodes)
    if any(_is_datasheet_detail_node(node) for node in cited_nodes):
        return []
    # Prefer detail tools attached to cited datasheet search hits, then fall back to
    # any available datasheet search evidence from the current run.
    candidates = [
        node for node in cited_nodes
        if str(node.get("type") or "") in {"datasheet_chunk", "datasheet_excerpt"}
        and isinstance(node.get("detail_tool"), Mapping)
    ]
    if not candidates:
        candidates = [
            node for node in evidence_nodes or []
            if isinstance(node, Mapping)
            and str(node.get("type") or "") in {"datasheet_chunk", "datasheet_excerpt"}
            and isinstance(node.get("detail_tool"), Mapping)
        ]
    return _detail_actions_from_nodes(
        candidates,
        source="quantitative_claim_detail_required",
        reason="最终回答包含规格书/电气参数类定量结论，但没有引用 datasheet detail 原文证据；应先打开 chunk/page 原文再确认。",
        limit=limit,
        priority=14,
    )


def _answer_missing_target_actions(answer: object,
                                   evidence_goal_contract: Mapping[str, object] | None,
                                   *,
                                   limit: int = 4) -> list[dict]:
    contract = evidence_goal_contract if isinstance(evidence_goal_contract, Mapping) else {}
    targets = [
        target for target in contract.get("covered_targets") or []
        if isinstance(target, Mapping) and target.get("value")
    ]
    if len(targets) < 2:
        return []
    answer_text = str(answer or "")
    missing = [
        target for target in targets
        if not _contains_target_token(answer_text, target.get("normalized") or target.get("value"))
    ]
    return _target_repair_actions_from_targets(
        missing,
        source="answer_missing_target_coverage",
        reason="最终回答未逐项覆盖用户问题中的目标对象/料号/型号；用聚焦工具补一轮证据，推动模型按目标逐项回答。",
        title_prefix="聚焦补齐回答遗漏目标",
        priority=16,
        limit=limit,
    )


def _answer_target_citation_actions(answer: object,
                                    citations: Sequence[Mapping[str, object]] | None,
                                    evidence_nodes: Sequence[Mapping[str, object]] | None,
                                    evidence_goal_contract: Mapping[str, object] | None,
                                    *,
                                    limit: int = 4) -> list[dict]:
    contract = evidence_goal_contract if isinstance(evidence_goal_contract, Mapping) else {}
    targets = [
        target for target in contract.get("covered_targets") or []
        if isinstance(target, Mapping) and target.get("value")
    ]
    if len(targets) < 2:
        return []
    answer_text = str(answer or "")
    cited_blobs = [_node_search_blob(node) for node in _valid_citation_nodes(citations, evidence_nodes)]
    missing_citations = [
        target for target in targets
        if _contains_target_token(answer_text, target.get("normalized") or target.get("value"))
        and not any(_contains_target_token(blob, target.get("normalized") or target.get("value")) for blob in cited_blobs)
    ]
    return _target_repair_actions_from_targets(
        missing_citations,
        source="answer_target_citation_missing",
        reason="最终回答提到了用户目标，但 citation 未覆盖该目标对应 evidence；用聚焦工具补一轮证据，推动模型为每个目标绑定 citation。",
        title_prefix="补齐目标 citation",
        priority=17,
        limit=limit,
    )


def is_low_effort_answer(answer: object) -> bool:
    text = str(answer or "").strip()
    if not text:
        return True
    lowered = text.lower()
    return any(marker.lower() in lowered for marker in EARLY_STOP_MARKERS)


def build_perseverance_retry_note(*,
                                  step_type: str,
                                  answer: object = "",
                                  tool_call_count: int = 0,
                                  max_tool_calls: int = 0,
                                  playbook_plan: Mapping[str, object] | None = None,
                                  tool_result_contracts: Sequence[Mapping[str, object]] | None = None,
                                  task_ledger: Mapping[str, object] | None = None,
                                  evidence_node_count: int = 0,
                                  citation_count: int = 0,
                                  allow_needs_user_input: bool = True) -> str:
    """Return a retry note when the model stops before exhausting safe evidence paths."""

    if int(max_tool_calls or 0) <= int(tool_call_count or 0):
        return ""

    first_tools = _recommended_tools_from_playbook(playbook_plan)
    next_tools = _recommended_tools_from_contracts(tool_result_contracts)
    ledger_tools = _recommended_tools_from_task_ledger(task_ledger)
    recommended = []
    for name in [*next_tools, *ledger_tools, *first_tools]:
        if name and name not in recommended:
            recommended.append(name)
    has_safe_next_step = bool(recommended)
    ledger_open = _task_ledger_has_open_tool_work(task_ledger)
    seeded_suffix = _seeded_retry_suffix(playbook_plan=playbook_plan, task_ledger=task_ledger)

    if step_type == "needs_user_input":
        if allow_needs_user_input and tool_call_count == 0 and has_safe_next_step:
            return (
                "你还没有调用任何本地只读工具就要求用户补充。请先沿 playbook 推荐工具取证，"
                f"可优先尝试：{', '.join(recommended[:4])}。{seeded_suffix}"
                "只有这些工具仍无法提供必要证据时，才输出 needs_user_input。"
            )
        if allow_needs_user_input and ledger_open and has_safe_next_step and evidence_node_count == 0:
            return (
                "task_ledger 仍有可执行的安全取证路径，且当前没有有效证据节点。"
                f"请先沿任务账本继续取证，可优先尝试：{', '.join(recommended[:4])}。{seeded_suffix}"
            )
        return ""

    if step_type != "final_answer":
        return ""

    low_effort = is_low_effort_answer(answer)
    incomplete = _has_incomplete_contract(tool_result_contracts)
    if evidence_node_count == 0 and int(citation_count or 0) == 0 and ledger_open and has_safe_next_step:
        return (
            "task_ledger 显示仍有可执行的本地只读取证路径，但当前还没有有效证据节点。"
            f"请不要直接给最终回答，先调用推荐工具取证，可优先尝试：{', '.join(recommended[:4])}。{seeded_suffix}"
        )
    if tool_call_count == 0 and has_safe_next_step and low_effort:
        return (
            "不要在未尝试本地只读工具前直接拒绝、说信息不足或给最终结论。"
            f"请先调用推荐工具取证，可优先尝试：{', '.join(recommended[:4])}。{seeded_suffix}"
        )
    if incomplete and has_safe_next_step and low_effort:
        return (
            "已有工具结果显示 preview/partial/truncated，不能直接放弃。"
            f"请沿 tool_result_contract.recommended_next_tools 继续取证，可优先尝试：{', '.join(recommended[:4])}。{seeded_suffix}"
        )
    return ""


def build_final_answer_quality_gate(*,
                                    answer: object,
                                    citations: Sequence[Mapping[str, object]] = (),
                                    proposed_actions: Sequence[Mapping[str, object]] = (),
                                    evidence_nodes: Sequence[Mapping[str, object]] = (),
                                    tool_result_contracts: Sequence[Mapping[str, object]] | None = None,
                                    task_ledger: Mapping[str, object] | None = None,
                                    evidence_goal_contract: Mapping[str, object] | None = None) -> dict:
    """Build a local final-answer self-check card for trace and future retries."""

    reasons: list[dict] = []
    score = 100
    answer_text = str(answer or "").strip()
    valid_citations = [item for item in citations or [] if isinstance(item, Mapping) and item.get("valid")]
    invalid_citations = [item for item in citations or [] if isinstance(item, Mapping) and item.get("valid") is False]
    evidence_count = len([item for item in evidence_nodes or [] if isinstance(item, Mapping)])
    incomplete_contracts = [
        item for item in tool_result_contracts or []
        if isinstance(item, Mapping) and str(item.get("completeness") or "").lower() in {"preview", "partial", "truncated"}
    ]
    progress = task_ledger.get("progress") if isinstance(task_ledger, Mapping) and isinstance(task_ledger.get("progress"), Mapping) else {}
    blocked_count = int(progress.get("blocked") or 0)
    next_tools = _recommended_tools_from_task_ledger(task_ledger)
    next_tool_actions = _recommended_tool_actions_from_task_ledger(task_ledger)
    evidence_goal_contract = evidence_goal_contract if isinstance(evidence_goal_contract, Mapping) else {}
    evidence_goal_status = str(evidence_goal_contract.get("status") or "").lower()
    target_status = str(evidence_goal_contract.get("target_status") or "").lower()
    missing_goal_types = list_of_strings(evidence_goal_contract.get("missing_evidence_types"), limit=8, text_limit=120)
    missing_target_values = [
        target.get("value")
        for target in evidence_goal_contract.get("missing_targets") or []
        if isinstance(target, Mapping)
    ]
    missing_targets = list_of_strings(
        missing_target_values,
        limit=8,
        text_limit=120,
    )
    citation_detail_actions = _citation_detail_required_actions(valid_citations, evidence_nodes)
    quantitative_detail_actions = _quantitative_claim_detail_actions(answer, valid_citations, evidence_nodes)
    answer_target_actions = _answer_missing_target_actions(answer, evidence_goal_contract)
    target_citation_actions = _answer_target_citation_actions(answer, valid_citations, evidence_nodes, evidence_goal_contract)
    connection_review_phase_status = str(evidence_goal_contract.get("connection_review_phase_status") or "").lower()
    missing_connection_review_phase_titles = list_of_strings(
        [
            phase.get("title") or phase.get("id")
            for phase in evidence_goal_contract.get("missing_connection_review_phases") or []
            if isinstance(phase, Mapping)
        ],
        limit=6,
        text_limit=120,
    )

    def add(reason_id: str, severity: str, message: str, penalty: int) -> None:
        nonlocal score
        reasons.append({
            "id": reason_id,
            "severity": severity,
            "message": _preview(message, 260),
        })
        score -= penalty

    if not answer_text:
        add("empty_answer", "fail", "最终回答为空。", 45)
    elif is_low_effort_answer(answer_text):
        add("low_effort_answer", "warn", "最终回答包含无法判断/信息不足等低努力话术。", 20)

    if evidence_count > 0 and not valid_citations:
        add("missing_valid_citation", "warn", "已有本地 evidence，但最终回答没有有效引用。", 25)
    if invalid_citations:
        add("invalid_citation", "warn", f"存在 {len(invalid_citations)} 个无效 evidence 引用。", 15)
    if incomplete_contracts:
        add("incomplete_tool_result", "warn", "仍存在 preview/partial/truncated 工具结果，结论可能需要继续聚合或回拉详情。", 15)
    if blocked_count:
        add("blocked_ledger_item", "warn", f"任务账本仍有 {blocked_count} 个 blocked 项，需要用户或人工补充。", 15)
    if next_tools and not valid_citations:
        add("open_next_actions", "warn", "任务账本仍有推荐下一步工具，且当前回答缺少有效证据引用。", 10)
    if evidence_goal_status in {"missing", "partial"} and missing_goal_types:
        add("missing_evidence_goal", "warn", f"当前 playbook 所需 evidence 类型尚未出现：{', '.join(missing_goal_types) or 'unknown'}。", 22)
    if target_status in {"missing", "partial"} and missing_targets:
        add("missing_target_coverage", "warn", f"用户问题中的目标尚未出现在 evidence 中：{', '.join(missing_targets)}。", 18)
    if connection_review_phase_status in {"missing", "partial"} and missing_connection_review_phase_titles:
        add("missing_connection_review_phase", "warn", f"连接 × datasheet 反查阶段尚未补齐：{', '.join(missing_connection_review_phase_titles)}。", 20)
    if answer_target_actions:
        add("answer_missing_target_coverage", "warn", f"最终回答未逐项覆盖 {len(answer_target_actions)} 组已取证目标。", 12)
    if target_citation_actions:
        add("answer_target_citation_missing", "warn", f"最终回答提到了目标对象，但 citation 未逐项覆盖 {len(target_citation_actions)} 组目标证据。", 14)
    if _has_quantitative_spec_claim(answer_text) and quantitative_detail_actions:
        add("quantitative_claim_detail_required", "warn", f"最终回答包含定量规格/电气参数结论，但缺少 datasheet detail 原文 citation。", 18)
    if citation_detail_actions:
        add("citation_detail_required", "warn", f"存在 {len(citation_detail_actions)} 个 citation 仍停留在搜索摘要层，需要打开 detail chunk 后再确认。", 18)

    score = max(0, min(100, score))
    severities = {item["severity"] for item in reasons}
    if "fail" in severities or score < 45:
        status = "fail"
    elif reasons:
        status = "warn"
    else:
        status = "pass"

    reason_ids = {item.get("id") for item in reasons}
    repair_actions: list[dict] = []
    if "missing_valid_citation" in reason_ids:
        repair_actions.append(_repair_action(
            "revise_answer",
            "补充有效 evidence citation",
            severity="warn",
            reason="已有本地 evidence，但最终回答没有引用有效证据。应重写结论并引用现有 evidence id。",
            source="missing_valid_citation",
            priority=20,
        ))
    if "invalid_citation" in reason_ids:
        repair_actions.append(_repair_action(
            "revise_answer",
            "替换或删除无效 citation",
            severity="warn",
            reason="最终回答引用了不存在的 evidence id。应从 trace 中选择有效 evidence，或移除无法支持的结论。",
            source="invalid_citation",
            priority=30,
        ))
    if "citation_detail_required" in reason_ids:
        repair_actions.extend(citation_detail_actions)
    if "quantitative_claim_detail_required" in reason_ids:
        repair_actions.extend(quantitative_detail_actions)
    if "answer_missing_target_coverage" in reason_ids:
        repair_actions.extend(answer_target_actions)
    if "answer_target_citation_missing" in reason_ids:
        repair_actions.extend(target_citation_actions)
    if "incomplete_tool_result" in reason_ids:
        for index, contract in enumerate(tool_result_contracts or [], start=1):
            if not isinstance(contract, Mapping):
                continue
            completeness = str(contract.get("completeness") or "").lower()
            if completeness not in {"preview", "partial", "truncated"}:
                continue
            scope = _preview(contract.get("scope_summary") or f"contract-{index}", 180)
            aggregation = contract.get("aggregation_tool")
            if isinstance(aggregation, Mapping) and aggregation.get("name"):
                aggregation_name = str(aggregation.get("name") or "")
                aggregation_matches_goal = (
                    aggregation_name != "summarize_schematic_page_count"
                    or "schematic_page_count" in missing_goal_types
                )
                if aggregation_matches_goal:
                    repair_actions.append(_repair_action(
                        "tool_call",
                        f"聚合截断结果：{scope}",
                        severity="warn",
                        tool=aggregation_name,
                        args=aggregation.get("args") if isinstance(aggregation.get("args"), Mapping) else {},
                        reason=f"工具结果完整性为 {completeness}，统计/覆盖类结论前应先调用聚合工具。",
                        source=f"incomplete_tool_result-{index}",
                        priority=25,
                    ))
            detail = contract.get("detail_tool")
            if isinstance(detail, Mapping) and detail.get("name"):
                repair_actions.append(_repair_action(
                    "tool_call",
                    f"读取原始详情：{scope}",
                    severity="warn",
                    tool=str(detail.get("name") or ""),
                    args=detail.get("args") if isinstance(detail.get("args"), Mapping) else {},
                    reason=f"工具结果完整性为 {completeness}，高风险或不确定结论前应回拉 detail。",
                    source=f"incomplete_tool_result-{index}",
                    priority=35,
                ))
            for tool in list_of_strings(contract.get("recommended_next_tools"), limit=8, text_limit=100):
                repair_actions.append(_repair_action(
                    "tool_call",
                    f"继续推荐取证：{scope}",
                    severity="warn",
                    tool=tool,
                    reason="工具结果仍是 preview/partial/truncated，不能把预览当完整事实；应调用推荐工具聚合或回拉详情。",
                    source=f"incomplete_tool_result-{index}",
                    priority=45,
                ))
    if "blocked_ledger_item" in reason_ids:
        repair_actions.append(_repair_action(
            "ask_user",
            "向用户补充 blocked 任务所需信息",
            severity="warn",
            reason="任务账本仍有 blocked 项，当前本地工具无法继续补齐，需要结构化向用户追问。",
            source="blocked_ledger_item",
            priority=45,
        ))
    if "open_next_actions" in reason_ids:
        for action in next_tool_actions:
            repair_actions.append(_repair_action(
                "tool_call",
                f"继续任务账本下一步：{action.get('tool')}",
                severity="warn",
                tool=str(action.get("tool") or ""),
                args=action.get("args") if isinstance(action.get("args"), Mapping) else None,
                reason=action.get("reason") or "任务账本仍有安全的本地只读下一步，最终回答前应继续取证。",
                source="open_next_actions",
                priority=40,
            ))
    if "missing_evidence_goal" in reason_ids:
        for action in list(evidence_goal_contract.get("repair_actions") or [])[:8]:
            if not isinstance(action, Mapping) or not action.get("tool"):
                continue
            if str(action.get("source") or "") == "missing_target_coverage":
                continue
            repair_actions.append(_repair_action(
                "tool_call",
                action.get("title") or f"补齐证据目标：{action.get('tool')}",
                severity="warn",
                tool=str(action.get("tool") or ""),
                args=action.get("args") if isinstance(action.get("args"), Mapping) else None,
                reason=action.get("reason") or "证据目标契约要求补齐当前 playbook 所需 evidence 后再下结论。",
                source="missing_evidence_goal",
                priority=int(action.get("priority") or 42),
            ))
    if "missing_target_coverage" in reason_ids:
        for action in list(evidence_goal_contract.get("repair_actions") or [])[:8]:
            if not isinstance(action, Mapping) or not action.get("tool"):
                continue
            if str(action.get("source") or "") != "missing_target_coverage":
                continue
            repair_actions.append(_repair_action(
                "tool_call",
                action.get("title") or f"补齐目标对象取证：{action.get('tool')}",
                severity="warn",
                tool=str(action.get("tool") or ""),
                args=action.get("args") if isinstance(action.get("args"), Mapping) else None,
                reason=action.get("reason") or "问题目标覆盖契约要求补齐用户问题中的具体对象/料号/型号 evidence。",
                source="missing_target_coverage",
                priority=int(action.get("priority") or 18),
            ))
    if "missing_connection_review_phase" in reason_ids:
        for action in list(evidence_goal_contract.get("connection_review_repair_actions") or evidence_goal_contract.get("repair_actions") or [])[:8]:
            if not isinstance(action, Mapping) or not action.get("tool"):
                continue
            if str(action.get("source") or "") != "missing_connection_review_phase":
                continue
            repair_actions.append(_repair_action(
                "tool_call",
                action.get("title") or f"补齐连接反查阶段：{action.get('tool')}",
                severity="warn",
                tool=str(action.get("tool") or ""),
                args=action.get("args") if isinstance(action.get("args"), Mapping) else None,
                reason=action.get("reason") or "连接 × datasheet 反查必须同时覆盖连接 evidence、元件身份和 datasheet detail/gap。",
                source="missing_connection_review_phase",
                priority=int(action.get("priority") or 16),
            ))
    if status == "fail" and not repair_actions:
        repair_actions.append(_repair_action(
            "manual_review",
            "人工复核最终回答",
            severity="fail",
            reason="质量门禁失败但没有可自动推荐的安全修复路径，需要人工复核 trace。",
            source="quality_gate_fail",
            priority=90,
        ))

    return {
        "version": "final-answer-quality-gate/v1",
        "status": status,
        "score": score,
        "reason_count": len(reasons),
        "reasons": reasons[:10],
        "valid_citation_count": len(valid_citations),
        "invalid_citation_count": len(invalid_citations),
        "evidence_node_count": evidence_count,
        "proposed_action_count": len([item for item in proposed_actions or [] if isinstance(item, Mapping)]),
        "recommended_next_tools": next_tools[:8],
        "evidence_goal_contract": dict(evidence_goal_contract or {}),
        "repair_action_count": len(repair_actions),
        "repair_actions": repair_actions[:12],
        "notes": "本地 runtime 质量门禁用于复盘 final_answer，不代表业务结论。第一版不强制阻断 warn。",
    }
