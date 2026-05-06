# -*- coding: utf-8 -*-
"""Local read-only review harness for PSTX reports."""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import List, Optional

from pstx_harness.model import MockHarnessModelProvider
from pstx_harness.report_tools import (
    DEFAULT_TOOL_ORDER,
    HarnessToolContext,
    HarnessToolRegistry,
    build_default_harness_registry,
)


class HarnessError(ValueError):
    """Raised when harness input cannot be accepted safely."""


@dataclass(frozen=True)
class HarnessRunRequest:
    task: str = "full_review"
    question: str = ""
    max_rows_per_table: int = 12
    include_model: bool = True

    @classmethod
    def from_mapping(cls, value: Optional[dict]) -> "HarnessRunRequest":
        value = value or {}
        task = str(value.get("task") or "full_review").strip() or "full_review"
        question = str(value.get("question") or "").strip()
        include_model = _as_bool(value.get("include_model", True), True)
        try:
            max_rows = int(value.get("max_rows_per_table", 12))
        except (TypeError, ValueError) as exc:
            raise HarnessError("max_rows_per_table 必须是数字。") from exc
        request = cls(
            task=task,
            question=question[:2000],
            max_rows_per_table=max_rows,
            include_model=include_model,
        )
        request.validate()
        return request

    def validate(self) -> None:
        if self.task != "full_review":
            raise HarnessError("第一版 harness 仅支持 task=full_review。")
        if self.max_rows_per_table < 1 or self.max_rows_per_table > 100:
            raise HarnessError("max_rows_per_table 必须在 1 到 100 之间。")


def _as_bool(value, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "on", "y"}:
        return True
    if text in {"0", "false", "no", "off", "n"}:
        return False
    return default


def _extract_balanced_json(text: str) -> Optional[dict]:
    content = str(text or "").strip()
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


def _safe_str(value, limit: int = 700) -> str:
    text = "" if value is None else str(value).strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _normalize_priorities(value, fallback: List[dict]) -> List[dict]:
    allowed_targets = {"bom", "network", "drc", "csa", "resistor", "derating", "summary"}
    allowed_severity = {"high", "medium", "low"}
    rows = []
    if isinstance(value, list):
        for item in value[:5]:
            if not isinstance(item, dict):
                continue
            target = str(item.get("target") or "summary").strip()
            severity = str(item.get("severity") or "medium").strip().lower()
            rows.append({
                "title": _safe_str(item.get("title") or "审查建议", 80),
                "body": _safe_str(item.get("body") or item.get("evidence") or "", 700),
                "target": target if target in allowed_targets else "summary",
                "severity": severity if severity in allowed_severity else "medium",
            })
    return rows or fallback


def _normalize_checklist(value, fallback: List[dict]) -> List[dict]:
    allowed_targets = {"bom", "network", "drc", "csa", "resistor", "derating", "summary"}
    allowed_status = {"pass", "covered_no_findings", "covered_with_findings", "needs_review", "manual_only"}
    allowed_severity = {"high", "medium", "low"}
    rows = []
    if isinstance(value, list):
        for item in value[:12]:
            if not isinstance(item, dict):
                continue
            target = str(item.get("target") or "summary").strip()
            status = str(item.get("status") or "needs_review").strip().lower()
            severity = str(item.get("severity") or "medium").strip().lower()
            rows.append({
                "item": _safe_str(item.get("item") or item.get("title") or "审查项", 80),
                "status": status if status in allowed_status else "needs_review",
                "evidence": _safe_str(item.get("evidence") or item.get("body") or "", 700),
                "target": target if target in allowed_targets else "summary",
                "severity": severity if severity in allowed_severity else "medium",
            })
    return rows or fallback


def _normalize_manual_review(value, fallback: List[dict]) -> List[dict]:
    allowed_targets = {"bom", "network", "drc", "csa", "resistor", "derating", "summary"}
    rows = []
    if isinstance(value, list):
        for item in value[:8]:
            if not isinstance(item, dict):
                continue
            target = str(item.get("target") or "summary").strip()
            rows.append({
                "topic": _safe_str(item.get("topic") or item.get("title") or "人工复核项", 80),
                "reason": _safe_str(item.get("reason") or item.get("boundary") or "", 700),
                "target": target if target in allowed_targets else "summary",
            })
    return rows or fallback


def _local_fallback(evidence_packs: List[dict]) -> dict:
    active = [pack for pack in evidence_packs if int(pack.get("issue_count") or 0) > 0]
    priorities = [
        {
            "title": pack.get("title", "审查项"),
            "body": pack.get("summary", ""),
            "target": pack.get("target", "summary"),
            "severity": pack.get("severity", "medium"),
        }
        for pack in active[:5]
    ]
    if not priorities:
        priorities = [{
            "title": "本地 harness 未发现高优先级项",
            "body": "当前证据包没有明显计数异常，建议按报告分区做抽样复核。",
            "target": "summary",
            "severity": "low",
        }]
    checklist = [
        {
            "item": pack.get("title", "审查项"),
            "status": "needs_review" if int(pack.get("issue_count") or 0) else "covered_no_findings",
            "evidence": pack.get("summary", ""),
            "target": pack.get("target", "summary"),
            "severity": pack.get("severity", "medium"),
        }
        for pack in evidence_packs[:12]
    ]
    return {
        "summary": f"本地 harness 已执行 {len(evidence_packs)} 个只读证据工具，模型不可用时仍可基于证据包继续复核。",
        "priorities": priorities,
        "review_checklist": checklist,
        "manual_review": [{
            "topic": "自动审查边界",
            "reason": "候选网络、电压、电阻上下拉和降额结论仍需结合规格书与设计意图确认。",
            "target": "summary",
        }],
    }


def _build_prompt(request: HarnessRunRequest, report: dict, evidence_packs: List[dict]) -> str:
    payload = {
        "project_name": report.get("project_name"),
        "task": request.task,
        "question": request.question,
        "evidence_packs": evidence_packs,
    }
    return (
        "你是硬件原理图审查 harness 的模型接口。\n"
        "本地工具已经固定执行完毕，你只能基于 evidence_packs 输出审查建议，不能请求或执行任何工具。\n"
        "如果信息不足，请写需人工确认，不要编造。\n"
        "只输出 JSON 对象，字段为 summary、priorities、review_checklist、manual_review。\n"
        f"输入：\n{json.dumps(payload, ensure_ascii=False, indent=2)}"
    )


def run_harness_review(report: dict,
                       bundle: dict,
                       request: HarnessRunRequest,
                       model_provider=None,
                       registry: Optional[HarnessToolRegistry] = None) -> dict:
    request.validate()
    registry = registry or build_default_harness_registry()
    context = HarnessToolContext(report=report, bundle=bundle, request=request)
    evidence_packs: List[dict] = []
    tool_runs: List[dict] = []
    for tool_name in DEFAULT_TOOL_ORDER:
        try:
            pack = registry.run(tool_name, context)
        except Exception as exc:
            tool_runs.append({"tool": tool_name, "ok": False, "error": str(exc)})
            continue
        evidence_packs.append(pack)
        tool_runs.append({
            "tool": tool_name,
            "ok": True,
            "evidence_pack_id": pack.get("id", tool_name),
            "issue_count": pack.get("issue_count", 0),
        })

    fallback = _local_fallback(evidence_packs)
    model_metadata = {
        "included": bool(request.include_model),
        "provider": "none",
        "ok": True,
    }
    parsed = None
    model_answer = ""
    if request.include_model:
        provider = model_provider or MockHarnessModelProvider()
        prompt = _build_prompt(request, report, evidence_packs)
        try:
            response = provider.generate(prompt, inputs={
                "project_name": report.get("project_name") or bundle.get("project_name") or "",
                "task": request.task,
                "question": request.question,
                "evidence_packs": evidence_packs,
            })
            model_answer = response.answer
            parsed = _extract_balanced_json(model_answer)
            model_metadata.update({
                "provider": response.provider,
                "mode": response.mode,
                "ok": True,
                **dict(response.metadata or {}),
            })
        except Exception as exc:
            model_metadata.update({
                "provider": provider.__class__.__name__,
                "ok": False,
                "error": str(exc),
            })

    summary = fallback["summary"]
    if parsed and isinstance(parsed.get("summary"), str) and parsed["summary"].strip():
        summary = parsed["summary"].strip()[:1200]
    elif model_answer.strip():
        summary = model_answer.strip()[:1200]

    return {
        "ok": True,
        "mode": "local-harness",
        "task": request.task,
        "question": request.question,
        "summary": summary,
        "priorities": _normalize_priorities(parsed.get("priorities") if parsed else None, fallback["priorities"]),
        "review_checklist": _normalize_checklist(
            parsed.get("review_checklist") if parsed else None,
            fallback["review_checklist"],
        ),
        "manual_review": _normalize_manual_review(
            parsed.get("manual_review") if parsed else None,
            fallback["manual_review"],
        ),
        "evidence_packs": evidence_packs,
        "tool_runs": tool_runs,
        "model_metadata": model_metadata,
        "safeguards": [
            "Harness 工具第一版全部只读，不写文件、不修改 PSTX、不主动联网同步外部系统。",
            "Aster 只作为模型接口接收本地证据包，不允许直接执行本地工具。",
            "模型输出中的工具调用 JSON 会被当作普通文本/无效结构处理，不会触发执行。",
        ],
    }


def build_harness_status(*, model_status: Optional[dict] = None) -> dict:
    registry = build_default_harness_registry()
    return {
        "ok": True,
        "mode": "local-harness",
        "tool_count": len(registry.list_tools()),
        "tools": registry.list_tools(),
        "model_provider": model_status or {},
        "safeguards": [
            "本地 harness 固定执行白名单只读工具。",
            "Aster 只作为 LLM Provider，不执行工具、不接触本地文件系统操作。",
            "第一版不开放模型自主 tool-calling。",
        ],
    }
