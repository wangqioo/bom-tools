# -*- coding: utf-8 -*-
"""Deterministic evaluation runner for the local harness agent."""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from pstx_harness.review import HarnessError
from pstx_harness.report_agent import HarnessAgentRequest, run_harness_agent
from pstx_harness.model import HarnessModelResponse, MockHarnessModelProvider


class AgentEvalError(ValueError):
    """Raised when an eval request is invalid."""


@dataclass(frozen=True)
class AgentEvalCase:
    case_id: str
    title: str
    description: str
    request: Dict[str, object] = field(default_factory=dict)
    provider_script: List[str] = field(default_factory=list)
    provider_kind: str = "script"
    expected_tools: List[str] = field(default_factory=list)
    forbidden_tools: List[str] = field(default_factory=list)
    expected_ok: Optional[bool] = True
    expected_stopped_reason: str = ""
    require_valid_citation: bool = False
    require_invalid_citation: bool = False
    expected_error_contains: str = ""
    expected_tool_error_recovery_count: Optional[int] = None


class ScriptedEvalProvider:
    provider = "agent-eval-scripted"
    mode = "eval"

    def __init__(self, answers: List[str]):
        self.answers = list(answers)
        self.calls: List[dict] = []

    def generate_agent_step(self, prompt: str, *, inputs: Optional[dict] = None) -> HarnessModelResponse:
        self.calls.append({"prompt_chars": len(prompt), "inputs": inputs or {}})
        answer = self.answers.pop(0) if self.answers else '{"final_answer":"eval fallback"}'
        return HarnessModelResponse(
            answer=answer,
            provider=self.provider,
            mode=self.mode,
            metadata={"eval_provider": "scripted", "call_index": len(self.calls)},
        )


class LoopEvalProvider:
    provider = "agent-eval-loop"
    mode = "eval"

    def __init__(self, tool_name: str = "list_report_tables"):
        self.tool_name = tool_name

    def generate_agent_step(self, prompt: str, *, inputs: Optional[dict] = None) -> HarnessModelResponse:
        payload = {
            "tool_call": {
                "name": self.tool_name,
                "args": {},
                "reason": "eval loop keeps requesting the same tool",
            }
        }
        return HarnessModelResponse(
            answer=json.dumps(payload, ensure_ascii=False),
            provider=self.provider,
            mode=self.mode,
            metadata={"eval_provider": "loop"},
        )


def sample_eval_report() -> dict:
    return {
        "project_name": "agent_eval_board",
        "ratio_limit": 70,
        "include_depop": False,
        "metrics": [
            {"label": "DEPOP 总数", "value": 1},
            {"label": "BOM圈问题", "value": 1},
            {"label": "降额不合格", "value": 1},
        ],
        "sections": [
            {
                "id": "drc",
                "title": "设计检查",
                "tables": [
                    {
                        "id": "missing_value",
                        "title": "缺少 VALUE",
                        "count": 1,
                        "columns": ["位号", "真实页"],
                        "rows": [{"位号": "R1", "真实页": "PAGE12", "问题": "缺少 VALUE"}],
                    },
                    {
                        "id": "bom_option_components",
                        "title": "BOM_OPTION 元件",
                        "count": 1,
                        "columns": ["位号", "BOM_OPTION", "页面"],
                        "rows": [{"位号": "R2", "BOM_OPTION": "DEPOP", "页面": "PAGE12"}],
                    },
                ],
            },
            {
                "id": "network",
                "title": "网络分析",
                "tables": [
                    {
                        "id": "page_mapping_rows",
                        "title": "逻辑页/真实页映射检查",
                        "count": 1,
                        "rows": [{"逻辑页": "PAGE242", "真实页": "PAGE518", "是否一一对应": "是"}],
                    }
                ],
            },
            {
                "id": "resistor",
                "title": "电阻检查",
                "tables": [
                    {
                        "id": "chip_pin_rows",
                        "title": "芯片 Pin 电阻状态",
                        "count": 1,
                        "rows": [{"芯片位号": "U1", "引脚": "GPIO1", "状态": "候选判断"}],
                    }
                ],
            },
        ],
    }


def _json_payload(value: dict) -> str:
    return json.dumps(value, ensure_ascii=False)


def default_eval_cases() -> List[AgentEvalCase]:
    return [
        AgentEvalCase(
            case_id="mock_quick_scan",
            title="Mock 快速扫描",
            description="默认 mock provider 应读取缺少 VALUE 表并给出有效 evidence citation。",
            request={"profile": "quick_scan", "max_steps": 4},
            provider_kind="mock",
            expected_tools=["get_table_rows"],
            expected_ok=True,
            expected_stopped_reason="final_answer",
            require_valid_citation=True,
        ),
        AgentEvalCase(
            case_id="invalid_json_retry",
            title="非法 JSON 重试",
            description="模型第一次输出非法 JSON 时，harness 应重试并接受后续 final answer。",
            request={"profile": "quick_scan", "max_steps": 2},
            provider_script=[
                "not json at all",
                _json_payload({"final_answer": "retry recovered", "confidence": "low"}),
            ],
            expected_tools=[],
            expected_ok=True,
            expected_stopped_reason="final_answer",
        ),
        AgentEvalCase(
            case_id="unknown_tool_rejected",
            title="未知工具恢复",
            description="模型请求未知工具时，本地 harness 应把失败作为观察结果反馈给模型，并允许一次安全恢复。",
            request={"profile": "quick_scan"},
            provider_script=[
                _json_payload({"tool_call": {"name": "unsafe_tool", "args": {}, "reason": "try unsafe"}}),
            ],
            expected_tools=["unsafe_tool"],
            expected_ok=True,
            expected_stopped_reason="final_answer",
            expected_tool_error_recovery_count=1,
        ),
        AgentEvalCase(
            case_id="profile_blocks_file_read",
            title="Profile 禁止文件读取",
            description="bom_depop profile 不允许读取项目文件，即使模型请求 read_project_text 也必须拒绝。",
            request={"profile": "bom_depop"},
            provider_script=[
                _json_payload({
                    "tool_call": {
                        "name": "read_project_text",
                        "args": {"path": "packaged/pstxprt.dat"},
                        "reason": "should be blocked by profile",
                    }
                }),
            ],
            expected_ok=False,
            expected_stopped_reason="tool_error",
            expected_error_contains="profile bom_depop",
        ),
        AgentEvalCase(
            case_id="invalid_citation_flagged",
            title="无效 Citation 标记",
            description="模型引用不存在 evidence id 时，harness 应标记 invalid 并保留 fallback evidence。",
            request={"profile": "quick_scan", "max_steps": 4},
            provider_script=[
                _json_payload({"tool_call": {"name": "list_report_tables", "args": {}, "reason": "collect tables"}}),
                _json_payload({"final_answer": "bad citation", "citations": [{"id": "ev-not-found"}]}),
            ],
            expected_tools=["list_report_tables"],
            expected_ok=True,
            expected_stopped_reason="final_answer",
            require_invalid_citation=True,
        ),
        AgentEvalCase(
            case_id="max_steps_limit",
            title="最大轮数安全终止",
            description="模型持续请求工具时，harness 到达 max_steps 后应安全终止。",
            request={"profile": "quick_scan", "max_steps": 1, "max_tool_calls": 8},
            provider_kind="loop",
            expected_tools=["list_report_tables"],
            expected_ok=True,
            expected_stopped_reason="max_steps",
        ),
    ]


def list_agent_eval_cases() -> List[dict]:
    return [
        {
            "case_id": item.case_id,
            "title": item.title,
            "description": item.description,
            "profile": str(item.request.get("profile") or "quick_scan"),
            "expected_tools": item.expected_tools,
            "expected_stopped_reason": item.expected_stopped_reason,
        }
        for item in default_eval_cases()
    ]


def _provider_for_case(case: AgentEvalCase):
    if case.provider_kind == "mock":
        return MockHarnessModelProvider()
    if case.provider_kind == "loop":
        return LoopEvalProvider()
    return ScriptedEvalProvider(case.provider_script)


def _evaluate_case(case: AgentEvalCase) -> dict:
    failures: List[str] = []
    try:
        request = HarnessAgentRequest.from_mapping(case.request)
    except HarnessError as exc:
        return {
            "case_id": case.case_id,
            "title": case.title,
            "passed": False,
            "failures": [str(exc)],
            "payload": {},
        }
    payload = run_harness_agent(sample_eval_report(), {}, request, model_provider=_provider_for_case(case))
    tool_names = [str(item.get("tool") or "") for item in payload.get("tool_calls", [])]
    stopped_reason = str((payload.get("model_metadata") or {}).get("stopped_reason") or "")
    tool_error_recovery_count = int((payload.get("model_metadata") or {}).get("tool_error_recovery_count") or 0)
    answer = str(payload.get("answer") or "")
    valid_citations = [item for item in payload.get("citations", []) if item.get("valid")]
    invalid_citations = [item for item in payload.get("citations", []) if not item.get("valid")]

    if case.expected_ok is not None and bool(payload.get("ok")) != case.expected_ok:
        failures.append(f"ok 期望 {case.expected_ok}，实际 {payload.get('ok')}")
    if case.expected_stopped_reason and stopped_reason != case.expected_stopped_reason:
        failures.append(f"stopped_reason 期望 {case.expected_stopped_reason}，实际 {stopped_reason}")
    if (
        case.expected_tool_error_recovery_count is not None
        and tool_error_recovery_count != case.expected_tool_error_recovery_count
    ):
        failures.append(
            "tool_error_recovery_count 期望 "
            f"{case.expected_tool_error_recovery_count}，实际 {tool_error_recovery_count}"
        )
    for tool_name in case.expected_tools:
        if tool_name not in tool_names:
            failures.append(f"缺少期望工具调用：{tool_name}")
    for tool_name in case.forbidden_tools:
        if tool_name in tool_names:
            failures.append(f"出现禁止工具调用：{tool_name}")
    if case.require_valid_citation and not valid_citations:
        failures.append("缺少有效 citation")
    if case.require_invalid_citation and not invalid_citations:
        failures.append("缺少 invalid citation 标记")
    if case.expected_error_contains and case.expected_error_contains not in answer:
        failures.append(f"错误信息未包含：{case.expected_error_contains}")

    return {
        "case_id": case.case_id,
        "title": case.title,
        "description": case.description,
        "profile": request.profile,
        "passed": not failures,
        "failures": failures,
        "metrics": {
            "tool_call_count": len(tool_names),
            "valid_citation_count": len(valid_citations),
            "invalid_citation_count": len(invalid_citations),
            "stopped_reason": stopped_reason,
            "tool_error_recovery_count": tool_error_recovery_count,
        },
        "tool_calls": tool_names,
        "trace_summary": payload.get("trace_summary", {}),
        "answer_preview": answer[:260],
    }


def run_agent_eval(case_ids: Optional[List[str]] = None) -> dict:
    cases = default_eval_cases()
    if case_ids:
        wanted = set(str(item).strip() for item in case_ids if str(item).strip())
        known = {item.case_id for item in cases}
        unknown = sorted(wanted - known)
        if unknown:
            raise AgentEvalError(f"未知 eval case：{', '.join(unknown)}")
        cases = [item for item in cases if item.case_id in wanted]
    results = [_evaluate_case(case) for case in cases]
    passed_count = sum(1 for item in results if item.get("passed"))
    failed_count = len(results) - passed_count
    score = round((passed_count / len(results)) * 100, 2) if results else 0.0
    return {
        "ok": failed_count == 0,
        "mode": "local-agent-eval",
        "case_count": len(results),
        "passed_count": passed_count,
        "failed_count": failed_count,
        "score": score,
        "cases": results,
        "safeguards": [
            "Eval Center 只使用本地 deterministic provider，不调用真实 Aster。",
            "Eval Center 只验证 Agent 行为边界，不修改 PSTX 或项目文件。",
            "Eval 判断工具轨迹、citation 和安全边界，不比较 LLM 文案。",
        ],
    }


def build_agent_eval_status() -> dict:
    cases = list_agent_eval_cases()
    return {
        "ok": True,
        "mode": "local-agent-eval",
        "case_count": len(cases),
        "cases": cases,
        "capabilities": ["deterministic_eval", "tool_trace_assertion", "citation_assertion", "safety_boundary_assertion"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Run local PSTX harness agent eval cases.")
    parser.add_argument("--case", action="append", default=[], help="Run a specific eval case id. Can be repeated.")
    args = parser.parse_args()
    payload = run_agent_eval(args.case or None)
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if payload.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
