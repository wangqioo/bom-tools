# -*- coding: utf-8 -*-
"""Model-provider adapters for the local PSTX harness."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Callable, Dict, Optional


@dataclass
class HarnessModelResponse:
    answer: str
    provider: str
    mode: str = "mock"
    metadata: Dict[str, object] = field(default_factory=dict)


class MockHarnessModelProvider:
    """Deterministic local model provider used when Aster is unavailable."""

    provider = "local-harness-mock"
    mode = "mock"

    def generate(self, prompt: str, *, inputs: Optional[dict] = None) -> HarnessModelResponse:
        inputs = inputs or {}
        evidence_packs = list(inputs.get("evidence_packs") or [])
        active = [pack for pack in evidence_packs if int(pack.get("issue_count") or 0) > 0]
        top_titles = "、".join(pack.get("title", pack.get("id", "")) for pack in active[:4]) or "常规审查项"
        payload = {
            "summary": f"本地 harness 已汇总 {len(evidence_packs)} 个只读证据包，建议优先复核 {top_titles}。",
            "priorities": [
                {
                    "title": pack.get("title", "审查项"),
                    "body": pack.get("summary", "该证据包需要工程复核。"),
                    "target": pack.get("target", "summary"),
                    "severity": pack.get("severity", "medium"),
                }
                for pack in (active[:5] or evidence_packs[:1])
            ],
            "review_checklist": [
                {
                    "item": pack.get("title", "审查项"),
                    "status": "needs_review" if int(pack.get("issue_count") or 0) else "covered_no_findings",
                    "evidence": pack.get("summary", ""),
                    "target": pack.get("target", "summary"),
                    "severity": pack.get("severity", "medium"),
                }
                for pack in evidence_packs[:10]
            ],
            "manual_review": [
                {
                    "topic": "人工复核边界",
                    "reason": "Harness 只汇总当前报告中的事实和候选项，无法替代原理图设计意图判断。",
                    "target": "summary",
                }
            ],
        }
        return HarnessModelResponse(
            answer=json.dumps(payload, ensure_ascii=False),
            provider=self.provider,
            mode=self.mode,
            metadata={"prompt_chars": len(prompt), "evidence_pack_count": len(evidence_packs)},
        )

    def generate_agent_step(self, prompt: str, *, inputs: Optional[dict] = None) -> HarnessModelResponse:
        inputs = inputs or {}
        observations = list(inputs.get("observations") or [])
        question = str(inputs.get("question") or "")
        profile = str(inputs.get("agent_profile") or "")
        capability_profiles = set(str(item) for item in inputs.get("capability_profiles") or [])
        if not observations:
            if profile == "review_checklist_qa" or "review_checklist_qa" in capability_profiles:
                query = question.strip()[:120] or "review checklist"
                payload = {
                    "tool_call": {
                        "name": "search_review_checklists",
                        "args": {"query": query, "limit": 5},
                        "reason": "先按用户问题搜索 ref_checklist 审查经验清单。",
                    }
                }
            elif profile == "agent_ref_qa" or "agent_ref_qa" in capability_profiles:
                query = question.strip()[:120] or "Agent Lab"
                payload = {
                    "tool_call": {
                        "name": "search_agent_ref_pdfs",
                        "args": {"query": query, "limit": 5},
                        "reason": "先按用户问题搜索 Agent Lab ref PDF 索引。",
                    }
                }
            elif profile == "feishu_bom_qa" or "feishu_bom_qa" in capability_profiles:
                query = question.strip()[:120] or "HQ"
                payload = {
                    "tool_call": {
                        "name": "search_feishu_cache_rows",
                        "args": {"query": query, "limit": 5},
                        "reason": "先按用户问题搜索本地飞书缓存。",
                    }
                }
            elif profile == "dfmea_prep" or "dfmea_prep" in capability_profiles:
                if any(keyword in question.lower() for keyword in ["规格书", "datasheet", "pdf", "手册"]):
                    payload = {
                        "tool_call": {
                            "name": "summarize_dfmea_datasheet_coverage",
                            "args": {"limit": 8},
                            "reason": "先汇总关键器件的本地规格书覆盖和证据缺口。",
                        }
                    }
                else:
                    payload = {
                        "tool_call": {
                            "name": "summarize_dfmea_readiness",
                            "args": {},
                            "reason": "先汇总元件身份卡和 DFMEA 输入准备度。",
                        }
                    }
            elif "越权" in question or "outside" in question:
                payload = {
                    "tool_call": {
                        "name": "read_project_text",
                        "args": {"path": "../outside-secret.txt"},
                        "reason": "演示路径边界拦截。",
                    }
                }
            elif "文件" in question or "list_project_files" in question:
                payload = {
                    "tool_call": {
                        "name": "list_project_files",
                        "args": {"limit": 20},
                        "reason": "先列出可读取的项目文件。",
                    }
                }
            else:
                payload = {
                    "tool_call": {
                        "name": "get_table_rows",
                        "args": {"table_id": "missing_value", "limit": 5},
                        "reason": "先读取缺少 VALUE 的具体行作为审查证据。",
                    }
                }
        else:
            evidence_ids = []
            missing_context_ids = []
            missing_fields = []
            for observation in observations:
                evidence_ids.extend(str(item) for item in observation.get("evidence_node_ids", []) if item)
                for node in observation.get("evidence_nodes", []) or []:
                    node_id = str(node.get("id") or "")
                    if node_id:
                        evidence_ids.append(node_id)
                    if node.get("type") == "missing_context" and node_id:
                        missing_context_ids.append(node_id)
                    for field_name in node.get("missing_fields", []) or []:
                        if field_name not in missing_fields:
                            missing_fields.append(field_name)
            evidence_ids = list(dict.fromkeys(evidence_ids))
            project_context = inputs.get("project_context") if isinstance(inputs.get("project_context"), dict) else {}
            context_answers = list(project_context.get("answers") or []) + list(inputs.get("context_answers") or [])
            if (profile == "dfmea_prep" or "dfmea_prep" in capability_profiles) and missing_context_ids and not context_answers:
                payload = {
                    "needs_user_input": {
                        "reason": "部分关键器件缺少 DFMEA 准备信息，需要用户补充后再继续。",
                        "missing_fields": missing_fields[:8] or ["hq_no", "spec", "feishu_match"],
                        "related_evidence_ids": missing_context_ids[:3],
                        "questions": [
                            {
                                "question_id": "dfmea-missing-context-1",
                                "question": "请补充缺失关键器件的 HQ 料号、规格型号或芯片类别；如果暂时无法确认，也请说明人工待查。",
                                "applies_to": {"field": "hq_no/spec/chip_type"},
                                "missing_fields": missing_fields[:8] or ["hq_no", "spec", "feishu_match"],
                                "related_evidence_ids": missing_context_ids[:3],
                            }
                        ],
                    }
                }
            elif profile == "review_checklist_qa" or "review_checklist_qa" in capability_profiles:
                final_answer = "本地 mock agent 已完成一次 review checklist 检索；请把命中的历史问题模式和当前报告证据结合复核。"
                payload = {
                    "final_answer": final_answer,
                    "confidence": "mock",
                    "citations": [
                        {
                            "id": evidence_id,
                            "note": "mock provider 引用最近一次 review checklist 工具返回的证据节点。",
                        }
                        for evidence_id in evidence_ids[:3]
                    ],
                    "proposed_actions": [
                        {
                            "title": "将 checklist 命中项映射到当前项目证据",
                            "reason": "历史 review 问题只能作为模式参考，仍需当前报告、位号、网络或页面 evidence 支撑。",
                            "priority": "manual_review",
                        }
                    ],
                }
            elif profile == "agent_ref_qa" or "agent_ref_qa" in capability_profiles:
                final_answer = "本地 mock agent 已完成一次 ref PDF 检索；请以返回的 PDF 片段和 evidence 引用为准。"
                payload = {
                    "final_answer": final_answer,
                    "confidence": "mock",
                    "citations": [
                        {
                            "id": evidence_id,
                            "note": "mock provider 引用最近一次 ref PDF 工具返回的证据节点。",
                        }
                        for evidence_id in evidence_ids[:3]
                    ],
                    "proposed_actions": [
                        {
                            "title": "继续补充 ref PDF 样本",
                            "reason": "mock provider 只验证 harness 链路，真实回答质量取决于索引内容和 Aster 配置。",
                            "priority": "manual_review",
                        }
                    ],
                }
            elif profile == "feishu_bom_qa" or "feishu_bom_qa" in capability_profiles:
                final_answer = "本地 mock agent 已完成一次飞书缓存搜索，请以返回的物料证据为准；无命中时建议补充 HQ 料号、规格型号、PI 或选型顺序。"
                payload = {
                    "final_answer": final_answer,
                    "confidence": "mock",
                    "citations": [
                        {
                            "id": evidence_id,
                            "note": "mock provider 引用最近一次工具返回的证据节点。",
                        }
                        for evidence_id in evidence_ids[:3]
                    ],
                    "proposed_actions": [
                        {
                            "title": "人工复核高优先级候选",
                            "reason": "mock provider 只验证 harness 链路，不替代真实工程判断。",
                            "priority": "manual_review",
                        }
                    ],
                }
            else:
                if profile == "dfmea_prep" or "dfmea_prep" in capability_profiles:
                    final_answer = "本地 mock agent 已完成 DFMEA 准备度扫描；第一阶段只输出元件身份、证据缺口和人工补充建议，不生成正式 DFMEA 风险结论。"
                else:
                    final_answer = "本地 mock agent 已完成一次工具观察，建议优先复核已返回的表格证据。"
                payload = {
                    "final_answer": final_answer,
                    "confidence": "mock",
                    "citations": [
                        {
                            "id": evidence_id,
                            "note": "mock provider 引用最近一次工具返回的证据节点。",
                        }
                        for evidence_id in evidence_ids[:3]
                    ],
                    "proposed_actions": [
                        {
                            "title": "人工复核高优先级候选",
                            "reason": "mock provider 只验证 harness 链路，不替代真实工程判断。",
                            "priority": "manual_review",
                        }
                    ],
                }
        return HarnessModelResponse(
            answer=json.dumps(payload, ensure_ascii=False),
            provider=self.provider,
            mode=self.mode,
            metadata={"prompt_chars": len(prompt), "agent_observation_count": len(observations)},
        )


class AsterHarnessModelProvider:
    """Aster-backed model provider. It never executes local tools."""

    def __init__(self, ask_model: Optional[Callable[..., dict]] = None):
        self._ask_model = ask_model

    def generate(self, prompt: str, *, inputs: Optional[dict] = None) -> HarnessModelResponse:
        ask_model = self._ask_model
        if ask_model is None:
            from pstx_integrations.aster.service import ask_aster_model

            ask_model = ask_aster_model
        payload = ask_model(prompt, inputs=inputs or {})
        return HarnessModelResponse(
            answer=str(payload.get("answer") or ""),
            provider=str(payload.get("provider") or "aster"),
            mode=str(payload.get("mode") or "live"),
            metadata=dict(payload.get("metadata") or {}),
        )

    def generate_agent_step(self, prompt: str, *, inputs: Optional[dict] = None) -> HarnessModelResponse:
        return self.generate(prompt, inputs=inputs)
