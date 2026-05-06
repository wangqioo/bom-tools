# -*- coding: utf-8 -*-
"""Configuration and request types for the report harness agent."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from pstx_agent_runtime import (
    list_public_profiles,
    profile_config as runtime_profile_config,
)
from pstx_harness.skill_tools import HARNESS_SKILL_TOOL_NAMES
from pstx_harness.review import HarnessError


HARNESS_AGENT_MAX_STEPS = 24
HARNESS_AGENT_MAX_TOOL_CALLS = 48
HARNESS_AGENT_MAX_SUBAGENTS = 6
HARNESS_AGENT_DEFAULT_MAX_SUBAGENTS = 2
HARNESS_AGENT_MODEL_OBSERVATION_LIMIT = 6
HARNESS_AGENT_MODEL_NODE_LIMIT = 8
HARNESS_AGENT_MODEL_JSON_BUDGET = 36000
HARNESS_AGENT_MODEL_TEXT_LIMIT = 900
HARNESS_AGENT_MAX_TOOL_BATCH_CALLS = 4

PROJECT_MEMORY_TOOL_NAMES = (
    "list_project_memory_evidence",
    "get_project_memory_evidence",
    "batch_get_project_memory_evidence",
)

HARNESS_AGENT_CAPABILITY_RULES = [
    ("connection_datasheet_review", ["datasheet 连接", "规格书 连接", "mineru 连接", "反查连接", "连接是否", "连接关系是否", "网表证据", "接口电平", "电源域", "power sequence", "reset timing", "strap", "clock requirement"]),
    ("review_checklist_qa", ["review checklist", "ref_checklist", "检查清单", "审查清单", "review经验", "review 问题", "changelist", "历史问题", "真实review"]),
    ("agent_ref_qa", ["ref", "参考资料", "资料库", "文档", "manual", "能力边界", "agent lab"]),
    ("datasheet_qa", ["datasheet", "pdf", "规格书", "手册", "芯片型号", "参数", "absolute maximum", "recommended operating", "electrical characteristics"]),
    ("dfmea_prep", ["dfmea", "失效", "失效模式", "后果", "测试方案", "准备度", "规格书", "datasheet", "pdf", "手册", "芯片型号"]),
    ("chip_topology", ["拓扑", "连接关系", "芯片级", "大芯片", "电平转换", "level shifter", "translator", "互联", "连接到", "连接了哪些", "review task", "part name", "part_name", "服务器", "server"]),
    ("document_search", ["文档", "资料", "搜索", "关键词", "段落", "上下文", "excerpt", "document"]),
    ("feishu_bom_qa", ["飞书", "缓存", "hq", "hq料号", "物料", "规格型号", "part number", "pi", "选型顺序"]),
    ("bom_depop", ["bom_option", "depop", "dnp", "不贴", "贴装", "打圈", "画圈", "pop"]),
    ("page_mapping", ["页", "page", "page.map", "module_order", "主模块页", "真实页", "用户看到", "sch_1", "页码", "原始文件", "源文件", "raw", "source", "追溯", "底层证据", "grep", "搜索原始"]),
    ("resistor_bias", ["串阻", "上下拉", "上拉", "下拉", "od", "oc", "pin", "引脚", "偏置", "电阻"]),
    ("derating", ["降额", "derating", "电容", "耐压", "ac耦合", "ac coupling"]),
    ("csa_geometry", ["csa", "dot", "circle", "arc", "几何", "十字", "坐标"]),
]

HARNESS_AGENT_PROFILES: Dict[str, dict] = {
    "auto": {
        "title": "智能自动",
        "description": "根据用户问题组合多个审查能力，不再把复合问题提前收窄到单一 profile。",
        "tools": [],
        "default_question": "请根据当前项目和用户问题自动组合审查能力，给出证据和人工复核建议。",
        "max_steps": 12,
        "max_tool_calls": 24,
        "subagent_profiles": [],
    },
    "quick_scan": {
        "title": "快速扫描",
        "description": "先用少量表格证据定位最值得人工复核的风险点。",
        "tools": ["list_report_tables", "summarize_schematic_page_count", "get_table_rows", "summarize_table_column_values", "batch_get_table_rows", "batch_query_report_entities", "get_evidence_pack", "search_project_text", "trace_project_source", "read_project_text"],
        "default_question": "请快速扫描当前报告，找出最需要优先复核的审查项。",
        "max_steps": 6,
        "max_tool_calls": 10,
        "subagent_profiles": ["bom_depop", "page_mapping", "resistor_bias"],
    },
    "bom_depop": {
        "title": "BOM/DEPOP",
        "description": "聚焦 BOM_OPTION、DEPOP 和画圈覆盖关系。",
        "tools": ["get_evidence_pack", "get_table_rows", "summarize_table_column_values", "batch_get_table_rows", "query_report_entity", "batch_query_report_entities"],
        "default_question": "请检查 BOM_OPTION/DEPOP 及画圈覆盖是否存在需要人工复核的地方。",
        "max_steps": 8,
        "max_tool_calls": 14,
        "subagent_profiles": [],
    },
    "page_mapping": {
        "title": "页码映射",
        "description": "聚焦主模块页、页码、page.map、module_order 和页码统计一致性；对用户提到的页默认按页码理解。",
        "tools": ["get_evidence_pack", "summarize_schematic_page_count", "get_table_rows", "summarize_table_column_values", "batch_get_table_rows", "batch_query_report_entities", "list_project_files", "search_project_text", "trace_project_source", "read_project_text"],
        "default_question": "请检查主模块页到页码的映射是否存在不一致或证据不足。",
        "max_steps": 8,
        "max_tool_calls": 14,
        "subagent_profiles": [],
    },
    "resistor_bias": {
        "title": "电阻与偏置",
        "description": "聚焦芯片 Pin、串阻、上下拉和重复偏置候选。",
        "tools": ["get_evidence_pack", "get_table_rows", "summarize_table_column_values", "batch_get_table_rows", "query_report_entity", "batch_query_report_entities"],
        "default_question": "请检查芯片 Pin 的串阻、上下拉和偏置候选是否需要人工复核。",
        "max_steps": 8,
        "max_tool_calls": 14,
        "subagent_profiles": [],
    },
    "derating": {
        "title": "电容降额",
        "description": "聚焦电容降额不合格、无法判断和电压来源边界。",
        "tools": ["get_evidence_pack", "get_table_rows", "summarize_table_column_values", "batch_get_table_rows", "query_report_entity", "batch_query_report_entities"],
        "default_question": "请检查电容降额结果，区分明确风险和需要人工确认的边界。",
        "max_steps": 8,
        "max_tool_calls": 14,
        "subagent_profiles": [],
    },
    "csa_geometry": {
        "title": "CSA 几何规范",
        "description": "聚焦 DOT 十字交叉、画圈对象和 CSA 页面几何证据。",
        "tools": ["get_evidence_pack", "get_table_rows", "summarize_table_column_values", "batch_get_table_rows", "batch_query_report_entities", "list_project_files", "search_project_text", "trace_project_source", "read_project_text"],
        "default_question": "请检查 CSA 几何规范候选，说明哪些需要看页面坐标复核。",
        "max_steps": 8,
        "max_tool_calls": 14,
        "subagent_profiles": [],
    },
    "feishu_bom_qa": {
        "title": "飞书缓存问答",
        "description": "只读搜索本地飞书 BOM 缓存，回答 HQ 料号、规格型号、PI、选型顺序和扩展字段问题。",
        "tools": ["feishu_bom", "list_feishu_cache_libraries", "search_feishu_cache_rows", "batch_search_feishu_cache_rows", "get_feishu_cache_row"],
        "default_question": "请围绕本地飞书缓存回答物料库问题；如果关键词不足，请提示补充 HQ 料号、规格型号、PI 或选型顺序。",
        "max_steps": 8,
        "max_tool_calls": 14,
        "subagent_profiles": [],
    },
    "datasheet_qa": {
        "title": "Datasheet PDF 问答",
        "description": "只读检索本地 datasheet PDF chunk，围绕规格书参数、工作条件、型号说明做可引用问答。",
        "tools": [
            "list_datasheet_sources",
            "list_datasheet_review_templates",
            "get_datasheet_review_template",
            "list_datasheet_documents",
            "search_datasheet_chunks",
            "batch_search_datasheet_chunks",
            "search_datasheet_parameters",
            "get_datasheet_parameter",
            "get_datasheet_chunk",
            "get_datasheet_page_excerpt",
            "search_datasheets",
            "get_datasheet_excerpt",
        ],
        "default_question": "请先检索本地 datasheet PDF 证据，再回答用户的规格书/参数问题；高风险或定量参数必须读取 detail chunk 后再下结论。",
        "max_steps": 10,
        "max_tool_calls": 18,
        "subagent_profiles": [],
    },
    "dfmea_prep": {
        "title": "DFMEA 准备",
        "description": "构建元件身份卡、物料证据和 DFMEA 输入缺口；第一阶段不输出正式 DFMEA 风险结论。",
        "tools": [
            "summarize_dfmea_readiness",
            "list_component_identity_cards",
            "get_component_identity_card",
            "batch_get_component_identity_cards",
            "search_component_identity_cards",
            "list_feishu_cache_libraries",
            "search_feishu_cache_rows",
            "batch_search_feishu_cache_rows",
            "get_feishu_cache_row",
            "list_datasheet_sources",
            "list_datasheet_review_templates",
            "get_datasheet_review_template",
            "list_datasheet_documents",
            "search_datasheet_chunks",
            "batch_search_datasheet_chunks",
            "search_datasheet_parameters",
            "get_datasheet_parameter",
            "get_datasheet_chunk",
            "get_datasheet_page_excerpt",
            "search_datasheets",
            "get_datasheet_excerpt",
            "match_component_datasheets",
            "batch_match_component_datasheets",
            "summarize_dfmea_datasheet_coverage",
        ],
        "default_question": "请评估当前项目做 DFMEA 的输入准备度，列出可优先分析的芯片/连接器和缺失上下文。",
        "max_steps": 10,
        "max_tool_calls": 18,
        "subagent_profiles": [],
    },
    "connection_datasheet_review": {
        "title": "连接 × Datasheet 反查",
        "description": "先读原理图/网表/拓扑连接 evidence，再按相关元件读取 MinerU-backed datasheet detail，反查电源、接口、复位、时钟和 strap 连接风险。",
        "tools": [
            "list_harness_skills",
            "select_harness_skills",
            "get_harness_skill",
            "list_business_dictionary",
            "get_evidence_pack",
            "batch_query_report_entities",
            "query_report_entity",
            "summarize_llm_topology_netlist",
            "summarize_topology_review_tasks",
            "query_llm_topology_netlist",
            "batch_query_llm_topology_netlist",
            "get_llm_topology_node",
            "get_llm_topology_edge",
            "get_topology_review_task",
            "batch_expand_topology_review_tasks",
            "list_component_identity_cards",
            "search_component_identity_cards",
            "get_component_identity_card",
            "batch_get_component_identity_cards",
            "summarize_dfmea_datasheet_coverage",
            "list_datasheet_sources",
            "list_datasheet_review_templates",
            "get_datasheet_review_template",
            "list_datasheet_documents",
            "match_component_datasheets",
            "batch_match_component_datasheets",
            "search_datasheet_parameters",
            "get_datasheet_parameter",
            "search_datasheet_chunks",
            "batch_search_datasheet_chunks",
            "get_datasheet_chunk",
            "get_datasheet_page_excerpt",
            "search_datasheets",
            "get_datasheet_excerpt",
            "search_project_text",
            "trace_project_source",
            "read_project_text",
        ],
        "default_question": "请先解读用户问题，再读取原理图网表/拓扑连接证据；按相关元器件匹配并读取 MinerU-backed datasheet detail；最后用 datasheet 事实反查电源、接口电平、复位、时钟、strap 或 level shifting 连接是否存在风险。",
        "max_steps": 14,
        "max_tool_calls": 28,
        "subagent_profiles": [],
    },
    "chip_topology": {
        "title": "芯片级拓扑",
        "description": "抽取 U/PU/XU 等芯片层面的模糊连接关系，帮助理解大芯片、电平转换、电源管理等芯片间互联。",
        "tools": [
            "list_business_dictionary",
            "summarize_llm_topology_netlist",
            "summarize_topology_review_tasks",
            "query_llm_topology_netlist",
            "batch_query_llm_topology_netlist",
            "get_llm_topology_node",
            "get_llm_topology_edge",
            "get_topology_review_task",
            "batch_expand_topology_review_tasks",
            "list_component_identity_cards",
            "get_component_identity_card",
            "batch_get_component_identity_cards",
            "search_component_identity_cards",
            "search_project_text",
            "trace_project_source",
            "read_project_text",
        ],
        "default_question": "请从芯片级拓扑角度概括当前项目的大芯片、电平转换、电源管理等芯片之间的主要连接关系，并指出需要人工复核的模糊关系。",
        "max_steps": 10,
        "max_tool_calls": 18,
        "subagent_profiles": [],
    },
    "document_search": {
        "title": "文档搜索",
        "description": "只读搜索本地 harness 文档根目录，先按关键词命中文档片段，再读取命中位置附近上下文。",
        "tools": ["list_document_search_sources", "search_documents", "batch_search_documents", "get_document_excerpt"],
        "default_question": "请先搜索本地文档关键词，再读取关键命中片段附近上下文，基于文档 evidence 回答用户问题。",
        "max_steps": 8,
        "max_tool_calls": 14,
        "subagent_profiles": [],
    },
    "agent_ref_qa": {
        "title": "ref PDF 问答",
        "description": "只读检索 Agent Lab 的 ref/ PDF 资料，用于测试 Agent 当前参考资料读取和证据引用能力。",
        "tools": ["list_agent_ref_sources", "search_agent_ref_pdfs", "get_agent_ref_pdf_excerpt"],
        "default_question": "请基于 ref/ 文件夹中的本地 PDF 资料回答用户问题；没有命中时说明无命中并建议换关键词。",
        "max_steps": 8,
        "max_tool_calls": 14,
        "subagent_profiles": [],
    },
    "review_checklist_qa": {
        "title": "Review checklist",
        "description": "只读检索 ref_checklist/ 中的真实原理图 review 问题、Excel 记录和 changelist，用历史问题模式辅助当前项目审查。",
        "tools": ["list_review_checklist_sources", "search_review_checklists", "get_review_checklist_excerpt", "batch_query_report_entities", "get_evidence_pack", "summarize_table_column_values"],
        "default_question": "请先检索 ref_checklist/ 审查经验清单，再把命中的历史问题模式迁移到当前原理图 review；没有命中时说明无命中并按当前报告证据继续审查。",
        "max_steps": 10,
        "max_tool_calls": 18,
        "subagent_profiles": [],
    },
    "full_review": {
        "title": "完整审查",
        "description": "允许读取全部只读证据工具，适合综合问题和长链路复核。",
        "tools": ["*"],
        "default_question": "请综合审查当前报告，列出优先级、证据和人工复核建议。",
        "max_steps": 12,
        "max_tool_calls": 24,
        "subagent_profiles": ["bom_depop", "page_mapping", "resistor_bias", "derating", "csa_geometry"],
    },
}


def _append_global_profile_tools(profile: dict) -> dict:
    item = dict(profile)
    tools = list(item.get("tools") or [])
    if "*" not in tools:
        for tool_name in HARNESS_SKILL_TOOL_NAMES:
            if tool_name not in tools:
                tools.append(tool_name)
        item["tools"] = tools
    return item


def list_harness_agent_profiles() -> List[dict]:
    return [
        _append_global_profile_tools(profile)
        for profile in list_public_profiles(HARNESS_AGENT_PROFILES, include_subagents=True)
    ]


def profile_config(profile: str) -> dict:
    return runtime_profile_config(HARNESS_AGENT_PROFILES, profile, default_profile="quick_scan")


def default_subagent_profiles(profile: str) -> List[str]:
    configured = list(profile_config(profile).get("subagent_profiles") or [])
    if configured:
        return configured
    if profile in {"quick_scan", "full_review"}:
        return ["bom_depop", "page_mapping", "resistor_bias"]
    return []


@dataclass(frozen=True)
class HarnessAgentRequest:
    profile: str = "quick_scan"
    question: str = ""
    max_steps: int = 6
    max_tool_calls: int = 10
    max_rows_per_table: int = 12
    debug: bool = False
    enable_subagents: bool = False
    subagent_profiles: Tuple[str, ...] = ()
    max_subagents: int = HARNESS_AGENT_DEFAULT_MAX_SUBAGENTS
    context_answers: Tuple[dict, ...] = ()
    continue_agent_run_id: str = ""

    @classmethod
    def from_mapping(cls, value: Optional[dict]) -> "HarnessAgentRequest":
        value = value or {}
        profile = str(value.get("profile") or "quick_scan").strip() or "quick_scan"
        config = HARNESS_AGENT_PROFILES.get(profile)
        if config is None:
            raise HarnessError(f"未知 agent profile：{profile}")
        try:
            request = cls(
                profile=profile,
                question=str(value.get("question") or config["default_question"]).strip()[:2000],
                max_steps=int(value.get("max_steps", config["max_steps"])),
                max_tool_calls=int(value.get("max_tool_calls", config["max_tool_calls"])),
                max_rows_per_table=int(value.get("max_rows_per_table", 12)),
                debug=_as_bool(value.get("debug", False), False),
                enable_subagents=_as_bool(value.get("enable_subagents", value.get("include_subagents", False)), False),
                subagent_profiles=_parse_subagent_profiles(value.get("subagent_profiles")),
                max_subagents=int(value.get("max_subagents", HARNESS_AGENT_DEFAULT_MAX_SUBAGENTS)),
                context_answers=_parse_context_answers(value.get("context_answers")),
                continue_agent_run_id=str(value.get("continue_agent_run_id") or "").strip()[:80],
            )
        except (TypeError, ValueError) as exc:
            raise HarnessError("max_steps、max_tool_calls、max_rows_per_table、max_subagents 必须是数字，context_answers 必须是对象数组。") from exc
        if request.enable_subagents and not request.subagent_profiles:
            request = HarnessAgentRequest(
                profile=request.profile,
                question=request.question,
                max_steps=request.max_steps,
                max_tool_calls=request.max_tool_calls,
                max_rows_per_table=request.max_rows_per_table,
                debug=request.debug,
                enable_subagents=request.enable_subagents,
                subagent_profiles=tuple(config.get("subagent_profiles") or default_subagent_profiles(request.profile)),
                max_subagents=request.max_subagents,
                context_answers=request.context_answers,
                continue_agent_run_id=request.continue_agent_run_id,
            )
        request.validate()
        return request

    def validate(self) -> None:
        if self.profile not in HARNESS_AGENT_PROFILES:
            raise HarnessError(f"未知 agent profile：{self.profile}")
        if self.max_steps < 1 or self.max_steps > HARNESS_AGENT_MAX_STEPS:
            raise HarnessError(f"max_steps 必须在 1 到 {HARNESS_AGENT_MAX_STEPS} 之间。")
        if self.max_tool_calls < 0 or self.max_tool_calls > HARNESS_AGENT_MAX_TOOL_CALLS:
            raise HarnessError(f"max_tool_calls 必须在 0 到 {HARNESS_AGENT_MAX_TOOL_CALLS} 之间。")
        if self.max_rows_per_table < 1 or self.max_rows_per_table > 100:
            raise HarnessError("max_rows_per_table 必须在 1 到 100 之间。")
        if self.max_subagents < 0 or self.max_subagents > HARNESS_AGENT_MAX_SUBAGENTS:
            raise HarnessError(f"max_subagents 必须在 0 到 {HARNESS_AGENT_MAX_SUBAGENTS} 之间。")
        for profile in self.subagent_profiles:
            if profile not in HARNESS_AGENT_PROFILES:
                raise HarnessError(f"未知 subagent profile：{profile}")
            if profile == "full_review":
                raise HarnessError("subagent_profiles 暂不允许包含 full_review，避免并行审查范围过大。")
        for item in self.context_answers:
            if not str(item.get("question_id") or "").strip():
                raise HarnessError("context_answers 中每一项都必须包含 question_id。")
            if not str(item.get("answer") or "").strip():
                raise HarnessError("context_answers 中每一项都必须包含 answer。")


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


def _parse_subagent_profiles(value) -> Tuple[str, ...]:
    if value is None or value == "":
        return ()
    if isinstance(value, str):
        items = [item.strip() for item in value.replace("；", ",").replace(";", ",").split(",")]
    elif isinstance(value, (list, tuple, set)):
        items = [str(item).strip() for item in value]
    else:
        items = [str(value).strip()]
    result = []
    for item in items:
        if item and item not in result:
            result.append(item)
    return tuple(result)


def _parse_context_answers(value) -> Tuple[dict, ...]:
    if value is None or value == "":
        return ()
    if not isinstance(value, list):
        raise ValueError("context_answers must be a list")
    answers = []
    for index, item in enumerate(value[:24], start=1):
        if not isinstance(item, dict):
            raise ValueError("context_answers item must be object")
        question_id = str(item.get("question_id") or item.get("id") or f"q-{index}").strip()[:120]
        answer = str(item.get("answer") or "").strip()[:4000]
        applies_to = item.get("applies_to") if isinstance(item.get("applies_to"), dict) else {}
        answers.append({
            "question_id": question_id,
            "answer": answer,
            "applies_to": {
                str(key)[:80]: str(value if value is not None else "").replace("\r", " ").replace("\n", " ").strip()[:240]
                for key, value in list(applies_to.items())[:12]
            },
        })
    return tuple(answers)
