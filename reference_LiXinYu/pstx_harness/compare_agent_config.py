# -*- coding: utf-8 -*-
"""Configuration and request types for the compare agent."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

from pstx_agent_runtime import (
    list_public_profiles,
    profile_config as runtime_profile_config,
)
from pstx_harness.skill_tools import HARNESS_SKILL_TOOL_NAMES
from pstx_harness.review import HarnessError


COMPARE_AGENT_MAX_STEPS = 24
COMPARE_AGENT_MAX_TOOL_CALLS = 48
COMPARE_AGENT_MODEL_OBSERVATION_LIMIT = 8
COMPARE_AGENT_MODEL_NODE_LIMIT = 10
COMPARE_AGENT_MODEL_JSON_BUDGET = 42000
COMPARE_AGENT_MAX_TOOL_BATCH_CALLS = 4
COMPARE_AGENT_CAPABILITY_RULES = [
    ("compare_datasheet_qa", ["datasheet", "pdf", "规格书", "手册", "芯片型号", "参数", "absolute maximum", "recommended operating", "electrical characteristics"]),
    ("compare_cadence_pages", ["第", "页", "page", "page1", "page.map", "sch_1", "csa", "csv", "cadence", "原始文件", "用户看到"]),
    ("compare_pin_net", ["pin", "net", "网络", "引脚", "连接", "串阻", "上下拉", "连接器"]),
    ("compare_bom_feishu", ["bom", "飞书", "hq", "hq料号", "pi", "规格", "选型顺序", "料号", "part number"]),
    ("compare_key_devices", ["芯片", "连接器", "关键器件", "pu", "xu", "u", "新增", "删除", "少了", "多了"]),
    ("compare_page_mapping", ["页码映射", "主模块页", "真实页", "逻辑页", "module_order", "page.map", "页面统计"]),
]


COMPARE_AGENT_PROFILES: Dict[str, dict] = {
    "auto": {
        "title": "智能自动",
        "description": "根据用户问题组合多个 compare 审查能力，不再把复合问题提前收窄到单一 profile。",
        "tools": [],
        "default_question": "请根据当前 A/B 对比和用户问题自动组合对比能力，给出证据和人工复核建议。",
        "max_steps": 14,
        "max_tool_calls": 28,
    },
    "compare_quick_scan": {
        "title": "对比快速扫描",
        "description": "快速定位 A/B 对比中最高优先级的工程风险。",
        "tools": ["list_compare_sections", "summarize_compare_risks", "get_compare_section_rows", "query_compare_diff", "batch_query_compare_diff", "batch_get_compare_rows"],
        "default_question": "请快速扫描当前项目对比，列出最需要优先人工复核的差异。",
        "max_steps": 8,
        "max_tool_calls": 14,
    },
    "compare_key_devices": {
        "title": "关键器件差异",
        "description": "聚焦芯片、PU、XU、连接器和非 R/C/L 关键器件增删。",
        "tools": ["list_compare_sections", "get_compare_section_rows", "query_compare_diff", "batch_query_compare_diff", "get_compare_row", "batch_get_compare_rows"],
        "default_question": "请重点检查关键器件新增、删除和关键属性变化。",
        "max_steps": 10,
        "max_tool_calls": 18,
    },
    "compare_pin_net": {
        "title": "Pin/Net 差异",
        "description": "聚焦关键器件和 R/C/L 的 Pin-Net 连接变化。",
        "tools": ["list_compare_sections", "get_compare_section_rows", "query_compare_diff", "batch_query_compare_diff", "get_compare_row", "batch_get_compare_rows"],
        "default_question": "请重点检查芯片、连接器和 R/C/L 的 Pin-Net 连接变化风险。",
        "max_steps": 12,
        "max_tool_calls": 22,
    },
    "compare_bom_feishu": {
        "title": "BOM/飞书差异",
        "description": "聚焦 HQ 料号、规格、PI、选型顺序变化，并可按变化料号/型号检索本地 datasheet 证据。",
        "tools": ["list_compare_sections", "get_compare_section_rows", "query_compare_diff", "batch_query_compare_diff", "get_compare_row", "batch_get_compare_rows", "list_datasheet_review_templates", "get_datasheet_review_template", "search_datasheet_chunks", "batch_search_datasheet_chunks", "search_datasheet_parameters", "get_datasheet_parameter", "get_datasheet_chunk"],
        "default_question": "请重点检查 HQ 料号、规格、PI 和选型顺序变化。",
        "max_steps": 10,
        "max_tool_calls": 18,
    },
    "compare_datasheet_qa": {
        "title": "对比规格书证据",
        "description": "在项目对比上下文中复用本地 datasheet PDF 证据库，按 A/B 差异中的 HQ/型号/参数检索规格书证据。",
        "tools": [
            "list_compare_sections",
            "query_compare_diff",
            "batch_query_compare_diff",
            "get_compare_row",
            "list_datasheet_review_templates",
            "get_datasheet_review_template",
            "list_datasheet_documents",
            "search_datasheet_chunks",
            "batch_search_datasheet_chunks",
            "search_datasheet_parameters",
            "get_datasheet_parameter",
            "get_datasheet_chunk",
            "get_datasheet_page_excerpt",
        ],
        "default_question": "请结合 A/B 对比差异和本地 datasheet PDF 证据，回答用户关于规格书、参数或芯片型号的问题。",
        "max_steps": 12,
        "max_tool_calls": 24,
    },
    "compare_page_mapping": {
        "title": "页码映射差异",
        "description": "聚焦主模块页、页码和页码检查表变化；对用户提到的页默认按页码理解。",
        "tools": [
            "list_compare_sections",
            "get_compare_section_rows",
            "query_compare_diff",
            "batch_query_compare_diff",
            "batch_get_compare_rows",
            "list_compare_project_files",
            "read_compare_project_text",
        ],
        "default_question": "请重点检查 A/B 项目主模块页/页码映射和页码统计差异。",
        "max_steps": 12,
        "max_tool_calls": 22,
    },
    "compare_cadence_pages": {
        "title": "Cadence 页语义比对",
        "description": "按页码读取 sch_1/pageX.csv|csa，做 Cadence 页面对象和连接拓扑语义级比对。",
        "tools": [
            "resolve_compare_page_range",
            "compare_cadence_page_semantics",
            "get_cadence_page_object",
            "batch_get_cadence_page_objects",
            "get_cadence_page_raw_excerpt",
            "query_compare_diff",
            "batch_query_compare_diff",
        ],
        "default_question": "请按页码比对指定页范围内的 Cadence CSV/CSA 语义差异。",
        "max_steps": 14,
        "max_tool_calls": 28,
    },
    "compare_full_review": {
        "title": "完整项目对比审查",
        "description": "允许读取全部 compare 只读工具和受限项目文件。",
        "tools": ["*"],
        "default_question": "请综合审查当前 A/B 项目差异，列出优先级、证据和人工复核建议。",
        "max_steps": 24,
        "max_tool_calls": 48,
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


def list_compare_agent_profiles() -> List[dict]:
    return [
        _append_global_profile_tools(profile)
        for profile in list_public_profiles(COMPARE_AGENT_PROFILES)
    ]


def profile_config(profile: str) -> dict:
    return runtime_profile_config(COMPARE_AGENT_PROFILES, profile, default_profile="compare_quick_scan")


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


@dataclass(frozen=True)
class CompareAgentRequest:
    profile: str = "compare_quick_scan"
    question: str = ""
    max_steps: int = 8
    max_tool_calls: int = 14
    detail_limit: int = 500
    debug: bool = False

    @classmethod
    def from_mapping(cls, value: Optional[dict]) -> "CompareAgentRequest":
        value = value or {}
        profile = str(value.get("profile") or "compare_quick_scan").strip() or "compare_quick_scan"
        config = COMPARE_AGENT_PROFILES.get(profile)
        if config is None:
            raise HarnessError(f"未知 compare agent profile：{profile}")
        try:
            request = cls(
                profile=profile,
                question=str(value.get("question") or config["default_question"]).strip()[:2200],
                max_steps=int(value.get("max_steps", config["max_steps"])),
                max_tool_calls=int(value.get("max_tool_calls", config["max_tool_calls"])),
                detail_limit=int(value.get("detail_limit", 500)),
                debug=_as_bool(value.get("debug", False), False),
            )
        except (TypeError, ValueError) as exc:
            raise HarnessError("max_steps、max_tool_calls、detail_limit 必须是数字。") from exc
        request.validate()
        return request

    def validate(self) -> None:
        if self.profile not in COMPARE_AGENT_PROFILES:
            raise HarnessError(f"未知 compare agent profile：{self.profile}")
        if self.max_steps < 1 or self.max_steps > COMPARE_AGENT_MAX_STEPS:
            raise HarnessError(f"max_steps 必须在 1 到 {COMPARE_AGENT_MAX_STEPS} 之间。")
        if self.max_tool_calls < 0 or self.max_tool_calls > COMPARE_AGENT_MAX_TOOL_CALLS:
            raise HarnessError(f"max_tool_calls 必须在 0 到 {COMPARE_AGENT_MAX_TOOL_CALLS} 之间。")
        if self.detail_limit < 1 or self.detail_limit > 5000:
            raise HarnessError("detail_limit 必须在 1 到 5000 之间。")
