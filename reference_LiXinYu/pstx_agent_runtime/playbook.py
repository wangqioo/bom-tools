# -*- coding: utf-8 -*-
"""Deterministic playbook planning and tool-result contracts for PSTX agents."""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Mapping, Sequence


@dataclass(frozen=True)
class AgentPlaybook:
    id: str
    title: str
    triggers: tuple[str, ...]
    capability_profiles: tuple[str, ...]
    preferred_tools: tuple[str, ...]
    preferred_batch_tools: tuple[str, ...] = ()
    required_evidence: tuple[str, ...] = ()
    output_rules: tuple[str, ...] = ()
    anti_patterns: tuple[str, ...] = ()

    def matches(self, text: str, *, capability_profiles: Sequence[object] = ()) -> bool:
        normalized = str(text or "").lower()
        upper = normalized.upper()
        if any(str(profile) in self.capability_profiles for profile in capability_profiles or []):
            return True
        return any(str(token).lower() in normalized or str(token).upper() in upper for token in self.triggers)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "title": self.title,
            "triggers": list(self.triggers),
            "capability_profiles": list(self.capability_profiles),
            "preferred_tools": list(self.preferred_tools),
            "preferred_batch_tools": list(self.preferred_batch_tools),
            "required_evidence": list(self.required_evidence),
            "output_rules": list(self.output_rules),
            "anti_patterns": list(self.anti_patterns),
        }


@dataclass(frozen=True)
class AgentPlaybookPlan:
    selected_playbooks: tuple[AgentPlaybook, ...]
    allowed_tools: tuple[str, ...]
    recommended_first_tools: tuple[str, ...]
    evidence_goals: tuple[str, ...]
    anti_patterns: tuple[str, ...]
    planner_warnings: tuple[str, ...] = ()
    seeded_tool_calls: tuple[dict, ...] = ()

    def to_dict(self) -> dict:
        return {
            "selected_playbooks": [item.to_dict() for item in self.selected_playbooks],
            "allowed_tools": list(self.allowed_tools),
            "recommended_first_tools": list(self.recommended_first_tools),
            "evidence_goals": list(self.evidence_goals),
            "anti_patterns": list(self.anti_patterns),
            "planner_warnings": list(self.planner_warnings),
            "seeded_tool_calls": [dict(item) for item in self.seeded_tool_calls],
        }


@dataclass(frozen=True)
class ToolResultContract:
    completeness: str
    recommended_next_tools: tuple[str, ...] = ()
    detail_tool: dict | None = None
    aggregation_tool: dict | None = None
    scope_summary: str = ""

    def to_dict(self) -> dict:
        payload = {
            "completeness": self.completeness,
            "recommended_next_tools": list(self.recommended_next_tools),
            "scope_summary": self.scope_summary,
        }
        if self.detail_tool:
            payload["detail_tool"] = dict(self.detail_tool)
        if self.aggregation_tool:
            payload["aggregation_tool"] = dict(self.aggregation_tool)
        return payload


REPORT_AGENT_PLAYBOOKS: tuple[AgentPlaybook, ...] = (
    AgentPlaybook(
        id="schematic_page_count",
        title="原理图总页数统计",
        triggers=("原理图", "schematic", "module_order", "module_order.dat", "用户看到的总页"),
        capability_profiles=("page_mapping", "full_review"),
        preferred_tools=("summarize_schematic_page_count",),
        required_evidence=("schematic_page_count",),
        output_rules=("回答原理图总页数必须以 module_order(.dat) 页范围为准；page_rows 只能说明有记录/有元件的页面数。",),
        anti_patterns=("不要用 page_rows 的行数或唯一页面数回答原理图总页数。",),
    ),
    AgentPlaybook(
        id="table_column_aggregation",
        title="表格统计聚合",
        triggers=("统计", "唯一", "总数", "总页数", "top", "count", "多少", "覆盖范围", "page_rows", "表格"),
        capability_profiles=("quick_scan", "page_mapping", "full_review"),
        preferred_tools=("list_report_tables", "summarize_table_column_values"),
        preferred_batch_tools=("batch_get_table_rows",),
        required_evidence=("table_column_summary",),
        output_rules=("遇到截断表格 preview 时先聚合列值，不要用分页行样本直接给统计结论。",),
        anti_patterns=("不要为了统计唯一值循环调用 get_table_rows 拉全表。",),
    ),
    AgentPlaybook(
        id="source_file_drilldown",
        title="分析结果到原始文件追溯",
        triggers=("原始文件", "源文件", "raw", "source", "source trace", "追溯", "底层证据", "grep", "搜索原始", "pstxprt", "pstxnet", "csa", "csv"),
        capability_profiles=(),
        preferred_tools=("trace_project_source", "search_project_text", "read_project_text"),
        required_evidence=("source_trace", "file_excerpt"),
        output_rules=("当用户要求追到底层或原始文件时，先用 trace_project_source 通过位号/网络/页码/报告行定位 line-number excerpt；若用户要求 grep/广搜或 trace 无命中，再用 search_project_text 限定路径/后缀跨文件搜索；最后按 detail_tool 精读必要窗口。",),
        anti_patterns=("不要整文件读取大 PSTX/Cadence 文件来碰运气；不要把报告摘要当作原始文件证据。",),
    ),
    AgentPlaybook(
        id="schematic_datasheet_connection_review",
        title="原理图连接 × Datasheet 反查",
        triggers=("datasheet连接", "datasheet 连接", "规格书连接", "规格书 连接", "mineru连接", "反查连接", "连接是否", "连接风险", "网表证据", "接口电平", "电源域", "power sequence", "reset timing", "strap", "clock requirement"),
        capability_profiles=("connection_datasheet_review", "full_review"),
        preferred_tools=(
            "list_datasheet_sources",
            "list_datasheet_review_templates",
            "summarize_llm_topology_netlist",
            "summarize_topology_review_tasks",
            "batch_query_llm_topology_netlist",
            "batch_get_component_identity_cards",
            "batch_match_component_datasheets",
            "search_datasheet_parameters",
            "batch_search_datasheet_chunks",
            "get_datasheet_parameter",
            "get_datasheet_chunk",
            "get_llm_topology_edge",
            "get_llm_topology_node",
            "trace_project_source",
        ),
        required_evidence=(
            "llm_topology_edge",
            "llm_topology_node",
            "component_identity",
            "datasheet_match",
            "datasheet_parameter",
            "datasheet_chunk",
            "source_trace",
        ),
        output_rules=(
            "先说明用户问题里的位号、网络、接口或电源域目标；没有目标时先批量查询报告实体或拓扑摘要。",
            "原理图连接结论必须引用 topology edge/node、pin-net、source trace 或报告实体 evidence。",
            "datasheet 定量/电气事实必须读取 parameter/chunk/detail；search snippet 只能作为 locator。",
            "用 datasheet 事实反查连接时，至少覆盖电源 rail、IO 电平/电源域、reset/enable/clock/strap 中与问题相关的项。",
            "MinerU/datasheet 索引缺失、PDF 未命中或 detail 不足时输出 evidence gap，不猜 pass/fail。",
        ),
        anti_patterns=(
            "不要只看 datasheet 摘要就判断原理图连接正确。",
            "不要只看拓扑边就推断接口电平兼容。",
            "不要把 absolute maximum 当 recommended operating。",
            "不要把 topology 网表当完整电气签核网表。",
        ),
    ),
    AgentPlaybook(
        id="report_entity_batch_lookup",
        title="报告实体批量查询",
        triggers=("位号", "refdes", "网络", "net", "hq", "料号", "多个", "批量", "pin", "引脚"),
        capability_profiles=("quick_scan", "bom_depop", "page_mapping", "resistor_bias", "derating", "full_review"),
        preferred_tools=("batch_query_report_entities", "query_report_entity"),
        preferred_batch_tools=("batch_query_report_entities",),
        required_evidence=("component", "net", "table_row"),
        output_rules=("多个对象必须优先批量查询，并逐项说明 found/missing/needs_context。",),
        anti_patterns=("不要对多个位号或网络逐条反复 tool_call。",),
    ),
    AgentPlaybook(
        id="chip_level_topology",
        title="芯片级大拓扑取证",
        triggers=("拓扑", "网表", "语义网表", "llm topology", "芯片级", "连接关系", "大芯片", "电平转换", "level shifter", "translator", "互联", "连接到", "连接了哪些", "part name", "part_name", "服务器", "server", "review task"),
        capability_profiles=("chip_topology", "full_review"),
        preferred_tools=("summarize_llm_topology_netlist", "summarize_topology_review_tasks", "query_llm_topology_netlist", "get_topology_review_task", "get_llm_topology_edge", "get_llm_topology_node"),
        preferred_batch_tools=("batch_query_llm_topology_netlist", "batch_expand_topology_review_tasks", "batch_get_component_identity_cards"),
        required_evidence=("llm_topology_summary", "llm_topology_review_task", "llm_topology_edge", "llm_topology_node", "component_identity"),
        output_rules=("LLM 拓扑网表只表示 IC 节点之间的共享信号关系和一跳无源桥摘要；需要说明它是无方向、模糊、用于 review 定位的证据。", "PART_NAME 只能作为服务器项目器件身份提示，不能替代 datasheet/飞书/pin-net evidence。"),
        anti_patterns=("不要把 R/C/L 无源件路径扩展成芯片拓扑；不要在没有 pin/net 证据时推断电气方向或协议方向。",),
    ),
    AgentPlaybook(
        id="local_document_search",
        title="本地文档关键词取证",
        triggers=("文档", "资料", "关键词", "段落", "上下文", "搜索文档", "document", "excerpt"),
        capability_profiles=("document_search", "full_review"),
        preferred_tools=("list_document_search_sources", "search_documents", "get_document_excerpt"),
        preferred_batch_tools=("batch_search_documents",),
        required_evidence=("document_match", "document_excerpt"),
        output_rules=("回答文档内容问题必须先搜索文档 evidence；命中后如需解释上下文，应读取 get_document_excerpt。",),
        anti_patterns=("不要只凭文件名或搜索 snippet 下最终结论；不要读取白名单文档根目录外的文件。",),
    ),
    AgentPlaybook(
        id="feishu_material_qa",
        title="飞书物料缓存问答",
        triggers=("飞书", "hq料号", "hq", "物料", "规格型号", "part number", "pi", "选型顺序", "缓存"),
        capability_profiles=("feishu_bom_qa", "dfmea_prep", "full_review"),
        preferred_tools=("list_feishu_cache_libraries", "search_feishu_cache_rows", "get_feishu_cache_row"),
        preferred_batch_tools=("batch_search_feishu_cache_rows",),
        required_evidence=("feishu_material", "material_match"),
        output_rules=("回答物料问题必须引用本地缓存 evidence；无命中时说明无命中和建议关键词。",),
        anti_patterns=("不要在未命中缓存时凭经验补全物料字段。",),
    ),
    AgentPlaybook(
        id="dfmea_preparation",
        title="DFMEA 准备取证",
        triggers=("dfmea", "失效", "失效模式", "后果", "测试", "规格书", "datasheet", "pdf", "芯片类别"),
        capability_profiles=("dfmea_prep", "datasheet_qa", "full_review"),
        preferred_tools=("summarize_dfmea_readiness", "search_component_identity_cards", "search_datasheet_chunks", "get_datasheet_chunk", "match_component_datasheets"),
        preferred_batch_tools=("batch_get_component_identity_cards", "batch_match_component_datasheets", "batch_search_feishu_cache_rows", "batch_search_datasheet_chunks"),
        required_evidence=("component_identity", "datasheet_chunk", "datasheet_match", "missing_context"),
        output_rules=("第一阶段只输出 DFMEA 准备度、证据缺口和人工补充问题，不输出正式风险表。",),
        anti_patterns=("不要把 Aster 候选当作已确认失效模式库结论。",),
    ),
    AgentPlaybook(
        id="datasheet_pdf_qa",
        title="Datasheet PDF 问答取证",
        triggers=("datasheet", "pdf", "规格书", "手册", "芯片型号", "参数", "absolute maximum", "recommended operating", "electrical characteristics"),
        capability_profiles=("datasheet_qa", "dfmea_prep", "full_review"),
        preferred_tools=("list_datasheet_documents", "search_datasheet_chunks", "get_datasheet_chunk", "get_datasheet_page_excerpt"),
        preferred_batch_tools=("batch_search_datasheet_chunks",),
        required_evidence=("datasheet_chunk", "datasheet_excerpt", "datasheet_gap"),
        output_rules=("回答规格书问题必须先检索本地 PDF evidence；高风险、定量、电气极限结论必须读取 detail chunk/page 后再回答。",),
        anti_patterns=("不要只凭搜索 snippet 给定量参数结论；不要在未命中本地 PDF 时编造 datasheet 内容。",),
    ),
    AgentPlaybook(
        id="agent_ref_pdf_qa",
        title="Agent Lab ref PDF 问答",
        triggers=("ref", "参考资料", "资料库", "文档", "manual", "手册", "能力边界", "agent lab"),
        capability_profiles=("agent_ref_qa", "full_review"),
        preferred_tools=("list_agent_ref_sources", "search_agent_ref_pdfs", "get_agent_ref_pdf_excerpt"),
        required_evidence=("agent_ref_excerpt",),
        output_rules=("回答 ref PDF 问题必须引用 ref PDF evidence；无命中时说明无命中和建议关键词。",),
        anti_patterns=("不要在未检索 ref PDF 时凭经验回答文档内容。",),
    ),
    AgentPlaybook(
        id="review_checklist_experience",
        title="Review checklist 经验迁移",
        triggers=("review checklist", "ref_checklist", "检查清单", "审查清单", "review经验", "review 问题", "changelist", "历史问题", "真实review"),
        capability_profiles=("review_checklist_qa", "quick_scan", "full_review"),
        preferred_tools=("list_review_checklist_sources", "search_review_checklists", "get_review_checklist_excerpt"),
        preferred_batch_tools=("batch_query_report_entities",),
        required_evidence=("review_checklist_excerpt", "table_row", "component", "net"),
        output_rules=("参考 checklist 时必须区分历史问题模式和当前项目证据；最终建议要同时引用 checklist evidence 和当前报告 evidence。",),
        anti_patterns=("不要把历史 checklist 命中直接当作当前项目已发生的问题。",),
    ),
)


COMPARE_AGENT_PLAYBOOKS: tuple[AgentPlaybook, ...] = (
    AgentPlaybook(
        id="compare_diff_batch_lookup",
        title="项目差异批量定位",
        triggers=("对比", "差异", "不同", "新增", "删除", "变化", "多个", "芯片", "连接器", "网络", "pin", "net"),
        capability_profiles=("compare_quick_scan", "compare_key_devices", "compare_pin_net", "compare_bom_feishu", "compare_full_review"),
        preferred_tools=("list_compare_sections", "query_compare_diff", "summarize_compare_risks"),
        preferred_batch_tools=("batch_query_compare_diff", "batch_get_compare_rows"),
        required_evidence=("compare_diff", "compare_component", "compare_net"),
        output_rules=("多个差异对象必须优先批量查询，并按高风险/需人工复核分组。",),
        anti_patterns=("不要只读首屏 preview 就断言没有差异。",),
    ),
    AgentPlaybook(
        id="cadence_page_semantic_compare",
        title="Cadence 页级语义比对",
        triggers=("第", "页", "page", "sch_1", "csa", "csv", "cadence", "原始文件", "用户看到的真实页"),
        capability_profiles=("compare_cadence_pages", "compare_page_mapping", "compare_full_review"),
        preferred_tools=("resolve_compare_page_range", "compare_cadence_page_semantics", "get_cadence_page_object"),
        preferred_batch_tools=("batch_get_cadence_page_objects",),
        required_evidence=("cadence_page_model", "cadence_topology_diff", "cadence_graphic_object"),
        output_rules=("页码范围必须解释为用户看到的真实页，并引用 cadence_* evidence。",),
        anti_patterns=("不要把子模块内部页或逻辑页当作用户页范围。",),
    ),
    AgentPlaybook(
        id="compare_bom_feishu_material",
        title="对比 BOM/飞书物料差异",
        triggers=("飞书", "hq", "hq料号", "pi", "规格", "选型顺序", "bom", "料号", "part number"),
        capability_profiles=("compare_bom_feishu", "compare_datasheet_qa", "compare_full_review"),
        preferred_tools=("list_compare_sections", "query_compare_diff", "search_datasheet_chunks", "get_datasheet_chunk"),
        preferred_batch_tools=("batch_query_compare_diff", "batch_get_compare_rows", "batch_search_datasheet_chunks"),
        required_evidence=("compare_feishu_material", "compare_diff", "datasheet_chunk"),
        output_rules=("飞书字段差异要自然融入元件/Pin-Net 差异说明，不作为独立孤立结论。",),
        anti_patterns=("不要只比较 HQ 料号而忽略 PI、规格和选型顺序；不要把料号变化直接等同为规格书参数变化。",),
    ),
    AgentPlaybook(
        id="compare_datasheet_pdf_qa",
        title="对比规格书证据取证",
        triggers=("datasheet", "pdf", "规格书", "手册", "芯片型号", "参数", "absolute maximum", "recommended operating", "electrical characteristics"),
        capability_profiles=("compare_datasheet_qa", "compare_bom_feishu", "compare_full_review"),
        preferred_tools=("query_compare_diff", "search_datasheet_chunks", "get_datasheet_chunk", "get_datasheet_page_excerpt"),
        preferred_batch_tools=("batch_query_compare_diff", "batch_search_datasheet_chunks"),
        required_evidence=("compare_diff", "datasheet_chunk", "datasheet_excerpt", "datasheet_gap"),
        output_rules=("对比规格书问题必须同时区分 A/B 差异 evidence 和 datasheet evidence；定量参数必须读取 detail chunk/page 后再回答。",),
        anti_patterns=("不要只凭 compare row 的料号/型号变化推断电气参数；不要未命中本地 PDF 时编造规格书内容。",),
    ),
)


def _dedupe(items: Sequence[object], *, limit: int = 80) -> tuple[str, ...]:
    result: list[str] = []
    for item in items:
        text = str(item or "").strip()
        if text and text not in result:
            result.append(text)
        if len(result) >= limit:
            break
    return tuple(result)


def _question_entities(text: object) -> dict:
    source = str(text or "")
    refdes: list[str] = []
    hq_codes: list[str] = []
    keywords: list[str] = []
    for match in re.findall(r"\bHQ[0-9A-Z]{3,}\b", source, flags=re.IGNORECASE):
        hq_codes.append(match.upper())
    # Keep this conservative: component refs are short letter prefixes followed
    # by digits, optionally with submodule suffixes such as U46A10 or PC16A10.
    for match in re.findall(r"\b(?:P?[RUCL]\d+[A-Z]?\d*|P?C\d+[A-Z]?\d*|PU\d+[A-Z]?\d*|XU\d+[A-Z]?\d*|U\d+[A-Z]?\d*|J\d+[A-Z]?\d*|CN\d+[A-Z]?\d*)\b", source, flags=re.IGNORECASE):
        token = match.upper()
        if token.startswith("HQ"):
            continue
        refdes.append(token)
    # Net/spec-like tokens. Avoid adding tiny words or pure numbers.
    for match in re.findall(r"\b[A-Za-z][A-Za-z0-9_./+-]{3,}\b", source):
        token = match.strip()
        upper = token.upper()
        if upper in {"PAGE", "SCH_1", "CADENCE", "DFMEA", "BOM", "PIN", "NET", "PDF"}:
            continue
        if upper in {item.upper() for item in [*refdes, *hq_codes]}:
            continue
        if re.fullmatch(r"[A-Z]{1,3}\d+[A-Z]?\d*", upper):
            continue
        keywords.append(token)
    merged_keywords = _dedupe([*refdes, *hq_codes, *keywords], limit=20)
    return {
        "refdes": _dedupe(refdes, limit=20),
        "hq_codes": _dedupe(hq_codes, limit=20),
        "keywords": merged_keywords,
    }


def _page_range_from_question(text: object) -> tuple[int, int] | None:
    source = str(text or "")
    for pattern in (
        r"第\s*(\d+)\s*[-~—到至]\s*(\d+)\s*页",
        r"page\s*(\d+)\s*[-~—到至]\s*(\d+)",
        r"页\s*(\d+)\s*[-~—到至]\s*(\d+)",
    ):
        match = re.search(pattern, source, flags=re.IGNORECASE)
        if not match:
            continue
        start = max(1, int(match.group(1)))
        end = max(1, int(match.group(2)))
        if end < start:
            start, end = end, start
        return start, min(end, start + 59)
    single = re.search(r"第\s*(\d+)\s*页|page\s*(\d+)", source, flags=re.IGNORECASE)
    if single:
        page = max(1, int(single.group(1) or single.group(2)))
        return page, page
    return None


def _connection_review_terms(text: object) -> tuple[str, ...]:
    """Conservative datasheet/search terms for connection back-check questions."""
    source = str(text or "")
    lower = source.lower()
    terms: list[str] = []
    if any(token in lower for token in ("i2c", "scl", "sda", "接口电平", "电平兼容", "io", "i/o", "level")):
        terms.extend(["I2C", "interface voltage", "I/O voltage", "VIH VIL"])
    if any(token in lower for token in ("spi", "miso", "mosi", "sclk", "cs")):
        terms.extend(["SPI", "interface timing", "VIH VIL"])
    if any(token in lower for token in ("uart", "tx", "rx")):
        terms.extend(["UART", "interface voltage", "VIH VIL"])
    if any(token in lower for token in ("usb", "dp", "dn")):
        terms.extend(["USB", "differential input", "recommended operating"])
    if any(token in lower for token in ("pcie", "pci express", "clkreq", "perst")):
        terms.extend(["PCIe", "PERST", "CLKREQ", "reference clock"])
    if any(token in lower for token in ("电源", "供电", "rail", "power", "vdd", "vcc", "电源域")):
        terms.extend(["recommended operating voltage", "power supply", "power sequence"])
    if any(token in lower for token in ("reset", "复位", "rst", "por")):
        terms.extend(["reset timing", "power-on reset", "reset input"])
    if any(token in lower for token in ("clock", "时钟", "clk", "osc")):
        terms.extend(["clock requirement", "input clock", "clock timing"])
    if any(token in lower for token in ("strap", "boot", "启动", "配置脚", "采样")):
        terms.extend(["strap", "boot configuration", "sampling timing"])
    if any(token in lower for token in ("enable", "en", "pgood", "power good")):
        terms.extend(["enable timing", "power good", "PGOOD"])
    return _dedupe(terms, limit=16)


def _is_schematic_page_count_question(text: object) -> bool:
    source = str(text or "").lower()
    has_schematic = any(token in source for token in ("原理图", "schematic", "module_order", "module_order.dat"))
    has_count_intent = any(token in source for token in ("多少页", "几页", "总页", "页数", "一共有", "总共", "count"))
    return has_schematic and has_count_intent


def _is_project_grep_question(text: object) -> bool:
    source = str(text or "").lower()
    return any(token in source for token in ("grep", "搜索原始", "搜原始", "全文搜索", "跨文件搜索", "find in files"))


def _seed_call(name: str, args: Mapping[str, object], *, reason: str) -> dict:
    return {
        "name": name,
        "args": dict(args),
        "reason": reason,
        "source": "playbook_seed",
    }


def _seeded_tool_calls(*,
                       question: object,
                       matched_playbooks: Sequence[AgentPlaybook],
                       allowed_tools: Sequence[str]) -> tuple[dict, ...]:
    entities = _question_entities(question)
    allowed = set(allowed_tools or [])
    playbook_ids = {playbook.id for playbook in matched_playbooks}
    refdes = list(entities["refdes"])
    hq_codes = list(entities["hq_codes"])
    keywords = list(entities["keywords"])
    calls: list[dict] = []

    if (
        "schematic_page_count" in playbook_ids
        and "summarize_schematic_page_count" in allowed
        and _is_schematic_page_count_question(question)
    ):
        calls.append(_seed_call(
            "summarize_schematic_page_count",
            {},
            reason="本地 playbook 识别到原理图总页数问题，必须先读取 module_order(.dat) 页范围统计。",
        ))
    if "report_entity_batch_lookup" in playbook_ids and "batch_query_report_entities" in allowed and keywords:
        calls.append(_seed_call(
            "batch_query_report_entities",
            {"queries": keywords[:20], "limit_per_query": 10},
            reason="本地 playbook 从用户问题中提取多个位号/网络/HQ 关键词，优先批量查询报告实体。",
        ))
    if "source_file_drilldown" in playbook_ids and ({"trace_project_source", "search_project_text"} & allowed):
        page_range = _page_range_from_question(question)
        if keywords and _is_project_grep_question(question) and "search_project_text" in allowed:
            calls.append(_seed_call(
                "search_project_text",
                {"query": " ".join(keywords[:6]), "context_lines": 2, "limit": 12},
                reason="本地 playbook 识别到原始文件 grep/跨文件搜索问题，先在白名单项目文本中搜索行级片段。",
            ))
        elif keywords and "trace_project_source" in allowed:
            query = " ".join(keywords[:4])
            kind = "refdes" if refdes else ("net" if not hq_codes else "text")
            calls.append(_seed_call(
                "trace_project_source",
                {"query": query, "kind": kind, "limit": 8},
                reason="本地 playbook 识别到原始文件追溯问题，先用实体关键词定位 PSTX/Cadence line-number excerpt。",
            ))
        elif keywords and "search_project_text" in allowed:
            calls.append(_seed_call(
                "search_project_text",
                {"query": " ".join(keywords[:6]), "context_lines": 2, "limit": 12},
                reason="本地 playbook 识别到原始文件追溯问题，但当前只允许 grep 工具，先做受控跨文件搜索。",
            ))
        elif page_range and "trace_project_source" in allowed:
            calls.append(_seed_call(
                "trace_project_source",
                {"query": f"PAGE{page_range[0]}", "kind": "page", "limit": 8},
                reason="本地 playbook 识别到页级原始文件追溯问题，先定位对应 pageX.csv/csa 和页码映射片段。",
            ))
    summarize_topology_tool = (
        "summarize_llm_topology_netlist"
        if "summarize_llm_topology_netlist" in allowed
        else "summarize_chip_topology"
    )
    batch_topology_tool = (
        "batch_query_llm_topology_netlist"
        if "batch_query_llm_topology_netlist" in allowed
        else "batch_query_chip_topology"
    )
    connection_review_matched = "schematic_datasheet_connection_review" in playbook_ids
    if "chip_level_topology" in playbook_ids and not connection_review_matched and summarize_topology_tool in allowed:
        focus = refdes[0] if len(refdes) == 1 else ""
        args = {"limit": 30}
        if focus:
            args["focus_refdes"] = focus
        calls.append(_seed_call(
            summarize_topology_tool,
            args,
            reason="本地 playbook 识别到芯片级拓扑问题，先生成 LLM 拓扑网表摘要和证据卡。",
        ))
    if "chip_level_topology" in playbook_ids and not connection_review_matched and "summarize_topology_review_tasks" in allowed:
        args = {"limit": 20}
        if len(refdes) == 1:
            args["focus_refdes"] = refdes[0]
        calls.append(_seed_call(
            "summarize_topology_review_tasks",
            args,
            reason="本地 playbook 识别到拓扑 review 问题，先生成可执行复核任务队列，避免模型只看散乱边表。",
        ))
    if "chip_level_topology" in playbook_ids and not connection_review_matched and batch_topology_tool in allowed:
        topology_queries = _dedupe([*refdes, *keywords], limit=20)
        if topology_queries:
            calls.append(_seed_call(
                batch_topology_tool,
                {"queries": list(topology_queries), "limit_per_query": 8},
                reason="本地 playbook 从用户问题中提取位号/角色/网络关键词，批量查询 LLM 拓扑网表。",
            ))
    if connection_review_matched:
        connection_terms = list(_connection_review_terms(question))
        topology_queries = _dedupe([*refdes, *keywords, *connection_terms], limit=20)
        datasheet_queries = _dedupe([*hq_codes, *connection_terms, *keywords], limit=20)
        if "list_datasheet_sources" in allowed:
            calls.append(_seed_call(
                "list_datasheet_sources",
                {},
                reason="本地 playbook 识别到 datasheet 反查连接问题，先确认 MinerU-backed 本地规格书索引状态。",
            ))
        if "batch_query_llm_topology_netlist" in allowed and topology_queries:
            calls.append(_seed_call(
                "batch_query_llm_topology_netlist",
                {"queries": list(topology_queries), "limit_per_query": 8},
                reason="本地 playbook 从用户问题中提取位号/网络/接口关键词，先批量读取原理图网表/拓扑连接 evidence。",
            ))
        if "batch_get_component_identity_cards" in allowed and refdes:
            calls.append(_seed_call(
                "batch_get_component_identity_cards",
                {"refdes_list": refdes[:20]},
                reason="本地 playbook 已识别连接反查对象位号，先读取对应元器件身份卡、HQ、型号、pin-net 和 power/interface nets。",
            ))
        if "batch_match_component_datasheets" in allowed and refdes:
            calls.append(_seed_call(
                "batch_match_component_datasheets",
                {"refdes_list": refdes[:20], "limit_per_component": 4},
                reason="本地 playbook 已识别连接反查对象位号，按元器件身份匹配 MinerU-backed datasheet 候选。",
            ))
        if "search_datasheet_parameters" in allowed and datasheet_queries:
            calls.append(_seed_call(
                "search_datasheet_parameters",
                {"query": " ".join(list(datasheet_queries)[:8]), "limit": 16},
                reason="本地 playbook 根据接口/电源/复位/时钟关键词检索 datasheet 参数卡，避免只看搜索 snippet 下定量结论。",
            ))
        if "batch_search_datasheet_chunks" in allowed and datasheet_queries:
            calls.append(_seed_call(
                "batch_search_datasheet_chunks",
                {"queries": list(datasheet_queries)[:12], "limit_per_query": 6},
                reason="本地 playbook 根据用户问题中的接口、电源、复位、时钟或 strap 关键词批量检索 MinerU-backed datasheet chunk。",
            ))
        if "trace_project_source" in allowed:
            trace_query = " ".join([*refdes, *keywords][:4])
            if trace_query:
                calls.append(_seed_call(
                    "trace_project_source",
                    {"query": trace_query, "kind": "refdes" if refdes else "text", "limit": 8},
                    reason="本地 playbook 为连接反查保留到底层 PSTX/Cadence 原始文件的 line-number excerpt 路径。",
                ))
    if "local_document_search" in playbook_ids and "search_documents" in allowed:
        calls.append(_seed_call(
            "search_documents",
            {"query": str(question or "")[:240], "limit": 10},
            reason="本地 playbook 识别到文档关键词搜索问题，先在本地文档根目录搜索命中片段。",
        ))
    if "feishu_material_qa" in playbook_ids and "batch_search_feishu_cache_rows" in allowed:
        queries = hq_codes or keywords
        if queries:
            calls.append(_seed_call(
                "batch_search_feishu_cache_rows",
                {"queries": queries[:20], "limit_per_query": 10},
                reason="本地 playbook 从用户问题中提取物料关键词，优先批量查询飞书缓存。",
            ))
    if "dfmea_preparation" in playbook_ids and "batch_get_component_identity_cards" in allowed and refdes:
        calls.append(_seed_call(
            "batch_get_component_identity_cards",
            {"refdes_list": refdes[:20]},
            reason="本地 playbook 从用户问题中提取多个位号，优先批量读取 DFMEA 元件身份卡。",
        ))
    if "datasheet_pdf_qa" in playbook_ids and not connection_review_matched and "batch_search_datasheet_chunks" in allowed:
        queries = _dedupe([*hq_codes, *keywords], limit=20)
        if queries:
            calls.append(_seed_call(
                "batch_search_datasheet_chunks",
                {"queries": list(queries), "limit_per_query": 8},
                reason="本地 playbook 从用户问题中提取 datasheet/HQ/型号关键词，优先批量检索 PDF chunk。",
            ))
    if "compare_datasheet_pdf_qa" in playbook_ids and "batch_search_datasheet_chunks" in allowed:
        queries = _dedupe([*hq_codes, *keywords], limit=20)
        if queries:
            calls.append(_seed_call(
                "batch_search_datasheet_chunks",
                {"queries": list(queries), "limit_per_query": 8},
                reason="本地 compare playbook 从用户问题中提取 datasheet/HQ/型号关键词，复用 PDF chunk 证据库。",
            ))
    if "cadence_page_semantic_compare" in playbook_ids and "compare_cadence_page_semantics" in allowed:
        page_range = _page_range_from_question(question)
        if page_range:
            page_start, page_end = page_range
            calls.append(_seed_call(
                "compare_cadence_page_semantics",
                {
                    "page_start": page_start,
                    "page_end": page_end,
                    "include_raw_unknown": True,
                    "coordinate_tolerance": 0,
                    "max_diff_items": 48,
                },
                reason="本地 playbook 已解析用户看到的真实页范围，直接构建 Cadence 页级语义差异证据。",
            ))
    if "compare_diff_batch_lookup" in playbook_ids and "batch_query_compare_diff" in allowed and keywords:
        calls.append(_seed_call(
            "batch_query_compare_diff",
            {"queries": keywords[:20], "limit_per_query": 10},
            reason="本地 playbook 从用户问题中提取对比关键词，优先批量查询项目差异。",
        ))
    return tuple(calls)


def build_playbook_plan(*,
                        question: str,
                        capability_profiles: Sequence[object],
                        allowed_tools: Sequence[object],
                        playbooks: Sequence[AgentPlaybook]) -> AgentPlaybookPlan:
    text = str(question or "")
    matched = [
        playbook
        for playbook in playbooks
        if playbook.matches(text, capability_profiles=capability_profiles)
    ]
    recommended = _dedupe([
        tool
        for playbook in matched
        for tool in [*playbook.preferred_batch_tools, *playbook.preferred_tools]
    ])
    allowed = _dedupe(allowed_tools)
    allowed_set = set(allowed)
    allowed_recommended = tuple(tool for tool in recommended if tool in allowed_set)
    missing_recommended = tuple(tool for tool in recommended if tool not in allowed_set)
    warnings = tuple(
        f"推荐工具不在当前 profile 白名单中：{tool}"
        for tool in missing_recommended
    )
    seeded_tool_calls = _seeded_tool_calls(
        question=question,
        matched_playbooks=matched,
        allowed_tools=allowed,
    )
    return AgentPlaybookPlan(
        selected_playbooks=tuple(matched),
        allowed_tools=allowed,
        recommended_first_tools=allowed_recommended,
        evidence_goals=_dedupe([goal for playbook in matched for goal in playbook.required_evidence]),
        anti_patterns=_dedupe([pattern for playbook in matched for pattern in playbook.anti_patterns]),
        planner_warnings=warnings,
        seeded_tool_calls=seeded_tool_calls,
    )


def _as_int(value: object, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def build_tool_result_contract(tool_name: str, result: Mapping[str, object] | None) -> ToolResultContract:
    result = dict(result or {})
    name = str(tool_name or result.get("id") or "")
    if result.get("ok") is False or result.get("error"):
        return ToolResultContract(
            completeness="error",
            scope_summary=str(result.get("summary") or result.get("error") or name)[:500],
        )

    total_rows = _as_int(result.get("total_rows"), -1)
    returned_rows = len(result.get("rows") or []) if isinstance(result.get("rows"), list) else -1
    has_more = bool(result.get("has_more") or result.get("truncated") or result.get("input_truncated"))
    completeness = "complete"
    recommended: list[str] = []
    detail_tool = None
    aggregation_tool = None

    if has_more:
        completeness = "truncated"
    elif total_rows >= 0 and returned_rows >= 0 and returned_rows < total_rows:
        completeness = "partial"
    elif name in {"get_table_rows", "get_compare_section_rows"} and total_rows >= 0:
        completeness = "complete" if returned_rows >= total_rows else "preview"

    if name == "list_report_tables":
        recommended.extend(["batch_get_table_rows", "get_table_rows", "summarize_table_column_values"])
    elif name == "list_compare_sections":
        recommended.extend(["batch_query_compare_diff", "get_compare_section_rows"])
    elif name == "get_table_rows":
        table_id = str(result.get("table_id") or "")
        if result.get("next_offset") is not None:
            detail_tool = {
                "name": "get_table_rows",
                "args": {"table_id": table_id, "offset": result.get("next_offset"), "limit": result.get("limit")},
            }
            recommended.append("get_table_rows")
        if result.get("aggregation_hint") or has_more:
            if table_id == "page_rows":
                aggregation_tool = {
                    "name": "summarize_schematic_page_count",
                    "args": {},
                }
                recommended.insert(0, "summarize_schematic_page_count")
            else:
                aggregation_tool = {
                    "name": "summarize_table_column_values",
                    "args": {"table_id": table_id, "column": "<需要统计的列名>"},
                }
                recommended.insert(0, "summarize_table_column_values")
    elif name == "summarize_schematic_page_count":
        completeness = "complete" if result.get("available", True) else "error"
    elif name == "trace_project_source":
        completeness = str(result.get("completeness") or ("complete" if result.get("source_hits") else "missing"))
        if result.get("source_hits"):
            recommended.append("read_project_text")
            if isinstance(result.get("detail_tool"), Mapping):
                detail_tool = dict(result.get("detail_tool") or {})
        else:
            recommended.extend(["search_project_text", "list_project_files", "read_project_text"])
    elif name == "search_project_text":
        completeness = str(result.get("completeness") or ("complete" if result.get("source_hits") else "missing"))
        if result.get("source_hits"):
            recommended.append("read_project_text")
            if isinstance(result.get("detail_tool"), Mapping):
                detail_tool = dict(result.get("detail_tool") or {})
        else:
            recommended.extend(["trace_project_source", "list_project_files", "read_project_text"])
    elif name == "read_project_text":
        completeness = "truncated" if result.get("truncated") else ("complete" if result.get("content") else "missing")
        if result.get("truncated"):
            recommended.append("read_project_text")
    elif name == "batch_get_table_rows":
        recommended.append("summarize_table_column_values")
        truncated_items = [
            item for item in result.get("items") or []
            if isinstance(item, Mapping) and (item.get("has_more") or item.get("truncated"))
        ]
        if truncated_items:
            first_item = truncated_items[0]
            table_id = str(first_item.get("table_id") or "")
            offset = first_item.get("next_offset")
            if offset is None:
                offset = _as_int(first_item.get("offset"), 0) + _as_int(first_item.get("limit"), 20)
            detail_tool = {
                "name": "get_table_rows",
                "args": {
                    "table_id": table_id,
                    "offset": offset,
                    "limit": first_item.get("limit") or result.get("limit_per_request") or 20,
                },
            }
            recommended.insert(0, "get_table_rows")
    elif name in {"search_feishu_cache_rows", "batch_search_feishu_cache_rows"}:
        if name == "search_feishu_cache_rows" and result.get("rows"):
            recommended.append("get_feishu_cache_row")
    elif name == "list_project_memory_evidence":
        if result.get("cards"):
            recommended.append("get_project_memory_evidence")
    elif name in {"list_component_identity_cards", "search_component_identity_cards", "batch_get_component_identity_cards"}:
        recommended.append("get_component_identity_card")
    elif name in {"summarize_llm_topology_netlist", "query_llm_topology_netlist", "batch_query_llm_topology_netlist"}:
        recommended.append("summarize_topology_review_tasks")
        recommended.append("get_llm_topology_edge")
        recommended.append("get_llm_topology_node")
    elif name in {"summarize_topology_review_tasks", "batch_expand_topology_review_tasks"}:
        recommended.append("get_topology_review_task")
    elif name == "get_topology_review_task":
        recommended.append("get_llm_topology_edge")
        recommended.append("get_llm_topology_node")
    elif name in {"summarize_chip_topology", "query_chip_topology", "batch_query_chip_topology"}:
        recommended.append("get_chip_topology_edge")
    elif name in {"search_documents", "batch_search_documents"}:
        recommended.append("get_document_excerpt")
    elif name == "list_document_search_sources":
        recommended.append("search_documents")
    elif name in {"search_datasheet_chunks", "batch_search_datasheet_chunks"}:
        recommended.append("get_datasheet_chunk")
    elif name == "list_datasheet_documents":
        recommended.append("search_datasheet_chunks")
    elif name in {"search_datasheets", "match_component_datasheets", "batch_match_component_datasheets"}:
        recommended.append("get_datasheet_chunk")
        recommended.append("get_datasheet_excerpt")
    elif name == "get_datasheet_chunk":
        completeness = "complete" if not result.get("truncated") else "truncated"
    elif name == "list_agent_ref_sources":
        recommended.append("search_agent_ref_pdfs")
    elif name == "search_agent_ref_pdfs":
        if result.get("matches"):
            recommended.append("get_agent_ref_pdf_excerpt")
    elif name == "get_compare_section_rows":
        section_id = str(result.get("section_id") or "")
        if result.get("next_offset") is not None:
            detail_tool = {
                "name": "get_compare_section_rows",
                "args": {"section_id": section_id, "offset": result.get("next_offset"), "limit": result.get("limit")},
            }
            recommended.append("get_compare_section_rows")
    elif name in {"query_compare_diff", "batch_query_compare_diff", "batch_get_compare_rows"}:
        recommended.append("get_compare_row")
    elif name in {"compare_cadence_page_semantics", "batch_get_cadence_page_objects"}:
        recommended.append("get_cadence_page_object")
        recommended.append("get_cadence_page_raw_excerpt")

    if result.get("items_truncated") or result.get("page_results_truncated") or result.get("values_truncated"):
        completeness = "truncated"
    if not recommended and completeness in {"truncated", "partial", "preview"}:
        recommended.append(name)

    scope_parts = []
    for key in ("table_id", "section_id", "column", "query", "focus_refdes", "edge_count", "page_start", "page_end", "total_rows", "unique_count"):
        if key in result and result.get(key) not in {None, ""}:
            scope_parts.append(f"{key}={result.get(key)}")
    scope = "；".join(scope_parts) or str(result.get("summary") or name)[:500]
    return ToolResultContract(
        completeness=completeness,
        recommended_next_tools=_dedupe(recommended),
        detail_tool=detail_tool,
        aggregation_tool=aggregation_tool,
        scope_summary=scope,
    )
