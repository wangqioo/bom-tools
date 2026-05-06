# -*- coding: utf-8 -*-
"""LLM-readable datasheet review templates.

The templates are intentionally deterministic and compact. They do not parse a
PDF by themselves; they tell an agent what to look for, which evidence must be
re-read in detail, and how datasheet facts should map back to schematic review.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Dict, Iterable, List


SCHEMA_VERSION = "pstx-datasheet-review-template.v1"


def _field(field_id: str,
           label: str,
           *,
           required: bool = False,
           evidence: str = "datasheet_chunk",
           schematic_evidence: Iterable[str] = ()) -> dict:
    return {
        "field_id": field_id,
        "label": label,
        "required": bool(required),
        "expected_evidence": evidence,
        "expected_schematic_evidence": list(schematic_evidence),
    }


_COMMON_OUTPUT_CONTRACT = [
    "不要只基于 search snippet 下定量结论；电压、电流、温度、时序、极限参数必须读取 detail chunk 或参数卡详情。",
    "回答必须区分：datasheet 明确证据、原理图已验证证据、仍需人工补充的证据。",
    "每个结论引用 evidence id、文档页码、chunk/parameter locator；证据不足时输出 missing_context。",
    "不要把整本 PDF 或整页长文本塞入模型上下文；先用模板/参数卡规划，再按需读取原文。",
]


TEMPLATES: Dict[str, dict] = {
    "common_chip": {
        "template_id": "common_chip",
        "schema_version": SCHEMA_VERSION,
        "title": "通用芯片 Datasheet 审查模板",
        "category": "generic",
        "applies_to": ["IC", "MCU", "ASIC", "FPGA", "bridge", "sensor", "interface chip"],
        "llm_goal": "把规格书事实转换成原理图 review 检查点，优先确认身份、供电、接口、时钟复位和封装热设计。",
        "search_queries": [
            "ordering information part number package",
            "recommended operating conditions supply voltage",
            "absolute maximum ratings",
            "pin description function voltage domain",
            "application information layout recommendations",
        ],
        "extraction_sections": [
            {
                "section_id": "identity",
                "title": "器件身份与版本",
                "fields": [
                    _field("part_number", "型号/Part Number", required=True, schematic_evidence=["HQ料号", "规格型号", "VALUE"]),
                    _field("package", "封装/Package", schematic_evidence=["封装", "PCB footprint"]),
                    _field("ordering_code", "Ordering Code/订货编码", schematic_evidence=["HQ料号", "飞书规格"]),
                    _field("document_revision", "文档版本/Revision"),
                ],
                "review_questions": [
                    "BOM/HQ 料号、规格型号、封装是否与 datasheet 的 ordering/package 匹配？",
                    "如果型号存在多个温度等级、封装或速度等级，当前原理图对应的是哪一个？",
                ],
            },
            {
                "section_id": "power",
                "title": "供电与极限参数",
                "fields": [
                    _field("recommended_supply", "推荐工作电压", required=True, evidence="datasheet_parameter", schematic_evidence=["电源网络", "供电源头", "电源域"]),
                    _field("absolute_max_supply", "绝对最大额定电压", evidence="datasheet_parameter", schematic_evidence=["电源网络", "保护电路"]),
                    _field("power_consumption", "功耗/电流预算", schematic_evidence=["电源树", "电源芯片负载能力"]),
                    _field("decoupling_requirement", "去耦/旁路电容要求", schematic_evidence=["去耦电容", "页码", "封装附近布局"]),
                ],
                "review_questions": [
                    "每个电源 pin 的网络名、电压域和供电来源是否落在 recommended operating 范围内？",
                    "绝对最大值是否被用作设计目标？如果是，需要标为风险。",
                    "去耦电容数量、容值、耐压和靠近芯片的布局要求是否有原理图/PCB 证据？",
                ],
            },
            {
                "section_id": "interfaces",
                "title": "接口与电压域",
                "fields": [
                    _field("pin_function", "Pin Function/Pin Description", required=True, schematic_evidence=["pin/net", "位号", "页码"]),
                    _field("io_voltage_domain", "IO 电压域", schematic_evidence=["电平转换", "上拉电压", "接口网络"]),
                    _field("strap_pull", "Strap/Boot 上下拉要求", schematic_evidence=["上下拉电阻", "默认状态"]),
                    _field("reserved_nc", "Reserved/NC/DNU 要求", schematic_evidence=["未连接网络", "NC pin"]),
                ],
                "review_questions": [
                    "跨芯片接口两侧电压域是否一致？不一致时是否有电平转换或容忍说明？",
                    "I2C/SPI/UART/PCIe/USB/MIPI 等接口是否满足上拉、端接、AC 耦合、差分极性和时钟要求？",
                    "strap/reserved/NC pin 是否按 datasheet 推荐处理？",
                ],
            },
        ],
        "review_playbook": [
            "先确认器件身份，再查供电推荐范围和绝对最大值。",
            "对每个电源域读取 detail chunk/参数卡，映射到原理图电源网络和拓扑节点。",
            "对跨芯片接口读取 pin function、电压域和应用章节，结合芯片级拓扑边复核。",
            "对缺失原理图证据的字段输出待人工补充，不要补猜。",
        ],
        "required_evidence": ["component_identity", "datasheet_chunk", "datasheet_parameter", "llm_topology_node", "llm_topology_edge"],
        "red_flags": [
            "只看到 absolute maximum，没看到 recommended operating，却直接判定可用。",
            "型号/封装/温度等级不一致。",
            "接口两侧电压域不明或跨域没有 level shifter/容忍证据。",
            "reserved/strap pin 没有明确处理。",
        ],
        "output_contract": _COMMON_OUTPUT_CONTRACT,
    },
    "power_regulator": {
        "template_id": "power_regulator",
        "schema_version": SCHEMA_VERSION,
        "title": "电源芯片/LDO/BUCK Datasheet 审查模板",
        "category": "power",
        "applies_to": ["LDO", "BUCK", "BOOST", "PMIC", "load switch", "power regulator"],
        "llm_goal": "把电源芯片规格书参数转成原理图电源完整性、稳定性、保护和热风险检查点。",
        "search_queries": [
            "recommended operating input voltage output current dropout",
            "output capacitor ESR stability",
            "enable threshold power good soft start",
            "thermal resistance junction temperature",
            "application circuit layout",
        ],
        "extraction_sections": [
            {
                "section_id": "regulator_limits",
                "title": "输入/输出/负载能力",
                "fields": [
                    _field("input_voltage", "输入电压范围", required=True, evidence="datasheet_parameter", schematic_evidence=["输入网络", "上游电源"]),
                    _field("output_voltage", "输出电压/反馈设置", required=True, schematic_evidence=["输出网络", "反馈电阻"]),
                    _field("output_current", "输出电流能力", evidence="datasheet_parameter", schematic_evidence=["负载芯片", "电源树"]),
                    _field("dropout_or_duty", "压差/占空比限制", evidence="datasheet_parameter", schematic_evidence=["Vin/Vout 差值"]),
                ],
                "review_questions": [
                    "输入范围、输出设定、负载电流是否同时满足 worst-case？",
                    "反馈电阻、采样网络和输出命名是否能证明目标电压？",
                ],
            },
            {
                "section_id": "stability_protection",
                "title": "稳定性/保护/控制",
                "fields": [
                    _field("output_capacitor", "输出电容与 ESR 稳定性", schematic_evidence=["输出电容", "容值", "ESR/封装"]),
                    _field("enable_threshold", "EN/OE 阈值", schematic_evidence=["EN 网络", "上拉/下拉"]),
                    _field("pgood_reset", "PGOOD/Reset 时序", schematic_evidence=["PGOOD 网络", "后级 reset/enable"]),
                    _field("thermal_limit", "结温/热阻/封装散热", evidence="datasheet_parameter", schematic_evidence=["负载电流", "封装", "散热路径"]),
                ],
                "review_questions": [
                    "输出电容是否满足 datasheet 对容值/ESR/数量的稳定性要求？",
                    "EN/PGOOD 是否满足阈值、默认状态和上电时序？",
                    "功耗和热阻是否提示热风险，需要人工复核 PCB 散热？",
                ],
            },
        ],
        "review_playbook": [
            "读取输入/输出/电流/热参数卡。",
            "匹配原理图电源网络和下游关键负载。",
            "检查输出电容稳定性、EN/PGOOD 上下拉、反馈采样。",
            "对电流预算和热设计输出证据缺口。",
        ],
        "required_evidence": ["component_identity", "datasheet_parameter", "datasheet_chunk", "llm_topology_edge"],
        "red_flags": ["输出电容不满足稳定性范围", "EN 浮空或默认态不明", "负载电流接近上限", "反馈网络无法证明输出电压"],
        "output_contract": _COMMON_OUTPUT_CONTRACT,
    },
    "complex_chip": {
        "template_id": "complex_chip",
        "schema_version": SCHEMA_VERSION,
        "title": "复杂芯片/大芯片 Datasheet 审查模板",
        "category": "complex_chip",
        "applies_to": ["SOC", "FPGA", "GPU", "ASIC", "large IC", "processor", "bridge"],
        "llm_goal": "把复杂芯片规格书拆成供电域、时序、接口、热设计和模式配置，配合芯片级拓扑做 review。",
        "search_queries": [
            "power rail voltage tolerance power sequence",
            "recommended operating conditions absolute maximum ratings",
            "power consumption current AC noise ripple",
            "power up sequence reset timing",
            "power down sequence timing T1 T2 T3 T4",
            "pin description voltage domain",
            "clock requirements reset strap boot mode",
            "electrical characteristics IO input output threshold",
            "thermal characteristics junction temperature package",
        ],
        "extraction_sections": [
            {
                "section_id": "rails_sequence",
                "title": "多电源域与上电时序",
                "fields": [
                    _field("power_rail_voltage", "各电源 rail 电压/噪声", required=True, evidence="datasheet_parameter", schematic_evidence=["电源网络", "拓扑节点", "去耦"]),
                    _field("power_budget_current", "各 rail 电流/功耗预算", evidence="datasheet_parameter", schematic_evidence=["电源芯片负载能力", "电源树"]),
                    _field("power_sequence_timing", "上电/下电/Reset 时序", evidence="datasheet_parameter", schematic_evidence=["EN/PGOOD/RESET 网络", "电源芯片连接"]),
                    _field("rail_grouping", "电源域分组/同源要求", schematic_evidence=["电源树", "网络别名"]),
                ],
                "review_questions": [
                    "每个 rail 是否有明确电压、容差、噪声/纹波要求和对应原理图网络？",
                    "多个 rail 的上电顺序是否能从 EN/PGOOD/Reset 拓扑证明？",
                ],
            },
            {
                "section_id": "interfaces",
                "title": "高速/低速接口与电平域",
                "fields": [
                    _field("interface_voltage_domain", "接口电压域", schematic_evidence=["拓扑边", "电平转换", "上拉电源"]),
                    _field("io_threshold", "IO 输入/输出阈值与容忍", evidence="datasheet_parameter", schematic_evidence=["接口两端电压域", "上拉电源"]),
                    _field("clock_reset", "Clock/Reset 要求", schematic_evidence=["clock 网络", "串阻/端接", "reset 上拉"]),
                    _field("high_speed_requirements", "高速接口 AC 耦合/端接/极性", schematic_evidence=["差分对", "AC 耦合电容", "端接电阻"]),
                    _field("debug_boot_strap", "调试/启动配置 pin", schematic_evidence=["strap 电阻", "下载接口", "默认状态"]),
                ],
                "review_questions": [
                    "芯片级拓扑中的接口边是否能解释接口类型、电压域和 review 风险？",
                    "PCIE/PCE/P5E、USB、MIPI、LVDS、DDR 等别名是否被归一为同一业务接口？",
                    "Clock/Reset/Boot strap 是否有默认状态和时序证据？",
                ],
            },
            {
                "section_id": "thermal_package",
                "title": "封装、温度、散热",
                "fields": [
                    _field("package_options", "封装/尺寸/焊盘", schematic_evidence=["封装", "BOM"]),
                    _field("thermal_characteristic", "热阻/结温", evidence="datasheet_parameter", schematic_evidence=["功耗", "散热路径"]),
                    _field("operating_environment", "工作环境温度/湿度", evidence="datasheet_parameter"),
                ],
                "review_questions": [
                    "选型的封装和速度/温度等级是否与项目环境一致？",
                    "热参数是否提示需要 DFMEA 或 PCB 散热复核？",
                ],
            },
        ],
        "review_playbook": [
            "先按型号/封装确认身份，再检索 rail/sequence/interface/thermal 四类证据。",
            "复杂芯片优先建立 rail 清单：电压、容差/噪声、最大电流、上电组、关断组、Reset/PGOOD 依赖。",
            "用 topology-netlist 查询该芯片的 hub 边和接口分组。",
            "对 power sequence 读取完整时序表/图，不能只引用搜索 snippet；若出现 T1/T2/T3/T4 等相对时序，必须保存每个条件。",
            "对每条高风险接口读取 pin function/detail chunk，而不是只看拓扑摘要。",
            "输出 review checklist：电源域、时序、接口电平、高速约束、clock/reset、strap、热。",
        ],
        "required_evidence": ["component_identity", "datasheet_parameter", "datasheet_chunk", "llm_topology_node", "llm_topology_edge"],
        "red_flags": [
            "多电源域没有电压/时序证据。",
            "只抽到 rail 电压但没有电流/功耗或噪声/纹波限制。",
            "接口跨电压域但无电平转换/容忍证据。",
            "高速接口缺 AC 耦合、端接、差分极性或参考时钟要求。",
            "strap/debug pin 默认状态不明。",
        ],
        "output_contract": _COMMON_OUTPUT_CONTRACT,
    },
    "level_shifter": {
        "template_id": "level_shifter",
        "schema_version": SCHEMA_VERSION,
        "title": "电平转换/接口桥 Datasheet 审查模板",
        "category": "interface",
        "applies_to": ["level shifter", "translator", "buffer", "redriver", "bridge"],
        "llm_goal": "确认两侧电压域、方向/OE、上拉、接口速率和信号容忍条件。",
        "search_queries": [
            "VCCA VCCB voltage range",
            "OE enable input threshold",
            "direction control auto direction pull-up",
            "maximum data rate capacitance",
            "application schematic level translation",
        ],
        "extraction_sections": [
            {
                "section_id": "voltage_domains",
                "title": "双侧电压域",
                "fields": [
                    _field("side_a_voltage", "A 侧电源/IO 电压", required=True, schematic_evidence=["VCCA", "A侧网络", "上拉电源"]),
                    _field("side_b_voltage", "B 侧电源/IO 电压", required=True, schematic_evidence=["VCCB", "B侧网络", "上拉电源"]),
                    _field("io_tolerance", "IO 容忍/方向限制", schematic_evidence=["接口拓扑", "连接芯片"]),
                ],
                "review_questions": [
                    "A/B 两侧供电是否对应拓扑两侧芯片的 IO 电压域？",
                    "双向/单向器件是否被用在正确方向和速率？",
                ],
            },
            {
                "section_id": "control_and_pull",
                "title": "控制脚与上拉",
                "fields": [
                    _field("oe_enable", "OE/EN 控制默认态", schematic_evidence=["OE/EN 网络", "上下拉"]),
                    _field("pullup_requirement", "上拉要求/阻值范围", schematic_evidence=["上拉电阻", "上拉电源"]),
                    _field("speed_capacitance", "速率/负载电容限制", schematic_evidence=["接口类型", "连接器/线长"]),
                ],
                "review_questions": [
                    "OE 默认态是否安全？上电前后是否会误驱动？",
                    "I2C 等开漏接口上拉是否在正确电压域且阻值合理？",
                ],
            },
        ],
        "review_playbook": [
            "先确认 VCCA/VCCB 和 A/B pin 分组。",
            "查询 topology edge，确认两侧芯片和接口类型。",
            "读取 OE/方向/速率章节，再核对原理图上拉和控制网络。",
        ],
        "required_evidence": ["datasheet_chunk", "component_identity", "llm_topology_edge"],
        "red_flags": ["A/B 侧电压域反接", "OE 浮空", "上拉接错电源域", "接口速率超过器件规格"],
        "output_contract": _COMMON_OUTPUT_CONTRACT,
    },
}


def _compact_template(template: dict, *, include_questions: bool) -> dict:
    item = deepcopy(template)
    if not include_questions:
        for section in item.get("extraction_sections", []) or []:
            section.pop("review_questions", None)
        item.pop("review_playbook", None)
        item.pop("red_flags", None)
    return item


def list_datasheet_review_templates(category: str = "", *, include_questions: bool = True) -> dict:
    """Return LLM-readable datasheet review templates."""

    category = str(category or "").strip().lower()
    templates: List[dict] = []
    for template in TEMPLATES.values():
        if category and category not in {
            str(template.get("category", "")).lower(),
            str(template.get("template_id", "")).lower(),
        }:
            applies = {str(item).lower() for item in template.get("applies_to", [])}
            if category not in applies:
                continue
        templates.append(_compact_template(template, include_questions=include_questions))
    return {
        "ok": True,
        "schema_version": SCHEMA_VERSION,
        "category": category,
        "template_count": len(templates),
        "templates": templates,
        "recommended_default": "common_chip",
        "summary": f"返回 {len(templates)} 个 datasheet 审查模板。",
    }


def get_datasheet_review_template(template_id: str) -> dict:
    """Read one datasheet review template by id."""

    key = str(template_id or "").strip().lower()
    if not key:
        return {"ok": False, "error": "get_datasheet_review_template 需要 template_id。"}
    template = TEMPLATES.get(key)
    if not template:
        return {
            "ok": False,
            "error": f"未知 datasheet 审查模板：{template_id}。",
            "available_template_ids": sorted(TEMPLATES.keys()),
        }
    return {
        "ok": True,
        "schema_version": SCHEMA_VERSION,
        "template": _compact_template(template, include_questions=True),
        "summary": f"读取 datasheet 审查模板：{template['title']}。",
    }


def datasheet_review_template_catalog() -> dict:
    """Return compact catalog for CLI/schema discovery."""

    return {
        "schema_version": SCHEMA_VERSION,
        "templates": [
            {
                "template_id": template["template_id"],
                "title": template["title"],
                "category": template["category"],
                "applies_to": list(template.get("applies_to", []) or []),
            }
            for template in TEMPLATES.values()
        ],
    }
