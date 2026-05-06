# -*- coding: utf-8 -*-
"""Synthetic report fixture used for UI-only report debugging."""

from __future__ import annotations

from datetime import datetime
from typing import Any, Dict, List

from pstx_webapp.report_tables import build_report_table, build_section_cards, build_top_insights


def _rows(prefix: str, count: int, *, kind: str = "候选判断") -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for index in range(1, count + 1):
        rows.append({
            "位号": f"{prefix}{index}",
            "页码": f"PAGE{100 + index}",
            "网络": f"NET_{prefix}_{index:02d}",
            "状态": "⚠ 待复核" if index % 4 else "✅ 通过",
            "结论类型": kind if index % 4 else "确定结论",
            "说明": f"Debug fixture row {index}: 用于观察报告页表格、筛选、折叠和密度效果。",
        })
    return rows


def build_debug_report_payload() -> Dict[str, Any]:
    """Return a small but visually representative report payload."""

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dataset_map = {
        "bom_normal_merged": build_report_table("bom_normal_merged", "贴装 BOM", _rows("U", 18, kind="物料")),
        "bom_depop_merged": build_report_table("bom_depop_merged", "DEPOP BOM", _rows("R", 5, kind="DEPOP")),
        "bom_option_components": build_report_table("bom_option_components", "BOM_OPTION 元件", _rows("C", 7)),
        "bom_option_circle_coverage": build_report_table("bom_option_circle_coverage", "BOM_OPTION 打圈覆盖明细", _rows("D", 3)),
        "power_net_rows": build_report_table("power_net_rows", "候选电源网络", _rows("PWR", 6)),
        "gnd_net_rows": build_report_table("gnd_net_rows", "候选 GND 网络", _rows("GND", 3)),
        "diff_pair_rows": build_report_table("diff_pair_rows", "候选差分对", _rows("DP", 4)),
        "single_node_rows": build_report_table("single_node_rows", "单节点网络概览", _rows("SN", 2)),
        "page_rows": build_report_table("page_rows", "页码元件分布", [
            {"页码": f"PAGE{page}", "元件数量": 18 + page % 7, "主要类型": "U/R/C", "备注": "Debug 页码分布"}
            for page in range(101, 113)
        ]),
        "page_mapping_rows": build_report_table("page_mapping_rows", "主模块页/页码映射检查", _rows("PM", 0)),
        "module_scope_rows": build_report_table("module_scope_rows", "模块范围汇总", [
            {"模块": "main", "类型": "主模块", "起始页码": "PAGE1", "结束页码": "PAGE96", "元件数量": 3200},
            {"模块": "i2c_repeater", "类型": "子模块", "起始页码": "PAGE177", "结束页码": "PAGE210", "元件数量": 520},
        ]),
        "module_component_rows": build_report_table("module_component_rows", "模块元件索引", _rows("XA", 10)),
        "missing_hq_code": build_report_table("missing_hq_code", "缺少料号", _rows("U", 6)),
        "missing_value": build_report_table("missing_value", "缺少 VALUE", _rows("R", 0)),
        "missing_package": build_report_table("missing_package", "缺少封装", _rows("C", 2)),
        "tbd_attrs": build_report_table("tbd_attrs", "TBD 待确认属性", _rows("L", 4)),
        "single_pin_nets": build_report_table("single_pin_nets", "单端候选网络", _rows("N", 8)),
        "unnamed_nets": build_report_table("unnamed_nets", "未命名网络", _rows("UN", 5)),
        "bom_option_typos": build_report_table("bom_option_typos", "BOM_OPTION 候选拼写", _rows("B", 0)),
        "bom_option_circle_issues": build_report_table("bom_option_circle_issues", "BOM_OPTION 打圈覆盖问题", _rows("BC", 3)),
        "csa_summary_rows": build_report_table("csa_summary_rows", "CSA 页级汇总", _rows("CSA", 5)),
        "cadence_connectivity_rows": build_report_table("cadence_connectivity_rows", "Cadence 连接语义页摘要", [
            {"页码": "PAGE114", "PAGE_NUMBER": "PAGE144", "WIRE": 42, "DOT": 8, "连接组": 16, "网络标签": 21, "端口": 4, "跨页连接": 2, "Bus": 1, "No Connect": 0, "未绑定语义": 1, "未知行": 3, "解析状态": "ok"},
            {"页码": "PAGE177", "PAGE_NUMBER": "PAGE1", "WIRE": 38, "DOT": 6, "连接组": 14, "网络标签": 18, "端口": 6, "跨页连接": 3, "Bus": 2, "No Connect": 1, "未绑定语义": 0, "未知行": 2, "解析状态": "ok"},
        ]),
        "csa_dot_cross_rows": build_report_table("csa_dot_cross_rows", "CSA DOT四向十字交叉", _rows("DOT", 2)),
        "divider_risks": build_report_table("divider_risks", "串阻分压候选风险", _rows("RS", 4)),
        "dup_pullups": build_report_table("dup_pullups", "重复上拉候选", _rows("RU", 3)),
        "dup_pulldowns": build_report_table("dup_pulldowns", "重复下拉候选", _rows("RD", 0)),
        "chip_pin_rows": build_report_table("chip_pin_rows", "芯片 Pin 电阻状态", _rows("U", 22)),
        "derating": build_report_table("derating", "电容降额结果", _rows("C", 9)),
    }
    section_specs = [
        ("bom", "BOM 视图", "贴装、DEPOP、BOM_OPTION 与打圈覆盖。", ["bom_normal_merged", "bom_depop_merged", "bom_option_components", "bom_option_circle_coverage"]),
        ("network", "网络分析", "电源、地、差分、单节点与页码分布。", ["power_net_rows", "gnd_net_rows", "diff_pair_rows", "single_node_rows", "page_rows", "page_mapping_rows"]),
        ("module", "模块视角", "主模块与子模块拆分后的元件索引。", ["module_scope_rows", "module_component_rows"]),
        ("drc", "设计检查", "缺属性、命名异常、单端网络与打圈问题。", ["missing_hq_code", "missing_value", "missing_package", "tbd_attrs", "single_pin_nets", "unnamed_nets", "bom_option_typos", "bom_option_circle_issues"]),
        ("csa", "规范检查", "CSA 几何对象候选。", ["csa_summary_rows", "cadence_connectivity_rows", "csa_dot_cross_rows"]),
        ("resistor", "电阻检查", "重复上下拉、串阻与芯片 Pin。", ["divider_risks", "dup_pullups", "dup_pulldowns", "chip_pin_rows"]),
        ("derating", "电容降额", "降额比值、原因代码与无法判断项。", ["derating"]),
    ]
    sections = []
    for section_id, title, lead, table_ids in section_specs:
        tables = [dataset_map[table_id] for table_id in table_ids]
        sections.append({
            "id": section_id,
            "title": title,
            "lead": lead,
            "tables": tables,
            "total_rows": sum(table["count"] for table in tables),
        })
    section_cards = build_section_cards(sections)
    warnings = ["Debug fixture：该页面只用于 UI 检查，不代表真实项目审查结果。"]
    return {
        "run_id": "debug-report",
        "project_name": "Debug UI 假项目 / Report Fixture",
        "generated_at": generated_at,
        "ratio_limit": 70.0,
        "include_depop": True,
        "include_total_bom": False,
        "depop_count": 42,
        "excluded_depop_count": 0,
        "warnings": warnings,
        "input_files": [
            {"label": "pstxprt.dat", "filename": "debug/packaged/pstxprt.dat", "size": "128000"},
            {"label": "pstxnet.dat", "filename": "debug/packaged/pstxnet.dat", "size": "96000"},
            {"label": "pstxref.dat", "filename": "debug/packaged/pstxref.dat", "size": "24000"},
        ],
        "metrics": [
            {"label": "贴装种类", "value": 128, "tone": "neutral", "target": "bom", "caption": "BOM 视图"},
            {"label": "贴装总数", "value": 5320, "tone": "neutral", "target": "bom", "caption": "贴装器件总量"},
            {"label": "DEPOP 总数", "value": 42, "tone": "muted", "target": "bom", "caption": "去装配器件"},
            {"label": "BOM圈问题", "value": 3, "tone": "warning", "target": "drc", "caption": "打圈覆盖"},
            {"label": "子模块数", "value": 7, "tone": "neutral", "target": "module", "caption": "module_order"},
            {"label": "网络总数", "value": 1840, "tone": "neutral", "target": "network", "caption": "网络总览"},
            {"label": "DRC 总数", "value": 28, "tone": "warning", "target": "drc", "caption": "设计检查"},
            {"label": "降额不合格", "value": 9, "tone": "warning", "target": "derating", "caption": "电容复核"},
            {"label": "电阻候选", "value": 12, "tone": "neutral", "target": "resistor", "caption": "偏置/串阻"},
            {"label": "规范候选", "value": 2, "tone": "warning", "target": "csa", "caption": "CSA 几何"},
        ],
        "top_insights": build_top_insights(
            drc_total=28,
            derating_fail=9,
            resistor_kind_counts={"候选判断": 12},
            csa_candidate_total=2,
            warnings=warnings,
            section_cards=section_cards,
        ),
        "section_cards": section_cards,
        "summary_lines": [
            "DEPOP 排查：开启，Debug fixture 保留 DEPOP 项用于观察 UI。",
            "模块视角：识别到 7 个子模块实例，可按导航查看模块分区。",
            "明细表默认折叠，点击分区内表格按钮后按需渲染。",
        ],
        "sections": sections,
        "feishu_hq_links": [],
    }
