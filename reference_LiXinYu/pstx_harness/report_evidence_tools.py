# -*- coding: utf-8 -*-
"""Fixed evidence-pack tools for report review harnesses."""

from __future__ import annotations

from typing import List, Optional

from pstx_harness.tool_core import HarnessToolContext
from pstx_harness.report_tool_utils import _as_int, _compact_rows
from pstx_knowledge.feishu_cache import build_feishu_bom_status


def _metric_value(report: dict, label: str, default=0):
    for metric in report.get("metrics", []) or []:
        if metric.get("label") == label:
            return metric.get("value", default)
    return default


def _find_table(report: dict, table_id: str) -> dict:
    for section in report.get("sections", []) or []:
        for table in section.get("tables", []) or []:
            if table.get("id") == table_id:
                return table
    return {}


def _table_pack(report: dict, table_id: str, limit: int) -> dict:
    table = _find_table(report, table_id)
    return {
        "id": table_id,
        "title": table.get("title") or table_id,
        "count": _as_int(table.get("count", len(table.get("rows", []) or []))),
        "kind_counts": dict(table.get("kind_counts") or {}),
        "sample_rows": _compact_rows(table.get("rows", []) or [], limit),
    }


def _pack(pack_id: str,
          title: str,
          target: str,
          summary: str,
          tables: List[dict],
          *,
          metrics: List[dict] = None,
          notes: List[str] = None,
          severity: str = "medium") -> dict:
    issue_count = sum(_as_int(table.get("count", 0)) for table in tables)
    return {
        "id": pack_id,
        "title": title,
        "target": target,
        "summary": summary,
        "severity": severity if issue_count else "low",
        "issue_count": issue_count,
        "metrics": metrics or [],
        "tables": tables,
        "notes": notes or [],
        "readonly": True,
    }


def _max_rows(context: HarnessToolContext) -> int:
    return _as_int(getattr(context.request, "max_rows_per_table", 12), 12)


def _bom_depop_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    report = context.report
    limit = _max_rows(context)
    tables = [
        _table_pack(report, "bom_option_components", limit),
        _table_pack(report, "bom_option_circle_issues", limit),
        _table_pack(report, "bom_option_circle_coverage", limit),
    ]
    depop_total = _metric_value(report, "DEPOP 总数", 0)
    circle_issues = _metric_value(report, "BOM圈问题", 0)
    summary = f"DEPOP 总数 {depop_total}，BOM_OPTION 打圈问题 {circle_issues}，include_depop={report.get('include_depop')}。"
    return _pack(
        "bom_depop",
        "BOM/DEPOP 与 BOM_OPTION 打圈",
        "bom",
        summary,
        tables,
        metrics=[
            {"label": "DEPOP 总数", "value": depop_total},
            {"label": "BOM圈问题", "value": circle_issues},
        ],
        severity="high" if _as_int(circle_issues) else "medium",
    )


def _page_mapping_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    report = context.report
    limit = _max_rows(context)
    tables = [
        _table_pack(report, "page_mapping_rows", limit),
        _table_pack(report, "page_rows", limit),
    ]
    summary = "汇总主模块页/页码映射检查和各页码元件统计，用于确认报告定位页码是否可追溯。"
    return _pack("page_mapping", "主模块页/页码映射", "network", summary, tables, severity="medium")


def _drc_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    report = context.report
    limit = _max_rows(context)
    ids = [
        "missing_hq_code",
        "missing_value",
        "missing_package",
        "tbd_attrs",
        "single_pin_nets",
        "unnamed_nets",
        "bom_option_typos",
    ]
    tables = [_table_pack(report, table_id, limit) for table_id in ids]
    summary = f"设计检查聚合 {sum(_as_int(table.get('count')) for table in tables)} 条属性、网络和命名类结果。"
    return _pack("drc", "DRC 属性/网络命名", "drc", summary, tables, severity="high")


def _resistor_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    report = context.report
    limit = _max_rows(context)
    ids = ["divider_risks", "dup_pullups", "dup_pulldowns", "chip_pin_rows"]
    tables = [_table_pack(report, table_id, limit) for table_id in ids]
    summary = "汇总芯片 Pin 电阻状态、串阻分压和重复上下拉候选，所有候选仍需结合芯片规格人工复核。"
    return _pack("resistor", "芯片 Pin / 串阻 / 上下拉", "resistor", summary, tables, severity="medium")


def _derating_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    report = context.report
    limit = _max_rows(context)
    tables = [_table_pack(report, "derating", limit)]
    fail_count = _metric_value(report, "降额不合格", 0)
    summary = f"电容降额阈值 {report.get('ratio_limit', '')}%，不合格 {fail_count}；无法确认工作电压来源时保持人工复核。"
    return _pack(
        "derating",
        "电容降额与人工复核边界",
        "derating",
        summary,
        tables,
        metrics=[{"label": "降额不合格", "value": fail_count}],
        severity="high" if _as_int(fail_count) else "medium",
    )


def _csa_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    report = context.report
    limit = _max_rows(context)
    ids = ["csa_summary_rows", "csa_dot_cross_rows"]
    tables = [_table_pack(report, table_id, limit) for table_id in ids]
    summary = "汇总 CSA 页级统计和 DOT 四向十字交叉；几何候选需要按页面坐标核对。"
    return _pack("csa", "CSA 几何规范", "csa", summary, tables, severity="medium")


def _feishu_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    status = build_feishu_bom_status(load_runtime=True)
    issue_count = 0 if status.get("available") else 1
    return {
        "id": "feishu_bom",
        "title": "Feishu BOM 缓存状态",
        "target": "bom",
        "summary": status.get("error") or f"飞书 BOM 缓存 {status.get('cache_count', 0)} 条，库数量 {status.get('library_count', 0)}。",
        "severity": "low",
        "issue_count": issue_count,
        "metrics": [
            {"label": "configured", "value": bool(status.get("configured"))},
            {"label": "cache_count", "value": _as_int(status.get("cache_count", 0))},
            {"label": "library_count", "value": _as_int(status.get("library_count", 0))},
        ],
        "tables": [],
        "notes": status.get("safeguards", []),
        "status": status,
        "readonly": True,
    }
