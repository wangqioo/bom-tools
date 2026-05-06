# -*- coding: utf-8 -*-
"""Mock-only Aster review summary for PSTX reports.

This module intentionally does not call real Aster intranet endpoints. It keeps
the product/API shape ready while the integration is still experimental.
"""

from __future__ import annotations

from typing import Dict, List


def _as_int(value, default: int = 0) -> int:
    try:
        return int(value or default)
    except (TypeError, ValueError):
        return default


def _metric_value(report: dict, label: str, default=0):
    for metric in report.get('metrics', []) or []:
        if metric.get('label') == label:
            return metric.get('value', default)
    return default


def _section_map(report: dict) -> Dict[str, dict]:
    return {section.get('id', ''): section for section in report.get('sections', []) or []}


def _table_count(section: dict, table_id: str) -> int:
    for table in section.get('tables', []) or []:
        if table.get('id') == table_id:
            return _as_int(table.get('count', 0))
    return 0


def _priority(title: str, body: str, target: str, severity: str) -> dict:
    return {
        'title': title,
        'body': body,
        'target': target,
        'severity': severity,
    }


def _checklist_item(item: str, status: str, evidence: str, target: str, severity: str) -> dict:
    return {
        'item': item,
        'status': status,
        'evidence': evidence,
        'target': target,
        'severity': severity,
    }


def build_aster_mock_summary(report: dict, bundle: dict) -> dict:
    """Build a deterministic local mock response for the future Aster summary API."""
    sections = _section_map(report)
    drc_section = sections.get('drc', {})
    resistor_section = sections.get('resistor', {})
    derating_section = sections.get('derating', {})
    csa_section = sections.get('csa', {})

    drc_total = _as_int(_metric_value(report, 'DRC 总数', 0))
    derating_fail = _as_int(_metric_value(report, '降额不合格', 0))
    resistor_candidates = _as_int(_metric_value(report, '电阻候选', 0))
    resistor_unknown = _as_int(_metric_value(report, '电阻无法判断', 0))
    csa_candidates = _as_int(_metric_value(report, '规范候选', 0))

    priorities: List[dict] = []
    if drc_total:
        priorities.append(_priority(
            '先看设计检查',
            f'当前有 {drc_total} 条 DRC/属性/命名类结果，建议优先确认缺料号、缺 VALUE、BOM_OPTION 和单端网络。',
            'drc',
            'high',
        ))
    if derating_fail:
        priorities.append(_priority(
            '复核电容降额不合格',
            f'当前有 {derating_fail} 条电容降额不合格，需要确认工作电压来源是否可靠。',
            'derating',
            'high',
        ))
    if resistor_candidates or resistor_unknown:
        priorities.append(_priority(
            '复核电阻偏置/串阻候选',
            f'当前有 {resistor_candidates} 条电阻候选、{resistor_unknown} 条无法判断，适合结合芯片 pin 状态逐项确认。',
            'resistor',
            'medium',
        ))
    if csa_candidates:
        priorities.append(_priority(
            '复核 CSA 几何规范候选',
            f'当前有 {csa_candidates} 条 CSA 几何候选，重点查看 DOT 四向十字交叉和画圈对象原始行号。',
            'csa',
            'medium',
        ))
    if not priorities:
        priorities.append(_priority(
            '本地 mock 未发现高优先级项',
            '当前聚合指标较平稳，可按 BOM、网络、设计检查、规范检查的顺序做抽样复核。',
            'summary',
            'low',
        ))

    section_focus = [
        {
            'section': '设计检查',
            'target': 'drc',
            'rows': drc_section.get('total_rows', 0),
            'reason': f"缺料号 {_table_count(drc_section, 'missing_hq_code')}，缺 VALUE {_table_count(drc_section, 'missing_value')}，BOM_OPTION {_table_count(drc_section, 'bom_option_components')}，打圈问题 {_table_count(drc_section, 'bom_option_circle_issues')}",
        },
        {
            'section': '电阻检查',
            'target': 'resistor',
            'rows': resistor_section.get('total_rows', 0),
            'reason': f"串阻分压 {_table_count(resistor_section, 'divider_risks')}，重复上下拉 {_table_count(resistor_section, 'dup_pullups') + _table_count(resistor_section, 'dup_pulldowns')}，OD/OC {_table_count(resistor_section, 'od_missing')}",
        },
        {
            'section': '电容降额',
            'target': 'derating',
            'rows': derating_section.get('total_rows', 0),
            'reason': f"不合格 {derating_fail}，阈值 {report.get('ratio_limit', '')}%",
        },
        {
            'section': '规范检查',
            'target': 'csa',
            'rows': csa_section.get('total_rows', 0),
            'reason': f"DOT 四向十字 {_table_count(csa_section, 'csa_dot_cross_rows')}，画圈对象 {_table_count(csa_section, 'csa_circle_rows')}",
        },
    ]
    section_focus.sort(key=lambda item: _as_int(item.get('rows', 0)), reverse=True)

    review_checklist = [
        _checklist_item(
            'BOM 与装配状态',
            'needs_review' if (_table_count(drc_section, 'bom_option_components') or _table_count(drc_section, 'bom_option_circle_issues')) else 'covered_no_findings',
            f"BOM_OPTION {_table_count(drc_section, 'bom_option_components')}，打圈问题 {_table_count(drc_section, 'bom_option_circle_issues')}，DEPOP 总数 {_metric_value(report, 'DEPOP 总数', 0)}，include_depop={report.get('include_depop')}",
            'bom',
            'medium',
        ),
        _checklist_item(
            '网络分类与页码映射',
            'needs_review' if _table_count(sections.get('network', {}), 'page_mapping_rows') else 'covered_with_findings',
            f"网络总数 {_metric_value(report, '网络总数', 0)}，单节点 {_table_count(sections.get('network', {}), 'single_node_rows')}，页码映射 {_table_count(sections.get('network', {}), 'page_mapping_rows')}",
            'network',
            'medium',
        ),
        _checklist_item(
            '属性与命名 DRC',
            'needs_review' if drc_total else 'covered_no_findings',
            f"缺料号 {_table_count(drc_section, 'missing_hq_code')}，缺 VALUE {_table_count(drc_section, 'missing_value')}，未命名网络 {_table_count(drc_section, 'unnamed_nets')}",
            'drc',
            'high' if drc_total else 'low',
        ),
        _checklist_item(
            '芯片 Pin 与电阻网络',
            'needs_review' if resistor_candidates or resistor_unknown else 'covered_no_findings',
            f"芯片 Pin {_table_count(resistor_section, 'chip_pin_rows')}，串阻分压 {_table_count(resistor_section, 'divider_risks')}，OD/OC {_table_count(resistor_section, 'od_missing')}",
            'resistor',
            'medium',
        ),
        _checklist_item(
            '电容降额',
            'needs_review' if derating_fail else 'covered_no_findings',
            f"降额不合格 {derating_fail}，总行数 {derating_section.get('total_rows', 0)}，阈值 {report.get('ratio_limit', '')}%",
            'derating',
            'high' if derating_fail else 'low',
        ),
        _checklist_item(
            'CSA 几何规范',
            'needs_review' if csa_candidates else 'covered_no_findings',
            f"DOT 四向十字 {_table_count(csa_section, 'csa_dot_cross_rows')}，画圈对象 {_table_count(csa_section, 'csa_circle_rows')}",
            'csa',
            'medium',
        ),
    ]

    manual_review = [
        {
            'topic': '电平/电压推断',
            'reason': '网络名 token 只能作为候选，不能替代真实电源拓扑、上拉/下拉和器件特性确认。',
            'target': 'derating',
        },
        {
            'topic': 'OD/OC 与上下拉',
            'reason': '缺少芯片手册或明确 OD/OC 属性时，AI 只能提示人工确认，不应下确定缺陷结论。',
            'target': 'resistor',
        },
        {
            'topic': 'CSA 几何候选',
            'reason': '几何对象不等价于网络短接，需结合原理图页面和设计规范人工复核。',
            'target': 'csa',
        },
    ]

    top = priorities[0]
    summary = (
        f"Mock Aster 建议：优先处理“{top['title']}”。"
        f"该结果基于本地报告聚合指标生成，没有访问真实 Aster，也没有上传项目文件。"
    )

    return {
        'ok': True,
        'mode': 'mock',
        'provider': 'local-aster-mock',
        'project_name': report.get('project_name') or bundle.get('project_name') or '未命名项目',
        'summary': summary,
        'priorities': priorities[:5],
        'section_focus': section_focus,
        'review_checklist': review_checklist,
        'manual_review': manual_review,
        'safeguards': [
            '当前为 mock-only 实验链路，不访问真实 Aster 内网地址。',
            '前端不接触 appSecret、apiKey、accessToken。',
            '真实 Aster 接入前需要增加显式开关、脱敏预览和服务端凭据配置。',
        ],
    }
