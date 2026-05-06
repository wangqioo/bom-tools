# -*- coding: utf-8 -*-
"""Report table and section-card view helpers."""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from pstx_core.page_resolution import USER_VISIBLE_REAL_PAGE_LABEL
from pstx_rules.result_meta import count_result_kinds as _count_result_kinds

METRIC_TARGETS = {
    '贴装种类': 'bom',
    '贴装总数': 'bom',
    'DEPOP 总数': 'bom',
    '总BOM 总数': 'bom',
    'BOM圈问题': 'drc',
    '子模块数': 'module',
    '网络总数': 'network',
    'DRC 总数': 'drc',
    '降额不合格': 'derating',
    '电阻候选': 'resistor',
    '电阻无法判断': 'resistor',
    '规范候选': 'csa',
}


SECTION_LAYOUT = [
    {
        'id': 'bom',
        'title': 'BOM 视图',
        'lead': '展示贴装、去装配、变体配置与 BOM_OPTION 覆盖明细，支持快速复核物料范围。',
        'tables': [
            ('贴装 BOM', 'bom_normal_merged'),
            ('DEPOP BOM', 'bom_depop_merged'),
            ('BOM_OPTION 元件', 'bom_option_components'),
            ('BOM_OPTION 打圈覆盖明细', 'bom_option_circle_coverage'),
        ],
    },
    {
        'id': 'network',
        'title': '网络分析',
        'lead': '按网络视角汇总候选电源、接地、差分对、单节点网络概览与页码分布。',
        'tables': [
            ('候选电源网络', 'power_net_rows'),
            ('候选 GND 网络', 'gnd_net_rows'),
            ('候选差分对', 'diff_pair_rows'),
            ('单节点网络概览', 'single_node_rows'),
            ('页码元件分布', 'page_rows'),
            ('主模块页/页码映射检查', 'page_mapping_rows'),
        ],
    },
    {
        'id': 'module',
        'title': '模块视角',
        'lead': '基于 module_order(.dat) 拆出主模块和子模块实例，辅助按模块范围复核元器件与父级 Symbol 定位。',
        'tables': [
            ('模块范围汇总', 'module_scope_rows'),
            ('模块元件索引', 'module_component_rows'),
        ],
    },
    {
        'id': 'drc',
        'title': '设计检查',
        'lead': '集中展示需要复核的属性缺失、命名异常和 BOM_OPTION 异常项；物料清单类明细放在 BOM 视图。',
        'tables': [
            ('缺少料号', 'missing_hq_code'),
            ('缺少 VALUE', 'missing_value'),
            ('缺少封装', 'missing_package'),
            ('TBD 待确认属性', 'tbd_attrs'),
            ('单端候选网络', 'single_pin_nets'),
            ('未命名网络', 'unnamed_nets'),
            ('BOM_OPTION 候选拼写', 'bom_option_typos'),
            ('BOM_OPTION 打圈覆盖问题', 'bom_option_circle_issues'),
        ],
    },
    {
        'id': 'csa',
        'title': '规范检查',
        'lead': '扫描 sch_1/page*.csa/csv 几何对象与连接语义，复核带 DOT 四向十字交叉等规范候选。',
        'tables': [
            ('CSA 页级汇总', 'csa_summary_rows'),
            ('Cadence 连接语义页摘要', 'cadence_connectivity_rows'),
            ('CSA DOT四向十字交叉', 'csa_dot_cross_rows'),
        ],
    },
    {
        'id': 'resistor',
        'title': '电阻检查',
        'lead': '面向偏置、串阻以及芯片引脚关联的规则检查结果。',
        'tables': [
            ('串阻分压候选风险', 'divider_risks'),
            ('重复上拉候选', 'dup_pullups'),
            ('重复下拉候选', 'dup_pulldowns'),
            ('芯片 Pin 电阻状态', 'chip_pin_rows'),
        ],
    },
    {
        'id': 'derating',
        'title': '电容降额',
        'lead': '汇总工作电压推断、降额比值、原因代码与无法判断项。',
        'tables': [
            ('电容降额结果', 'derating'),
        ],
    },
]

TABLE_DISPLAY_LEVEL_LABELS = {
    'focus': '重点',
    'review': '常规复核',
    'info': '信息概览',
    'debug': 'Debug / 证据明细',
}

TABLE_DISPLAY_LEVELS = {
    # Project facts and navigational summaries.
    'bom_total_merged': 'info',
    'bom_normal_merged': 'info',
    'bom_depop_merged': 'info',
    'power_net_rows': 'info',
    'gnd_net_rows': 'info',
    'diff_pair_rows': 'info',
    'single_node_rows': 'info',
    'page_rows': 'info',
    'module_scope_rows': 'info',
    'csa_summary_rows': 'info',
    'cadence_connectivity_rows': 'info',
    # Focus findings shown first because they usually require manual confirmation.
    'missing_hq_code': 'focus',
    'missing_value': 'focus',
    'missing_package': 'focus',
    'single_pin_nets': 'focus',
    'unnamed_nets': 'focus',
    'bom_option_circle_issues': 'focus',
    'csa_dot_cross_rows': 'focus',
    'divider_risks': 'focus',
    'derating': 'focus',
    # Regular review findings.
    'bom_option_components': 'review',
    'tbd_attrs': 'review',
    'bom_option_typos': 'review',
    'dup_pullups': 'review',
    'dup_pulldowns': 'review',
    # Detail evidence that is useful during debugging, but should not dominate scanning.
    'bom_option_circle_coverage': 'debug',
    'page_mapping_rows': 'debug',
    'module_component_rows': 'debug',
    'chip_pin_rows': 'debug',
}

TABLE_TRUST_DEFAULTS_BY_LEVEL = {
    'focus': {
        'label': '规则候选',
        'tone': 'candidate',
        'note': '由规则筛出的优先复核项，需要结合原始表格和设计意图人工确认。',
    },
    'review': {
        'label': '规则候选',
        'tone': 'candidate',
        'note': '由规则筛出的常规复核项，不能直接等同于设计错误。',
    },
    'info': {
        'label': '信息统计',
        'tone': 'info',
        'note': '用于理解项目范围、数量和导航，不代表问题结论。',
    },
    'debug': {
        'label': '证据明细',
        'tone': 'evidence',
        'note': '用于追溯解析、页码映射或原始证据链，默认按需展开。',
    },
}

TABLE_TRUST_OVERRIDES = {
    'missing_hq_code': {
        'label': '明确异常',
        'tone': 'issue',
        'note': '源属性缺失可直接定位，通常需要补料号或确认 DEPOP 口径。',
    },
    'missing_value': {
        'label': '明确异常',
        'tone': 'issue',
        'note': 'VALUE 字段缺失可直接定位，通常需要补属性或确认器件定义。',
    },
    'missing_package': {
        'label': '明确异常',
        'tone': 'issue',
        'note': '封装字段缺失可直接定位，通常需要补属性或确认器件定义。',
    },
    'tbd_attrs': {
        'label': '待确认项',
        'tone': 'candidate',
        'note': '属性中含 TBD/待确认语义，需要人工确认是否允许保留。',
    },
    'bom_option_components': {
        'label': '配置清单',
        'tone': 'info',
        'note': '列出 BOM_OPTION 相关元件，主要用于变体范围复核。',
    },
    'bom_option_circle_coverage': {
        'label': '证据明细',
        'tone': 'evidence',
        'note': '记录画圈覆盖证据和候选坐标，用于追溯 BOM_OPTION 覆盖判断。',
    },
    'cadence_connectivity_rows': {
        'label': '语义摘要',
        'tone': 'evidence',
        'note': 'Cadence 页级连接语义摘要，只作为取证目录，不替代完整电气拓扑。',
    },
    'csa_summary_rows': {
        'label': '信息统计',
        'tone': 'info',
        'note': 'CSA 页级数量统计，用于定位规范候选所在页面。',
    },
    'csa_dot_cross_rows': {
        'label': '规则候选',
        'tone': 'candidate',
        'note': 'DOT 四向十字交叉是规范候选，需结合页面语义和原始线段复核。',
    },
    'derating': {
        'label': '规则结果',
        'tone': 'candidate',
        'note': '降额规则结果需结合工作电压来源、器件规格和设计上下文确认。',
    },
    'chip_pin_rows': {
        'label': '证据明细',
        'tone': 'evidence',
        'note': '芯片 Pin 与电阻状态索引，用于追溯连接和规则候选。',
    },
}


def trust_profile_for_table(table_id: str, level: str) -> dict:
    profile = dict(TABLE_TRUST_DEFAULTS_BY_LEVEL.get(level, TABLE_TRUST_DEFAULTS_BY_LEVEL['info']))
    profile.update(TABLE_TRUST_OVERRIDES.get(table_id, {}))
    return profile


def build_report_table(table_id: str,
                       title: str,
                       rows: List[dict],
                       *,
                       default_hidden_columns: Optional[List[str]] = None,
                       sort_profiles: Optional[List[dict]] = None,
                       default_sort_mode: str = 'column',
                       display_level: Optional[str] = None) -> dict:
    columns = list(rows[0].keys()) if rows else []
    if USER_VISIBLE_REAL_PAGE_LABEL in columns and '页面' in columns:
        columns = [column for column in columns if column != '页面']
    level = display_level or TABLE_DISPLAY_LEVELS.get(table_id, 'info')
    trust = trust_profile_for_table(table_id, level)
    return {
        'id': table_id,
        'title': title,
        'count': len(rows),
        'columns': columns,
        'rows': rows,
        'kind_counts': dict(_count_result_kinds(rows)),
        'default_hidden_columns': default_hidden_columns or [],
        'sort_profiles': sort_profiles or [{'id': 'column', 'label': '字段排序'}],
        'default_sort_mode': default_sort_mode,
        'display_level': level,
        'display_level_label': TABLE_DISPLAY_LEVEL_LABELS.get(level, TABLE_DISPLAY_LEVEL_LABELS['info']),
        'trust_label': trust['label'],
        'trust_tone': trust['tone'],
        'trust_note': trust['note'],
    }


def build_top_insights(
    *,
    drc_total: int,
    derating_fail: int,
    resistor_kind_counts,
    csa_candidate_total: int,
    warnings: List[str],
    section_cards: List[dict],
) -> List[dict]:
    insights = []
    if drc_total:
        insights.append({
            'title': '优先处理 DRC 检查项',
            'body': f'当前共有 {drc_total} 项 DRC 结果，建议先进入设计检查分区复核。',
            'tone': 'warning',
            'target': 'drc',
        })
    if derating_fail:
        insights.append({
            'title': '存在降额不合格项',
            'body': f'当前识别到 {derating_fail} 项降额不满足阈值，建议优先确认电容工作电压依据。',
            'tone': 'danger',
            'target': 'derating',
        })
    resistor_candidates = resistor_kind_counts.get('候选判断', 0)
    if resistor_candidates:
        insights.append({
            'title': '电阻相关候选项较多',
            'body': f'当前有 {resistor_candidates} 项电阻候选结果，建议结合偏置和串阻关系优先筛查。',
            'tone': 'neutral',
            'target': 'resistor',
        })
    if csa_candidate_total:
        insights.append({
            'title': '发现 CSA 几何规范候选项',
            'body': f'当前有 {csa_candidate_total} 项 CSA 几何候选结果，建议进入规范检查分区核对页面坐标和原始行号。',
            'tone': 'warning',
            'target': 'csa',
        })
    if warnings:
        insights.append({
            'title': '存在补充说明',
            'body': warnings[0],
            'tone': 'neutral',
            'target': 'summary',
        })
    if not insights:
        top_section = max(section_cards, key=lambda item: item.get('rows', 0), default=None)
        section_name = top_section['title'] if top_section else '本次报告'
        insights.append({
            'title': '结果已生成，可按分区复核',
            'body': f'当前未发现需要优先弹出的高风险项，建议从 {section_name} 开始浏览。',
            'tone': 'ok',
            'target': top_section['id'] if top_section else 'summary',
        })
    return insights[:4]


def build_section_cards(sections: List[dict]) -> List[dict]:
    cards = []
    for section in sections:
        non_empty_tables = [table for table in section['tables'] if table['count'] > 0]
        level_counts = {}
        for table in non_empty_tables:
            level = table.get('display_level') or 'info'
            level_counts[level] = level_counts.get(level, 0) + 1
        top_table = max(section['tables'], key=lambda table: table['count'], default=None)
        top_label = top_table['title'] if top_table and top_table['count'] else '暂无重点表'
        top_value = top_table['count'] if top_table else 0
        tone = 'ok' if section['total_rows'] == 0 else ('warning' if section['id'] in {'drc', 'derating', 'resistor', 'csa'} else 'neutral')
        cards.append({
            'id': section['id'],
            'title': section['title'],
            'rows': section['total_rows'],
            'active_tables': len(non_empty_tables),
            'top_label': top_label,
            'top_value': top_value,
            'tone': tone,
            'lead': section['lead'],
            'level_counts': level_counts,
        })
    return cards


_RELATED_REFDES_KEYS = (
    '位号',
    '元件',
    '器件',
    'LOCATION',
    'refdes',
    'RefDes',
    '芯片位号',
)
_RELATED_NET_KEYS = (
    '网络',
    '网络名',
    '网名',
    '连接网络',
    '节点网络',
    'net',
    'net_name',
    'Net',
)
_RELATED_PAGE_KEYS = (
    '页码',
    '用户看到的真实页',
    '主模块页',
    '页面',
    '真实页',
    '父级Symbol页码',
)

_RECOMMENDED_ACTIONS = {
    'focus': '优先展开原表，按位号、网络或页码复核证据，并记录是否需要人工处理。',
    'review': '按模块、页码和位号筛选后复核，确认是否为规则候选或真实问题。',
    'info': '作为项目范围、统计和导航信息使用；需要定位时再展开明细。',
    'debug': '仅在追溯解析、页码映射或证据链时展开，避免干扰常规审查。',
}


def _clean_plan_value(value: Any) -> str:
    if value is None:
        return ''
    if isinstance(value, (list, tuple, set)):
        parts = [_clean_plan_value(item) for item in value]
        return ', '.join(part for part in parts if part)
    text = str(value).strip()
    if not text or text.lower() in {'none', 'nan', 'null'} or text in {'-', '--', '无'}:
        return ''
    return text


def _sample_related_values(rows: List[dict], keys, *, limit: int = 8) -> List[str]:
    values: List[str] = []
    seen = set()
    for row in rows:
        if not isinstance(row, dict):
            continue
        for key in keys:
            if key not in row:
                continue
            text = _clean_plan_value(row.get(key))
            if not text:
                continue
            # A few tables merge refdes/pages into one cell; split common separators
            # so cards remain scannable without losing the original table evidence.
            parts = [part.strip() for part in text.replace('，', ',').replace('、', ',').split(',')]
            for part in parts:
                if not part or part in seen:
                    continue
                seen.add(part)
                values.append(part)
                if len(values) >= limit:
                    return values
    return values


def _kind_count_summary(table: dict) -> str:
    kind_counts = table.get('kind_counts') or {}
    if not kind_counts:
        return ''
    return '，'.join(f'{label} {value}' for label, value in list(kind_counts.items())[:4])


def _plan_item_from_table(section: dict, table: dict) -> dict:
    level = table.get('display_level') or TABLE_DISPLAY_LEVELS.get(table.get('id'), 'info')
    count = int(table.get('count') or 0)
    rows = table.get('rows') or []
    kind_summary = _kind_count_summary(table)
    summary = f'{table.get("title", table.get("id"))} 当前共有 {count} 条记录'
    if kind_summary:
        summary += f'，类型分布：{kind_summary}'
    summary += '。'
    return {
        'id': f'{level}-{section.get("id", "section")}-{table.get("id", "table")}',
        'title': table.get('title') or table.get('id') or '未命名审查项',
        'summary': summary,
        'level': level,
        'level_label': TABLE_DISPLAY_LEVEL_LABELS.get(level, TABLE_DISPLAY_LEVEL_LABELS['info']),
        'category': section.get('title') or section.get('id') or '报告',
        'section_id': section.get('id'),
        'table_id': table.get('id'),
        'count': count,
        'related_refdes': _sample_related_values(rows, _RELATED_REFDES_KEYS),
        'related_nets': _sample_related_values(rows, _RELATED_NET_KEYS),
        'related_pages': _sample_related_values(rows, _RELATED_PAGE_KEYS),
        'evidence_sources': [{
            'type': 'table',
            'table_id': table.get('id'),
            'title': table.get('title'),
            'row_count': count,
            'section_id': section.get('id'),
        }],
        'recommended_action': _RECOMMENDED_ACTIONS.get(level, _RECOMMENDED_ACTIONS['info']),
        'trust_label': table.get('trust_label') or trust_profile_for_table(table.get('id'), level)['label'],
        'trust_tone': table.get('trust_tone') or trust_profile_for_table(table.get('id'), level)['tone'],
        'trust_note': table.get('trust_note') or trust_profile_for_table(table.get('id'), level)['note'],
        'target': section.get('id'),
        'target_table_id': table.get('id'),
        'kind_counts': table.get('kind_counts') or {},
    }


def build_table_display_policy(sections: List[dict]) -> List[dict]:
    """Return a stable policy describing how each raw table should be presented."""
    policies = []
    for section in sections:
        for table in section.get('tables', []):
            level = table.get('display_level') or TABLE_DISPLAY_LEVELS.get(table.get('id'), 'info')
            count = int(table.get('count') or 0)
            policies.append({
                'table_id': table.get('id'),
                'title': table.get('title'),
                'section_id': section.get('id'),
                'section_title': section.get('title'),
                'level': level,
                'level_label': TABLE_DISPLAY_LEVEL_LABELS.get(level, TABLE_DISPLAY_LEVEL_LABELS['info']),
                'trust_label': table.get('trust_label') or trust_profile_for_table(table.get('id'), level)['label'],
                'trust_tone': table.get('trust_tone') or trust_profile_for_table(table.get('id'), level)['tone'],
                'trust_note': table.get('trust_note') or trust_profile_for_table(table.get('id'), level)['note'],
                'count': count,
                'default_collapsed': level in {'info', 'debug'} or count <= 0,
                'shown_in_plan': count > 0,
            })
    return policies


def build_review_plan(sections: List[dict]) -> Dict[str, Any]:
    """Convert report tables into layered, human-oriented review tasks.

    The raw tables remain the source of truth. This planner only adds display
    intent, evidence pointers and a cleaner first-screen ordering.
    """
    focus_items: List[dict] = []
    review_items_by_section: Dict[str, dict] = {}
    info_items: List[dict] = []
    debug_items: List[dict] = []
    hidden_table_ids: List[str] = []
    trust_counts: Dict[str, int] = {}

    for section in sections:
        for table in section.get('tables', []):
            level = table.get('display_level') or TABLE_DISPLAY_LEVELS.get(table.get('id'), 'info')
            count = int(table.get('count') or 0)
            if level == 'debug' or count <= 0:
                hidden_table_ids.append(table.get('id'))
            if count <= 0:
                continue
            item = _plan_item_from_table(section, table)
            trust_label = item.get('trust_label') or table.get('trust_label') or ''
            if trust_label:
                trust_counts[trust_label] = trust_counts.get(trust_label, 0) + count
            if level == 'focus':
                focus_items.append(item)
            elif level == 'review':
                section_id = section.get('id') or 'review'
                group = review_items_by_section.setdefault(section_id, {
                    'id': f'review-{section_id}',
                    'title': section.get('title') or section_id,
                    'summary': section.get('lead') or '',
                    'level': 'review',
                    'category': section.get('title') or section_id,
                    'section_id': section_id,
                    'items': [],
                    'count': 0,
                    'target': section_id,
                })
                group['items'].append(item)
                group['count'] += count
            elif level == 'debug':
                debug_items.append(item)
            else:
                info_items.append(item)

    review_groups = list(review_items_by_section.values())
    for group in review_groups:
        group['item_count'] = len(group.get('items') or [])
        group['summary'] = group['summary'] or f'{group["title"]} 有 {group["count"]} 条常规复核记录。'

    return {
        'summary': {
            'focus_count': len(focus_items),
            'review_group_count': len(review_groups),
            'review_item_count': sum(group.get('item_count', 0) for group in review_groups),
            'info_count': len(info_items),
            'debug_count': len(debug_items),
            'hidden_table_count': len(hidden_table_ids),
            'trust_counts': trust_counts,
        },
        'focus_items': focus_items,
        'review_groups': review_groups,
        'info_items': info_items,
        'debug_items': debug_items,
        'hidden_table_ids': hidden_table_ids,
    }
