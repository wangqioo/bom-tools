# -*- coding: utf-8 -*-
"""Compare payload section/table view model builders."""

from __future__ import annotations

import re
from typing import List, Optional

from pstx_webapp.report_tables import build_report_table


def build_compare_section(section_id: str,
                          title: str,
                          lead: str,
                          diff: dict,
                          *,
                          priority: str = 'normal',
                          group: str = 'detail',
                          default_hidden_columns: Optional[List[str]] = None,
                          sort_profiles: Optional[List[dict]] = None) -> dict:
    rows = diff.get('rows', []) or []
    table = build_report_table(
        f'compare_{section_id}',
        title,
        rows,
        default_hidden_columns=default_hidden_columns or [],
        sort_profiles=sort_profiles or [{'id': 'column', 'label': '字段排序'}],
    )
    table['count'] = diff.get('total_rows', len(rows))
    table['displayed_count'] = len(rows)
    table['default_density'] = 'comfortable'
    return {
        'id': section_id,
        'title': title,
        'lead': lead,
        'priority': priority,
        'group': group,
        'diff': diff,
        'table': table,
    }


def safe_compare_section_id(value: str) -> str:
    text = re.sub(r'[^0-9A-Za-z_-]+', '_', str(value or 'section')).strip('_')
    return text[:80] or 'section'


def build_compare_sections(payload: dict) -> List[dict]:
    sections = [
        build_compare_section(
            'overview',
            '指标差异',
            '项目级元件数、网络数、DRC 数量和报告指标变化。',
            {
                'title': '指标差异',
                'key_label': '指标',
                'added_count': 0,
                'removed_count': 0,
                'changed_count': len(payload.get('overview', []) or []),
                'rows': payload.get('overview', []) or [],
                'total_rows': len(payload.get('overview', []) or []),
                'truncated': False,
            },
            priority='high',
            group='overview',
        ),
        build_compare_section(
            'net_view',
            'Net 视角变化',
            '按左侧网络到右侧网络聚合 Pin/Net 证据，优先看网络迁移、关键器件和 R/C/L 影响范围。',
            payload.get('net_view_diff', {}),
            priority='critical',
            group='net',
            default_hidden_columns=['变化说明'],
        ),
        build_compare_section(
            'key_components',
            '关键器件增删',
            '芯片、连接器和其他非 R/C/L 关键器件的新增和删除。',
            payload.get('key_component_diff', {}),
            priority='high',
            group='device',
        ),
        build_compare_section(
            'key_pin_nets',
            '关键器件 Pin/Net 连接差异',
            '芯片、PU、XU、连接器等关键器件逐 pin 对比网络、引脚名、页码和飞书信息。',
            payload.get('key_pin_net_diff', {}),
            priority='critical',
            group='net',
        ),
        build_compare_section(
            'passive_pin_nets',
            'R/C/L Pin/Net 连接差异',
            '电阻、电容、电感等无源件逐 pin 对比连接网络变化。',
            payload.get('passive_pin_net_diff', {}),
            priority='normal',
            group='net',
        ),
        build_compare_section(
            'components',
            '元件属性差异',
            '全量位号属性差异，包含料号、值、封装、BOM_OPTION、页码和飞书 PI/选型顺序。',
            payload.get('component_diff', {}),
            priority='normal',
            group='parts',
        ),
        build_compare_section(
            'nets',
            '网络节点明细',
            '按网络名展示新增、删除和节点连接列表变化，作为 Net 视角的原始节点清单。',
            payload.get('net_diff', {}),
            priority='normal',
            group='net',
        ),
    ]
    for diff in payload.get('table_diffs', []) or []:
        table_id = str(diff.get('id') or diff.get('title') or 'report_table')
        sections.append(build_compare_section(
            f"report_table_{safe_compare_section_id(table_id)}",
            f"检查结果表：{diff.get('title') or table_id}",
            '报告中已有检查表的逐行差异，用于追踪 DRC、电阻、降额、CSA、BOM_OPTION 和主模块页/页码映射等审查输出变化。',
            diff,
            priority='report',
            group='report',
        ))
    return sections
