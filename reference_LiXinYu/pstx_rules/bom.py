# -*- coding: utf-8 -*-
"""BOM detail and merged BOM builders."""

from typing import Dict, List, Tuple

from pstx_core.page_resolution import _component_page_fields
from pstx_rules.common import (
    COMP_TYPE_CN,
    _TYPE_ORDER,
    _is_depop_option,
    _natural_sort_key,
)

def build_bom(components: Dict):
    detail_normal, detail_depop = [], []
    for comp in components.values():
        ctype = comp.get('comp_type', '')
        row = {
            '位号':          comp['refdes'],
            '料号':          comp.get('hq_code', ''),
            '描述':          comp.get('part_name', ''),
            '值':            comp.get('value', ''),
            '封装':          comp.get('package', ''),
            '耐压/额定电压': comp.get('voltage', ''),
            '额定功率':      comp.get('power', ''),
            '精度':          comp.get('tolerance', ''),
            '材质':          comp.get('material', ''),
            '类型':          COMP_TYPE_CN.get(ctype, ctype),
            '_ctype':        ctype,
            'ROOM':          comp.get('room', ''),
        }
        row.update(_component_page_fields(comp))
        (detail_depop if _is_depop_option(comp.get('bom_option', '')) else detail_normal).append(row)

    def _merge(detail):
        if not detail:
            return []
        groups = {}
        for row in detail:
            key = (
                ('pn', row['料号']) if row['料号'] else
                ('desc', row['描述'], row['值'], row['封装'], row['耐压/额定电压'],
                 row['额定功率'], row['精度'], row['材质'], row['类型'])
            )
            if key not in groups:
                groups[key] = {
                    '料号': row['料号'], '位号列表': [], '数量': 0,
                    '描述': row['描述'], '值': row['值'], '封装': row['封装'],
                    '耐压': row['耐压/额定电压'], '额定功率': row['额定功率'],
                    '精度': row['精度'], '材质': row['材质'],
                    '类型': row['类型'], '_ctype': row['_ctype'],
                }
            groups[key]['位号列表'].append(row['位号'])
            groups[key]['数量'] += 1
        merged = list(groups.values())
        merged.sort(key=lambda r: (
            _TYPE_ORDER.index(r['_ctype']) if r['_ctype'] in _TYPE_ORDER else 99,
            r['料号'],
            r['描述'],
            r['值'],
            r['封装'],
        ))
        for i, r in enumerate(merged, 1):
            r['序号'] = i
            r['位号列表'] = ', '.join(sorted(r['位号列表'], key=_natural_sort_key))
            del r['_ctype']
        return merged

    def _clean(rows):
        return [{k: v for k, v in r.items() if k != '_ctype'} for r in rows]

    return _clean(detail_normal), _clean(detail_depop), _merge(detail_normal), _merge(detail_depop)


def _bom_type_order_from_label(type_label: str) -> int:
    for type_key, label in COMP_TYPE_CN.items():
        if label == type_label:
            return _TYPE_ORDER.index(type_key) if type_key in _TYPE_ORDER else 99
    return 99


def build_total_bom(detail_normal: List[dict], detail_depop: List[dict]) -> Tuple[List[dict], List[dict]]:
    """Build total BOM detail/merged rows from mounted and DEPOP detail rows."""
    total_detail = (
        [{**row, 'BOM状态': '贴装'} for row in (detail_normal or [])]
        + [{**row, 'BOM状态': 'DEPOP'} for row in (detail_depop or [])]
    )
    if not total_detail:
        return [], []

    groups: Dict[tuple, dict] = {}
    for row in total_detail:
        key = (
            ('pn', row.get('料号', '')) if row.get('料号', '') else
            (
                'desc',
                row.get('描述', ''),
                row.get('值', ''),
                row.get('封装', ''),
                row.get('耐压/额定电压', ''),
                row.get('额定功率', ''),
                row.get('精度', ''),
                row.get('材质', ''),
                row.get('类型', ''),
            )
        )
        if key not in groups:
            groups[key] = {
                '料号': row.get('料号', ''),
                '位号列表': [],
                '数量': 0,
                '贴装数量': 0,
                'DEPOP数量': 0,
                'BOM状态': '',
                '描述': row.get('描述', ''),
                '值': row.get('值', ''),
                '封装': row.get('封装', ''),
                '耐压': row.get('耐压/额定电压', ''),
                '额定功率': row.get('额定功率', ''),
                '精度': row.get('精度', ''),
                '材质': row.get('材质', ''),
                '类型': row.get('类型', ''),
            }
        group = groups[key]
        group['位号列表'].append(row.get('位号', ''))
        group['数量'] += 1
        if row.get('BOM状态') == 'DEPOP':
            group['DEPOP数量'] += 1
        else:
            group['贴装数量'] += 1

    merged = list(groups.values())
    for row in merged:
        states = []
        if row['贴装数量']:
            states.append('贴装')
        if row['DEPOP数量']:
            states.append('DEPOP')
        row['BOM状态'] = ' / '.join(states)
        row['位号列表'] = ', '.join(sorted((item for item in row['位号列表'] if item), key=_natural_sort_key))
    merged.sort(key=lambda row: (
        _bom_type_order_from_label(row.get('类型', '')),
        row.get('料号', ''),
        row.get('描述', ''),
        row.get('值', ''),
        row.get('封装', ''),
    ))
    for index, row in enumerate(merged, 1):
        row['序号'] = index
    return total_detail, merged
