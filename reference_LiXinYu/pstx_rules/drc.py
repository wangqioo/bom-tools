# -*- coding: utf-8 -*-
"""General DRC rule checks."""

import re
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

from pstx_core.page_resolution import USER_VISIBLE_REAL_PAGE_LABEL, _component_page_fields, component_user_visible_page
from pstx_rules.common import (
    COMP_TYPE_CN,
    _display_bom_option,
    _first_net_token,
    _is_depop_option,
    _matches_prefix_with_boundary,
    _natural_sort_key,
    _normalize_bom_option,
    _token_is_ground,
    _token_is_power,
)
from pstx_rules.result_meta import meta_fields as _meta_fields, with_meta as _with_meta

_VALID_BOM_OPTIONS = {'', 'DEPOP', 'OPTION', 'MAIN_PLD', 'MAIN', 'ALT', 'DNP'}
_FUZZY_KEYWORDS    = sorted(opt for opt in _VALID_BOM_OPTIONS if opt)


def _edit_distance(a: str, b: str) -> int:
    a, b = a.upper(), b.upper()
    if a == b: return 0
    if not a:  return len(b)
    if not b:  return len(a)
    dp = list(range(len(b) + 1))
    for i, ca in enumerate(a):
        prev = dp[:]
        dp[0] = i + 1
        for j, cb in enumerate(b):
            dp[j+1] = min(prev[j] + (0 if ca == cb else 1), dp[j]+1, prev[j+1]+1)
    return dp[len(b)]


def _net_nodes_page_summary(nodes: List[dict], components: Dict[str, dict]) -> Tuple[str, str]:
    pages: List[str] = []
    endpoints: List[str] = []
    for node in nodes or []:
        refdes = str(node.get('refdes', '')).strip()
        pin = str(node.get('pin_name') or node.get('pin') or '').strip()
        comp = components.get(refdes, {})
        page = component_user_visible_page(comp)
        if page and page not in pages:
            pages.append(page)
        endpoint = f'{refdes}.{pin}' if pin else refdes
        if endpoint and endpoint not in endpoints:
            endpoints.append(endpoint)
    return ', '.join(pages), ', '.join(endpoints)


def check_drc(components: Dict,
              nets: Dict,
              *,
              option_components_source: Optional[Dict] = None,
              single_pin_components: Optional[Dict] = None,
              single_pin_nets: Optional[Dict] = None) -> dict:
    missing_hq, missing_val, missing_pkg, tbd_attrs, single_pin, unnamed = [], [], [], [], [], []
    bom_option_components = []
    option_source = option_components_source if option_components_source is not None else components
    single_scope_components = single_pin_components if single_pin_components is not None else components
    single_scope_nets = single_pin_nets if single_pin_nets is not None else nets
    for refdes, comp in components.items():
        ctype = comp.get('comp_type', '')
        if ctype == 'TESTPOINT':
            continue
        page_label = component_user_visible_page(comp)
        page_fields = _component_page_fields(comp)
        base = {'位号': refdes, '类型': COMP_TYPE_CN.get(ctype, ctype), '页面': page_label}
        base.update(page_fields)
        if not comp.get('hq_code'):
            missing_hq.append(_with_meta(base.copy(), 'confirmed', 'high', 'high', 'missing_hq_code'))
        if not comp.get('value'):
            missing_val.append(_with_meta(base.copy(), 'confirmed', 'high', 'high', 'missing_value'))
        if not comp.get('package'):
            missing_pkg.append(_with_meta(base.copy(), 'confirmed', 'high', 'high', 'missing_package'))
        for attr in ('voltage', 'current', 'power'):
            val = comp.get(attr, '')
            if val and 'TBD' in val.upper():
                tbd_attrs.append(_with_meta({
                    '位号': refdes, '属性': attr.upper(), '当前值': val,
                    '类型': COMP_TYPE_CN.get(ctype, ctype), '页面': page_label
                }, 'confirmed', 'medium', 'high', 'tbd_attribute'))
                tbd_attrs[-1].update(page_fields)
    active_refdes = set(components.keys())
    for net_name, nodes in single_scope_nets.items():
        if len(nodes) == 1:
            n = nodes[0]
            if n.get('refdes') not in active_refdes:
                continue
            comp = single_scope_components.get(n['refdes'], components.get(n['refdes'], {}))
            if comp.get('comp_type') != 'TESTPOINT' and not re.search(r'^UNNAMED_', net_name, re.I):
                single_pin.append(_with_meta({
                    '网络名': net_name, '连接元件': n['refdes'],
                    '引脚': n['pin_name'], '页面': component_user_visible_page(comp)
                }, 'candidate', 'medium', 'low', 'single_pin_net'))
                single_pin[-1].update(_component_page_fields(comp))
        if re.search(r'^UNNAMED_', net_name, re.I):
            active_nodes = [node for node in nodes if node.get('refdes') in active_refdes]
            if not active_nodes:
                continue
            page_summary, endpoint_summary = _net_nodes_page_summary(active_nodes, single_scope_components)
            unnamed.append(_with_meta({'网络名': net_name, '节点数': len(active_nodes),
                                      '连接点': endpoint_summary, '页面': page_summary,
                                      USER_VISIBLE_REAL_PAGE_LABEL: page_summary},
                                      'candidate', 'medium', 'high', 'unnamed_net'))
    option_map: Dict[str, List[str]] = defaultdict(list)
    for refdes, comp in option_source.items():
        ctype = comp.get('comp_type', '')
        if ctype == 'TESTPOINT':
            continue
        bom_option = _normalize_bom_option(comp.get('bom_option', ''))
        if bom_option:
            bom_option_components.append({
                '位号': refdes,
                '类型': COMP_TYPE_CN.get(ctype, ctype),
                'BOM_OPTION值': bom_option,
                '是否DEPOP': '是' if _is_depop_option(bom_option) else '否',
                '页面': component_user_visible_page(comp),
            })
            bom_option_components[-1].update(_component_page_fields(comp))
        option_map[_normalize_bom_option(comp.get('bom_option'))].append(refdes)
    typos = []
    for val, refs in sorted(option_map.items()):
        if val in _VALID_BOM_OPTIONS:
            continue
        min_d   = min(_edit_distance(val, kw) for kw in _FUZZY_KEYWORDS)
        nearest = min(_FUZZY_KEYWORDS, key=lambda kw: _edit_distance(val, kw))
        typos.append(_with_meta({
            '实际填写值': val, '疑似应为': nearest if min_d <= 2 else '未知',
            '编辑距离': min_d, '使用该值的位号': ', '.join(sorted(refs, key=_natural_sort_key)),
            '数量': len(refs), '风险': '❌ 疑似拼错' if min_d <= 2 else '⚠ 未知值'
        }, 'candidate', 'medium', 'medium' if min_d <= 2 else 'low', 'bom_option_typo'))
    return {
        'missing_hq_code': missing_hq, 'missing_value': missing_val,
        'missing_package': missing_pkg, 'tbd_attrs': tbd_attrs,
        'single_pin_nets': single_pin, 'unnamed_nets': unnamed,
        'bom_option_typos': typos,
        'bom_option_components': sorted(
            bom_option_components,
            key=lambda r: (_natural_sort_key(r['位号']), _natural_sort_key(r['页面'])),
        ),
    }


_DRC_REPORT_KEYS = ['bom_option_components', 'bom_option_circle_coverage']


# ══════════════════════════════════════════════════════════
# 五、电容降额分析
# ══════════════════════════════════════════════════════════

_SAFE_VOLTAGE_EXAMPLES: List[Tuple[str, float]] = [
    ('P12V_AUX', 12.0),
    ('12V_MAIN', 12.0),
    ('P5V_STBY', 5.0),
    ('P3V3_AON', 3.3),
    ('P1V8_S0', 1.8),
    ('P1V05_RTC', 1.05),
    ('GND', 0.0),
]
