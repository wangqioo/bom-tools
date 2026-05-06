# -*- coding: utf-8 -*-
"""Capacitor derating checks."""

import re
from typing import Dict, List, Optional, Tuple

from pstx_core.page_resolution import _component_page_fields, component_user_visible_page
from pstx_rules.common import (
    COMP_TYPE_CN,
    _collect_component_nets,
    _find_ac_coupling_candidates,
    _first_net_token,
    _is_depop_option,
    _matches_prefix_with_boundary,
    _natural_sort_key,
    _net_is_gnd,
    _parse_voltage_token,
    _token_is_ground,
    _unique_component_nets,
)
from pstx_rules.result_meta import meta_fields as _meta_fields

_SAFE_VOLTAGE_EXAMPLES = '例如 P1V8=1.8、VDD_3V3=3.3、VBAT=4.4'

def _match_custom_voltage(net_name: str,
                          custom_volt_map: Optional[Dict[str, float]]) -> Optional[float]:
    if not custom_volt_map:
        return None
    best: Optional[Tuple[int, float]] = None
    for key, volt in custom_volt_map.items():
        prefix = str(key).strip().upper()
        if prefix and _matches_prefix_with_boundary(net_name, prefix):
            if best is None or len(prefix) > best[0]:
                best = (len(prefix), float(volt))
    return best[1] if best else None


def _infer_voltage(net_name: str) -> Optional[float]:
    token = _first_net_token(net_name)
    if _token_is_ground(token):
        return 0.0
    return _parse_voltage_token(token)


def _collect_global_max_voltage(nets: Dict,
                                custom_volt_map: Optional[Dict[str, float]] = None
                                ) -> Tuple[Optional[float], str, str]:
    max_voltage: Optional[float] = None
    max_net = ''
    max_source = ''
    for net_name in nets:
        v = _match_custom_voltage(net_name, custom_volt_map)
        source = 'custom_map' if v is not None else ''
        if v is None:
            v = _infer_voltage(net_name)
            if v is not None:
                source = 'net_token'
        if v is None or v <= 0:
            continue
        if max_voltage is None or float(v) > max_voltage:
            max_voltage = float(v)
            max_net = net_name
            max_source = source
    return max_voltage, max_net, max_source


def analyze_derating(components: Dict, nets: Dict,
                     pct: float = 70.0,
                     custom_volt_map: Optional[Dict[str, float]] = None) -> List[dict]:
    """pct: 工作电压上限占额定电压的百分比，默认 70（即工作电压 ≤ 额定 × 70% 视为合格）"""
    comp_nets = _collect_component_nets(nets)
    ac_coupling_caps = _find_ac_coupling_candidates(components, comp_nets, nets)
    global_max_v, global_max_net, global_max_source = _collect_global_max_voltage(nets, custom_volt_map)

    rows = []
    for refdes, comp in components.items():
        ctype = comp.get('comp_type', '')
        if ctype not in ('CAP', 'CAP_POL'):
            continue
        connected_nets = _unique_component_nets(comp_nets, refdes)
        rated_str = comp.get('voltage', '')
        source_type = ''
        if not rated_str:
            status, derating, max_v, from_net = '⚪ 无额定电压', None, None, ''
            meta = _meta_fields('indeterminate', 'medium', 'high', 'missing_rated_voltage')
        else:
            m = re.match(r'([\d.]+)\s*V', rated_str.strip(), re.I)
            rated_v = float(m.group(1)) if m else None
            if rated_v is None:
                status, derating, max_v, from_net = '⚪ 无法解析额定电压', None, None, ''
                meta = _meta_fields('indeterminate', 'medium', 'high', 'unparsed_rated_voltage')
            elif global_max_v is not None and global_max_v <= 12.0 and rated_v >= 50.0:
                max_v = global_max_v
                from_net = global_max_net
                derating = rated_v / max_v if max_v else None
                source_type = '全局最大电压(自定义映射)' if global_max_source == 'custom_map' else '全局最大电压(网络名 token)'
                status = f'✅ 合格 (全局最大电压 {max_v:.1f}V ≤ 12V，50V 高耐压器件直接通过)'
                meta = _meta_fields(
                    'confirmed' if global_max_source == 'custom_map' else 'candidate',
                    'low',
                    'high' if global_max_source == 'custom_map' else 'medium',
                    'global_max_voltage_under_12v_high_rated_cap',
                )
            elif refdes in ac_coupling_caps:
                ac_info = ac_coupling_caps.get(refdes, {})
                status, derating, max_v = '✅ 低风险通过（差分同极性 AC 耦合）', None, None
                from_net = ' ↔ '.join(ac_info.get('nets', []))
                source_type = '差分同极性 AC 耦合'
                meta = _meta_fields('confirmed', 'low', 'medium', 'ac_coupling_same_polarity_diff_pair')
            else:
                max_v, from_net = None, ''
                known_nets: List[Tuple[str, float, str]] = []
                ground_present = False
                for net_name in connected_nets:
                    if _net_is_gnd(net_name):
                        ground_present = True
                    v = _match_custom_voltage(net_name, custom_volt_map)
                    source = 'custom_map' if v is not None else ''
                    if v is None:
                        v = _infer_voltage(net_name)
                        if v is not None:
                            source = 'net_token'
                    if v is None:
                        continue
                    if v == 0:
                        ground_present = True
                    known_nets.append((net_name, float(v), source))

                positives: Dict[float, Tuple[str, str]] = {}
                for net_name, v, source in known_nets:
                    if v > 0:
                        positives.setdefault(round(v, 6), (net_name, source))

                if not ground_present:
                    status, derating = '⚪ 无法判断（未连接地）', None
                    meta = _meta_fields('indeterminate', 'low', 'high', 'no_ground_reference')
                elif not positives:
                    status, derating = '⚪ 无法推断工作电压', None
                    meta = _meta_fields('indeterminate', 'low', 'high', 'no_positive_voltage_evidence')
                elif len(positives) > 1:
                    status, derating = '⚪ 无法判断（连接多个不同电位）', None
                    meta = _meta_fields('indeterminate', 'medium', 'high', 'multiple_positive_rails')
                else:
                    rounded_v, (from_net, source) = next(iter(positives.items()))
                    max_v = rounded_v
                    source_type = '自定义映射' if source == 'custom_map' else '网络首 token'
                    usage_pct = max_v / rated_v * 100        # 工作电压占额定的 %
                    derating  = rated_v / max_v              # 仍保留降额比供参考
                    if usage_pct <= pct:
                        status = f'✅ 合格 ({usage_pct:.0f}% ≤ {pct:.0f}%)'
                    else:
                        status = f'❌ 不合格 ({usage_pct:.0f}% > {pct:.0f}%)'
                    if source == 'custom_map':
                        meta = _meta_fields(
                            'confirmed',
                            'high' if status.startswith('❌') else 'low',
                            'high',
                            'custom_voltage_map',
                        )
                    else:
                        meta = _meta_fields(
                            'candidate',
                            'medium' if status.startswith('❌') else 'low',
                            'medium',
                            'single_positive_rail_token',
                        )
        rows.append({
            '位号':            refdes,
            '值':              comp.get('value', ''),
            '封装':            comp.get('package', ''),
            '类型':            COMP_TYPE_CN.get(ctype, ctype),
            '额定电压':        rated_str,
            '推断工作电压(V)': str(max_v) if max_v is not None else '',
            '推断来源网络':    from_net,
            '推断来源类型':    source_type,
            '所有连接网络':    ', '.join(connected_nets),
            '降额比':          f'{derating:.2f}' if derating is not None else '',
            '状态':            status,
            **meta,
            '页面':            component_user_visible_page(comp),
            'DEPOP':           'Y' if _is_depop_option(comp.get('bom_option', '')) else '',
        })
        rows[-1].update(_component_page_fields(comp))
    rows.sort(key=lambda r: (
        0 if r['状态'].startswith('❌') else 1 if r.get('结论类型') == '候选判断' else 2,
        {'高': 0, '中': 1, '低': 2}.get(r.get('严重级别', ''), 9),
        _natural_sort_key(r.get('位号', '')),
    ))
    return rows
