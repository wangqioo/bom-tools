# -*- coding: utf-8 -*-
"""Shared rule helpers for PSTX analysis.

Keep this module free of Web/Harness imports; rule modules can reuse these
small pure helpers without depending on the historical analyzer entrypoint.
"""

import ntpath
import os
import re
from collections import defaultdict, deque
from pathlib import Path, PureWindowsPath
from typing import Dict, List, Optional, Tuple

_XY_POINT_RE = re.compile(r"\((-?\d+(?:\.\d+)?)\s*,?\s*(-?\d+(?:\.\d+)?)\)")
_NUMBER_VALUE_RE = re.compile(r"-?\d+(?:\.\d+)?")


def _parse_xy_point(value: object) -> Optional[Tuple[float, float]]:
    text = str(value or '').strip()
    if not text:
        return None
    match = _XY_POINT_RE.search(text)
    if match:
        return float(match.group(1)), float(match.group(2))
    nums = _NUMBER_VALUE_RE.findall(text)
    if len(nums) >= 2:
        return float(nums[0]), float(nums[1])
    return None


def _format_measure(value: Optional[float]) -> str:
    if value is None:
        return ''
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f'{value:.1f}'


def _format_ratio(value: Optional[float]) -> str:
    if value is None:
        return ''
    return f'{value:.2f}'


def _format_percent(value: Optional[float]) -> str:
    if value is None:
        return ''
    return f'{max(0.0, min(100.0, value)):.0f}%'


def _looks_like_windows_path(path_text: str) -> bool:
    return bool(re.match(r'^[A-Za-z]:[\\/]', path_text)) or '\\' in path_text


def _infer_project_root_from_data_paths(*paths: str) -> str:
    raw_paths = [str(raw_path or '').strip().strip('"') for raw_path in paths if str(raw_path or '').strip()]
    if not raw_paths:
        return ''

    windows_mode = any(_looks_like_windows_path(path_text) for path_text in raw_paths)
    candidates = []
    for path_text in raw_paths:
        try:
            if windows_mode:
                candidates.append(PureWindowsPath(path_text))
            else:
                candidates.append(Path(path_text).expanduser().resolve())
        except OSError:
            continue
    if not candidates:
        return ''

    for path in candidates:
        parent = path.parent
        if parent.name.lower() == 'packaged':
            return str(parent.parent)

    parent_strings = [str(path.parent) for path in candidates]
    try:
        if windows_mode:
            common_parent = PureWindowsPath(ntpath.commonpath(parent_strings))
        else:
            common_parent = Path(os.path.commonpath(parent_strings))
    except ValueError:
        common_parent = candidates[0].parent
    if common_parent.name.lower() == 'packaged':
        return str(common_parent.parent)
    return str(common_parent)


def _split_net_tokens(net_name: str) -> List[str]:
    return [tok for tok in re.split(r'[_./-]+', (net_name or '').upper()) if tok]


def _first_net_token(net_name: str) -> str:
    tokens = _split_net_tokens(net_name)
    return tokens[0] if tokens else (net_name or '').upper()


def _matches_prefix_with_boundary(name: str, prefix: str) -> bool:
    if not prefix:
        return False
    name = (name or '').upper()
    prefix = prefix.upper()
    if not name.startswith(prefix):
        return False
    return len(name) == len(prefix) or name[len(prefix)] in '_./-'


def _parse_voltage_token(token: str) -> Optional[float]:
    m = re.fullmatch(r'P?(\d+)V(\d*)', token.upper())
    if not m:
        return None
    int_part, frac_part = m.groups()
    return float(f'{int_part}.{frac_part}') if frac_part else float(int_part)


_POWER_TOKEN_RE = re.compile(
    r'(?:VCC|VDD|VBAT|VCORE|VCCIO|PVDD|PVCC|AVDD|DVDD|VBUS)[A-Z0-9]*',
    re.IGNORECASE,
)
# Ground nets often carry domain prefixes/suffixes, for example AGND1,
# GNDA, VSSA, or AVSS. Treat them as terminal ground nodes in resistor walks.
_GROUND_TOKEN_RE = re.compile(
    r'(?:[A-Z0-9]*GND[A-Z0-9]*|[A-Z0-9]*VSS[A-Z0-9]*|0V|0)',
    re.IGNORECASE,
)


def _token_is_power(token: str) -> bool:
    return _parse_voltage_token(token) is not None or bool(_POWER_TOKEN_RE.fullmatch(token))


def _token_is_ground(token: str) -> bool:
    return bool(_GROUND_TOKEN_RE.fullmatch(token))


def _net_is_power(net: str) -> bool:
    return _token_is_power(_first_net_token(net))


def _net_is_gnd(net: str) -> bool:
    return _token_is_ground(_first_net_token(net))


_DIFF_SUFFIX_PAIRS = [
    ('_P', '_N'),
    ('_DP', '_DN'),
    ('.P', '.N'),
    ('_TXPLUS', '_TXMINUS'),
    ('_RXPLUS', '_RXMINUS'),
]


def _get_diff_net_info(net_name: str, upper_name_map: Dict[str, str]) -> Optional[Dict[str, str]]:
    upper_name = (net_name or '').upper()
    for pos_suffix, neg_suffix in _DIFF_SUFFIX_PAIRS:
        pos_upper = pos_suffix.upper()
        neg_upper = neg_suffix.upper()
        if upper_name.endswith(pos_upper):
            partner = upper_name_map.get(upper_name[:-len(pos_upper)] + neg_upper)
            if partner:
                return {
                    'base': net_name[:-len(pos_suffix)],
                    'polarity': 'P',
                    'partner': partner,
                }
        elif upper_name.endswith(neg_upper):
            partner = upper_name_map.get(upper_name[:-len(neg_upper)] + pos_upper)
            if partner:
                return {
                    'base': net_name[:-len(neg_suffix)],
                    'polarity': 'N',
                    'partner': partner,
                }
    return None


def _collect_diff_pairs(nets: Dict) -> Dict[str, dict]:
    diff_pairs: Dict[str, dict] = {}
    upper_name_map = {name.upper(): name for name in nets}
    for net_name in nets:
        info = _get_diff_net_info(net_name, upper_name_map)
        if not info:
            continue
        base = info['base']
        if info['polarity'] == 'P':
            diff_pairs[base] = {'P': net_name, 'N': info['partner']}
        elif base not in diff_pairs:
            diff_pairs[base] = {'P': info['partner'], 'N': net_name}
    return diff_pairs


def _collect_component_nets(nets: Dict) -> Dict[str, List[str]]:
    comp_nets: Dict[str, List[str]] = defaultdict(list)
    for net_name, nodes in nets.items():
        for node in nodes:
            comp_nets[node['refdes']].append(net_name)
    return comp_nets


def _unique_component_nets(comp_nets: Dict[str, List[str]], refdes: str) -> List[str]:
    return list(dict.fromkeys(comp_nets.get(refdes, [])))


def _find_ac_coupling_candidates(components: Dict,
                                 comp_nets: Dict[str, List[str]],
                                 nets: Dict) -> Dict[str, dict]:
    upper_name_map = {name.upper(): name for name in nets}
    diff_info_map = {
        net_name: info
        for net_name in nets
        if (info := _get_diff_net_info(net_name, upper_name_map))
    }
    cap_pairs: Dict[str, Tuple[str, str]] = {}
    caps_by_pair: Dict[frozenset, List[str]] = defaultdict(list)

    for refdes, comp in components.items():
        if comp.get('comp_type') not in ('CAP', 'CAP_POL'):
            continue
        unique_nets = _unique_component_nets(comp_nets, refdes)
        if len(unique_nets) != 2:
            continue
        net_a, net_b = unique_nets
        if _net_is_power(net_a) or _net_is_power(net_b) or _net_is_gnd(net_a) or _net_is_gnd(net_b):
            continue
        cap_pairs[refdes] = (net_a, net_b)
        caps_by_pair[frozenset((net_a, net_b))].append(refdes)

    candidates: Dict[str, dict] = {}
    for refdes, (net_a, net_b) in cap_pairs.items():
        info_a = diff_info_map.get(net_a)
        info_b = diff_info_map.get(net_b)
        if not info_a or not info_b:
            continue
        if info_a['polarity'] != info_b['polarity']:
            continue
        partner_pair = frozenset((info_a['partner'], info_b['partner']))
        mirror_caps = sorted(
            (cap for cap in caps_by_pair.get(partner_pair, []) if cap != refdes),
            key=_natural_sort_key,
        )
        candidates[refdes] = {
            'nets': (net_a, net_b),
            'mirror_nets': sorted(partner_pair, key=_natural_sort_key),
            'mirror_caps': mirror_caps,
            'polarity': info_a['polarity'],
        }
    return candidates


def _natural_sort_key(value: str):
    parts = re.split(r'(\d+)', str(value or '').upper())
    return [int(p) if p.isdigit() else p for p in parts]


def natural_sort_key(value: str):
    """Public natural sort key for view/query layers."""
    return _natural_sort_key(value)


def _normalize_bom_option(value: str) -> str:
    return (value or '').strip().upper()


def _display_bom_option(value: str) -> str:
    normalized = _normalize_bom_option(value)
    return normalized or '默认'


def display_bom_option(value: str) -> str:
    """Public BOM_OPTION display formatter."""
    return _display_bom_option(value)


def _is_depop_option(bom_option: str) -> bool:
    return _normalize_bom_option(bom_option) in {'DEPOP', 'DNP'}


def _clone_components(components: Dict[str, dict], allowed_nets: Optional[set] = None) -> Dict[str, dict]:
    cloned: Dict[str, dict] = {}
    for refdes, comp in components.items():
        cloned_comp = dict(comp)
        nets_map = dict(comp.get('nets', {}))
        if allowed_nets is not None:
            nets_map = {pin: net for pin, net in nets_map.items() if net in allowed_nets}
        cloned_comp['nets'] = nets_map
        cloned[refdes] = cloned_comp
    return cloned


def _collect_depop_refdes(components: Dict[str, dict]) -> List[str]:
    return sorted(
        [refdes for refdes, comp in components.items() if _is_depop_option(comp.get('bom_option', ''))],
        key=_natural_sort_key,
    )


def _build_analysis_scope(components: Dict[str, dict],
                          nets: Dict[str, List[dict]],
                          *,
                          include_depop: bool) -> Tuple[Dict[str, dict], Dict[str, List[dict]], List[str], List[str]]:
    depop_refdes = _collect_depop_refdes(components)
    if include_depop or not depop_refdes:
        active_nets = {
            net_name: [dict(node) for node in nodes]
            for net_name, nodes in nets.items()
        }
        active_components = _clone_components(components, set(active_nets.keys()))
        return active_components, active_nets, depop_refdes, []

    excluded = set(depop_refdes)
    active_nets: Dict[str, List[dict]] = {}
    for net_name, nodes in nets.items():
        filtered_nodes = [dict(node) for node in nodes if node.get('refdes') not in excluded]
        if filtered_nodes:
            active_nets[net_name] = filtered_nodes

    active_components = _clone_components(
        {refdes: comp for refdes, comp in components.items() if refdes not in excluded},
        set(active_nets.keys()),
    )
    return active_components, active_nets, depop_refdes, depop_refdes[:]


# ══════════════════════════════════════════════════════════
# 二、BOM 分析
# ══════════════════════════════════════════════════════════

COMP_TYPE_CN = {
    'CAP': '电容', 'CAP_POL': '电解/钽电容', 'RES': '电阻',
    'IND': '电感/磁珠', 'IC': 'IC 芯片', 'CONN': '连接器',
    'DIODE': '二极管', 'LED': 'LED', 'FET': 'MOS/FET',
    'BJT': '三极管', 'XTAL': '晶振', 'FUSE': '保险丝',
    'SWITCH': '开关', 'TESTPOINT': '测试点', 'TRANSFORMER': '变压器',
}
_TYPE_ORDER = list(COMP_TYPE_CN.keys())


def component_type_label(comp_type: str) -> str:
    """Return the user-facing component type label."""
    return COMP_TYPE_CN.get(comp_type, comp_type)
