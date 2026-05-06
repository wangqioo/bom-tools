# -*- coding: utf-8 -*-
"""Pull-up, pull-down, series resistor and chip pin bias checks."""

import re
from collections import defaultdict, deque
from typing import Dict, List, Optional, Tuple

from pstx_core.page_resolution import USER_VISIBLE_REAL_PAGE_LABEL, _component_page_fields, component_user_visible_page
from pstx_rules.common import (
    _display_bom_option,
    _first_net_token,
    _is_depop_option,
    _natural_sort_key,
    _parse_voltage_token,
    _token_is_ground,
    _token_is_power,
)
from pstx_rules.result_meta import meta_fields as _meta_fields, with_meta as _with_meta

def _parse_ohms(value_str: str) -> Optional[float]:
    """解析电阻值字符串为欧姆数，支持 k/M/R/Ω 后缀，如 10k→10000, 4.7k→4700, 100R→100"""
    if not value_str:
        return None
    s = re.sub(r'\s', '', value_str.upper())
    s = s.replace('Ω', 'R').replace('Ω', 'R').replace('欧', 'R')
    s = re.sub(r'OHMS?$', 'R', s)
    s = re.sub(r'([KMG])R$', r'\1', s)
    m = re.match(r'^([\d.]+)([KMGR]?)$', s)
    if not m:
        embedded = re.match(r'^(\d+)([KMGR])(\d+)$', s)
        if not embedded:
            return None
        val = float(f'{embedded.group(1)}.{embedded.group(3)}')
        return val * {'K': 1e3, 'M': 1e6, 'G': 1e9, 'R': 1}.get(embedded.group(2), 1)
    val = float(m.group(1))
    return val * {'K': 1e3, 'M': 1e6, 'G': 1e9, 'R': 1, '': 1}.get(m.group(2), 1)


def _net_is_power(net: str) -> bool:
    return _token_is_power(_first_net_token(net))


def _net_is_gnd(net: str) -> bool:
    return _token_is_ground(_first_net_token(net))


_CHIP_REFDES_RE = re.compile(r'^(?:XU|PU|U)[A-Z0-9]+$', re.IGNORECASE)
_REFDES_SUFFIX_RE = re.compile(r'(?<=\d)([A-Z]+\d+)$', re.IGNORECASE)


def _is_chip_component(refdes: str, comp: Dict) -> bool:
    return comp.get('comp_type') == 'IC' and bool(_CHIP_REFDES_RE.match(refdes or ''))


def _extract_refdes_suffix_group(refdes: str) -> str:
    match = _REFDES_SUFFIX_RE.search((refdes or '').strip().upper())
    return match.group(1) if match else ''


def _extract_pin_hierarchy_labels(pin_name: str) -> List[str]:
    raw = str(pin_name or '').strip()
    if '@' not in raw:
        return []

    labels: List[str] = []
    for segment in raw.split('@'):
        segment = segment.strip()
        if not segment:
            continue
        head = segment.split(':', 1)[0].strip()
        if '(' in head:
            head = head.split('(', 1)[0].strip()
        label = head.rsplit('.', 1)[-1].strip()
        if label:
            labels.append(label)
    return labels


def _extract_pin_submodule_info(pin_name: str) -> Tuple[str, str]:
    labels = _extract_pin_hierarchy_labels(pin_name)
    if len(labels) >= 2:
        return labels[-2], ' / '.join(labels[:-1])
    if labels:
        return labels[0], labels[0]
    return '', ''


def _format_entry_list(entries: List[dict], key: str) -> str:
    return ', '.join(dict.fromkeys(str(entry.get(key, '')) for entry in entries if entry.get(key, '') != ''))


def _merge_display_values(*values: str) -> str:
    merged: List[str] = []
    for value in values:
        for part in str(value or '').split(','):
            part = part.strip()
            if part and part not in merged:
                merged.append(part)
    return ', '.join(merged)


def _series_edge_sort_key(entry: Dict[str, object]):
    return (_natural_sort_key(entry.get('refdes', '')), _natural_sort_key(entry.get('other_net', '')))


def _series_bias_sort_key(entry: Dict[str, object]):
    return (
        int(entry.get('via_hop_count', 0) or 0),
        _natural_sort_key(entry.get('refdes', '')),
        _natural_sort_key(entry.get('source_net', '')),
    )


def _series_chain_field(chain: List[dict], key: str) -> str:
    return ' -> '.join(str(edge.get(key, '')) for edge in chain if edge.get(key, ''))


def _series_chain_pages(chain: List[dict], key: str) -> str:
    return ', '.join(
        dict.fromkeys(str(edge.get(key, '')) for edge in chain if edge.get(key, ''))
    )


def _series_chain_total_ohms(chain: List[dict]) -> Optional[float]:
    total = 0.0
    for edge in chain:
        ohms = edge.get('ohms')
        if ohms is None:
            return None
        total += float(ohms)
    return total


def _classify_resistor_usage(refdes: str, comp: Dict, ohms: Optional[float]) -> str:
    """Give series/bias rows a conservative usage hint without changing topology."""
    text = ' '.join(
        str(comp.get(key, '') or '')
        for key in ('part_name', 'value', 'package', 'material')
    ).upper()
    ref = (refdes or '').upper()
    if (ohms is not None and abs(float(ohms)) <= 0.5) or re.search(
        r'(^|[^A-Z0-9])(0R|0OHM|0Ω|JUMPER|LINK|SHORT)([^A-Z0-9]|$)',
        text,
    ):
        return '0R/跳线'
    if re.match(r'^(?:RN|RP|RA)\d+', ref) or re.search(
        r'ARRAY|RPACK|RES[_ -]?ARRAY|排阻|电阻阵列',
        text,
    ):
        return '排阻/电阻阵列'
    return '普通电阻'


def _bom_option_risk_fields(entries: List[dict], reason_code: str) -> Tuple[str, str, str, str]:
    options = [
        _display_bom_option(str(entry.get('bom_option', '') or ''))
        for entry in entries
    ]
    unique_options = list(dict.fromkeys(options))
    explicit_options = [option for option in unique_options if option != '默认']
    if explicit_options and '默认' not in unique_options and len(explicit_options) > 1:
        return (
            'low',
            'low',
            f'{reason_code}_bom_option_variant',
            '多个偏置分属不同 BOM_OPTION，可能是互斥装配；按候选低风险提示，需结合装配方案确认。',
        )
    if len(explicit_options) == 1 and '默认' not in unique_options:
        return (
            'medium',
            'medium',
            reason_code,
            f'多个偏置均属于 BOM_OPTION={explicit_options[0]}，同一装配条件下可能同时存在。',
        )
    if '默认' in unique_options and len(unique_options) > 1:
        return (
            'medium',
            'medium',
            reason_code,
            '存在默认装配与 BOM_OPTION 偏置混合，无法确认是否互斥，按中风险候选提示。',
        )
    return (
        'medium',
        'medium',
        reason_code,
        '多个偏置未体现互斥装配信息，按同一网络重复偏置候选提示。',
    )


def _merge_duplicate_bias_entries(entries: List[dict]) -> List[dict]:
    merged: Dict[Tuple[str, str, str], dict] = {}
    for entry in entries:
        key = (
            str(entry.get('refdes', '') or ''),
            str(entry.get('source_net', '') or ''),
            str(entry.get('reference_net', '') or ''),
        )
        if key not in merged:
            merged[key] = dict(entry)
            continue
        current = merged[key]
        for field in ('via_refdes_chain', 'via_usage_chain', 'via_net_chain', 'via_pages', 'via_mapped_pages', 'connection_mode', 'usage'):
            current[field] = _merge_display_values(current.get(field, ''), entry.get(field, ''))
    return sorted(merged.values(), key=lambda row: _natural_sort_key(row.get('refdes', '')))


def _duplicate_bias_rows(direct_map: Dict[str, list],
                         indirect_map: Dict[str, list],
                         *,
                         bias_kind: str) -> List[dict]:
    rows: List[dict] = []
    count_label = '上拉数量' if bias_kind == '上拉' else '下拉数量'
    reason_base = 'multiple_pullup_paths' if bias_kind == '上拉' else 'multiple_pulldown_paths'

    candidate_nets = sorted(set(direct_map.keys()) | set(indirect_map.keys()), key=_natural_sort_key)
    for sig_net in candidate_nets:
        entries: List[dict] = []
        for item in direct_map.get(sig_net, []):
            entries.append({
                **item,
                'source_net': sig_net,
                'reference_net': item.get('power_net', 'GND' if bias_kind == '下拉' else ''),
                'connection_mode': '直接',
                'via_refdes_chain': '',
                'via_net_chain': '',
                'via_pages': '',
                'via_mapped_pages': '',
            })
        for item in indirect_map.get(sig_net, []):
            entries.append({
                **item,
                'reference_net': item.get('power_net', 'GND' if bias_kind == '下拉' else ''),
                'connection_mode': '隔串阻',
            })

        group = _merge_duplicate_bias_entries(entries)
        if len(group) < 2:
            continue

        has_series_path = any(entry.get('connection_mode') != '直接' for entry in group)
        reason = f'{reason_base}_with_series' if has_series_path else reason_base
        severity, confidence, reason, option_note = _bom_option_risk_fields(group, reason)
        page_values = _merge_display_values(
            ', '.join(entry.get('via_pages', '') for entry in group),
            ', '.join(entry.get('page', '') for entry in group),
        )
        mapped_page_values = _merge_display_values(
            ', '.join(entry.get('via_mapped_pages', '') for entry in group),
            ', '.join(entry.get('mapped_page', '') for entry in group),
        )

        row_body = {
            '信号网络': sig_net,
            count_label: len(group),
            '位号': ', '.join(entry['refdes'] for entry in group),
            '阻值': ', '.join(entry.get('value', '') for entry in group),
            '连接方式': ', '.join(dict.fromkeys(entry.get('connection_mode', '') for entry in group if entry.get('connection_mode', ''))),
            '偏置所在网络': ', '.join(dict.fromkeys(entry.get('source_net', '') for entry in group if entry.get('source_net', ''))),
            '隔串阻链': ', '.join(dict.fromkeys(entry.get('via_refdes_chain', '') for entry in group if entry.get('via_refdes_chain', ''))),
            '隔串阻经过网络': ', '.join(dict.fromkeys(entry.get('via_net_chain', '') for entry in group if entry.get('via_net_chain', ''))),
            'BOM_OPTION': ', '.join(dict.fromkeys(_display_bom_option(entry.get('bom_option', '')) for entry in group)),
            '装配选项判断': option_note,
            '偏置电阻类型候选': ', '.join(dict.fromkeys(entry.get('usage', '') for entry in group if entry.get('usage', ''))),
            '串阻类型候选': ', '.join(dict.fromkeys(entry.get('via_usage_chain', '') for entry in group if entry.get('via_usage_chain', ''))),
            '页面': mapped_page_values or page_values,
        }
        if bias_kind == '上拉':
            row_body['上拉电源'] = ', '.join(dict.fromkeys(entry.get('reference_net', '') for entry in group if entry.get('reference_net', '')))
        else:
            row_body['下拉参考网络'] = 'GND'

        row = _with_meta(row_body, 'candidate', severity, confidence, reason)
        row[USER_VISIBLE_REAL_PAGE_LABEL] = row['页面']
        rows.append(row)
    return rows


MAX_SERIES_WALK_HOPS = 8
MAX_SERIES_WALK_RESULTS = 512


def _walk_series_paths(start_net: str, series_by_net: Dict[str, list]) -> List[dict]:
    if not start_net:
        return []

    queue = deque([{'net': start_net, 'chain': [], 'net_chain': [start_net]}])
    seen_paths = set()
    results: List[dict] = []

    while queue:
        state = queue.popleft()
        current_net = state['net']
        if len(state['chain']) >= MAX_SERIES_WALK_HOPS:
            continue
        for edge in sorted(series_by_net.get(current_net, []), key=_series_edge_sort_key):
            next_net = edge.get('other_net', '')
            if not next_net or next_net in state['net_chain']:
                continue
            if any(prev_edge.get('refdes', '') == edge.get('refdes', '') for prev_edge in state['chain']):
                continue
            next_chain = state['chain'] + [edge]
            next_net_chain = state['net_chain'] + [next_net]
            path_key = (
                next_net,
                tuple(
                    (chain_edge.get('refdes', ''), chain_edge.get('other_net', ''))
                    for chain_edge in next_chain
                ),
            )
            if path_key in seen_paths:
                continue
            seen_paths.add(path_key)
            results.append({
                'source_net': start_net,
                'target_net': next_net,
                'chain': next_chain,
                'net_chain': next_net_chain,
                'hop_count': len(next_chain),
                'via_refdes_chain': _series_chain_field(next_chain, 'refdes'),
                'via_value_chain': _series_chain_field(next_chain, 'value'),
                'via_usage_chain': _series_chain_field(next_chain, 'usage'),
                'via_net_chain': ' -> '.join(next_net_chain),
                'via_pages': _series_chain_pages(next_chain, 'page'),
                'via_mapped_pages': _series_chain_pages(next_chain, 'mapped_page'),
                'via_total_ohms': _series_chain_total_ohms(next_chain),
            })
            if len(results) >= MAX_SERIES_WALK_RESULTS:
                return results
            queue.append({
                'net': next_net,
                'chain': next_chain,
                'net_chain': next_net_chain,
            })

    return results


def _build_indirect_bias_maps(pullups: Dict[str, list],
                              pulldowns: Dict[str, list],
                              series_by_net: Dict[str, list]) -> Tuple[Dict[str, list], Dict[str, list]]:
    indirect_pullups: Dict[str, list] = defaultdict(list)
    indirect_pulldowns: Dict[str, list] = defaultdict(list)

    for start_net in sorted(series_by_net.keys(), key=_natural_sort_key):
        seen_keys = {'pullup': set(), 'pulldown': set()}
        for path in _walk_series_paths(start_net, series_by_net):
            remote_net = path['target_net']
            for bias_kind, direct_map, indirect_map in (
                ('pullup', pullups, indirect_pullups),
                ('pulldown', pulldowns, indirect_pulldowns),
            ):
                for bias in sorted(direct_map.get(remote_net, []), key=lambda row: _natural_sort_key(row.get('refdes', ''))):
                    dedupe_key = (remote_net, bias.get('refdes', ''), path['via_refdes_chain'])
                    if dedupe_key in seen_keys[bias_kind]:
                        continue
                    seen_keys[bias_kind].add(dedupe_key)
                    indirect_map[start_net].append({
                        **bias,
                        'source_net': remote_net,
                        'other_net': start_net,
                        'via_refdes': path['via_refdes_chain'],
                        'via_value': path['via_value_chain'],
                        'via_ohms': path['via_total_ohms'],
                        'via_refdes_chain': path['via_refdes_chain'],
                        'via_value_chain': path['via_value_chain'],
                        'via_usage_chain': path['via_usage_chain'],
                        'via_net_chain': path['via_net_chain'],
                        'via_hop_count': path['hop_count'],
                        'via_pages': path['via_pages'],
                        'via_mapped_pages': path['via_mapped_pages'],
                    })

    return dict(indirect_pullups), dict(indirect_pulldowns)


def _classify_series_bias_ratio(series_ohms: Optional[float],
                                bias_ohms: Optional[float]) -> Tuple[Optional[float], str, Dict[str, str]]:
    if series_ohms is None or bias_ohms is None or bias_ohms <= 0:
        return None, '⚪ 阻值缺失，无法计算', _meta_fields('indeterminate', 'low', 'high', 'missing_resistance_value')

    ratio = series_ohms / bias_ohms
    if bias_ohms < 1000 and ratio > 0.1:
        return ratio, '❌ 高风险', _meta_fields('candidate', 'high', 'medium', 'series_bias_ratio_high')
    if ratio >= 0.33:
        return ratio, '❌ 高风险', _meta_fields('candidate', 'high', 'medium', 'series_bias_ratio_high')
    if ratio > 0.1:
        return ratio, '⚠️ 关注', _meta_fields('candidate', 'medium', 'medium', 'series_bias_ratio_warn')
    return ratio, '✅ 正常', _meta_fields('candidate', 'low', 'medium', 'series_bias_ratio_ok')


def _analyze_resistors_multi_series(components: Dict, nets: Dict, *, exclude_depop: bool = True) -> dict:
    pullups: Dict[str, list] = defaultdict(list)
    pulldowns: Dict[str, list] = defaultdict(list)
    series_by_net: Dict[str, list] = defaultdict(list)
    node_lookup: Dict[Tuple[str, str], str] = {}

    for net_name, nodes in nets.items():
        for node in nodes:
            node_lookup[(node['refdes'], node['pin'])] = node.get('pin_name', node['pin'])

    for refdes, comp in components.items():
        if comp.get('comp_type') != 'RES':
            continue
        pin_nets = list(dict.fromkeys(comp.get('nets', {}).values()))
        if len(pin_nets) != 2:
            continue

        net_a, net_b = pin_nets[0], pin_nets[1]
        ohms = _parse_ohms(comp.get('value', ''))
        value = comp.get('value', '')
        page = component_user_visible_page(comp)
        mapped_page = page
        bom_option = comp.get('bom_option', '')
        usage = _classify_resistor_usage(refdes, comp, ohms)
        if exclude_depop and _is_depop_option(bom_option):
            continue

        a_pwr, b_pwr = _net_is_power(net_a), _net_is_power(net_b)
        a_gnd, b_gnd = _net_is_gnd(net_a), _net_is_gnd(net_b)

        if a_pwr and not b_pwr and not b_gnd:
            pullups[net_b].append({
                'refdes': refdes,
                'ohms': ohms,
                'value': value,
                'power_net': net_a,
                'page': page,
                'mapped_page': mapped_page,
                'bom_option': bom_option,
                'usage': usage,
            })
        elif b_pwr and not a_pwr and not a_gnd:
            pullups[net_a].append({
                'refdes': refdes,
                'ohms': ohms,
                'value': value,
                'power_net': net_b,
                'page': page,
                'mapped_page': mapped_page,
                'bom_option': bom_option,
                'usage': usage,
            })
        elif a_gnd and not b_gnd and not b_pwr:
            pulldowns[net_b].append({
                'refdes': refdes,
                'ohms': ohms,
                'value': value,
                'page': page,
                'mapped_page': mapped_page,
                'bom_option': bom_option,
                'usage': usage,
            })
        elif b_gnd and not a_gnd and not a_pwr:
            pulldowns[net_a].append({
                'refdes': refdes,
                'ohms': ohms,
                'value': value,
                'page': page,
                'mapped_page': mapped_page,
                'bom_option': bom_option,
                'usage': usage,
            })
        elif not a_pwr and not b_pwr and not a_gnd and not b_gnd:
            edge_a = {
                'refdes': refdes,
                'ohms': ohms,
                'value': value,
                'page': page,
                'mapped_page': mapped_page,
                'bom_option': bom_option,
                'other_net': net_b,
                'usage': usage,
            }
            edge_b = dict(edge_a, other_net=net_a)
            series_by_net[net_a].append(edge_a)
            series_by_net[net_b].append(edge_b)

    indirect_pullups, indirect_pulldowns = _build_indirect_bias_maps(pullups, pulldowns, series_by_net)

    dup_pullups = _duplicate_bias_rows(pullups, indirect_pullups, bias_kind='上拉')
    dup_pulldowns = _duplicate_bias_rows(pulldowns, indirect_pulldowns, bias_kind='下拉')

    divider_risks = []
    for bias_kind, indirect_map in (('上拉', indirect_pullups), ('下拉', indirect_pulldowns)):
        for affected_net, entries in sorted(indirect_map.items(), key=lambda item: _natural_sort_key(item[0])):
            for bias in sorted(entries, key=_series_bias_sort_key):
                ratio, status, meta = _classify_series_bias_ratio(bias.get('via_ohms'), bias.get('ohms'))
                ref_net = bias.get('power_net', '') if bias_kind == '上拉' else 'GND'
                row = {
                    '串阻位号': bias.get('via_refdes_chain', ''),
                    '串阻值': bias.get('via_value_chain', ''),
                    '串阻网络A': affected_net,
                    '串阻网络B': bias.get('source_net', ''),
                    '串阻经过网络': bias.get('via_net_chain', ''),
                    '串阻跳数': bias.get('via_hop_count', 0),
                    '偏置类型': bias_kind,
                    '偏置位号': bias['refdes'],
                    '偏置值': bias['value'],
                    '偏置所在网络': bias.get('source_net', ''),
                    '偏置参考网络': ref_net,
                    '受影响网络': affected_net,
                    '串阻类型候选': bias.get('via_usage_chain', ''),
                    '偏置电阻类型候选': bias.get('usage', ''),
                    '串/偏置比': f'{ratio:.3f}' if ratio is not None else '',
                    '偏置 < 1k': '是' if (bias.get('ohms') or 0) < 1000 else '否',
                    '说明': (
                        f'{bias_kind}位于 {bias.get("source_net", "")} 侧，'
                        f'通过 {bias.get("via_refdes_chain", "")} 影响 {affected_net}'
                    ),
                    '状态': status,
                    **meta,
                    '页面': _merge_display_values(bias.get('via_pages', ''), bias.get('page', '')),
                }
                row[USER_VISIBLE_REAL_PAGE_LABEL] = _merge_display_values(
                    bias.get('via_mapped_pages', ''),
                    bias.get('mapped_page', ''),
                )
                row['页面'] = row[USER_VISIBLE_REAL_PAGE_LABEL] or row.get('页面', '')
                divider_risks.append(row)
    divider_risks.sort(key=lambda row: (
        0 if row['状态'].startswith('❌') else 1 if row['状态'].startswith('⚠') else 2,
        _natural_sort_key(row.get('串阻位号', '')),
        _natural_sort_key(row.get('偏置位号', '')),
    ))

    chip_pin_rows = []
    for refdes, comp in sorted(components.items(), key=lambda item: _natural_sort_key(item[0])):
        if not _is_chip_component(refdes, comp):
            continue
        for pin, net_name in sorted(comp.get('nets', {}).items(), key=lambda item: _natural_sort_key(item[0])):
            pin_name = node_lookup.get((refdes, pin), pin)
            submodule, submodule_path = _extract_pin_submodule_info(pin_name)
            series_entries = sorted(series_by_net.get(net_name, []), key=_series_edge_sort_key)
            pullup_entries = sorted(pullups.get(net_name, []), key=lambda row: _natural_sort_key(row.get('refdes', '')))
            pulldown_entries = sorted(pulldowns.get(net_name, []), key=lambda row: _natural_sort_key(row.get('refdes', '')))
            indirect_pullup_entries = sorted(indirect_pullups.get(net_name, []), key=_series_bias_sort_key)
            indirect_pulldown_entries = sorted(indirect_pulldowns.get(net_name, []), key=_series_bias_sort_key)
            row = {
                '芯片位号': refdes,
                '引脚': pin,
                '引脚名': pin_name,
                '后缀组': _extract_refdes_suffix_group(refdes),
                '子模块': submodule,
                '子模块路径': submodule_path,
                '网络名': net_name,
                '有串阻': '是' if series_entries else '否',
                '串阻数量': len(series_entries),
                '串阻位号': _format_entry_list(series_entries, 'refdes'),
                '串阻另一端网络': _format_entry_list(series_entries, 'other_net'),
                '串阻类型候选': _format_entry_list(series_entries, 'usage'),
                '串阻BOM_OPTION': ', '.join(dict.fromkeys(_display_bom_option(entry.get('bom_option', '')) for entry in series_entries)),
                '有上拉': '是' if pullup_entries else '否',
                '上拉数量': len(pullup_entries),
                '上拉位号': _format_entry_list(pullup_entries, 'refdes'),
                '上拉电源': _format_entry_list(pullup_entries, 'power_net'),
                '上拉电阻类型候选': _format_entry_list(pullup_entries, 'usage'),
                '上拉BOM_OPTION': ', '.join(dict.fromkeys(_display_bom_option(entry.get('bom_option', '')) for entry in pullup_entries)),
                '隔串阻上拉数量': len(indirect_pullup_entries),
                '隔串阻上拉位号': _format_entry_list(indirect_pullup_entries, 'refdes'),
                '隔串阻上拉来源网络': _format_entry_list(indirect_pullup_entries, 'source_net'),
                '隔串阻上拉电源': _format_entry_list(indirect_pullup_entries, 'power_net'),
                '隔串阻上拉串阻链': _format_entry_list(indirect_pullup_entries, 'via_refdes_chain'),
                '隔串阻上拉串阻类型候选': _format_entry_list(indirect_pullup_entries, 'via_usage_chain'),
                '有下拉': '是' if pulldown_entries else '否',
                '下拉数量': len(pulldown_entries),
                '下拉位号': _format_entry_list(pulldown_entries, 'refdes'),
                '下拉电阻类型候选': _format_entry_list(pulldown_entries, 'usage'),
                '下拉BOM_OPTION': ', '.join(dict.fromkeys(_display_bom_option(entry.get('bom_option', '')) for entry in pulldown_entries)),
                '隔串阻下拉数量': len(indirect_pulldown_entries),
                '隔串阻下拉位号': _format_entry_list(indirect_pulldown_entries, 'refdes'),
                '隔串阻下拉来源网络': _format_entry_list(indirect_pulldown_entries, 'source_net'),
                '隔串阻下拉串阻链': _format_entry_list(indirect_pulldown_entries, 'via_refdes_chain'),
                '隔串阻下拉串阻类型候选': _format_entry_list(indirect_pulldown_entries, 'via_usage_chain'),
                '页面': component_user_visible_page(comp),
                '主模块页映射一一对应': comp.get('page_mapping_ok', ''),
            }
            row.update(_component_page_fields(comp))
            chip_pin_rows.append(row)

    return {
        'dup_pullups': dup_pullups,
        'dup_pulldowns': dup_pulldowns,
        'divider_risks': divider_risks,
        'chip_pin_rows': chip_pin_rows,
        'pullups': dict(pullups),
        'pulldowns': dict(pulldowns),
        'indirect_pullups': dict(indirect_pullups),
        'indirect_pulldowns': dict(indirect_pulldowns),
        'series_by_net': dict(series_by_net),
    }


def analyze_resistors(components: Dict, nets: Dict, *, exclude_depop: bool = True) -> dict:
    """检测上拉/下拉/串阻相关设计问题"""
    return _analyze_resistors_multi_series(components, nets, exclude_depop=exclude_depop)
