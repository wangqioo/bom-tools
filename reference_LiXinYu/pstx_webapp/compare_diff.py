# -*- coding: utf-8 -*-
"""Pure diff builders for project compare payloads."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from pstx_core.page_resolution import USER_VISIBLE_REAL_PAGE_LABEL, component_user_visible_page
from pstx_webapp.compare_view import refdes_category, refdes_category_label
from pstx_webapp.json_utils import compact_value, json_fingerprint
from pstx_webapp.report_feishu import feishu_link_value

MAX_COMPARE_DETAIL_ROWS = 200


def component_compare_value(comp: dict) -> dict:
    return {
        '类型': comp.get('comp_type', ''),
        '料号': comp.get('hq_code', ''),
        '值': comp.get('value', ''),
        '封装': comp.get('package', ''),
        'BOM_OPTION': comp.get('bom_option', ''),
        USER_VISIBLE_REAL_PAGE_LABEL: component_user_visible_page(comp),
        '网络': comp.get('nets', {}),
    }


def component_compare_value_with_feishu(comp: dict, link: Optional[dict]) -> dict:
    value = component_compare_value(comp)
    feishu = feishu_link_value(link)
    value.update({
        '飞书HQ料号': feishu['飞书HQ料号'],
        '飞书规格型号': feishu['飞书规格型号'],
        'PI': feishu['PI'],
        '选型顺序': feishu['选型顺序'],
        '飞书校对结论': feishu['飞书校对结论'],
    })
    return value


def net_compare_value(nodes: List[dict]) -> List[str]:
    return sorted(
        (
            f"{node.get('refdes', '')}:{node.get('pin', '')}:{node.get('pin_name', '')}"
            for node in nodes or []
        ),
        key=str.upper,
    )


def _net_name(value: object) -> str:
    return str(value or '').strip()


def _net_node_count(nets: dict, net_name: str) -> int:
    name = _net_name(net_name)
    if not name:
        return 0
    return len((nets or {}).get(name, []) or [])


def _pin_sample(row: dict) -> str:
    refdes = str(row.get('位号') or '').strip()
    pin = str(row.get('引脚') or '').strip()
    left_pin_name = str(row.get('左侧引脚名') or '').strip()
    right_pin_name = str(row.get('右侧引脚名') or '').strip()
    pin_name = right_pin_name or left_pin_name
    if refdes and pin and pin_name:
        return f'{refdes}:{pin}({pin_name})'
    if refdes and pin:
        return f'{refdes}:{pin}'
    return refdes or pin or 'pin'


def _net_transition_type(left_net: str, right_net: str, change_types: set) -> str:
    if left_net and right_net and left_net != right_net:
        return '网络迁移'
    if not left_net and right_net:
        return '新增连接'
    if left_net and not right_net:
        return '删除连接'
    if '引脚名变化' in ' / '.join(sorted(change_types)):
        return '同网引脚名变化'
    return '网络节点变化'


def _net_transition_sort_key(row: dict) -> Tuple[int, int, str, str]:
    type_rank = {
        '网络迁移': 0,
        '网络节点变化': 1,
        '新增连接': 2,
        '删除连接': 3,
        '新增网络': 4,
        '删除网络': 5,
        '同网引脚名变化': 6,
    }
    return (
        type_rank.get(str(row.get('类型') or ''), 9),
        -int(row.get('影响引脚数') or 0),
        str(row.get('左侧网络') or '').upper(),
        str(row.get('右侧网络') or '').upper(),
    )


def build_net_view_diff(left_nets: dict,
                        right_nets: dict,
                        key_pin_net_diff: dict,
                        passive_pin_net_diff: dict,
                        net_diff: dict,
                        *,
                        detail_limit: int = MAX_COMPARE_DETAIL_ROWS) -> dict:
    """Build a net-centric summary from existing compare evidence."""
    entries: Dict[Tuple[str, str], dict] = {}

    def ensure_entry(left_net: str, right_net: str) -> dict:
        key = (_net_name(left_net), _net_name(right_net))
        entry = entries.get(key)
        if entry is None:
            entry = {
                'left_net': key[0],
                'right_net': key[1],
                'key_refdes': set(),
                'passive_refdes': set(),
                'samples': [],
                'change_types': set(),
                'pin_count': 0,
                'network_notes': [],
                'sources': set(),
            }
            entries[key] = entry
        return entry

    def add_pin_rows(diff: dict, source: str) -> None:
        for row in diff.get('_all_rows', diff.get('rows', [])) or []:
            left_net = _net_name(row.get('左侧网络'))
            right_net = _net_name(row.get('右侧网络'))
            if not left_net and not right_net:
                continue
            entry = ensure_entry(left_net, right_net)
            refdes = str(row.get('位号') or '').strip()
            if source == 'key':
                entry['key_refdes'].add(refdes)
            else:
                entry['passive_refdes'].add(refdes)
            entry['pin_count'] += 1
            entry['change_types'].add(str(row.get('类型') or '连接变化').strip() or '连接变化')
            entry['sources'].add(source)
            sample = _pin_sample(row)
            if sample and sample not in entry['samples'] and len(entry['samples']) < 8:
                entry['samples'].append(sample)

    add_pin_rows(key_pin_net_diff or {}, 'key')
    add_pin_rows(passive_pin_net_diff or {}, 'passive')

    net_key_label = str((net_diff or {}).get('key_label') or '网络名')
    net_changes_by_name: Dict[str, List[str]] = {}
    for row in (net_diff or {}).get('_all_rows', (net_diff or {}).get('rows', [])) or []:
        net_name = _net_name(row.get(net_key_label) or row.get('网络名'))
        if not net_name:
            continue
        net_changes_by_name.setdefault(net_name, []).append(str(row.get('类型') or '变化'))

    covered_net_names = set()
    for entry in entries.values():
        left_net = entry['left_net']
        right_net = entry['right_net']
        notes = []
        if left_net and left_net in net_changes_by_name:
            notes.append(f"左侧 {left_net}: {'/'.join(net_changes_by_name[left_net])}")
            covered_net_names.add(left_net)
        if right_net and right_net != left_net and right_net in net_changes_by_name:
            notes.append(f"右侧 {right_net}: {'/'.join(net_changes_by_name[right_net])}")
            covered_net_names.add(right_net)
        entry['network_notes'].extend(notes)

    for row in (net_diff or {}).get('_all_rows', (net_diff or {}).get('rows', [])) or []:
        net_name = _net_name(row.get(net_key_label) or row.get('网络名'))
        if not net_name or net_name in covered_net_names:
            continue
        change_type = str(row.get('类型') or '变化')
        if change_type == '新增':
            entry = ensure_entry('', net_name)
            entry['network_notes'].append('右侧新增网络节点列表')
            entry['change_types'].add('新增网络')
        elif change_type == '删除':
            entry = ensure_entry(net_name, '')
            entry['network_notes'].append('左侧删除网络节点列表')
            entry['change_types'].add('删除网络')
        else:
            entry = ensure_entry(net_name, net_name)
            entry['network_notes'].append('同名网络节点列表变化')
            entry['change_types'].add('网络节点变化')
        entry['sources'].add('net_diff')

    rows = []
    added_count = 0
    removed_count = 0
    changed_count = 0
    for entry in entries.values():
        left_net = entry['left_net']
        right_net = entry['right_net']
        row_type = _net_transition_type(left_net, right_net, entry['change_types'])
        if '新增网络' in entry['change_types'] and not left_net:
            row_type = '新增网络'
        elif '删除网络' in entry['change_types'] and not right_net:
            row_type = '删除网络'
        if row_type in {'新增连接', '新增网络'}:
            added_count += 1
        elif row_type in {'删除连接', '删除网络'}:
            removed_count += 1
        else:
            changed_count += 1
        key_refdes = sorted(filter(None, entry['key_refdes']), key=str.upper)
        passive_refdes = sorted(filter(None, entry['passive_refdes']), key=str.upper)
        affected_refdes = sorted(set(key_refdes) | set(passive_refdes), key=str.upper)
        notes = sorted(set(filter(None, entry['network_notes'])))
        rows.append({
            '类型': row_type,
            '左侧网络': left_net,
            '右侧网络': right_net,
            '网络迁移': f"{left_net or '未连接'} -> {right_net or '未连接'}",
            '影响位号数': len(affected_refdes),
            '影响引脚数': entry['pin_count'],
            '关键器件数': len(key_refdes),
            'R/C/L数': len(passive_refdes),
            '关键器件样例': ', '.join(key_refdes[:6]),
            'R/C/L样例': ', '.join(passive_refdes[:6]),
            '样例引脚': ', '.join(entry['samples'][:6]),
            '左侧节点数': _net_node_count(left_nets, left_net),
            '右侧节点数': _net_node_count(right_nets, right_net),
            '网络节点变化': '; '.join(notes) or '来自 Pin/Net 连接差异',
            '变化说明': ', '.join(sorted(entry['change_types'])) or row_type,
        })
    rows.sort(key=_net_transition_sort_key)
    truncated = (
        len(rows) > detail_limit
        or bool((key_pin_net_diff or {}).get('truncated'))
        or bool((passive_pin_net_diff or {}).get('truncated'))
        or bool((net_diff or {}).get('truncated'))
    )
    return {
        'title': 'Net 视角变化',
        'key_label': '网络迁移',
        'added_count': added_count,
        'removed_count': removed_count,
        'changed_count': changed_count,
        'rows': rows[:detail_limit],
        'total_rows': len(rows),
        'truncated': truncated,
        'summary': {
            'transition_count': sum(1 for row in rows if row.get('类型') == '网络迁移'),
            'key_pin_count': sum(int(row.get('关键器件数') or 0) for row in rows),
            'passive_pin_count': sum(int(row.get('R/C/L数') or 0) for row in rows),
            'net_added_count': (net_diff or {}).get('added_count', 0),
            'net_removed_count': (net_diff or {}).get('removed_count', 0),
            'net_changed_count': (net_diff or {}).get('changed_count', 0),
        },
    }


def component_inventory_summary(refdes: str, comp: dict, link: Optional[dict] = None) -> str:
    summary = {
        '类别': refdes_category_label(refdes_category(refdes)),
        '类型': comp.get('comp_type', ''),
        '料号': comp.get('hq_code', ''),
        '值': comp.get('value', ''),
        '封装': comp.get('package', ''),
        USER_VISIBLE_REAL_PAGE_LABEL: component_user_visible_page(comp),
    }
    feishu = feishu_link_value(link)
    if any(feishu.values()):
        summary.update({
            '飞书规格型号': feishu['飞书规格型号'],
            'PI': feishu['PI'],
            '选型顺序': feishu['选型顺序'],
            '飞书校对结论': feishu['飞书校对结论'],
        })
    return compact_value(summary, 260)


def compare_component_inventory(left_components: dict,
                                right_components: dict,
                                *,
                                title: str,
                                predicate,
                                detail_limit: int = MAX_COMPARE_DETAIL_ROWS,
                                left_feishu: Optional[Dict[str, dict]] = None,
                                right_feishu: Optional[Dict[str, dict]] = None) -> dict:
    left_feishu = left_feishu or {}
    right_feishu = right_feishu or {}
    left_keys = {refdes for refdes in left_components if predicate(refdes)}
    right_keys = {refdes for refdes in right_components if predicate(refdes)}
    added = sorted(right_keys - left_keys, key=str.upper)
    removed = sorted(left_keys - right_keys, key=str.upper)
    rows = []
    for refdes in added:
        category = refdes_category(refdes)
        link = right_feishu.get(refdes.upper())
        feishu = feishu_link_value(link)
        rows.append({
            '类型': '新增',
            '位号': refdes,
            '器件类别': refdes_category_label(category),
            '左侧': '',
            '右侧': component_inventory_summary(refdes, right_components.get(refdes, {}), link),
            '右侧飞书校对': feishu['飞书校对结论'],
            '右侧飞书规格型号': feishu['飞书规格型号'],
            '右侧PI': feishu['PI'],
            '右侧选型顺序': feishu['选型顺序'],
            '变化字段': '新增关键器件',
        })
    for refdes in removed:
        category = refdes_category(refdes)
        link = left_feishu.get(refdes.upper())
        feishu = feishu_link_value(link)
        rows.append({
            '类型': '删除',
            '位号': refdes,
            '器件类别': refdes_category_label(category),
            '左侧': component_inventory_summary(refdes, left_components.get(refdes, {}), link),
            '右侧': '',
            '左侧飞书校对': feishu['飞书校对结论'],
            '左侧飞书规格型号': feishu['飞书规格型号'],
            '左侧PI': feishu['PI'],
            '左侧选型顺序': feishu['选型顺序'],
            '变化字段': '删除关键器件',
        })
    return {
        'title': title,
        'key_label': '位号',
        'added_count': len(added),
        'removed_count': len(removed),
        'changed_count': 0,
        'rows': rows[:detail_limit],
        'total_rows': len(rows),
        'truncated': len(rows) > detail_limit,
    }


def pin_sort_key(pin: str) -> Tuple[int, object, str]:
    text = str(pin or '').strip()
    if text.isdigit():
        return (0, int(text), text)
    return (1, text.upper(), text)


def pin_name_lookup(nets: dict, refdes: str) -> Dict[str, str]:
    lookup: Dict[str, str] = {}
    target = str(refdes or '').upper()
    for nodes in (nets or {}).values():
        for node in nodes or []:
            if str(node.get('refdes', '')).upper() != target:
                continue
            pin = str(node.get('pin', '')).strip()
            pin_name = str(node.get('pin_name', '')).strip()
            if pin and pin_name and pin not in lookup:
                lookup[pin] = pin_name
    return lookup


def component_pin_connections(refdes: str, comp: dict, nets: dict) -> Dict[str, dict]:
    pin_names = pin_name_lookup(nets, refdes)
    connections: Dict[str, dict] = {}
    for raw_pin, net_name in (comp.get('nets', {}) or {}).items():
        pin = str(raw_pin or '').strip()
        if not pin:
            continue
        connections[pin] = {
            'pin': pin,
            'pin_name': pin_names.get(pin, ''),
            'net': str(net_name or '').strip(),
        }
    for pin, pin_name in pin_names.items():
        connections.setdefault(pin, {'pin': pin, 'pin_name': pin_name, 'net': ''})
    return connections


def pin_connection_change_type(left_pin: Optional[dict], right_pin: Optional[dict]) -> str:
    if left_pin is None:
        return '新增连接'
    if right_pin is None:
        return '删除连接'
    changed = []
    if left_pin.get('net', '') != right_pin.get('net', ''):
        changed.append('网络变化')
    if left_pin.get('pin_name', '') != right_pin.get('pin_name', ''):
        changed.append('引脚名变化')
    return ' / '.join(changed) or '连接变化'


def compare_component_pin_nets(left_components: dict,
                               right_components: dict,
                               left_nets: dict,
                               right_nets: dict,
                               *,
                               title: str,
                               predicate,
                               detail_limit: int = MAX_COMPARE_DETAIL_ROWS,
                               include_all_rows: bool = False,
                               left_feishu: Optional[Dict[str, dict]] = None,
                               right_feishu: Optional[Dict[str, dict]] = None) -> dict:
    left_feishu = left_feishu or {}
    right_feishu = right_feishu or {}
    shared = sorted(
        {refdes for refdes in left_components if predicate(refdes)}
        & {refdes for refdes in right_components if predicate(refdes)},
        key=str.upper,
    )
    rows = []
    added_count = 0
    removed_count = 0
    changed_count = 0
    for refdes in shared:
        category = refdes_category(refdes)
        left_feishu_value = feishu_link_value(left_feishu.get(refdes.upper()))
        right_feishu_value = feishu_link_value(right_feishu.get(refdes.upper()))
        left_pins = component_pin_connections(refdes, left_components.get(refdes, {}), left_nets)
        right_pins = component_pin_connections(refdes, right_components.get(refdes, {}), right_nets)
        for pin in sorted(set(left_pins) | set(right_pins), key=pin_sort_key):
            left_pin = left_pins.get(pin)
            right_pin = right_pins.get(pin)
            if left_pin == right_pin:
                continue
            change_type = pin_connection_change_type(left_pin, right_pin)
            if left_pin is None:
                added_count += 1
            elif right_pin is None:
                removed_count += 1
            else:
                changed_count += 1
            rows.append({
                '类型': change_type,
                '位号': refdes,
                '器件类别': refdes_category_label(category),
                '引脚': pin,
                '左侧引脚名': (left_pin or {}).get('pin_name', ''),
                '右侧引脚名': (right_pin or {}).get('pin_name', ''),
                '左侧网络': (left_pin or {}).get('net', ''),
                '右侧网络': (right_pin or {}).get('net', ''),
                f'左侧{USER_VISIBLE_REAL_PAGE_LABEL}': component_user_visible_page(left_components.get(refdes, {})),
                f'右侧{USER_VISIBLE_REAL_PAGE_LABEL}': component_user_visible_page(right_components.get(refdes, {})),
                '左侧飞书校对': left_feishu_value['飞书校对结论'],
                '右侧飞书校对': right_feishu_value['飞书校对结论'],
                '左侧飞书规格型号': left_feishu_value['飞书规格型号'],
                '右侧飞书规格型号': right_feishu_value['飞书规格型号'],
                '左侧PI': left_feishu_value['PI'],
                '右侧PI': right_feishu_value['PI'],
                '左侧选型顺序': left_feishu_value['选型顺序'],
                '右侧选型顺序': right_feishu_value['选型顺序'],
            })
    result = {
        'title': title,
        'key_label': '位号',
        'added_count': added_count,
        'removed_count': removed_count,
        'changed_count': changed_count,
        'rows': rows[:detail_limit],
        'total_rows': len(rows),
        'truncated': len(rows) > detail_limit,
    }
    if include_all_rows:
        result['_all_rows'] = rows
    return result


def diff_named_maps(left_map: dict,
                    right_map: dict,
                    *,
                    title: str,
                    key_label: str,
                    value_builder=None,
                    detail_limit: int = MAX_COMPARE_DETAIL_ROWS,
                    include_all_rows: bool = False) -> dict:
    value_builder = value_builder or (lambda value: value)
    left_keys = set(left_map)
    right_keys = set(right_map)
    added = sorted(right_keys - left_keys, key=str.upper)
    removed = sorted(left_keys - right_keys, key=str.upper)
    shared = sorted(left_keys & right_keys, key=str.upper)
    rows = []
    for key in added:
        rows.append({'类型': '新增', key_label: key, '左侧': '', '右侧': compact_value(value_builder(right_map[key])), '变化字段': '新增'})
    for key in removed:
        rows.append({'类型': '删除', key_label: key, '左侧': compact_value(value_builder(left_map[key])), '右侧': '', '变化字段': '删除'})
    changed_count = 0
    for key in shared:
        left_value = value_builder(left_map[key])
        right_value = value_builder(right_map[key])
        if left_value == right_value:
            continue
        changed_count += 1
        changed_fields = []
        if isinstance(left_value, dict) and isinstance(right_value, dict):
            for field in sorted(set(left_value) | set(right_value)):
                if left_value.get(field) != right_value.get(field):
                    changed_fields.append(field)
        rows.append({
            '类型': '变化',
            key_label: key,
            '左侧': compact_value(left_value),
            '右侧': compact_value(right_value),
            '变化字段': ', '.join(changed_fields) or '内容变化',
        })
    result = {
        'title': title,
        'key_label': key_label,
        'added_count': len(added),
        'removed_count': len(removed),
        'changed_count': changed_count,
        'rows': rows[:detail_limit],
        'total_rows': len(rows),
        'truncated': len(rows) > detail_limit,
    }
    if include_all_rows:
        result['_all_rows'] = rows
    return result


def row_compare_key(row: dict) -> str:
    priority = [
        '位号', '网络名', '芯片位号', '引脚', '基础名', 'P端网络', 'N端网络',
        '使用该值的位号', '料号', '值', '封装', '主模块页', USER_VISIBLE_REAL_PAGE_LABEL,
        f'左侧{USER_VISIBLE_REAL_PAGE_LABEL}', f'右侧{USER_VISIBLE_REAL_PAGE_LABEL}', '页面', '原因代码', '状态',
    ]
    fields = [field for field in priority if field in row]
    if fields:
        return ' | '.join(f'{field}={row.get(field, "")}' for field in fields[:4])
    return json_fingerprint(row)


def table_rows_by_key(table: dict) -> Dict[str, dict]:
    rows_by_key: Dict[str, dict] = {}
    for index, row in enumerate(table.get('rows', []) or []):
        key = row_compare_key(row)
        if key in rows_by_key:
            key = f'{key} #{index + 1}'
        rows_by_key[key] = row
    return rows_by_key


def flatten_report_tables(report: dict) -> Dict[str, dict]:
    tables: Dict[str, dict] = {}
    for section in report.get('sections', []) or []:
        for table in section.get('tables', []) or []:
            tables[table.get('id') or table.get('title')] = table
    return tables


def compare_report_tables(left_report: dict, right_report: dict, *, detail_limit: int = MAX_COMPARE_DETAIL_ROWS) -> List[dict]:
    left_tables = flatten_report_tables(left_report)
    right_tables = flatten_report_tables(right_report)
    results = []
    for table_id in sorted(set(left_tables) | set(right_tables)):
        left_table = left_tables.get(table_id, {'title': table_id, 'rows': []})
        right_table = right_tables.get(table_id, {'title': table_id, 'rows': []})
        diff = diff_named_maps(
            table_rows_by_key(left_table),
            table_rows_by_key(right_table),
            title=right_table.get('title') or left_table.get('title') or table_id,
            key_label='对象',
            detail_limit=detail_limit,
        )
        diff['id'] = table_id
        if diff['added_count'] or diff['removed_count'] or diff['changed_count']:
            results.append(diff)
    return results
