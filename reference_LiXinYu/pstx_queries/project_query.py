# -*- coding: utf-8 -*-
"""Structured component/network query view models for analyzed projects."""

from typing import Dict, List

from pstx_core.page_resolution import USER_VISIBLE_REAL_PAGE_LABEL, component_user_visible_page
from pstx_rules.common import component_type_label, display_bom_option, natural_sort_key

def query_project_data(components: Dict,
                       nets: Dict,
                       mode: str,
                       keyword: str) -> dict:
    kw = (keyword or '').strip()
    if not kw:
        return {
            'title': '空查询',
            'lines': ['请输入位号或网络名。'],
            'mode': mode,
            'view': 'empty',
            'entity_type': '',
            'match_type': 'empty',
            'summary': {},
            'cards': [],
            'items': [],
        }

    lines: List[str] = []
    if mode == '位号':
        comp = components.get(kw)
        if comp is None:
            comp = next((value for refdes, value in components.items() if refdes.upper() == kw.upper()), None)
        if comp:
            value_text = comp.get('value', '') or comp.get('part_name', '')
            display_page = component_user_visible_page(comp)
            prop_items = [
                {'label': '位号', 'value': str(comp.get('refdes', kw))},
                {'label': '类型', 'value': str(component_type_label(comp.get('comp_type', '')))},
                {'label': '料号', 'value': str(comp.get('hq_code', ''))},
                {'label': '值', 'value': str(comp.get('value', ''))},
                {'label': '封装', 'value': str(comp.get('package', ''))},
                {'label': 'BOM_OPTION', 'value': display_bom_option(comp.get('bom_option', ''))},
                {'label': USER_VISIBLE_REAL_PAGE_LABEL, 'value': display_page},
                {'label': '主模块页映射一一对应', 'value': str(comp.get('page_mapping_ok', ''))},
                {'label': '主模块页来源', 'value': str(comp.get('page_source', ''))},
                {'label': 'ROOM', 'value': str(comp.get('room', ''))},
                {'label': 'DRAWING', 'value': str(comp.get('drawing', ''))},
            ]
            prop_items = [item for item in prop_items if item['value']]
            lines.append(f'◆ 元件：{comp.get("refdes", kw)}')
            for item in prop_items:
                lines.append(f'  {item["label"]:<16} {item["value"]}')
            lines += ['', '  引脚 -> 网络：']
            pin_rows = []
            for pin, net_name in sorted(comp.get('nets', {}).items(), key=lambda item: natural_sort_key(item[0])):
                lines.append(f'    pin {pin:<6} -> {net_name}')
                pin_rows.append({'pin': pin, 'net': net_name})
            return {
                'title': comp.get('refdes', kw),
                'lines': lines,
                'mode': mode,
                'view': 'component',
                'entity_type': 'component',
                'match_type': 'exact',
                'summary': {
                    'title': comp.get('refdes', kw),
                    'subtitle': value_text,
                    'meta': [
                        {'label': '封装', 'value': str(comp.get('package', ''))},
                        {'label': USER_VISIBLE_REAL_PAGE_LABEL, 'value': display_page},
                        {'label': '主模块页映射一一对应', 'value': str(comp.get('page_mapping_ok', ''))},
                        {'label': '主模块页来源', 'value': str(comp.get('page_source', ''))},
                        {'label': '料号', 'value': str(comp.get('hq_code', ''))},
                    ],
                },
                'cards': [
                    {'title': '元件属性', 'kind': 'properties', 'items': prop_items},
                    {'title': '引脚连接', 'kind': 'pins', 'items': pin_rows},
                ],
                'items': [],
            }

        matched = sorted(refdes for refdes in components if kw.upper() in refdes.upper())
        lines.append('未找到精确匹配，模糊结果：' if matched else f'未找到位号：{kw}')
        lines.extend(f'  {refdes}' for refdes in matched[:50])
        items = []
        for refdes in matched[:50]:
            comp = components.get(refdes, {})
            items.append({
                'title': refdes,
                'subtitle': str(comp.get('value', '') or comp.get('part_name', '')),
                'meta': [
                    {'label': '封装', 'value': str(comp.get('package', ''))},
                    {'label': USER_VISIBLE_REAL_PAGE_LABEL, 'value': component_user_visible_page(comp)},
                    {'label': '主模块页映射一一对应', 'value': str(comp.get('page_mapping_ok', ''))},
                ],
                'keyword': refdes,
            })
        return {
            'title': kw,
            'lines': lines,
            'mode': mode,
            'view': 'list',
            'entity_type': 'component',
            'match_type': 'fuzzy' if matched else 'missing',
            'summary': {
                'title': kw,
                'subtitle': '模糊匹配结果' if matched else '未找到位号',
                'meta': [{'label': '结果数', 'value': str(len(items))}],
            },
            'cards': [],
            'items': items,
        }

    nodes = nets.get(kw)
    exact_name = kw
    if nodes is None:
        exact_name = next((name for name in nets if name.upper() == kw.upper()), kw)
        nodes = nets.get(exact_name)
    if nodes:
        lines.append(f'◆ 网络：{exact_name}（{len(nodes)} 个节点）')
        node_rows = []
        for node in nodes:
            comp = components.get(node['refdes'], {})
            desc = comp.get('value', '') or comp.get('part_name', '')
            lines.append(f'  {node["refdes"]:<10} pin {node["pin"]:<6} ({node["pin_name"]:<12}) {desc}')
            node_rows.append({
                'refdes': node['refdes'],
                'pin': node['pin'],
                'pin_name': node['pin_name'],
                'desc': desc,
                USER_VISIBLE_REAL_PAGE_LABEL: component_user_visible_page(comp),
                '主模块页映射一一对应': str(comp.get('page_mapping_ok', '')),
            })
        return {
            'title': exact_name,
            'lines': lines,
            'mode': mode,
            'view': 'network',
            'entity_type': 'network',
            'match_type': 'exact',
            'summary': {
                'title': exact_name,
                'subtitle': f'{len(nodes)} 个连接节点',
                'meta': [
                    {'label': '节点数', 'value': str(len(nodes))},
                    {'label': '页码覆盖', 'value': str(len({row[USER_VISIBLE_REAL_PAGE_LABEL] for row in node_rows if row[USER_VISIBLE_REAL_PAGE_LABEL]}))},
                ],
            },
            'cards': [
                {'title': '网络节点', 'kind': 'nodes', 'items': node_rows},
            ],
            'items': [],
        }

    matched = sorted(name for name in nets if kw.upper() in name.upper())
    lines.append('未找到精确匹配，模糊结果：' if matched else f'未找到网络：{kw}')
    lines.extend(f'  {name}  ({len(nets[name])} nodes)' for name in matched[:50])
    items = [{
        'title': name,
        'subtitle': f'{len(nets[name])} 个节点',
        'meta': [],
        'keyword': name,
    } for name in matched[:50]]
    return {
        'title': kw,
        'lines': lines,
        'mode': mode,
        'view': 'list',
        'entity_type': 'network',
        'match_type': 'fuzzy' if matched else 'missing',
        'summary': {
            'title': kw,
            'subtitle': '模糊匹配结果' if matched else '未找到网络',
            'meta': [{'label': '结果数', 'value': str(len(items))}],
        },
        'cards': [],
        'items': items,
    }
