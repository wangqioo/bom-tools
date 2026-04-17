# -*- coding: utf-8 -*-
"""
PST 文件解析器
支持 pstxprt.dat / pstxnet.dat
"""

import re
from typing import Dict, List, Tuple


# ─────────────────────────────────────────────
# 工具函数
# ─────────────────────────────────────────────

def _join_continuations(text: str) -> str:
    """合并 PST 行续行符 ~ """
    lines = text.split('\n')
    result, buf = [], ''
    for line in lines:
        stripped = line.rstrip()
        if stripped.endswith('~'):
            buf += stripped[:-1]
        else:
            buf += line
            result.append(buf)
            buf = ''
    if buf:
        result.append(buf)
    return '\n'.join(result)


def _extract_attrs(text: str) -> Dict[str, str]:
    """从文本中提取所有 KEY='VALUE' 键值对（取首次出现）"""
    attrs = {}
    for m in re.finditer(r"\b([A-Z][A-Z0-9_]*)\s*=\s*'([^']*)'", text):
        key, val = m.group(1), m.group(2)
        if key not in attrs:
            attrs[key] = val
    return attrs


def _extract_page(drawing: str) -> str:
    """从 DRAWING 路径提取页面号，如 PAGE23"""
    m = re.search(r'(PAGE\d+)', drawing, re.IGNORECASE)
    return m.group(1).upper() if m else ''


def _get_comp_type(refdes: str, part_name: str) -> str:
    """根据 refdes 前缀或料号推断元件类型"""
    pn = part_name.lower()

    type_rules = [
        (['cap_pol'], 'CAP_POL'),
        (['cap_hdl', 'cap_'], 'CAP'),
        (['res_hdl', 'res_'], 'RES'),
        (['ind_hdl', 'ind_', 'ferrite', 'fer_hdl', 'fb_hdl'], 'IND'),
        (['osc_', 'crystal', 'xtal'], 'XTAL'),
        (['conn_', 'connector'], 'CONN'),
        (['led_'], 'LED'),
        (['diode', '_d_hdl'], 'DIODE'),
        (['mosfet', 'mos_', 'nmos', 'pmos', 'nfet', 'pfet'], 'FET'),
        (['bjt', 'transistor', 'npn', 'pnp'], 'BJT'),
        (['fuse'], 'FUSE'),
        (['sw_hdl', 'switch'], 'SWITCH'),
        (['testpoint', 'test_point', 'tp_hdl'], 'TESTPOINT'),
        (['transformer', 'xfmr'], 'TRANSFORMER'),
    ]
    for keywords, ctype in type_rules:
        if any(k in pn for k in keywords):
            return ctype

    prefix = (re.match(r'[A-Za-z]+', refdes) or re.match(r'', '')).group(0).upper()
    prefix_map = {
        'C': 'CAP', 'PC': 'CAP',
        'R': 'RES',
        'L': 'IND', 'FB': 'IND',
        'U': 'IC',
        'J': 'CONN', 'P': 'CONN', 'CN': 'CONN',
        'Q': 'FET',
        'D': 'DIODE',
        'LED': 'LED',
        'Y': 'XTAL',
        'F': 'FUSE',
        'SW': 'SWITCH',
        'TP': 'TESTPOINT',
        'T': 'TRANSFORMER',
    }
    return prefix_map.get(prefix, 'IC')


# ─────────────────────────────────────────────
# pstxprt.dat 解析
# ─────────────────────────────────────────────

def parse_pstxprt(content: str) -> Dict[str, dict]:
    """解析 pstxprt.dat，返回 {refdes: component_dict}"""
    text = _join_continuations(content)
    components = {}

    blocks = re.split(r'\nPART_NAME\n', text)

    for block in blocks[1:]:
        lines = block.split('\n')
        first = lines[0].strip()

        m = re.match(r"(\S+)\s+'([^']+)'", first)
        if not m:
            continue

        refdes = m.group(1)
        part_name = m.group(2)
        attrs = _extract_attrs(block)

        components[refdes] = {
            'refdes':     refdes,
            'part_name':  part_name,
            'hq_code':    attrs.get('HQ_CODE', ''),
            'value':      attrs.get('VALUE', ''),
            'package':    attrs.get('PACKAGE', ''),
            'material':   attrs.get('MATERIAL', ''),
            'tolerance':  attrs.get('TOLERANCE', ''),
            'voltage':    attrs.get('VOLTAGE', ''),
            'current':    attrs.get('CURRENT', ''),
            'power':      attrs.get('POWER', ''),
            'bom_option': attrs.get('BOM_OPTION', ''),
            'bom_cost':   attrs.get('BOM_COST', ''),
            'room':       attrs.get('ROOM', ''),
            'drawing':    attrs.get('DRAWING', ''),
            'page':       _extract_page(attrs.get('DRAWING', '')),
            'phys_page':  attrs.get('PHYS_PAGE', ''),
            'comp_type':  _get_comp_type(refdes, part_name),
        }

    return components


# ─────────────────────────────────────────────
# pstxnet.dat 解析
# ─────────────────────────────────────────────

def parse_pstxnet(content: str) -> Dict[str, List[dict]]:
    """解析 pstxnet.dat，返回 {net_name: [{refdes, pin, pin_name}, ...]}"""
    text = _join_continuations(content)
    nets = {}

    blocks = re.split(r'\nNET_NAME\n', text)

    for block in blocks[1:]:
        m = re.search(r"'([^']+)'", block)
        if not m:
            continue
        net_name = m.group(1)

        nodes = []
        node_re = re.compile(r'NODE_NAME\s+(\S+)\s+(\S+)')
        pin_name_re = re.compile(r"'([^']+)'\s*:")

        for node_m in node_re.finditer(block):
            refdes = node_m.group(1)
            pin_num = node_m.group(2)
            after = block[node_m.end(): node_m.end() + 200]
            pin_nm = pin_name_re.search(after)
            pin_name = pin_nm.group(1) if pin_nm else pin_num
            nodes.append({'refdes': refdes, 'pin': pin_num, 'pin_name': pin_name})

        if nodes:
            nets[net_name] = nodes

    return nets


# ─────────────────────────────────────────────
# 统一入口
# ─────────────────────────────────────────────

def parse_all(prt_content: str, net_content: str) -> Tuple[Dict, Dict, Dict]:
    """
    解析两个文件，返回 (components, nets, comp_nets)
    comp_nets: {refdes: {pin: net_name}}
    """
    components = parse_pstxprt(prt_content)
    nets = parse_pstxnet(net_content)

    comp_nets: Dict[str, Dict[str, str]] = {}
    for net_name, nodes in nets.items():
        for node in nodes:
            rd = node['refdes']
            if rd not in comp_nets:
                comp_nets[rd] = {}
            comp_nets[rd][node['pin']] = net_name

    for refdes, comp in components.items():
        comp['nets'] = comp_nets.get(refdes, {})

    return components, nets, comp_nets
