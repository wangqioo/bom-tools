# -*- coding: utf-8 -*-
"""原理图分析工具 (PSTX) — Blueprint"""

import os, uuid, re, json
from pathlib import Path
from collections import defaultdict, Counter

from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    render_template, request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _natural_sort_key,
    FEISHU_PRESET_TABLES,
)
from flask import Blueprint

pstx_bp = Blueprint('pstx_tool', __name__)


# ── 常量 ─────────────────────────────────────────────────────

_PAGE_TOKEN_RE = re.compile(
    r'(?<![A-Z0-9])PAGE(?:[_\-/ ]*)(\d+)([A-Z]?)(?![A-Z0-9])', re.IGNORECASE)
_PATH_SEGMENT_RE = re.compile(
    r'^(?P<head>.+?)\((?P<view>[^)]+)\)\s*:\s*(?P<tail>.+)$', re.IGNORECASE)
_SECTION_PATH_RE = re.compile(
    r'(?ims)^\s*SECTION_NUMBER\s+(?P<num>\d+)\s*\n\s*\'(?P<path>[^\']+)\'\s*:',)
_PAGE_NUMBER_LINE_RE = re.compile(
    r"""^\s*["']?PAGE_NUMBER["']?\s*(?:=|:)\s*["']?(?P<value>[A-Z0-9_./ -]+?)["']?\s*[;,]?\s*$""",
    re.IGNORECASE)
_OD_SKIP_PATTERNS = re.compile(
    r'\bPG\b|PGOOD|_PG_|_PGD\b|PG_N|PWRGD|POWER_GOOD|\bFAULT\b|_FAULT|VR_FAULT|\bALERT\b|_ALERT|SMBALERT|\bSDA\b|\bSCL\b|\bOC_N\b|_OC\b|\bPRSNT\b|\bPRESENT\b|\bINT_N\b|\bIRQ_N\b',
    re.IGNORECASE)
_DIFF_SUFFIX_PAIRS = [
    ('_P', '_N'), ('_DP', '_DN'), ('.P', '.N'),
    ('_TXPLUS', '_TXMINUS'), ('_RXPLUS', '_RXMINUS'),
]
_VALID_BOM_OPTIONS = {'', 'DEPOP', 'OPTION', 'MAIN_PLD', 'MAIN', 'ALT', 'DNP'}
_FUZZY_KEYWORDS = sorted(opt for opt in _VALID_BOM_OPTIONS if opt)
_GROUND_TOKEN_RE = re.compile(
    r'(?:[A-Z0-9]*GND[A-Z0-9]*|[A-Z0-9]*VSS[A-Z0-9]*|0V|0)', re.IGNORECASE)
_POWER_TOKEN_RE = re.compile(
    r'(?:VCC|VDD|VBAT|VCORE|VCCIO|PVDD|PVCC|AVDD|DVDD|VBUS)[A-Z0-9]*', re.IGNORECASE)
_CHIP_REFDES_RE = re.compile(r'^(?:XU|PU|U)[A-Z0-9]+$', re.IGNORECASE)
_OD_STRONG_TOKENS = {'SCL', 'SMBALERT', 'SMBDAT', 'SMBDATA', 'SMBCLK', 'OD', 'OC'}
_OD_WEAK_TOKENS = {'FAULT', 'IRQ', 'INT', 'PGOOD', 'PWROK', 'PWRGD', 'PRSNT', 'PRESENT'}
_MODULE_ORDER_LINE_RE = re.compile(
    r'^\s*(?P<path>@\S+)\s+(?P<unk1>\d+)\s+(?P<unk2>\d+)\s+(?P<start>\d+)\s+(?P<count>\d+)\s+(?P<flag>\d+)\s*$')

COMP_TYPE_CN = {
    'CAP': '电容', 'CAP_POL': '电解/钽电容', 'RES': '电阻', 'IND': '电感/磁珠',
    'IC': 'IC 芯片', 'CONN': '连接器', 'DIODE': '二极管', 'LED': 'LED',
    'FET': 'MOS/FET', 'BJT': '三极管', 'XTAL': '晶振', 'FUSE': '保险丝',
    'SWITCH': '开关', 'TESTPOINT': '测试点', 'TRANSFORMER': '变压器',
}
_TYPE_ORDER = list(COMP_TYPE_CN.keys())

# ── Excel 导出样式 ──
_BL = PatternFill('solid', fgColor='1F4E79')
_OR = PatternFill('solid', fgColor='C55A11')
_GR = PatternFill('solid', fgColor='375623')
_GY = PatternFill('solid', fgColor='595959')
_RF = PatternFill('solid', fgColor='FFCCCC')
_WF = Font(color='FFFFFF', bold=True, size=10)
_BF = Font(bold=True, size=10)
_NF = Font(size=10)
_CA = Alignment(horizontal='center', vertical='center', wrap_text=True)
_LA = Alignment(horizontal='left', vertical='center', wrap_text=True)
_TH = Side(style='thin')
_BD = Border(left=_TH, right=_TH, top=_TH, bottom=_TH)


# ── 文本解析 ─────────────────────────────────────────────────

def _split_named_blocks(text, marker):
    return re.split(rf'(?:^|\n){re.escape(marker)}\n', text)[1:]


def _join_continuations(text):
    normalized = str(text or '').replace('\r\n', '\n').replace('\r', '\n')
    lines = normalized.split('\n')
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


def _extract_attrs(text):
    attrs = {}
    for m in re.finditer(r"\b([A-Z][A-Z0-9_]*)\s*=\s*'([^']*)'", text):
        key, val = m.group(1), m.group(2)
        if key not in attrs:
            attrs[key] = val
    return attrs


def _get_comp_type(refdes, part_name):
    pn = part_name.lower()
    type_rules = [
        (['cap_pol'], 'CAP_POL'), (['cap_hdl', 'cap_'], 'CAP'),
        (['res_hdl', 'res_'], 'RES'),
        (['ind_hdl', 'ind_', 'ferrite', 'fer_hdl', 'fb_hdl'], 'IND'),
        (['osc_', 'crystal', 'xtal'], 'XTAL'),
        (['conn_', 'connector'], 'CONN'), (['led_'], 'LED'),
        (['diode', '_d_hdl'], 'DIODE'),
        (['mosfet', 'mos_', 'nmos', 'pmos', 'nfet', 'pfet'], 'FET'),
        (['bjt', 'transistor', 'npn', 'pnp'], 'BJT'), (['fuse'], 'FUSE'),
        (['sw_hdl', 'switch'], 'SWITCH'),
        (['testpoint', 'test_point', 'tp_hdl'], 'TESTPOINT'),
        (['transformer', 'xfmr'], 'TRANSFORMER'),
    ]
    for keywords, ctype in type_rules:
        if any(k in pn for k in keywords):
            return ctype
    prefix = (re.match(r'[A-Za-z]+', refdes) or re.match(r'', '')).group(0).upper()
    prefix_map = {
        'C': 'CAP', 'PC': 'CAP', 'R': 'RES', 'L': 'IND', 'FB': 'IND',
        'U': 'IC', 'J': 'CONN', 'P': 'CONN', 'CN': 'CONN', 'Q': 'FET',
        'D': 'DIODE', 'LED': 'LED', 'Y': 'XTAL', 'F': 'FUSE', 'SW': 'SWITCH',
        'TP': 'TESTPOINT', 'T': 'TRANSFORMER',
    }
    return prefix_map.get(prefix, 'IC')


def _net_is_gnd(net):
    return bool(_GROUND_TOKEN_RE.fullmatch(
        net.upper().split('_')[0] if '_' in net else net.upper()))


def _net_is_power(net):
    t = net.upper().split('_')[0]
    return bool(re.fullmatch(r'P?\d+V\d*', t)) or bool(_POWER_TOKEN_RE.fullmatch(t))


def _is_depop_option(bo):
    return str(bo or '').strip().upper() in {'DEPOP', 'DNP'}


def _parse_ohms(value_str):
    if not value_str:
        return None
    s = re.sub(r'\s', '', value_str.upper()).replace('Ω', 'R').replace('OHM', 'R').replace('OHMS', 'R')
    m = re.match(r'^([\d.]+)([KMGR]?)$', s)
    if m:
        val = float(m.group(1))
        return val * {'K': 1e3, 'M': 1e6, 'G': 1e9, 'R': 1, '': 1}.get(m.group(2), 1)
    emb = re.match(r'^(\d+)([KMGR])(\d+)$', s)
    if emb:
        return float(f'{emb.group(1)}.{emb.group(3)}') * {'K': 1e3, 'M': 1e6, 'G': 1e9, 'R': 1}.get(
            emb.group(2), 1)
    return None


def _infer_voltage(net_name):
    if _OD_SKIP_PATTERNS.search(net_name):
        return None
    token = net_name.upper().split('_')[0]
    if _net_is_gnd(token):
        return 0.0
    m = re.fullmatch(r'P?(\d+)V(\d*)', token)
    if m:
        i, f = m.groups()
        return float(f'{i}.{f}') if f else float(i)
    return None


# ── 网络分析 ─────────────────────────────────────────────────

def _collect_diff_pairs(nets):
    diff_pairs = {}
    upper_map = {n.upper(): n for n in nets}
    for net_name in nets:
        u = net_name.upper()
        for ps, ns in _DIFF_SUFFIX_PAIRS:
            pu, nu = ps.upper(), ns.upper()
            if u.endswith(pu):
                p = upper_map.get(u[:-len(pu)] + nu)
                if p:
                    diff_pairs[net_name[:-len(ps)]] = {'P': net_name, 'N': p}
                break
            elif u.endswith(nu):
                p = upper_map.get(u[:-len(nu)] + pu)
                if p and net_name[:-len(ns)] not in diff_pairs:
                    diff_pairs[net_name[:-len(ns)]] = {'P': p, 'N': net_name}
                break
    return diff_pairs


def _edit_distance(a, b):
    a, b = a.upper(), b.upper()
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    dp = list(range(len(b) + 1))
    for i, ca in enumerate(a):
        prev = dp[:]
        dp[0] = i + 1
        for j, cb in enumerate(b):
            dp[j + 1] = min(prev[j] + (0 if ca == cb else 1), dp[j] + 1, prev[j + 1] + 1)
    return dp[len(b)]


def _is_chip_component(refdes, comp):
    return comp.get('comp_type') == 'IC' and bool(_CHIP_REFDES_RE.match(refdes or ''))


def _classify_series_bias_ratio(series_ohms, bias_ohms):
    if series_ohms is None or bias_ohms is None or bias_ohms <= 0:
        return None, '⚪ 阻值缺失'
    ratio = series_ohms / bias_ohms
    if bias_ohms < 1000 and ratio > 0.1:
        return ratio, '❌ 高风险'
    if ratio >= 0.33:
        return ratio, '❌ 高风险'
    if ratio > 0.1:
        return ratio, '⚠️ 关注'
    return ratio, '✅ 正常'


# ── 解析器核心 ───────────────────────────────────────────────

def _parse_pstxprt(content):
    text = _join_continuations(content)
    components = {}
    for block in _split_named_blocks(text, 'PART_NAME'):
        m = re.match(r"(\S+)\s+'([^']+)'", block.split('\n')[0].strip())
        if not m:
            continue
        refdes, part_name = m.group(1), m.group(2)
        attrs = _extract_attrs(block)
        phys_raw = attrs.get('PHYS_PAGE', '').strip()
        phys_page = f'PAGE{phys_raw}' if phys_raw.isdigit() else ''
        components[refdes] = {
            'refdes': refdes, 'part_name': part_name,
            'hq_code': attrs.get('HQ_CODE', ''), 'value': attrs.get('VALUE', ''),
            'package': attrs.get('PACKAGE', ''), 'material': attrs.get('MATERIAL', ''),
            'tolerance': attrs.get('TOLERANCE', ''), 'voltage': attrs.get('VOLTAGE', ''),
            'current': attrs.get('CURRENT', ''), 'power': attrs.get('POWER', ''),
            'bom_option': attrs.get('BOM_OPTION', ''), 'bom_cost': attrs.get('BOM_COST', ''),
            'room': attrs.get('ROOM', ''), 'drawing': attrs.get('DRAWING', ''),
            'page': phys_page or '', 'comp_type': _get_comp_type(refdes, part_name),
        }
    return components


def _parse_pstxnet(content):
    text = _join_continuations(content)
    nets = {}
    node_re = re.compile(r'NODE_NAME\s+(\S+)\s+(\S+)')
    pin_name_re = re.compile(r"'([^']+)'\s*:")
    for block in _split_named_blocks(text, 'NET_NAME'):
        m = re.search(r"'([^']+)'", block)
        if not m:
            continue
        net_name = m.group(1)
        nodes = []
        matches = list(node_re.finditer(block))
        for idx, nm in enumerate(matches):
            ns = matches[idx + 1].start() if idx + 1 < len(matches) else len(block)
            pn = pin_name_re.search(block[nm.end():ns])
            nodes.append({
                'refdes': nm.group(1), 'pin': nm.group(2),
                'pin_name': pn.group(1) if pn else nm.group(2),
            })
        if nodes:
            nets[net_name] = nodes
    return nets


def _parse_all(prt_content, net_content):
    comps, nets = _parse_pstxprt(prt_content), _parse_pstxnet(net_content)
    comp_nets = {}
    for net_name, nodes in nets.items():
        for node in nodes:
            comp_nets.setdefault(node['refdes'], {})[node['pin']] = net_name
    for refdes, comp in comps.items():
        comp['nets'] = comp_nets.get(refdes, {})
    return comps, nets, comp_nets


# ── BOM 构建 ─────────────────────────────────────────────────

def _build_bom(comps):
    dn, dd = [], []
    for comp in comps.values():
        ctype = comp.get('comp_type', '')
        row = {
            '位号': comp['refdes'], '料号': comp.get('hq_code', ''),
            '描述': comp.get('part_name', ''), '值': comp.get('value', ''),
            '封装': comp.get('package', ''), '耐压/额定电压': comp.get('voltage', ''),
            '额定功率': comp.get('power', ''), '精度': comp.get('tolerance', ''),
            '材质': comp.get('material', ''), '类型': COMP_TYPE_CN.get(ctype, ctype),
            '_ctype': ctype, '页面': comp.get('page', ''), 'ROOM': comp.get('room', ''),
        }
        (dd if _is_depop_option(comp.get('bom_option', '')) else dn).append(row)

    def _merge(detail):
        if not detail:
            return []
        groups = {}
        for r in detail:
            k = r['料号'] or r['描述']
            if k not in groups:
                groups[k] = {
                    '料号': r['料号'], '位号列表': [], '数量': 0,
                    '描述': r['描述'], '值': r['值'], '封装': r['封装'],
                    '耐压': r['耐压/额定电压'], '额定功率': r['额定功率'],
                    '精度': r['精度'], '材质': r['材质'], '类型': r['类型'],
                    '_ctype': r['_ctype'],
                }
            groups[k]['位号列表'].append(r['位号'])
            groups[k]['数量'] += 1
        merged = list(groups.values())
        merged.sort(key=lambda r: (
            _TYPE_ORDER.index(r['_ctype']) if r['_ctype'] in _TYPE_ORDER else 99, r['料号']))
        for i, r in enumerate(merged, 1):
            r['序号'] = i
            r['位号列表'] = ', '.join(sorted(r['位号列表'], key=_natural_sort_key))
            del r['_ctype']
        return merged

    def _clean(rows):
        return [{k: v for k, v in r.items() if k != '_ctype'} for r in rows]

    return _clean(dn), _clean(dd), _merge(dn), _merge(dd)


# ── 网络分析 ─────────────────────────────────────────────────

def _analyze_networks(nets, comps):
    single_node = {k: v for k, v in nets.items() if len(v) == 1}
    gnd_nets = {k: v for k, v in nets.items() if _net_is_gnd(k)}
    power_nets = {k: v for k, v in nets.items()
                  if _net_is_power(k) and k not in gnd_nets}
    diff_pairs = _collect_diff_pairs(nets)
    pc = Counter()
    for comp in comps.values():
        pc[comp.get('page', '') or 'UNKNOWN'] += 1
    return {
        'total': len(nets), 'single_node': single_node,
        'gnd_nets': gnd_nets, 'power_nets': power_nets,
        'diff_pairs': diff_pairs, 'page_counter': pc,
    }


# ── DRC 检查 ─────────────────────────────────────────────────

def _check_drc(comps, nets):
    missing_hq, missing_val, missing_pkg, tbd_attrs, single_pin, unnamed = (
        [], [], [], [], [], [])
    bom_opt_comps = []
    for refdes, comp in comps.items():
        ctype = comp.get('comp_type', '')
        if ctype == 'TESTPOINT':
            continue
        base = {'位号': refdes, '类型': COMP_TYPE_CN.get(ctype, ctype),
                '页面': comp.get('page', '')}
        if not comp.get('hq_code'):
            missing_hq.append(base.copy())
        if not comp.get('value'):
            missing_val.append(base.copy())
        if not comp.get('package'):
            missing_pkg.append(base.copy())
        for attr in ('voltage', 'current', 'power'):
            val = comp.get(attr, '')
            if val and 'TBD' in val.upper():
                tbd_attrs.append({
                    '位号': refdes, '属性': attr.upper(), '当前值': val,
                    '类型': COMP_TYPE_CN.get(ctype, ctype), '页面': comp.get('page', ''),
                })
        bo = str(comp.get('bom_option', '') or '').strip().upper()
        if bo:
            bom_opt_comps.append({
                '位号': refdes, '类型': COMP_TYPE_CN.get(ctype, ctype),
                'BOM_OPTION值': bo,
                '是否DEPOP': '是' if _is_depop_option(bo) else '否',
                '页面': comp.get('page', ''),
            })

    for net_name, nodes in nets.items():
        if len(nodes) == 1:
            n = nodes[0]
            c = comps.get(n['refdes'], {})
            if c.get('comp_type') != 'TESTPOINT' and not re.search(r'^UNNAMED_', net_name, re.I):
                single_pin.append({
                    '网络名': net_name, '连接元件': n['refdes'],
                    '引脚': n['pin_name'], '页面': c.get('page', ''),
                })
        if re.search(r'^UNNAMED_', net_name, re.I):
            unnamed.append({'网络名': net_name, '节点数': len(nodes)})

    risk_per_value = {}
    for val in set(
            str(comp.get('bom_option', '') or '').strip().upper()
            for comp in comps.values()):
        if not val:
            continue
        if val in _VALID_BOM_OPTIONS:
            risk_per_value[val] = '✅ 合法'
        else:
            risk_per_value[val] = ('❌ 疑似拼错' if min(
                _edit_distance(val, kw) for kw in _FUZZY_KEYWORDS) <= 2 else '⚠ 未知值')
    for item in bom_opt_comps:
        item['拼写风险'] = risk_per_value.get(item['BOM_OPTION值'], '')

    return {
        'missing_hq_code': missing_hq, 'missing_value': missing_val,
        'missing_package': missing_pkg, 'tbd_attrs': tbd_attrs,
        'single_pin_nets': single_pin, 'unnamed_nets': unnamed,
        'bom_option_components': sorted(
            bom_opt_comps, key=lambda r: _natural_sort_key(r['位号'])),
    }


# ── 降额分析 ─────────────────────────────────────────────────

def _analyze_derating(comps, nets, pct=70.0):
    comp_nets = {}
    for net_name, nodes in nets.items():
        for node in nodes:
            comp_nets.setdefault(node['refdes'], []).append(net_name)
    rows = []
    for refdes, comp in comps.items():
        ctype = comp.get('comp_type', '')
        if ctype not in ('CAP', 'CAP_POL'):
            continue
        if _is_depop_option(comp.get('bom_option', '')):
            continue
        nets_u = list(dict.fromkeys(comp_nets.get(refdes, [])))
        rated_str = comp.get('voltage', '')
        if not rated_str:
            rows.append({
                '位号': refdes, '值': comp.get('value', ''),
                '封装': comp.get('package', ''),
                '类型': COMP_TYPE_CN.get(ctype, ctype),
                '额定电压': rated_str, '状态': '⚪ 无额定电压',
                '页面': comp.get('page', ''),
            })
            continue
        m = re.match(r'([\d.]+)\s*V', rated_str.strip(), re.I)
        if not m:
            rows.append({
                '位号': refdes, '值': comp.get('value', ''),
                '封装': comp.get('package', ''),
                '类型': COMP_TYPE_CN.get(ctype, ctype),
                '额定电压': rated_str, '状态': '⚪ 无法解析额定电压',
                '页面': comp.get('page', ''),
            })
            continue
        rated_v = float(m.group(1))
        v_known = {}
        for net_name in nets_u:
            if _net_is_gnd(net_name):
                continue
            if _OD_SKIP_PATTERNS.search(net_name):
                continue
            v = _infer_voltage(net_name)
            if v and v > 0:
                v_known.setdefault(round(v, 6), net_name)
        if not v_known:
            rows.append({
                '位号': refdes, '值': comp.get('value', ''),
                '封装': comp.get('package', ''),
                '类型': COMP_TYPE_CN.get(ctype, ctype),
                '额定电压': rated_str, '状态': '⚪ 无法推断工作电压',
                '页面': comp.get('page', ''),
            })
            continue
        max_v, from_net = max(v_known.items(), key=lambda x: x[0])
        usage_pct = max_v / rated_v * 100
        status = (
            f'✅ 合格 ({usage_pct:.0f}% ≤ {pct:.0f}%)'
            if usage_pct <= pct
            else f'❌ 不合格 ({usage_pct:.0f}% > {pct:.0f}%)'
        )
        rows.append({
            '位号': refdes, '值': comp.get('value', ''),
            '封装': comp.get('package', ''),
            '类型': COMP_TYPE_CN.get(ctype, ctype),
            '额定电压': rated_str, '推断工作电压(V)': str(max_v),
            '推断来源网络': from_net, '降额比': f'{rated_v / max_v:.2f}',
            '状态': status, '页面': comp.get('page', ''),
        })
    rows.sort(key=lambda r: (
        0 if r['状态'].startswith('❌') else 1 if r['状态'].startswith('✅') else 2,
        _natural_sort_key(r.get('位号', '')),
    ))
    return rows


# ── 电阻分析 ─────────────────────────────────────────────────

def _analyze_resistors(comps, nets):
    pullups, pulldowns, series_by_net = (
        defaultdict(list), defaultdict(list), defaultdict(list))
    node_lookup = {}
    for net_name, nodes in nets.items():
        for node in nodes:
            node_lookup[(node['refdes'], node['pin'])] = node.get('pin_name', node['pin'])

    for refdes, comp in comps.items():
        if comp.get('comp_type') != 'RES':
            continue
        if _is_depop_option(comp.get('bom_option', '')):
            continue
        pin_nets = list(dict.fromkeys(comp.get('nets', {}).values()))
        if len(pin_nets) != 2:
            continue
        net_a, net_b = pin_nets
        ohms = _parse_ohms(comp.get('value', ''))
        val = comp.get('value', '')
        page = comp.get('page', '')
        a_pwr, b_pwr = _net_is_power(net_a), _net_is_power(net_b)
        a_gnd, b_gnd = _net_is_gnd(net_a), _net_is_gnd(net_b)
        entry = {'refdes': refdes, 'ohms': ohms, 'value': val, 'page': page}
        if a_pwr and not b_pwr and not b_gnd:
            pullups[net_b].append({**entry, 'power_net': net_a})
        elif b_pwr and not a_pwr and not a_gnd:
            pullups[net_a].append({**entry, 'power_net': net_b})
        elif a_gnd and not b_gnd and not b_pwr:
            pulldowns[net_b].append(entry.copy())
        elif b_gnd and not a_gnd and not a_pwr:
            pulldowns[net_a].append(entry.copy())
        elif not a_pwr and not b_pwr and not a_gnd and not b_gnd:
            series_by_net[net_a].append({**entry, 'other_net': net_b})
            series_by_net[net_b].append({**entry, 'other_net': net_a})

    dup_pu, dup_pd, div_risks = [], [], []
    for sig_net, pu in sorted(pullups.items()):
        if len(pu) < 2:
            continue
        g = sorted(pu, key=lambda r: _natural_sort_key(r.get('refdes', '')))
        dup_pu.append({
            '信号网络': sig_net, '上拉数量': len(g),
            '位号': ', '.join(r['refdes'] for r in g),
            '阻值': ', '.join(r['value'] for r in g),
            '上拉电源': ', '.join(dict.fromkeys(r['power_net'] for r in g)),
            '页面': ', '.join(dict.fromkeys(r['page'] for r in g)),
        })
    for sig_net, pd in sorted(pulldowns.items()):
        if len(pd) < 2:
            continue
        g = sorted(pd, key=lambda r: _natural_sort_key(r.get('refdes', '')))
        dup_pd.append({
            '信号网络': sig_net, '下拉数量': len(g),
            '位号': ', '.join(r['refdes'] for r in g),
            '阻值': ', '.join(r['value'] for r in g),
            '页面': ', '.join(dict.fromkeys(r['page'] for r in g)),
        })

    seen = set()
    for net_name in sorted(series_by_net.keys()):
        for sr in sorted(series_by_net[net_name],
                         key=lambda r: _natural_sort_key(r.get('refdes', ''))):
            o_net = sr['other_net']
            for bias_kind, bias_map in [('上拉', pullups), ('下拉', pulldowns)]:
                for bias in bias_map.get(o_net, []):
                    pk = (sr['refdes'], bias['refdes'], bias_kind, net_name, o_net)
                    if pk in seen:
                        continue
                    seen.add(pk)
                    ratio, status = _classify_series_bias_ratio(
                        sr['ohms'], bias.get('ohms'))
                    ref_net = bias.get('power_net', '') if bias_kind == '上拉' else 'GND'
                    pages = ', '.join(
                        dict.fromkeys(v for v in [sr.get('page', ''), bias.get('page', '')] if v))
                    div_risks.append({
                        '串阻位号': sr['refdes'], '串阻值': sr['value'],
                        '偏置类型': bias_kind, '偏置位号': bias['refdes'],
                        '偏置值': bias['value'], '偏置所在网络': o_net,
                        '偏置参考网络': ref_net,
                        '串/偏置比': f'{ratio:.3f}' if ratio else '',
                        '状态': status, '页面': pages,
                    })
    div_risks.sort(key=lambda r: (
        0 if r['状态'].startswith('❌') else 1,
        _natural_sort_key(r.get('串阻位号', '')),
    ))

    return {
        'dup_pullups': dup_pu, 'dup_pulldowns': dup_pd,
        'divider_risks': div_risks,
    }


# ── Excel 导出 ───────────────────────────────────────────────

def _xl_hdr(ws, row_idx, fill):
    for cell in ws[row_idx]:
        if cell.value is not None:
            cell.fill = fill
            cell.font = _WF
            cell.alignment = _CA
            cell.border = _BD


def _xl_autowidth(ws, mx=50):
    for col in ws.columns:
        vals = [str(c.value or '') for c in col]
        ws.column_dimensions[col[0].column_letter].width = min(
            max((len(v) for v in vals), default=8) + 2, mx)


def _xl_write_rows(ws, rows, fill, hl_col=None, freeze=True):
    if not rows:
        ws.append(['（无数据）'])
        return
    hdrs = list(rows[0].keys())
    ws.append(hdrs)
    _xl_hdr(ws, ws.max_row, fill)
    hl_idx = hdrs.index(hl_col) if hl_col in hdrs else None
    for row in rows:
        ws.append(list(row.values()))
        ri = ws.max_row
        red = hl_idx is not None and '❌' in str(ws.cell(ri, hl_idx + 1).value or '')
        for cell in ws[ri]:
            cell.border = _BD
            cell.alignment = _LA
            cell.font = _NF
        if red:
            cell.fill = _RF
    _xl_autowidth(ws)
    if freeze:
        ws.freeze_panes = 'A2'


def _xl_section(ws, title, fill):
    ws.append([title])
    for cell in ws[ws.max_row]:
        cell.fill = fill
        cell.font = _WF
        cell.border = _BD
    ws.append([])


def _export_pstx_excel(data, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    na = data.get('net_analysis', {})
    drc = data.get('drc', {})
    drt = data.get('derating', [])
    res = data.get('resistor_analysis', {})
    mn = data.get('bom_normal_merged', [])
    md = data.get('bom_depop_merged', [])

    ws = wb.create_sheet('概览')
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 16
    drc_total = sum(len(v) for v in drc.values() if isinstance(v, list))
    fail = sum(1 for r in drt if r.get('状态', '').startswith('❌'))
    for label, val in [
        ('贴装元件种类数', len(mn)),
        ('贴装元件总数', sum(r.get('数量', 0) for r in mn)),
        ('DEPOP 元件种类数', len(md)),
        ('DEPOP 元件总数', sum(r.get('数量', 0) for r in md)),
        ('网络总数', na.get('total', '')),
        ('单端网络数（疑似漏连）', len(na.get('single_node', {}))),
        ('电源网络数', len(na.get('power_nets', {}))),
        ('差分对数', len(na.get('diff_pairs', {}))),
        ('DRC 问题总数', drc_total),
        ('电容降额不合格数', fail),
    ]:
        ws.append([label, val])
    for row in ws.iter_rows():
        for cell in row:
            cell.border = _BD
            cell.font = _BF if cell.column == 1 else _NF
            cell.alignment = _LA

    ws = wb.create_sheet('BOM_贴装')
    _xl_write_rows(ws, mn, _BL)
    ws = wb.create_sheet('BOM_DEPOP')
    _xl_write_rows(ws, md, _OR)
    ws = wb.create_sheet('BOM_明细')
    all_d = [{'DEPOP': '', **r} for r in data.get('bom_normal_detail', [])] + \
            [{'DEPOP': 'Y', **r} for r in data.get('bom_depop_detail', [])]
    _xl_write_rows(ws, all_d, _GY)

    ws = wb.create_sheet('网络分析')
    ws.freeze_panes = None
    _xl_section(ws, '电源网络', _BL)
    _xl_write_rows(ws, [
        {'网络名': k, '节点数': len(v)}
        for k, v in sorted(na.get('power_nets', {}).items(), key=lambda x: -len(x[1]))
    ], _BL, freeze=False)
    ws.append([])
    _xl_section(ws, 'GND 网络', _GR)
    _xl_write_rows(ws, [
        {'网络名': k, '节点数': len(v)}
        for k, v in sorted(na.get('gnd_nets', {}).items(), key=lambda x: -len(x[1]))
    ], _GR, freeze=False)
    ws.append([])
    _xl_section(ws, '差分对', _OR)
    _xl_write_rows(ws, [
        {'基础名': b, 'P端网络': pr['P'], 'N端网络': pr['N']}
        for b, pr in sorted(na.get('diff_pairs', {}).items())
    ], _OR, freeze=False)
    ws.append([])
    _xl_section(ws, '单端网络（疑似漏连）', _GY)
    _xl_write_rows(ws, [
        {'网络名': k, '连接元件': v[0]['refdes'], '引脚': v[0]['pin_name']}
        for k, v in sorted(na.get('single_node', {}).items())
    ], _GY, freeze=False)
    ws.append([])
    _xl_section(ws, '各页面元件数', _BL)
    _xl_write_rows(ws, [
        {'页面': p, '元件数': c}
        for p, c in sorted(na.get('page_counter', {}).items())
    ], _BL, freeze=False)
    _xl_autowidth(ws)

    ws = wb.create_sheet('设计检查')
    ws.freeze_panes = None
    sections = [
        ('TBD 待确认属性', 'tbd_attrs', _OR),
        ('缺少料号', 'missing_hq_code', _RF),
        ('缺少 VALUE', 'missing_value', _RF),
        ('缺少封装', 'missing_package', _RF),
        ('单端网络', 'single_pin_nets', _GY),
        ('未命名网络', 'unnamed_nets', _GY),
        ('BOM_OPTION 元件清单（含拼写风险）', 'bom_option_components', _BL),
    ]
    for title, key, fill in sections:
        _xl_section(ws, title, fill)
        _xl_write_rows(ws, drc.get(key, []), fill, freeze=False)
        ws.append([])
    _xl_autowidth(ws)

    ws = wb.create_sheet('降额分析')
    _xl_write_rows(ws, drt, _BL, hl_col='状态')

    ws = wb.create_sheet('电阻检查')
    ws.freeze_panes = None
    for title, key, hl, fill in [
        ('串阻分压风险', 'divider_risks', '状态', _OR),
        ('重复上拉', 'dup_pullups', None, _BL),
        ('重复下拉', 'dup_pulldowns', None, _GY),
    ]:
        _xl_section(ws, title, fill)
        _xl_write_rows(ws, res.get(key, []), fill, hl_col=hl, freeze=False)
        ws.append([])
    _xl_autowidth(ws)

    wb.save(out_path)
    return out_path


# ── 路由 ─────────────────────────────────────────────────────

@pstx_bp.route('/pstx', methods=['GET', 'POST'])
def tool_pstx():
    if request.method == 'POST':
        prt_file = request.files.get('prt_file')
        net_file = request.files.get('net_file')
        if not prt_file or not net_file:
            return "请上传 pstxprt.dat 和 pstxnet.dat", 400
        pct = float(request.form.get('pct', 70.0))

        uid = str(uuid.uuid4())[:8]
        prt_path = os.path.join(UPLOAD_DIR, f"pstxprt_{uid}.dat")
        net_path = os.path.join(UPLOAD_DIR, f"pstxnet_{uid}.dat")
        out_path = os.path.join(OUTPUT_DIR, f"PSTX分析报告_{uid}.xlsx")
        prt_file.save(prt_path)
        net_file.save(net_path)

        try:
            for enc in ['utf-8-sig', 'utf-16', 'utf-16-le', 'utf-16-be', 'utf-8', 'gb18030', 'cp936']:
                try:
                    prt = Path(prt_path).read_bytes().decode(enc)
                    break
                except Exception:
                    continue
            else:
                prt = Path(prt_path).read_bytes().decode('utf-8', errors='replace')
            for enc in ['utf-8-sig', 'utf-16', 'utf-16-le', 'utf-16-be', 'utf-8', 'gb18030', 'cp936']:
                try:
                    net = Path(net_path).read_bytes().decode(enc)
                    break
                except Exception:
                    continue
            else:
                net = Path(net_path).read_bytes().decode('utf-8', errors='replace')
        except Exception as e:
            return jsonify({'success': False, 'error': f'文件读取失败：{e}'})

        comps, nets, _ = _parse_all(prt, net)
        dn, dd, mn, md = _build_bom(comps)
        na = _analyze_networks(nets, comps)
        drc = _check_drc(comps, nets)
        drt = _analyze_derating(comps, nets, pct)
        res = _analyze_resistors(comps, nets)

        _export_pstx_excel({
            'bom_normal_detail': dn, 'bom_depop_detail': dd,
            'bom_normal_merged': mn, 'bom_depop_merged': md,
            'net_analysis': na, 'drc': drc, 'derating': drt,
            'resistor_analysis': res,
        }, out_path)

        drc_total = sum(len(v) for v in drc.values() if isinstance(v, list))
        fail = sum(1 for r in drt if r.get('状态', '').startswith('❌'))
        return jsonify({
            'success': True,
            'components': len(comps), 'nets': len(nets),
            'bom_normal': len(mn), 'bom_depop': len(md),
            'drc_total': drc_total, 'derating_fail': fail,
            'download': f'/download/PSTX分析报告_{uid}.xlsx',
            'bom_detail': dn + dd,
            'net_analysis': {
                'total': na.get('total', 0),
                'power_nets': {k: len(v) for k, v in na.get('power_nets', {}).items()},
                'gnd_nets': {k: len(v) for k, v in na.get('gnd_nets', {}).items()},
                'diff_pairs': na.get('diff_pairs', {}),
                'single_node': {k: len(v) for k, v in na.get('single_node', {}).items()},
            },
            'drc': drc,
            'derating': drt,
            'resistor_analysis': {
                'divider_risks': res.get('divider_risks', []),
                'dup_pullups': res.get('dup_pullups', []),
                'dup_pulldowns': res.get('dup_pulldowns', []),
            },
        })
    return render_template('index.html', tables=FEISHU_PRESET_TABLES)
