# -*- coding: utf-8 -*-
"""
分析模块：BOM / 网络分析 / DRC / 电容降额
纯 Python，无 pandas 依赖
"""

import re
from collections import Counter, defaultdict
from typing import Dict, List, Optional, Tuple


COMP_TYPE_CN = {
    'CAP':         '电容',
    'CAP_POL':     '电解/钽电容',
    'RES':         '电阻',
    'IND':         '电感/磁珠',
    'IC':          'IC 芯片',
    'CONN':        '连接器',
    'DIODE':       '二极管',
    'LED':         'LED',
    'FET':         'MOS/FET',
    'BJT':         '三极管',
    'XTAL':        '晶振',
    'FUSE':        '保险丝',
    'SWITCH':      '开关',
    'TESTPOINT':   '测试点',
    'TRANSFORMER': '变压器',
}

_TYPE_ORDER = list(COMP_TYPE_CN.keys())


# ══════════════════════════════════════════════════════════
# 一、BOM 分析
# ══════════════════════════════════════════════════════════

def build_bom(components: Dict):
    """
    返回 (detail_normal, detail_depop, merged_normal, merged_depop)
    每项均为 list[dict]
    """
    detail_normal, detail_depop = [], []

    for comp in components.values():
        ctype = comp.get('comp_type', '')
        row = {
            '位号':          comp['refdes'],
            '料号':          comp.get('hq_code', ''),
            '描述':          comp.get('part_name', ''),
            '值':            comp.get('value', ''),
            '封装':          comp.get('package', ''),
            '耐压/额定电压':  comp.get('voltage', ''),
            '额定功率':      comp.get('power', ''),
            '精度':          comp.get('tolerance', ''),
            '材质':          comp.get('material', ''),
            '类型':          COMP_TYPE_CN.get(ctype, ctype),
            '_ctype':        ctype,
            '页面':          comp.get('page', ''),
            'ROOM':          comp.get('room', ''),
        }
        if comp.get('bom_option') == 'DEPOP':
            detail_depop.append(row)
        else:
            detail_normal.append(row)

    def _merge(detail: List[dict]) -> List[dict]:
        if not detail:
            return []
        groups: Dict[str, dict] = {}
        for row in detail:
            key = row['料号'] or row['描述']
            if key not in groups:
                groups[key] = {
                    '料号':    row['料号'],
                    '位号列表': [],
                    '数量':    0,
                    '描述':    row['描述'],
                    '值':      row['值'],
                    '封装':    row['封装'],
                    '耐压':    row['耐压/额定电压'],
                    '额定功率': row['额定功率'],
                    '精度':    row['精度'],
                    '材质':    row['材质'],
                    '类型':    row['类型'],
                    '_ctype':  row['_ctype'],
                }
            groups[key]['位号列表'].append(row['位号'])
            groups[key]['数量'] += 1

        merged = []
        for g in groups.values():
            g['位号列表'] = ', '.join(sorted(g['位号列表']))
            merged.append(g)

        merged.sort(key=lambda r: (
            _TYPE_ORDER.index(r['_ctype']) if r['_ctype'] in _TYPE_ORDER else 99,
            r['料号']
        ))
        for i, row in enumerate(merged, 1):
            row['序号'] = i
            del row['_ctype']
        return merged

    # 去掉内部排序字段后再返回 detail
    def _clean_detail(rows):
        out = []
        for r in rows:
            d = {k: v for k, v in r.items() if k != '_ctype'}
            out.append(d)
        return out

    return (
        _clean_detail(detail_normal),
        _clean_detail(detail_depop),
        _merge(detail_normal),
        _merge(detail_depop),
    )


# ══════════════════════════════════════════════════════════
# 二、网络分析
# ══════════════════════════════════════════════════════════

def analyze_networks(nets: Dict, components: Dict) -> dict:
    """分析网络：统计 / 单端 / 电源 / 差分对 / 各页分布"""
    total = len(nets)

    single_node = {k: v for k, v in nets.items() if len(v) == 1}

    gnd_nets = {k: v for k, v in nets.items()
                if re.search(r'GND|AGND|SGND|PGND|DGND', k, re.I)}

    power_nets = {k: v for k, v in nets.items()
                  if re.search(r'^P\d|^[0-9]+V|VCC|VDD|VBAT|VCORE|VCCIO|PVDD|AVDD|DVDD', k, re.I)
                  and k not in gnd_nets}

    diff_pairs: Dict[str, dict] = {}
    for net_name in nets:
        for suffix_p, suffix_n in [('_P', '_N'), ('_DP', '_DN'),
                                    ('.P', '.N'), ('_TXPLUS', '_TXMINUS'),
                                    ('_RXPLUS', '_RXMINUS')]:
            if net_name.endswith(suffix_p):
                base = net_name[:-len(suffix_p)]
                counterpart = base + suffix_n
                if counterpart in nets:
                    diff_pairs[base] = {'P': net_name, 'N': counterpart}
            elif net_name.endswith(suffix_n):
                base = net_name[:-len(suffix_n)]
                counterpart = base + suffix_p
                if counterpart in nets and base not in diff_pairs:
                    diff_pairs[base] = {'P': counterpart, 'N': net_name}

    page_counter: Counter = Counter()
    for comp in components.values():
        page_counter[comp.get('page', 'UNKNOWN')] += 1

    return {
        'total':        total,
        'single_node':  single_node,
        'gnd_nets':     gnd_nets,
        'power_nets':   power_nets,
        'diff_pairs':   diff_pairs,
        'page_counter': page_counter,
    }


# ══════════════════════════════════════════════════════════
# 三、DRC 设计检查
# ══════════════════════════════════════════════════════════

def check_drc(components: Dict, nets: Dict) -> dict:
    missing_hq_code = []
    missing_value   = []
    missing_package = []
    tbd_attrs       = []
    single_pin_nets = []
    unnamed_nets    = []

    for refdes, comp in components.items():
        ctype = comp.get('comp_type', '')
        if ctype == 'TESTPOINT':
            continue

        if not comp.get('hq_code'):
            missing_hq_code.append({
                '位号': refdes,
                '类型': COMP_TYPE_CN.get(ctype, ctype),
                '页面': comp.get('page', ''),
            })

        if not comp.get('value'):
            missing_value.append({
                '位号': refdes,
                '类型': COMP_TYPE_CN.get(ctype, ctype),
                '页面': comp.get('page', ''),
            })

        if not comp.get('package'):
            missing_package.append({
                '位号': refdes,
                '类型': COMP_TYPE_CN.get(ctype, ctype),
                '页面': comp.get('page', ''),
            })

        for attr in ('voltage', 'current', 'power'):
            val = comp.get(attr, '')
            if val and 'TBD' in val.upper():
                tbd_attrs.append({
                    '位号':  refdes,
                    '属性':  attr.upper(),
                    '当前值': val,
                    '类型':  COMP_TYPE_CN.get(ctype, ctype),
                    '页面':  comp.get('page', ''),
                })

    for net_name, nodes in nets.items():
        if len(nodes) == 1:
            node = nodes[0]
            comp = components.get(node['refdes'], {})
            single_pin_nets.append({
                '网络名':   net_name,
                '连接元件': node['refdes'],
                '引脚':     node['pin_name'],
                '页面':     comp.get('page', ''),
            })

        if re.search(r'^UNNAMED_', net_name, re.I):
            unnamed_nets.append({'网络名': net_name, '节点数': len(nodes)})

    return {
        'missing_hq_code': missing_hq_code,
        'missing_value':   missing_value,
        'missing_package': missing_package,
        'tbd_attrs':       tbd_attrs,
        'single_pin_nets': single_pin_nets,
        'unnamed_nets':    unnamed_nets,
        'bom_option_typos': check_bom_option_typos(components),
    }


# ──────────────────────────────────────────────────────────
# BOM_OPTION 拼写错误检测
# ──────────────────────────────────────────────────────────

_VALID_BOM_OPTIONS = {'', 'DEPOP', 'OPTION', 'MAIN_PLD', 'MAIN', 'ALT', 'DNP'}
_FUZZY_KEYWORDS    = ['DEPOP', 'OPTION']


def _edit_distance(a: str, b: str) -> int:
    a, b = a.upper(), b.upper()
    if a == b:   return 0
    if not a:    return len(b)
    if not b:    return len(a)
    dp = list(range(len(b) + 1))
    for i, ca in enumerate(a):
        prev = dp[:]
        dp[0] = i + 1
        for j, cb in enumerate(b):
            dp[j + 1] = min(prev[j] + (0 if ca == cb else 1),
                            dp[j] + 1, prev[j + 1] + 1)
    return dp[len(b)]


def check_bom_option_typos(components: Dict) -> List[dict]:
    option_map: Dict[str, List[str]] = defaultdict(list)
    for refdes, comp in components.items():
        val = (comp.get('bom_option') or '').strip().upper()
        option_map[val].append(refdes)

    rows = []
    for val, refdes_list in sorted(option_map.items()):
        if val in _VALID_BOM_OPTIONS:
            continue
        min_dist = min(_edit_distance(val, kw) for kw in _FUZZY_KEYWORDS)
        nearest  = min(_FUZZY_KEYWORDS, key=lambda kw: _edit_distance(val, kw))
        rows.append({
            '实际填写值':      val,
            '疑似应为':        nearest if min_dist <= 2 else '未知',
            '编辑距离':        min_dist,
            '使用该值的位号':  ', '.join(sorted(refdes_list)),
            '数量':           len(refdes_list),
            '风险':           '❌ 疑似拼错' if min_dist <= 2 else '⚠ 未知值',
        })
    return rows


# ══════════════════════════════════════════════════════════
# 四、电容降额分析
# ══════════════════════════════════════════════════════════

_VOLT_RULES: List[Tuple[str, float]] = [
    (r'P48V',                  48.0),
    (r'P24V',                  24.0),
    (r'P19V',                  19.0),
    (r'P15V',                  15.0),
    (r'P12V',                  12.0),
    (r'\b12V',                 12.0),
    (r'P9V',                    9.0),
    (r'P7V',                    7.4),
    (r'P5V(?!\d)',              5.0),
    (r'\b5V',                   5.0),
    (r'P3V3',                   3.3),
    (r'\b3V3',                  3.3),
    (r'P3V',                    3.3),
    (r'P2V5',                   2.5),
    (r'2V5',                    2.5),
    (r'P1V8',                   1.8),
    (r'1V8',                    1.8),
    (r'P1V5',                   1.5),
    (r'1V5',                    1.5),
    (r'P1V2',                   1.2),
    (r'1V2',                    1.2),
    (r'P1V05',                  1.05),
    (r'1V05',                   1.05),
    (r'P1V(?!\d)',              1.0),
    (r'1V0',                    1.0),
    (r'P0V9',                   0.9),
    (r'0V9',                    0.9),
    (r'P0V8',                   0.8),
    (r'GND',                    0.0),
    (r'AGND|PGND|DGND|SGND',   0.0),
]


def _infer_voltage(net_name: str) -> Optional[float]:
    for pattern, volt in _VOLT_RULES:
        if re.search(pattern, net_name, re.IGNORECASE):
            return volt
    return None


def _parse_rated_voltage(volt_str: str) -> Optional[float]:
    m = re.match(r'([\d.]+)\s*V', volt_str.strip(), re.IGNORECASE)
    return float(m.group(1)) if m else None


def analyze_derating(components: Dict,
                     nets: Dict,
                     ratio: float = 2.0,
                     custom_volt_map: Optional[Dict[str, float]] = None
                     ) -> List[dict]:
    """电容耐压降额分析，返回 list[dict]"""
    comp_nets: Dict[str, List[str]] = defaultdict(list)
    for net_name, nodes in nets.items():
        for node in nodes:
            comp_nets[node['refdes']].append(net_name)

    rows = []
    for refdes, comp in components.items():
        ctype = comp.get('comp_type', '')
        if ctype not in ('CAP', 'CAP_POL'):
            continue

        rated_v_str = comp.get('voltage', '')
        if not rated_v_str:
            rows.append(_derating_row(refdes, comp, '', '无额定电压', None, ''))
            continue

        rated_v = _parse_rated_voltage(rated_v_str)
        if rated_v is None:
            rows.append(_derating_row(refdes, comp, rated_v_str, '无法解析额定电压', None, ''))
            continue

        connected = comp_nets.get(refdes, [])
        max_v, from_net = None, ''

        for net_name in connected:
            v = None
            if custom_volt_map:
                for key, vv in custom_volt_map.items():
                    if key.upper() in net_name.upper():
                        v = vv
                        break
            if v is None:
                v = _infer_voltage(net_name)
            if v is not None and v > 0:
                if max_v is None or v > max_v:
                    max_v, from_net = v, net_name

        if max_v is None:
            status = '⚪ 无法推断工作电压'
            derating = None
        else:
            derating = rated_v / max_v
            if derating >= ratio:
                status = f'✅ 合格 ({derating:.1f}x)'
            else:
                status = f'❌ 不合格 ({derating:.2f}x < {ratio}x)'

        rows.append(_derating_row(refdes, comp, rated_v_str, status, derating, from_net,
                                   max_v, connected))

    rows.sort(key=lambda r: (0 if r['状态'].startswith('❌') else 1))
    return rows


def _derating_row(refdes, comp, rated_str, status, derating, from_net,
                   working_v=None, nets_list=None):
    ctype = comp.get('comp_type', '')
    return {
        '位号':            refdes,
        '值':              comp.get('value', ''),
        '封装':            comp.get('package', ''),
        '类型':            COMP_TYPE_CN.get(ctype, ctype),
        '额定电压':        rated_str or '',
        '推断工作电压(V)':  str(working_v) if working_v is not None else '',
        '推断来源网络':    from_net,
        '所有连接网络':    ', '.join(nets_list) if nets_list else '',
        '降额比':          f'{derating:.2f}' if derating is not None else '',
        '状态':            status,
        '页面':            comp.get('page', ''),
        'DEPOP':           'Y' if comp.get('bom_option') == 'DEPOP' else '',
    }


# ══════════════════════════════════════════════════════════
# 五、汇总统计
# ══════════════════════════════════════════════════════════

def build_summary(components: Dict, nets: Dict,
                  bom_normal_merged: List[dict],
                  bom_depop_merged: List[dict],
                  net_analysis: dict,
                  drc: dict,
                  derating: List[dict]) -> dict:
    """生成概览数字"""
    normal_total = sum(r.get('数量', 0) for r in bom_normal_merged)
    depop_total  = sum(r.get('数量', 0) for r in bom_depop_merged)
    drc_total    = sum(
        len(v) for v in drc.values() if isinstance(v, list)
    )
    derating_fail = sum(1 for r in derating if r['状态'].startswith('❌'))
    derating_unknown = sum(1 for r in derating if r['状态'].startswith('⚪'))

    return {
        '贴装元件种类': len(bom_normal_merged),
        '贴装元件总数': normal_total,
        'DEPOP种类':    len(bom_depop_merged),
        'DEPOP总数':    depop_total,
        '网络总数':     net_analysis.get('total', 0),
        'DRC问题数':    drc_total,
        '降额不合格数': derating_fail,
        '降额无法判断': derating_unknown,
    }
