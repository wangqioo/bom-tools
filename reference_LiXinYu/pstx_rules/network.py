# -*- coding: utf-8 -*-
"""Network topology summary rules."""

from collections import Counter
from typing import Dict, Optional

from pstx_core.page_resolution import USER_VISIBLE_REAL_PAGE_LABEL, component_user_visible_page
from pstx_rules.common import _collect_diff_pairs, _natural_sort_key, _net_is_gnd, _net_is_power
from pstx_rules.result_meta import with_meta as _with_meta

def analyze_networks(nets: Dict, components: Dict, *, single_node_topology_nets: Optional[Dict] = None) -> dict:
    topology_nets = single_node_topology_nets or nets
    single_node = {
        k: v for k, v in nets.items()
        if len(v) == 1 and len(topology_nets.get(k, v) or []) == 1
    }
    gnd_nets    = {k: v for k, v in nets.items()
                   if _net_is_gnd(k)}
    power_nets  = {k: v for k, v in nets.items()
                   if _net_is_power(k)
                   and k not in gnd_nets}
    diff_pairs = _collect_diff_pairs(nets)
    page_counter: Counter = Counter()
    for comp in components.values():
        page_label = component_user_visible_page(comp) or 'UNKNOWN'
        page_counter[page_label] += 1
    page_rows = [
        {USER_VISIBLE_REAL_PAGE_LABEL: page, '元件数': count}
        for page, count in sorted(
            page_counter.items(),
            key=lambda item: (item[0] == 'UNKNOWN', _natural_sort_key(item[0])),
        )
    ]
    power_rows = [
        _with_meta({'网络名': k, '节点数': len(v)}, 'candidate', 'low', 'medium', 'net_name_power_token')
        for k, v in sorted(power_nets.items(), key=lambda x: (-len(x[1]), x[0].upper()))
    ]
    gnd_rows = [
        _with_meta({'网络名': k, '节点数': len(v)}, 'candidate', 'low', 'high', 'net_name_ground_token')
        for k, v in sorted(gnd_nets.items(), key=lambda x: (-len(x[1]), x[0].upper()))
    ]
    diff_rows = [
        _with_meta({'基础名': b, 'P端网络': pr['P'], 'N端网络': pr['N']},
                   'candidate', 'low', 'medium', 'paired_name_suffix')
        for b, pr in sorted(diff_pairs.items(), key=lambda x: x[0].upper())
    ]
    single_rows = [
        _with_meta({'网络名': k, '连接元件': v[0]['refdes'], '引脚': v[0]['pin_name']},
                   'candidate', 'medium', 'low', 'single_node_net')
        for k, v in sorted(single_node.items(), key=lambda x: x[0].upper())
    ]
    return {
        'total': len(nets), 'single_node': single_node,
        'gnd_nets': gnd_nets, 'power_nets': power_nets,
        'diff_pairs': diff_pairs, 'page_counter': page_counter,
        'page_rows': page_rows,
        'power_net_rows': power_rows,
        'gnd_net_rows': gnd_rows,
        'diff_pair_rows': diff_rows,
        'single_node_rows': single_rows,
    }


# ══════════════════════════════════════════════════════════
# 四、DRC 设计检查
# ══════════════════════════════════════════════════════════
