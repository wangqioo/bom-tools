(function () {
  const SVG_NS = 'http://www.w3.org/2000/svg';
  const context = window.PSTX_TOPOLOGY_CONTEXT || {};
  const state = {
    topology: null,
    layoutSeed: 0,
    selected: null,
    edgeLabelMode: 'auto',
  };

  function roleClass(role) {
    const text = String(role || '').toLowerCase();
    if (text.includes('level')) return 'level';
    if (text.includes('power')) return 'power';
    if (text.includes('connector')) return 'connector';
    if (text.includes('processor') || text.includes('fpga') || text.includes('large')) return 'processor';
    if (text.includes('memory')) return 'memory';
    return 'chip';
  }

  function text(value, fallback = '—') {
    const raw = value == null ? '' : String(value).trim();
    return raw || fallback;
  }

  function escapeHtml(value) {
    return text(value, '').replace(/[&<>"']/g, (char) => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;',
    }[char]));
  }

  function listText(values, max = 5) {
    const items = Array.isArray(values) ? values.filter(Boolean) : [];
    if (!items.length) return '—';
    const shown = items.slice(0, max).join(', ');
    return items.length > max ? `${shown} …` : shown;
  }

  function shortText(value, max = 22) {
    const raw = text(value, '');
    return raw.length > max ? `${raw.slice(0, max - 1)}…` : raw;
  }

  function debugTopologyFixture() {
    return {
      ok: true,
      project_name: 'Debug Topology Fixture',
      topology: {
        schema_version: 'llm-topology.v1',
        summary: 'Debug fixture：展示芯片、连接器、电平转换和供电关系的拓扑视图。',
        counts: {
          total_node_count: 6,
          returned_node_count: 6,
          total_signal_edge_count: 5,
          returned_signal_edge_count: 5,
          total_supply_edge_count: 2,
          returned_supply_edge_count: 2,
          total_supply_group_count: 1,
          returned_supply_group_count: 1,
          visual_edge_count: 8,
        },
        summary_layer: {
          hubs: [{ refdes: 'U46', role: 'processor_or_fpga', connected_edge_count: 4 }],
          risk_edges: [],
        },
        business_view: {
          review_queue: [
            { item_id: 'review-i2c', title: 'I2C 跨电平转换', summary: '检查两侧上拉电压、OE 默认态和电平转换供电。', review_priority: 'high' },
            { item_id: 'review-clock', title: 'Clock 扇出链路', summary: '检查串阻/端接、REFCLK 命名一致性和页码定位。', review_priority: 'medium' },
          ],
        },
        nodes: [
          { refdes: 'U46', role: 'processor_or_fpga', user_visible_page: 'PAGE131', pin_count: 486, signal_net_count: 120, interface_groups: ['pcie', 'i2c', 'clock'], power_nets: ['P1V8', 'P3V3'], review_priority: 'high' },
          { refdes: 'U12', role: 'level_shifter', user_visible_page: 'PAGE88', pin_count: 20, signal_net_count: 8, interface_groups: ['i2c'], power_nets: ['P1V8', 'P3V3'], review_priority: 'high' },
          { refdes: 'PU3', role: 'power_management_ic', user_visible_page: 'PAGE42', pin_count: 32, signal_net_count: 6, interface_groups: ['power_control'], power_nets: ['VBAT'], review_priority: 'medium' },
          { refdes: 'J8', role: 'connector', user_visible_page: 'PAGE155', pin_count: 40, signal_net_count: 22, interface_groups: ['pcie', 'reset'], power_nets: ['P3V3'], review_priority: 'medium' },
          { refdes: 'U30', role: 'memory', user_visible_page: 'PAGE120', pin_count: 96, signal_net_count: 32, interface_groups: ['ddr'], power_nets: ['P1V2'], review_priority: 'medium' },
          { refdes: 'U61', role: 'clock_source', user_visible_page: 'PAGE98', pin_count: 10, signal_net_count: 4, interface_groups: ['clock'], power_nets: ['P3V3'], review_priority: 'low' },
        ],
        edges: [
          { edge_id: 'chip-edge-U46-U12', source_refdes: 'U46', target_refdes: 'U12', source_role: 'processor_or_fpga', target_role: 'level_shifter', interface_groups: ['i2c'], shared_net_count: 2, review_priority: 'high', review_score: 80, shared_nets: [{ net: 'I2C_SCL' }, { net: 'I2C_SDA' }], review_hints: ['检查两侧上拉电压', '检查 OE/EN 默认态'], summary: 'U46 与 U12 共享 I2C_SCL/I2C_SDA。' },
          { edge_id: 'chip-edge-U46-J8', source_refdes: 'U46', target_refdes: 'J8', source_role: 'processor_or_fpga', target_role: 'connector', interface_groups: ['pcie'], shared_net_count: 8, review_priority: 'high', review_score: 75, shared_nets: [{ net: 'PCE_TX0_P' }, { net: 'PCE_TX0_N' }], review_hints: ['检查 AC 耦合和差分极性', '检查 REFCLK/PERST#'], summary: 'U46 与 J8 形成 PCIe 连接。' },
          { edge_id: 'chip-edge-U46-U30', source_refdes: 'U46', target_refdes: 'U30', source_role: 'processor_or_fpga', target_role: 'memory', interface_groups: ['ddr'], shared_net_count: 24, review_priority: 'medium', review_score: 55, shared_nets: [{ net: 'DDR_DQ0' }], review_hints: ['检查 DDR 电压域和终端'], summary: 'U46 与 U30 共享 DDR 网络。' },
          { edge_id: 'chip-edge-U46-U61', source_refdes: 'U46', target_refdes: 'U61', source_role: 'processor_or_fpga', target_role: 'clock_source', interface_groups: ['clock'], shared_net_count: 1, review_priority: 'medium', review_score: 45, shared_nets: [{ net: 'REFCLK_100M' }], review_hints: ['检查串阻/端接和时钟扇出'], summary: 'U61 给 U46 提供参考时钟。' },
          { edge_id: 'chip-edge-U12-J8', source_refdes: 'U12', target_refdes: 'J8', source_role: 'level_shifter', target_role: 'connector', interface_groups: ['gpio'], shared_net_count: 2, review_priority: 'low', review_score: 20, shared_nets: [{ net: 'GPIO_EXP_INT' }], review_hints: ['检查默认态和外部接口保护'], summary: 'U12 与 J8 共享 GPIO 信号。' },
        ],
        supply_edges: [
          { edge_id: 'supply-edge-PU3-U46-P3V3', edge_kind: 'supply', source_refdes: 'PU3', target_refdes: 'U46', supply_net: 'P3V3', relation_label: '电源管理供电关系', review_priority: 'medium', review_hints: ['检查负载电流', '检查上电时序'], summary: 'PU3 通过 P3V3 给 U46 供电。' },
          { edge_id: 'supply-edge-PU3-U12-P1V8', edge_kind: 'supply', source_refdes: 'PU3', target_refdes: 'U12', supply_net: 'P1V8', relation_label: '电源管理供电关系', review_priority: 'medium', review_hints: ['检查电平转换两侧供电'], summary: 'PU3 通过 P1V8 给 U12 供电。' },
        ],
        supply_edge_groups: [
          { group_id: 'supply-group-PU3-P3V3', edge_kind: 'supply_group', source_refdes: 'PU3', supply_net: 'P3V3', voltage_domain: '3V3', target_count: 42, target_refdes_list: ['U46', 'U12', 'U30', 'U61', 'J8'], sample_target_refdes: 'U46', target_roles: [{ role: 'processor_or_fpga', count: 1 }, { role: 'peripheral_ic', count: 41 }], relation_label: '电源管理供电关系聚合', review_priority: 'medium', review_hints: ['这是供电关系聚合视图；切换全量模式查看每个负载。'], summary: 'PU3 通过 P3V3 给 42 个芯片/连接器节点提供供电关系。' },
        ],
      },
    };
  }

  function metricNode(label, value, hint = '') {
    const node = document.createElement('div');
    node.className = 'topology-metric';
    node.innerHTML = `<span>${escapeHtml(label)}</span><strong>${escapeHtml(value ?? '—')}</strong>${hint ? `<em>${escapeHtml(hint)}</em>` : ''}`;
    return node;
  }

  function renderMetrics(topology) {
    const host = document.getElementById('topology-metrics');
    if (!host) return;
    const counts = topology?.counts || {};
    const supplyGroups = counts.returned_supply_group_count ?? topology?.supply_edge_groups?.length ?? 0;
    const supplyHint = supplyGroups
      ? `组 ${supplyGroups} / 总 ${counts.total_supply_edge_count ?? topology?.supply_edge_count ?? 0}`
      : `总 ${counts.total_supply_edge_count ?? topology?.supply_edge_count ?? 0}`;
    host.replaceChildren(
      metricNode('节点', counts.returned_node_count ?? topology?.nodes?.length ?? 0, `总 ${counts.total_node_count ?? topology?.node_count ?? 0}`),
      metricNode('信号关系', counts.returned_signal_edge_count ?? topology?.edges?.length ?? 0, `总 ${counts.total_signal_edge_count ?? topology?.edge_count ?? 0}`),
      metricNode('供电关系', counts.returned_supply_edge_count ?? topology?.supply_edges?.length ?? 0, supplyHint),
      metricNode('状态', topology?.truncated ? '已截断' : '完整预览', topology?.include_connectors ? '含连接器' : '不含连接器'),
    );
  }

  function nodeRadius(degree, nodeCount) {
    if (nodeCount > 96) return Math.max(14, Math.min(22, 14 + Math.sqrt(degree || 0) * 2.2));
    if (nodeCount > 60) return Math.max(16, Math.min(26, 16 + Math.sqrt(degree || 0) * 2.8));
    return Math.max(22, Math.min(40, 22 + Math.sqrt(degree || 0) * 4));
  }

  function buildLayout(nodes, edges, width, height) {
    const degree = new Map(nodes.map((node) => [node.refdes, 0]));
    edges.forEach((edge) => {
      degree.set(edge.source_refdes, (degree.get(edge.source_refdes) || 0) + 1);
      degree.set(edge.target_refdes, (degree.get(edge.target_refdes) || 0) + 1);
    });
    const ordered = [...nodes].sort((a, b) => (degree.get(b.refdes) || 0) - (degree.get(a.refdes) || 0) || a.refdes.localeCompare(b.refdes));
    const positions = new Map();
    const cx = width / 2;
    const cy = height / 2;
    if (!ordered.length) return positions;
    const rotation = state.layoutSeed * 0.47;
    const assigned = new Set();
    const bucket = (name) => ordered.filter((node) => roleClass(node.role) === name && !assigned.has(node.refdes));
    const put = (node, x, y) => {
      positions.set(node.refdes, { x, y, degree: degree.get(node.refdes) || 0 });
      assigned.add(node.refdes);
    };
    const margin = nodes.length > 72 ? 44 : 64;
    const zoneGap = nodes.length > 72 ? 18 : 28;
    const hub = bucket('processor')[0] || ordered[0];
    put(hub, cx, cy);

    const placeArc = (items, start, end, rx, ry) => {
      items.forEach((node, index) => {
        const t = items.length === 1 ? 0.5 : index / Math.max(1, items.length - 1);
        const angle = start + (end - start) * t + rotation * 0.08;
        put(node, cx + Math.cos(angle) * rx, cy + Math.sin(angle) * ry);
      });
    };

    const placeZone = (items, zone) => {
      if (!items.length) return;
      if (items.length <= 6) {
        placeArc(items, zone.start, zone.end, zone.rx, zone.ry);
        return;
      }
      const columns = Math.max(1, Math.ceil(Math.sqrt((items.length * zone.w) / Math.max(1, zone.h))));
      const rows = Math.max(1, Math.ceil(items.length / columns));
      const cellW = zone.w / columns;
      const cellH = zone.h / rows;
      items.forEach((node, index) => {
        const row = Math.floor(index / columns);
        const col = index % columns;
        const snakeCol = row % 2 ? columns - 1 - col : col;
        const jitter = ((index % 3) - 1) * Math.min(8, cellW * 0.08);
        put(
          node,
          zone.x + cellW * (snakeCol + 0.5) + jitter,
          zone.y + cellH * (row + 0.5),
        );
      });
    };

    const leftW = Math.max(210, width * 0.28);
    const rightW = Math.max(210, width * 0.26);
    placeZone(bucket('power'), {
      x: margin,
      y: margin,
      w: leftW,
      h: height - margin * 2,
      start: Math.PI * 0.70,
      end: Math.PI * 1.30,
      rx: width * 0.34,
      ry: height * 0.34,
    });
    placeZone(bucket('level'), {
      x: Math.max(margin, width * 0.22),
      y: Math.max(margin, height * 0.60),
      w: Math.max(200, width * 0.30),
      h: Math.max(120, height * 0.28),
      start: Math.PI * 1.20,
      end: Math.PI * 1.72,
      rx: width * 0.30,
      ry: height * 0.36,
    });
    placeZone(bucket('memory'), {
      x: Math.min(width - rightW - margin, width * 0.55),
      y: margin,
      w: rightW,
      h: Math.max(130, height * 0.34),
      start: Math.PI * 0.24,
      end: Math.PI * 0.70,
      rx: width * 0.34,
      ry: height * 0.36,
    });
    placeZone(bucket('connector'), {
      x: width - rightW - margin,
      y: Math.max(margin, height * 0.32),
      w: rightW,
      h: Math.max(160, height * 0.44),
      start: Math.PI * -0.24,
      end: Math.PI * 0.24,
      rx: width * 0.40,
      ry: height * 0.40,
    });

    const extraProcessors = bucket('processor');
    placeZone(extraProcessors, {
      x: Math.max(margin, cx - width * 0.15),
      y: Math.max(margin, cy - height * 0.30),
      w: Math.max(180, width * 0.30),
      h: Math.max(120, height * 0.22),
      start: Math.PI * 1.78,
      end: Math.PI * 2.22,
      rx: width * 0.22,
      ry: height * 0.26,
    });

    const rest = ordered.filter((node) => !assigned.has(node.refdes));
    placeZone(rest, {
      x: margin + leftW + zoneGap,
      y: margin,
      w: Math.max(180, width - leftW - rightW - margin * 2 - zoneGap * 2),
      h: height - margin * 2,
      start: 0,
      end: Math.PI * 2,
      rx: Math.max(170, Math.min(width, height) * 0.38),
      ry: Math.max(170, Math.min(width, height) * 0.38),
    });

    const minDistance = nodes.length > 96 ? 31 : nodes.length > 60 ? 38 : 54;
    const list = ordered.filter((node) => positions.has(node.refdes));
    for (let iteration = 0; iteration < 8; iteration += 1) {
      for (let i = 0; i < list.length; i += 1) {
        for (let j = i + 1; j < list.length; j += 1) {
          const a = positions.get(list[i].refdes);
          const b = positions.get(list[j].refdes);
          if (!a || !b) continue;
          const dx = b.x - a.x;
          const dy = b.y - a.y;
          const distance = Math.hypot(dx, dy) || 0.001;
          if (distance >= minDistance) continue;
          const push = (minDistance - distance) / 2;
          const ux = dx / distance;
          const uy = dy / distance;
          a.x -= ux * push;
          a.y -= uy * push;
          b.x += ux * push;
          b.y += uy * push;
        }
      }
      list.forEach((node) => {
        const pos = positions.get(node.refdes);
        const radius = nodeRadius(pos?.degree || 0, nodes.length) + 14;
        pos.x = Math.max(radius, Math.min(width - radius, pos.x));
        pos.y = Math.max(radius, Math.min(height - radius, pos.y));
      });
    }
    return positions;
  }

  function supplyGroupTarget(group, positions) {
    const refs = Array.isArray(group.target_refdes_list) ? group.target_refdes_list : [];
    const points = refs.map((ref) => positions.get(ref)).filter(Boolean);
    if (!points.length && group.sample_target_refdes) {
      const sample = positions.get(group.sample_target_refdes);
      if (sample) points.push(sample);
    }
    if (!points.length) return null;
    return {
      x: points.reduce((sum, point) => sum + point.x, 0) / points.length,
      y: points.reduce((sum, point) => sum + point.y, 0) / points.length,
      degree: points.length,
    };
  }

  function shouldShowEdgeLabel(edge, index, total, nodeCount) {
    if (state.edgeLabelMode === 'off') return false;
    if (state.edgeLabelMode === 'all') return true;
    if (nodeCount > 60) {
      if (edge.edge_kind === 'supply_group') return index < 4;
      return edge.review_priority === 'high' && index < 8;
    }
    if (edge.review_priority === 'high') return true;
    if (edge.edge_kind === 'supply_group') return index < 8;
    return total <= 36 && index < 18;
  }

  function svgEl(name, attrs = {}) {
    const node = document.createElementNS(SVG_NS, name);
    Object.entries(attrs).forEach(([key, value]) => node.setAttribute(key, String(value)));
    return node;
  }

  function renderGraph(topology) {
    const host = document.getElementById('topology-graph');
    const detail = document.getElementById('topology-detail');
    if (!host) return;
    const nodes = Array.isArray(topology?.nodes) ? topology.nodes : [];
    const signalEdges = (Array.isArray(topology?.edges) ? topology.edges : []).map((edge) => ({ ...edge, edge_kind: edge.edge_kind || 'signal' }));
    const supplyEdges = (Array.isArray(topology?.supply_edges) ? topology.supply_edges : []).map((edge) => ({ ...edge, edge_kind: 'supply' }));
    const supplyGroups = (Array.isArray(topology?.supply_edge_groups) ? topology.supply_edge_groups : []).map((edge) => ({ ...edge, edge_kind: 'supply_group' }));
    const edges = [...signalEdges, ...supplyEdges, ...supplyGroups].filter((edge) => edge.source_refdes && (edge.target_refdes || edge.sample_target_refdes || edge.target_refdes_list));
    host.replaceChildren();
    if (!nodes.length) {
      host.innerHTML = '<p class="topology-empty">没有可显示的芯片/连接器拓扑节点。可以尝试勾选“显示连接器”或放宽过滤条件。</p>';
      return;
    }

    const width = Math.max(780, host.clientWidth || 960);
    const height = Math.max(560, Math.min(760, Math.round(width * 0.58)));
    const positions = buildLayout(nodes, edges, width, height);
    const denseGraph = nodes.length > 60;
    const svg = svgEl('svg', { viewBox: `0 0 ${width} ${height}`, class: `topology-svg ${denseGraph ? 'is-dense' : ''}`, role: 'img' });
    const edgeLayer = svgEl('g', { class: 'topology-edge-layer' });
    const nodeLayer = svgEl('g', { class: 'topology-node-layer' });

    edges.forEach((edge, index) => {
      const source = positions.get(edge.source_refdes);
      const target = edge.edge_kind === 'supply_group' ? supplyGroupTarget(edge, positions) : positions.get(edge.target_refdes);
      if (!source || !target) return;
      const line = svgEl('path', {
        d: `M ${source.x} ${source.y} C ${(source.x + target.x) / 2} ${source.y}, ${(source.x + target.x) / 2} ${target.y}, ${target.x} ${target.y}`,
        class: `topology-edge ${edge.edge_kind === 'supply' || edge.edge_kind === 'supply_group' ? 'is-supply' : 'is-signal'} ${edge.edge_kind === 'supply_group' ? 'is-group' : ''} ${edge.review_priority === 'high' ? 'is-high' : ''}`,
        tabindex: 0,
        'data-edge-id': edge.edge_id || edge.group_id || '',
      });
      line.addEventListener('click', () => selectEdge(edge));
      line.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          selectEdge(edge);
        }
      });
      edgeLayer.appendChild(line);

      if (shouldShowEdgeLabel(edge, index, edges.length, nodes.length)) {
        const labelText = edge.edge_kind === 'supply_group'
          ? `${text(edge.supply_net, 'SUPPLY')} x${text(edge.target_count, '?')}`
          : edge.edge_kind === 'supply'
            ? text(edge.supply_net, 'SUPPLY')
            : listText(edge.interface_groups, 2);
        const label = svgEl('text', {
          x: (source.x + target.x) / 2,
          y: (source.y + target.y) / 2 - 8,
          class: 'topology-edge-label',
          'text-anchor': 'middle',
        });
        label.textContent = labelText;
        edgeLayer.appendChild(label);
      }
    });

    let denseLabelCount = 0;
    nodes.forEach((node) => {
      const pos = positions.get(node.refdes);
      if (!pos) return;
      const group = svgEl('g', {
        class: `topology-node is-${roleClass(node.role)} ${node.review_priority === 'high' ? 'is-high' : ''}`,
        transform: `translate(${pos.x} ${pos.y})`,
        tabindex: 0,
        'data-refdes': node.refdes,
      });
      const radius = nodeRadius(pos.degree || 0, nodes.length);
      group.appendChild(svgEl('circle', { r: radius, class: 'topology-node-orb' }));
      const isHubLike = roleClass(node.role) === 'processor' || pos.degree >= 5;
      const isImportant = pos.degree >= 3 || node.review_priority === 'high';
      const showLabel = !denseGraph || isHubLike || (isImportant && denseLabelCount < 20);
      const title = svgEl('title');
      title.textContent = `${text(node.refdes)} · ${text(node.role, 'chip')}`;
      group.appendChild(title);
      if (showLabel) {
        if (denseGraph && !isHubLike) denseLabelCount += 1;
        const label = svgEl('text', { y: 5, class: 'topology-node-label', 'text-anchor': 'middle' });
        label.textContent = node.refdes;
        group.appendChild(label);
      }
      if (!denseGraph && showLabel) {
        const sub = svgEl('text', { y: radius + 18, class: 'topology-node-sub', 'text-anchor': 'middle' });
        sub.textContent = shortText(node.role, 20) || 'chip';
        group.appendChild(sub);
      }
      group.addEventListener('click', () => selectNode(node));
      group.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          selectNode(node);
        }
      });
      nodeLayer.appendChild(group);
    });

    svg.appendChild(edgeLayer);
    svg.appendChild(nodeLayer);
    host.appendChild(svg);
    if (!state.selected && detail) {
      const hub = [...nodes].sort((a, b) => (b.connected_edge_count || 0) - (a.connected_edge_count || 0))[0];
      if (hub) selectNode(hub);
    }
  }

  function rows(items) {
    return `<div class="topology-detail-rows">${items.map((item) => `
      <div><span>${escapeHtml(item[0])}</span><strong>${escapeHtml(item[1] || '—')}</strong></div>
    `).join('')}</div>`;
  }

  function selectNode(node) {
    state.selected = { type: 'node', id: node.refdes };
    const host = document.getElementById('topology-detail');
    if (!host) return;
    host.innerHTML = `
      <div class="topology-detail-title">
        <span class="topology-detail-badge ${roleClass(node.role)}">${escapeHtml(text(node.role, 'chip'))}</span>
        <h3>${escapeHtml(text(node.refdes))}</h3>
      </div>
      ${rows([
        ['页码', text(node.user_visible_page || node['页码'])],
        ['HQ 料号', text(node.hq_no)],
        ['规格 / Value', text(node.spec || node.value)],
        ['Pin 数', text(node.pin_count)],
        ['信号网络', text(node.signal_net_count)],
        ['接口组', listText(node.interface_groups, 8)],
        ['电源网', listText(node.power_nets, 8)],
      ])}
      <div class="topology-hints">
        <strong>风险标签</strong>
        <p>${escapeHtml(listText(node.risk_tags, 8))}</p>
      </div>
    `;
  }

  function selectEdge(edge) {
    state.selected = { type: 'edge', id: edge.edge_id || edge.group_id };
    const host = document.getElementById('topology-detail');
    if (!host) return;
    const isSupplyGroup = edge.edge_kind === 'supply_group';
    const netPreview = edge.edge_kind === 'supply' || isSupplyGroup
      ? text(edge.supply_net)
      : listText((edge.shared_nets || []).map((item) => item.net), 10);
    const targetTitle = isSupplyGroup
      ? `${text(edge.target_count, '?')} 个负载`
      : text(edge.target_refdes);
    host.innerHTML = `
      <div class="topology-detail-title">
        <span class="topology-detail-badge ${edge.edge_kind === 'supply' || isSupplyGroup ? 'power' : 'chip'}">${isSupplyGroup ? '供电组' : edge.edge_kind === 'supply' ? '供电' : '信号'}</span>
        <h3>${escapeHtml(text(edge.source_refdes))} ↔ ${escapeHtml(targetTitle)}</h3>
      </div>
      <p class="topology-detail-summary">${escapeHtml(text(edge.summary))}</p>
      ${rows([
        ['关系类型', text(edge.relation_label)],
        ['接口组', listText(edge.interface_groups, 6)],
        ['网络样本', netPreview],
        ['共享网络数', isSupplyGroup ? text(edge.target_count) : text(edge.shared_net_count)],
        ['目标角色', isSupplyGroup ? listText((edge.target_roles || []).map((item) => `${item.role}:${item.count}`), 8) : '—'],
        ['样本目标', isSupplyGroup ? listText(edge.target_refdes_list, 12) : text(edge.target_refdes)],
        ['优先级', text(edge.review_priority)],
        ['置信度', text(edge.confidence)],
      ])}
      <div class="topology-hints">
        <strong>Review hints</strong>
        <ul>${(edge.review_hints || edge.review_focus || []).slice(0, 8).map((item) => `<li>${escapeHtml(item)}</li>`).join('') || '<li>暂无。</li>'}</ul>
      </div>
    `;
  }

  function renderReviewQueue(topology) {
    const host = document.getElementById('topology-review-queue');
    if (!host) return;
    const items = topology?.business_view?.review_queue || [];
    if (!items.length) {
      host.innerHTML = '<p class="empty-state">暂无高优先级拓扑复核建议。</p>';
      return;
    }
    host.innerHTML = items.slice(0, 8).map((item) => `
      <button type="button" class="topology-review-item">
        <span>${escapeHtml(text(item.review_priority, 'low'))}</span>
        <strong>${escapeHtml(text(item.title || item.item_id))}</strong>
        <em>${escapeHtml(text(item.summary))}</em>
      </button>
    `).join('');
  }

  async function fetchTopology() {
    const form = document.getElementById('topology-controls');
    const status = document.getElementById('topology-status');
    state.edgeLabelMode = document.getElementById('topology-edge-label-mode')?.value || 'auto';
    if (context.debugFixture || document.body?.dataset.debugFixture === 'true') {
      return debugTopologyFixture();
    }
    const view = document.getElementById('topology-view')?.value || 'summary';
    const params = new URLSearchParams({
      include_connectors: document.getElementById('topology-include-connectors')?.checked ? '1' : '0',
      focus_refdes: document.getElementById('topology-focus-refdes')?.value || '',
      role_filter: document.getElementById('topology-role-filter')?.value || '',
      limit: document.getElementById('topology-limit')?.value || '120',
      view,
      supply_mode: document.getElementById('topology-supply-mode')?.value || 'grouped',
      supply_limit: '12',
      edge_label_mode: state.edgeLabelMode,
    });
    const response = await fetch(`/api/report/${context.runId}/topology?${params.toString()}`);
    const payload = await response.json();
    if (!response.ok || payload.ok === false) {
      throw new Error(payload.error || `拓扑读取失败：${response.status}`);
    }
    if (form) form.dataset.lastQuery = params.toString();
    if (status) status.textContent = '拓扑已生成。';
    return payload;
  }

  async function loadAndRender() {
    const status = document.getElementById('topology-status');
    if (status) status.textContent = '正在生成芯片级拓扑…';
    try {
      const payload = await fetchTopology();
      state.topology = payload.topology || {};
      renderMetrics(state.topology);
      renderGraph(state.topology);
      renderReviewQueue(state.topology);
      if (status) status.textContent = `${payload.project_name || context.projectName || '项目'}：${state.topology.summary || '拓扑生成完成。'}`;
    } catch (error) {
      if (status) status.textContent = error.message || String(error);
      const host = document.getElementById('topology-graph');
      if (host) host.innerHTML = `<p class="topology-empty topology-error">${escapeHtml(error.message || error)}</p>`;
    }
  }

  function boot() {
    const form = document.getElementById('topology-controls');
    const redraw = document.getElementById('topology-redraw');
    form?.addEventListener('submit', (event) => {
      event.preventDefault();
      state.selected = null;
      loadAndRender();
    });
    redraw?.addEventListener('click', () => {
      state.layoutSeed += 1;
      renderGraph(state.topology);
    });
    loadAndRender();
  }

  document.addEventListener('DOMContentLoaded', () => {
    window.PSTXApp?.bootCommon?.();
    boot();
  });
}());
