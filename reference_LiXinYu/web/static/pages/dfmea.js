document.addEventListener('DOMContentLoaded', () => {
  window.PSTXApp?.bootCommon?.();

  const context = window.PSTX_DFMEA_CONTEXT || {};
  const runId = context.runId || document.body.dataset.runId || '';
  const debugFixture = Boolean(context.debugFixture || document.body.dataset.debugFixture === 'true');
  if (!runId) return;

  const filterForm = document.getElementById('dfmea-filter-form');
  const groupForm = document.getElementById('dfmea-group-form');
  const pendingHost = document.getElementById('dfmea-pending-list');
  const groupsHost = document.getElementById('dfmea-groups');
  const pendingCount = document.getElementById('dfmea-pending-count');
  const selectedCount = document.getElementById('dfmea-selected-count');
  const clearSelection = document.getElementById('dfmea-clear-selection');
  const cancelEdit = document.getElementById('dfmea-cancel-edit');
  const statusNode = document.getElementById('dfmea-form-status');

  let state = {
    pending: [],
    groups: [],
    selected: new Set(),
    editingGroup: null,
    debugNextGroupId: 3,
    tableSort: { key: 'page', direction: 'asc' },
    tableFilters: {},
    renderLimit: 240,
    displayRows: [],
  };
  let pendingTableShell = null;
  let pendingDelegationBound = false;

  const RENDER_BATCH_SIZE = 240;
  const RENDER_BATCH_GROWTH = 360;

  const PENDING_COLUMNS = [
    { key: 'select', label: '选择', sortable: false },
    { key: 'refdes', label: '位号' },
    { key: 'page', label: '页码' },
    { key: 'refdes_type', label: '类型' },
    { key: 'hq_no', label: 'HQ料号' },
    { key: 'value', label: '值/规格' },
    { key: 'package', label: '封装' },
    { key: 'bom_option', label: 'BOM_OPTION' },
  ];

  const DEBUG_PENDING = [
    {
      refdes: 'U46',
      page: 'PAGE130, PAGE131',
      refdes_type: 'U',
      category: 'large_ic',
      hq_no: 'HQ11112042009',
      value: 'LCMXO3LF-9400C-5BG484C',
      package: 'BGA484',
      bom_option: '',
      is_depop: false,
    },
    {
      refdes: 'PU12',
      page: 'PAGE42',
      refdes_type: 'PU',
      category: 'power_ic',
      hq_no: 'HQ2200DEBUG',
      value: 'Buck Regulator',
      package: 'QFN',
      bom_option: '',
      is_depop: false,
    },
    {
      refdes: 'R120',
      page: 'PAGE42',
      refdes_type: 'R',
      category: 'passive',
      hq_no: '',
      value: '10k',
      package: '0402',
      bom_option: '',
      is_depop: false,
    },
    {
      refdes: 'C88',
      page: 'PAGE43',
      refdes_type: 'C',
      category: 'passive',
      hq_no: 'HQ1710101A3Q0',
      value: '0.5pF',
      package: '0201',
      bom_option: 'DEPOP',
      is_depop: true,
    },
    {
      refdes: 'J5',
      page: 'PAGE88',
      refdes_type: 'J',
      category: 'connector',
      hq_no: 'HQCONNDEBUG',
      value: 'Board Connector',
      package: 'CONN',
      bom_option: '',
      is_depop: false,
    },
  ];

  const DEBUG_GROUPS = [
    {
      id: 1,
      refdes: ['U46', 'PU12'],
      refdes_text: 'U46, PU12',
      pages: ['PAGE42', 'PAGE130', 'PAGE131'],
      pages_text: 'PAGE42, PAGE130, PAGE131',
      function_requirement: '完成主控逻辑与电源转换，维持关键电源时序。',
      failure_mode: '输出异常、通信失效、配置加载失败。',
      failure_effect: '系统无法启动或关键接口不可用。',
      failure_cause: '供电裕量不足、焊接异常、配置链路异常。',
      prevention_detection: '降额检查、ICT/FCT、上电时序测试。',
      updated_at: '2026-04-29T04:00:00',
      components: [],
    },
    {
      id: 2,
      refdes: ['J5'],
      refdes_text: 'J5',
      pages: ['PAGE88'],
      pages_text: 'PAGE88',
      function_requirement: '',
      failure_mode: '',
      failure_effect: '',
      failure_cause: '',
      prevention_detection: '',
      updated_at: '2026-04-29T04:05:00',
      components: [],
    },
  ];

  DEBUG_GROUPS[0].components = DEBUG_PENDING.filter((row) => ['U46', 'PU12'].includes(row.refdes));
  DEBUG_GROUPS[1].components = DEBUG_PENDING.filter((row) => row.refdes === 'J5');

  function field(form, name) {
    return form?.querySelector(`[name="${name}"]`);
  }

  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, (char) => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;',
    }[char]));
  }

  function setStatus(message, tone = 'neutral') {
    if (!statusNode) return;
    statusNode.textContent = message;
    statusNode.style.color = tone === 'error' ? 'var(--warn)' : tone === 'ok' ? 'var(--ok)' : 'var(--muted)';
  }

  function selectedRefdes() {
    return Array.from(state.selected);
  }

  function sortKeyFromControl(value) {
    const mapping = {
      page: 'page',
      type: 'refdes_type',
      refdes: 'refdes',
      hq: 'hq_no',
    };
    return mapping[value] || 'page';
  }

  function syncTableSortFromControl() {
    const sortSelect = field(filterForm, 'sort');
    state.tableSort = {
      key: sortKeyFromControl(sortSelect?.value || 'page'),
      direction: 'asc',
    };
  }

  function resetPendingRenderLimit() {
    state.renderLimit = RENDER_BATCH_SIZE;
  }

  function destroyPendingTableShell(message = '') {
    pendingTableShell = null;
    if (pendingHost) {
      pendingHost.innerHTML = message ? `<p class="query-empty">${escapeHtml(message)}</p>` : '';
    }
  }

  function setRefdesDisplay(refs) {
    const target = field(groupForm, 'refdes_text');
    if (!target) return;
    const text = refs.join(', ');
    target.value = text;
    target.textContent = text || '从左侧待排查表格勾选生成';
  }

  function updateSelectionUi() {
    const refs = selectedRefdes();
    if (selectedCount) selectedCount.textContent = `已选择：${refs.length}`;
    setRefdesDisplay(refs);
    document.querySelectorAll('.dfmea-pending-table input[type="checkbox"][data-refdes]').forEach((input) => {
      input.checked = state.selected.has(input.value);
    });
    if (!state.editingGroup) {
      setStatus(refs.length ? `准备保存 ${refs.length} 个位号为一个分组。` : '请选择待排查元器件。');
    }
  }

  function resetForm() {
    state.editingGroup = null;
    groupForm?.reset();
    state.selected = new Set();
    if (field(groupForm, 'group_id')) field(groupForm, 'group_id').value = '';
    if (cancelEdit) cancelEdit.hidden = true;
    updateSelectionUi();
    renderPending();
  }

  function debugVisiblePending() {
    const params = new URLSearchParams(new FormData(filterForm || undefined));
    const includeDepop = field(filterForm, 'include_depop')?.checked;
    const excludeRc = field(filterForm, 'exclude_rc')?.checked;
    const query = String(params.get('q') || '').trim().toLowerCase();
    const grouped = new Set(state.groups.flatMap((group) => group.refdes || []));
    return DEBUG_PENDING
      .filter((row) => !grouped.has(row.refdes))
      .filter((row) => includeDepop || !row.is_depop)
      .filter((row) => !excludeRc || !['R', 'C'].includes(String(row.refdes_type || '').toUpperCase()))
      .filter((row) => {
        if (!query) return true;
        return [row.refdes, row.page, row.hq_no, row.value, row.package, row.bom_option]
          .join(' ')
          .toLowerCase()
          .includes(query);
      });
  }

  function rowMatchesTableFilters(row) {
    return Object.entries(state.tableFilters).every(([key, value]) => {
      const query = String(value || '').trim().toLowerCase();
      if (!query) return true;
      return String(row[key] || '').toLowerCase().includes(query);
    });
  }

  function rowMatchesGlobalQuery(row) {
    const query = String(field(filterForm, 'q')?.value || '').trim().toLowerCase();
    if (!query) return true;
    return [row.refdes, row.page, row.refdes_type, row.category, row.hq_no, row.value, row.package, row.bom_option]
      .join(' ')
      .toLowerCase()
      .includes(query);
  }

  function compareRows(left, right, key) {
    if (key === 'page') {
      const leftNumber = Number(String(left.page || '').match(/\d+/)?.[0] || 999999);
      const rightNumber = Number(String(right.page || '').match(/\d+/)?.[0] || 999999);
      if (leftNumber !== rightNumber) return leftNumber - rightNumber;
    }
    return String(left[key] || '').localeCompare(String(right[key] || ''), 'zh-Hans-CN', { numeric: true, sensitivity: 'base' });
  }

  function pendingDisplayRows() {
    const rowsByRefdes = new Map();
    state.pending.forEach((row) => rowsByRefdes.set(String(row.refdes || '').toUpperCase(), { ...row, source: 'pending' }));
    if (state.editingGroup) {
      const groupRows = state.editingGroup.components?.length
        ? state.editingGroup.components
        : (state.editingGroup.refdes || []).map((refdes) => ({ refdes, page: state.editingGroup.pages_text || '', refdes_type: '', category: '', hq_no: '', value: '', package: '', bom_option: '', is_depop: false }));
      groupRows.forEach((row) => {
        rowsByRefdes.set(String(row.refdes || '').toUpperCase(), { ...row, source: 'editing', editing_member: true });
      });
    }
    let rows = Array.from(rowsByRefdes.values()).filter(rowMatchesGlobalQuery).filter(rowMatchesTableFilters);
    const { key, direction } = state.tableSort;
    rows.sort((left, right) => {
      const primary = compareRows(left, right, key || 'page');
      const fallback = key === 'page'
        ? compareRows(left, right, 'refdes_type') || compareRows(left, right, 'refdes')
        : compareRows(left, right, 'page') || compareRows(left, right, 'refdes');
      const value = primary || fallback;
      return direction === 'desc' ? -value : value;
    });
    return rows;
  }

  function ensurePendingTableShell() {
    if (!pendingHost) return null;
    if (pendingTableShell?.host === pendingHost && pendingHost.contains(pendingTableShell.tbody)) {
      return pendingTableShell;
    }
    const header = PENDING_COLUMNS.map((column) => {
      if (column.key === 'select') {
        return '<th class="dfmea-select-col">选择</th>';
      }
      return `<th><button type="button" class="dfmea-sort-btn" data-sort-key="${escapeHtml(column.key)}">${escapeHtml(column.label)}</button></th>`;
    }).join('');
    const filters = PENDING_COLUMNS.map((column) => {
      if (column.key === 'select') {
        return '<th><button type="button" class="ghost-btn dfmea-small-btn" data-action="select-visible">全选</button></th>';
      }
      return `<th><input class="dfmea-column-filter" data-filter-key="${escapeHtml(column.key)}" value="${escapeHtml(state.tableFilters[column.key] || '')}" placeholder="筛选${escapeHtml(column.label)}"></th>`;
    }).join('');
    pendingHost.innerHTML = `
      <div class="dfmea-table-wrap">
        <table class="dfmea-pending-table">
          <thead>
            <tr>${header}</tr>
            <tr class="dfmea-filter-row">${filters}</tr>
          </thead>
          <tbody></tbody>
        </table>
        <div class="dfmea-table-more" data-role="pending-more" hidden></div>
      </div>
    `;
    if (!pendingDelegationBound) {
      pendingHost.addEventListener('change', onPendingHostChange);
      pendingHost.addEventListener('click', onPendingHostClick);
      pendingHost.addEventListener('input', onPendingHostInput);
      pendingDelegationBound = true;
    }
    pendingTableShell = {
      host: pendingHost,
      tbody: pendingHost.querySelector('tbody'),
      more: pendingHost.querySelector('[data-role="pending-more"]'),
      filters: pendingHost.querySelectorAll('[data-filter-key]'),
      sortButtons: pendingHost.querySelectorAll('[data-sort-key]'),
    };
    return pendingTableShell;
  }

  function rowHtml(row) {
    const editing = row.editing_member ? '<span class="dfmea-tag">编辑中</span>' : '';
    const depop = row.is_depop ? '<span class="dfmea-tag is-warn">DEPOP</span>' : '';
    return `
      <tr class="${row.editing_member ? 'is-editing-member' : ''}">
        <td class="dfmea-select-col"><input type="checkbox" data-refdes="${escapeHtml(row.refdes)}" value="${escapeHtml(row.refdes)}"></td>
        <td><strong>${escapeHtml(row.refdes)}</strong> ${editing}</td>
        <td><span class="dfmea-component-page">${escapeHtml(row.page || '未识别页码')}</span></td>
        <td>${escapeHtml(row.refdes_type || row.category || '未知')}</td>
        <td>${escapeHtml(row.hq_no || '')}</td>
        <td>${escapeHtml(row.value || '')}</td>
        <td>${escapeHtml(row.package || '')}</td>
        <td>${depop || escapeHtml(row.bom_option || '')}</td>
      </tr>
    `;
  }

  function updatePendingSortButtons(shell) {
    shell.sortButtons.forEach((button) => {
      const active = state.tableSort.key === button.dataset.sortKey;
      button.classList.toggle('is-active', active);
      button.classList.toggle('is-asc', active && state.tableSort.direction !== 'desc');
      button.classList.toggle('is-desc', active && state.tableSort.direction === 'desc');
    });
  }

  function updatePendingFilterInputs(shell) {
    shell.filters.forEach((input) => {
      const value = state.tableFilters[input.dataset.filterKey] || '';
      if (document.activeElement !== input && input.value !== value) {
        input.value = value;
      }
    });
  }

  function onPendingHostChange(event) {
    const input = event.target?.closest?.('input[type="checkbox"][data-refdes]');
    if (!input) return;
    if (input.checked) {
      state.selected.add(input.value);
    } else {
      state.selected.delete(input.value);
    }
    updateSelectionUi();
  }

  function onPendingHostClick(event) {
    const sortButton = event.target?.closest?.('[data-sort-key]');
    if (sortButton) {
      const key = sortButton.dataset.sortKey;
      state.tableSort = {
        key,
        direction: state.tableSort.key === key && state.tableSort.direction === 'asc' ? 'desc' : 'asc',
      };
      resetPendingRenderLimit();
      renderPending();
      return;
    }
    if (event.target?.closest?.('[data-action="select-visible"]')) {
      state.displayRows.forEach((row) => state.selected.add(row.refdes));
      updateSelectionUi();
      return;
    }
    if (event.target?.closest?.('[data-action="load-more"]')) {
      state.renderLimit = Math.min(state.displayRows.length, state.renderLimit + RENDER_BATCH_GROWTH);
      renderPending();
    }
  }

  function onPendingHostInput(event) {
    const input = event.target?.closest?.('[data-filter-key]');
    if (!input) return;
    state.tableFilters[input.dataset.filterKey] = input.value;
    resetPendingRenderLimit();
    window.clearTimeout(window.__pstxDfmeaColumnFilterTimer);
    window.__pstxDfmeaColumnFilterTimer = window.setTimeout(renderPending, 120);
  }

  function renderPending() {
    const shell = ensurePendingTableShell();
    if (!shell) return;
    const rows = pendingDisplayRows();
    state.displayRows = rows;
    const visibleRows = rows.slice(0, Math.max(1, state.renderLimit));
    if (pendingCount) {
      const renderedText = rows.length > visibleRows.length ? ` / 已渲染：${visibleRows.length}` : '';
      pendingCount.textContent = `待排查：${state.pending.length} / 当前显示：${rows.length}${renderedText}`;
    }
    shell.tbody.innerHTML = visibleRows.map(rowHtml).join('') || `
      <tr class="dfmea-empty-row">
        <td colspan="${PENDING_COLUMNS.length}">
          <p class="query-empty">当前筛选条件下没有待排查元器件。筛选栏已保留，可以直接修改条件。</p>
        </td>
      </tr>
    `;
    updatePendingSortButtons(shell);
    updatePendingFilterInputs(shell);
    shell.more.hidden = rows.length <= visibleRows.length;
    shell.more.innerHTML = rows.length > visibleRows.length
      ? `<button type="button" class="ghost-btn inline-btn dfmea-load-more" data-action="load-more">加载更多 ${Math.min(RENDER_BATCH_GROWTH, rows.length - visibleRows.length)} 条</button><span>剩余 ${rows.length - visibleRows.length} 条未渲染</span>`
      : '';
    updateSelectionUi();
  }

  function fillGroupForm(group) {
    state.editingGroup = group;
    state.selected = new Set(group.refdes || []);
    field(groupForm, 'group_id').value = group.id || '';
    field(groupForm, 'function_requirement').value = group.function_requirement || '';
    field(groupForm, 'failure_mode').value = group.failure_mode || '';
    field(groupForm, 'failure_effect').value = group.failure_effect || '';
    field(groupForm, 'failure_cause').value = group.failure_cause || '';
    field(groupForm, 'prevention_detection').value = group.prevention_detection || '';
    if (cancelEdit) cancelEdit.hidden = false;
    updateSelectionUi();
    renderPending();
    setStatus(`正在编辑组 ${group.id}，保存后会覆盖该组内容。`, 'ok');
    groupForm?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderGroups() {
    if (!groupsHost) return;
    if (!state.groups.length) {
      groupsHost.innerHTML = '<p class="query-empty">暂无已保存分组。</p>';
      return;
    }
    groupsHost.innerHTML = state.groups.map((group) => {
      const emptyCount = ['function_requirement', 'failure_mode', 'failure_effect', 'failure_cause', 'prevention_detection']
        .filter((key) => !String(group[key] || '').trim()).length;
      const emptyBadge = emptyCount ? `<span class="dfmea-tag is-warn">待补 ${emptyCount} 项</span>` : '<span class="dfmea-tag is-ok">已补齐</span>';
      return `
        <article class="dfmea-group-card is-collapsed" data-group-id="${escapeHtml(group.id)}">
          <div class="dfmea-group-head">
            <div>
              <p class="eyebrow">GROUP ${escapeHtml(group.id)}</p>
              <h3>${escapeHtml(group.refdes_text || '未命名分组')}</h3>
              <p>页码：${escapeHtml(group.pages_text || '未识别')}</p>
            </div>
            <div class="dfmea-group-head-actions">
              ${emptyBadge}
              <button type="button" class="ghost-btn inline-btn" data-action="edit">编辑</button>
              <button type="button" class="ghost-btn inline-btn" data-action="delete">删除</button>
              <button type="button" class="ghost-btn inline-btn" data-action="toggle">展开</button>
            </div>
          </div>
          <dl class="dfmea-group-fields" hidden>
            <div><dt>功能/需求</dt><dd>${escapeHtml(group.function_requirement || '未填写')}</dd></div>
            <div><dt>潜在失效模式</dt><dd>${escapeHtml(group.failure_mode || '未填写')}</dd></div>
            <div><dt>潜在失效后果</dt><dd>${escapeHtml(group.failure_effect || '未填写')}</dd></div>
            <div><dt>潜在失效原因/机理</dt><dd>${escapeHtml(group.failure_cause || '未填写')}</dd></div>
            <div><dt>现有预防/探测方案</dt><dd>${escapeHtml(group.prevention_detection || '未填写')}</dd></div>
          </dl>
          <div class="dfmea-group-actions" hidden>
            <button type="button" class="ghost-btn inline-btn" data-action="edit">编辑</button>
            <button type="button" class="ghost-btn inline-btn" data-action="delete">删除并退回待排查池</button>
            <span>更新时间：${escapeHtml(group.updated_at || '')}</span>
          </div>
        </article>
      `;
    }).join('');
    groupsHost.querySelectorAll('.dfmea-group-card').forEach((card) => {
      const id = Number(card.dataset.groupId);
      const fields = card.querySelector('.dfmea-group-fields');
      const actions = card.querySelector('.dfmea-group-actions');
      const toggle = card.querySelector('[data-action="toggle"]');
      toggle?.addEventListener('click', () => {
        const collapsed = card.classList.toggle('is-collapsed');
        card.classList.toggle('is-expanded', !collapsed);
        if (fields) fields.hidden = collapsed;
        if (actions) actions.hidden = collapsed;
        toggle.textContent = collapsed ? '展开' : '收起';
      });
      card.querySelectorAll('[data-action="edit"]').forEach((button) => button.addEventListener('click', () => {
        const group = state.groups.find((item) => Number(item.id) === id);
        if (group) fillGroupForm(group);
      }));
      card.querySelectorAll('[data-action="delete"]').forEach((button) => button.addEventListener('click', async () => {
        if (!window.confirm(`确认删除 DFMEA 组 ${id}？组内位号会回到待排查池。`)) return;
        await deleteGroup(id);
      }));
    });
  }

  async function loadWorkbench() {
    if (debugFixture) {
      if (!state.groups.length) state.groups = DEBUG_GROUPS.map((group) => ({ ...group, refdes: [...group.refdes], pages: [...group.pages] }));
      state.pending = debugVisiblePending();
      state.selected = new Set([...state.selected].filter((refdes) => pendingDisplayRows().some((row) => row.refdes === refdes) || state.groups.some((group) => (group.refdes || []).includes(refdes))));
      resetPendingRenderLimit();
      renderPending();
      renderGroups();
      updateSelectionUi();
      setStatus('Debug fixture：当前为模拟数据，不会写入 SQLite。', 'ok');
      return;
    }
    const params = new URLSearchParams(new FormData(filterForm || undefined));
    params.delete('q');
    params.delete('sort');
    params.set('include_depop', field(filterForm, 'include_depop')?.checked ? '1' : '0');
    params.set('exclude_rc', field(filterForm, 'exclude_rc')?.checked ? '1' : '0');
    destroyPendingTableShell('正在读取元器件…');
    const response = await fetch(`/api/report/${runId}/dfmea/workbench?${params.toString()}`);
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || 'DFMEA 工作台读取失败');
    }
    const editingId = state.editingGroup?.id;
    state.pending = payload.pending_components || [];
    state.groups = payload.groups || [];
    if (editingId) {
      state.editingGroup = state.groups.find((group) => Number(group.id) === Number(editingId)) || state.editingGroup;
    }
    state.selected = new Set([...state.selected].filter((refdes) => pendingDisplayRows().some((row) => row.refdes === refdes)));
    resetPendingRenderLimit();
    renderPending();
    renderGroups();
    updateSelectionUi();
  }

  async function saveGroup(event) {
    event.preventDefault();
    const refs = selectedRefdes();
    if (!refs.length) {
      setStatus('请至少选择一个位号后再保存。', 'error');
      return;
    }
    const payload = {
      refdes: refs,
      function_requirement: field(groupForm, 'function_requirement')?.value || '',
      failure_mode: field(groupForm, 'failure_mode')?.value || '',
      failure_effect: field(groupForm, 'failure_effect')?.value || '',
      failure_cause: field(groupForm, 'failure_cause')?.value || '',
      prevention_detection: field(groupForm, 'prevention_detection')?.value || '',
    };
    const groupId = field(groupForm, 'group_id')?.value;
    if (debugFixture) {
      const pages = Array.from(new Set(refs.flatMap((refdes) => {
        const found = DEBUG_PENDING.find((row) => row.refdes === refdes);
        return String(found?.page || '').split(',').map((item) => item.trim()).filter(Boolean);
      }))).sort();
      const now = new Date().toISOString().slice(0, 19);
      if (groupId) {
        const index = state.groups.findIndex((group) => String(group.id) === String(groupId));
        if (index >= 0) {
          state.groups[index] = {
            ...state.groups[index],
            ...payload,
            id: Number(groupId),
            refdes: refs,
            refdes_text: refs.join(', '),
            pages,
            pages_text: pages.join(', '),
            components: refs.map((refdes) => ({ ...(DEBUG_PENDING.find((row) => row.refdes === refdes) || { refdes }) })),
            updated_at: now,
          };
        }
      } else {
        const nextId = state.debugNextGroupId++;
        state.groups.push({
          ...payload,
          id: nextId,
          refdes: refs,
          refdes_text: refs.join(', '),
          pages,
          pages_text: pages.join(', '),
          components: refs.map((refdes) => ({ ...(DEBUG_PENDING.find((row) => row.refdes === refdes) || { refdes }) })),
          created_at: now,
          updated_at: now,
        });
      }
      resetForm();
      await loadWorkbench();
      return;
    }
    const url = groupId
      ? `/api/report/${runId}/dfmea/groups/${encodeURIComponent(groupId)}`
      : `/api/report/${runId}/dfmea/groups`;
    const response = await fetch(url, {
      method: groupId ? 'PATCH' : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || '保存 DFMEA 分组失败');
    }
    setStatus(groupId ? `已更新组 ${groupId}。` : `已保存组 ${result.group_id}。`, 'ok');
    resetForm();
    await loadWorkbench();
  }

  async function deleteGroup(groupId) {
    if (debugFixture) {
      state.groups = state.groups.filter((group) => Number(group.id) !== Number(groupId));
      setStatus(`Debug fixture：已模拟删除组 ${groupId}。`, 'ok');
      resetForm();
      await loadWorkbench();
      return;
    }
    const response = await fetch(`/api/report/${runId}/dfmea/groups/${encodeURIComponent(groupId)}`, { method: 'DELETE' });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || '删除 DFMEA 分组失败');
    }
    setStatus(`已删除组 ${groupId}，组内位号已退回待排查池。`, 'ok');
    resetForm();
    await loadWorkbench();
  }

  filterForm?.addEventListener('submit', async (event) => {
    event.preventDefault();
    syncTableSortFromControl();
    resetPendingRenderLimit();
    try {
      await loadWorkbench();
    } catch (error) {
      destroyPendingTableShell(error.message || String(error));
      setStatus(error.message || String(error), 'error');
    }
  });

  field(filterForm, 'q')?.addEventListener('input', () => {
    window.clearTimeout(window.__pstxDfmeaSearchTimer);
    window.__pstxDfmeaSearchTimer = window.setTimeout(() => {
      resetPendingRenderLimit();
      renderPending();
    }, 120);
  });

  field(filterForm, 'include_depop')?.addEventListener('change', () => {
    resetPendingRenderLimit();
    filterForm?.requestSubmit();
  });
  field(filterForm, 'exclude_rc')?.addEventListener('change', () => {
    resetPendingRenderLimit();
    filterForm?.requestSubmit();
  });
  field(filterForm, 'sort')?.addEventListener('change', () => {
    syncTableSortFromControl();
    resetPendingRenderLimit();
    renderPending();
  });

  groupForm?.addEventListener('submit', async (event) => {
    try {
      await saveGroup(event);
    } catch (error) {
      setStatus(error.message || String(error), 'error');
    }
  });

  clearSelection?.addEventListener('click', () => {
    state.selected = new Set();
    updateSelectionUi();
  });

  cancelEdit?.addEventListener('click', resetForm);
  document.getElementById('dfmea-export-link')?.addEventListener('click', (event) => {
    if (!debugFixture) return;
    event.preventDefault();
    setStatus('Debug fixture：导出按钮仅用于样式预览，真实导出请从报告页进入 DFMEA 工作台。', 'ok');
  });

  loadWorkbench().catch((error) => {
    destroyPendingTableShell(error.message || String(error));
    setStatus(error.message || String(error), 'error');
  });
});
