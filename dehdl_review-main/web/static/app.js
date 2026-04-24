const revealObserver = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) {
      entry.target.classList.add('is-visible');
      revealObserver.unobserve(entry.target);
    }
  });
}, { threshold: 0.16 });

const SIDEBAR_STORAGE_KEY = 'pstx-report-sidebar-collapsed';
const TABLE_COLUMN_STORAGE_PREFIX = 'pstx-report-table-columns:';
const TABLE_WIDTH_STORAGE_PREFIX = 'pstx-report-table-widths:';
const DEFAULT_COLUMN_WIDTH = 168;
const MIN_COLUMN_WIDTH = 96;
const MAX_COLUMN_WIDTH = 640;
const TABLE_INITIAL_RENDER_LIMIT = 320;
const TABLE_RENDER_STEP = 320;
const TABLE_FILTER_DEBOUNCE_MS = 90;
const LONG_TEXT_COLUMN_HINTS = new Map([
  ['引脚名', 340],
  ['子模块路径', 340],
  ['说明', 320],
  ['判定依据', 320],
  ['连接元件', 260],
  ['使用该值的位号', 260],
  ['串阻位号', 220],
  ['上拉位号', 220],
  ['下拉位号', 220],
  ['串阻另一端网络', 240],
  ['上拉电源', 220],
  ['网络名', 220],
]);
const REFDES_PRIORITY_COLUMNS = [
  '芯片位号',
  '位号',
  '串阻位号',
  '上拉位号',
  '下拉位号',
  '连接元件',
  '使用该值的位号',
];

const tableMountObserver = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (!entry.isIntersecting) return;
    const body = entry.target;
    tableMountObserver.unobserve(body);
    body.dataset.pendingMount = '';
    if (body.pstxTableData) {
      mountTable(body, body.pstxTableData);
      body.pstxTableData = null;
    }
  });
}, { rootMargin: '720px 0px', threshold: 0.01 });

function bootReveals() {
  const nodes = document.querySelectorAll('[data-reveal]');
  if (!nodes.length) return;
  document.documentElement.classList.add('reveal-enabled');
  nodes.forEach((node) => {
    if (node.dataset.revealObserved === 'true') return;
    node.dataset.revealObserved = 'true';
    revealObserver.observe(node);
    requestAnimationFrame(() => {
      const rect = node.getBoundingClientRect();
      if (rect.top < window.innerHeight * 1.05 && rect.bottom > -40) {
        node.classList.add('is-visible');
      }
    });
  });
}

function setStatus(message, isError = false) {
  const node = document.getElementById('form-status');
  if (!node) return;
  node.textContent = message;
  node.style.color = isError ? 'var(--warn)' : 'var(--muted)';
}

function scrollToSection(id) {
  const target = document.getElementById(id);
  if (!target) return;
  target.scrollIntoView({ behavior: 'smooth', block: 'start' });
  target.classList.add('section-flash');
  window.setTimeout(() => target.classList.remove('section-flash'), 1400);
}

async function handleHomePage() {
  const form = document.getElementById('analyze-form');
  const mask = document.getElementById('loading-mask');
  if (!form) return;

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    setStatus('正在生成分析报告，请稍候…');
    mask.hidden = false;
    try {
      const response = await fetch('/api/analyze', {
        method: 'POST',
        body: new FormData(form),
      });
      const payload = await response.json();
      if (!response.ok || !payload.ok) {
        throw new Error(payload.error || '分析任务创建失败');
      }
      window.location.href = payload.redirect_url;
    } catch (error) {
      mask.hidden = true;
      setStatus(error.message || String(error), true);
    }
  });
}

function metricNode(metric) {
  const wrapper = document.createElement(metric.target ? 'button' : 'div');
  wrapper.className = `metric tone-${metric.tone || 'neutral'}${metric.target ? ' metric-action' : ''}`;
  if (metric.target) {
    wrapper.type = 'button';
    wrapper.addEventListener('click', () => scrollToSection(metric.target));
  }
  wrapper.innerHTML = `
    <span class="metric-icon" aria-hidden="true">${metricIconForLabel(metric.label)}</span>
    <div class="metric-copy">
      <div class="metric-label">${metric.label}</div>
      <div class="metric-value">${metric.value}</div>
      ${metric.caption ? `<div class="metric-caption">${metric.caption}</div>` : ''}
    </div>
  `;
  return wrapper;
}

function metricIconForLabel(label) {
  const text = String(label || '');
  if (text.includes('网络')) return 'N';
  if (text.includes('DRC') || text.includes('风险') || text.includes('不合格')) return '!';
  if (text.includes('无法') || text.includes('待人工')) return '?';
  if (text.includes('电阻')) return 'R';
  if (text.includes('电容')) return 'C';
  if (text.includes('DEPOP')) return 'D';
  return '#';
}

function badgeForValue(value) {
  const text = String(value ?? '');
  if (!text) return null;
  let klass = 'badge-muted';
  if (text.startsWith('❌') || text === '高' || text.includes('确定结论')) {
    klass = 'badge-danger';
  } else if (text.startsWith('⚠') || text === '中' || text.includes('候选判断')) {
    klass = 'badge-warning';
  } else if (text.startsWith('✅')) {
    klass = 'badge-ok';
  }
  const span = document.createElement('span');
  span.className = `cell-badge ${klass}`;
  span.textContent = text;
  return span;
}

function cellNode(column, value) {
  const td = document.createElement('td');
  const text = String(value ?? '');
  if (text.length > 18) {
    td.title = text;
  }
  if (['状态', '结论类型', '严重级别'].includes(column)) {
    const badge = badgeForValue(text);
    if (badge) {
      td.appendChild(badge);
      return td;
    }
  }
  td.textContent = text;
  return td;
}

function shortLabelForNav(text) {
  const label = String(text || '').trim();
  if (!label) return '';
  if (/^[A-Za-z0-9 _-]+$/.test(label)) {
    return label.slice(0, 3).toUpperCase();
  }
  return label.slice(0, 2);
}

function normalizeText(value) {
  return String(value ?? '').trim();
}

function parseSortableValue(value) {
  const text = normalizeText(value);
  if (!text) {
    return { kind: 'empty', text: '', number: Number.NaN };
  }
  const numericLike = text.replace(/,/g, '').replace(/%$/, '');
  if (/^-?\d+(\.\d+)?$/.test(numericLike)) {
    return { kind: 'number', text: text.toLowerCase(), number: Number(numericLike) };
  }
  return { kind: 'text', text: text.toLowerCase(), number: Number.NaN };
}

function compareTextNatural(left, right) {
  return normalizeText(left).localeCompare(normalizeText(right), 'zh-Hans-CN', {
    numeric: true,
    sensitivity: 'base',
  });
}

function compareParsedValues(left, right) {
  if (left.kind === 'number' && right.kind === 'number') {
    return left.number - right.number;
  }
  return compareTextNatural(left.text, right.text);
}

function tableStorageKey(tableData) {
  return `${TABLE_COLUMN_STORAGE_PREFIX}${tableData.id || tableData.title}`;
}

function tableWidthStorageKey(tableData) {
  return `${TABLE_WIDTH_STORAGE_PREFIX}${tableData.id || tableData.title}`;
}

function getDefaultVisibleColumns(tableData) {
  const hidden = new Set(tableData.default_hidden_columns || []);
  const visible = tableData.columns.filter((column) => !hidden.has(column));
  return visible.length ? visible : tableData.columns.slice(0, 1);
}

function clampColumnWidth(width) {
  return Math.min(MAX_COLUMN_WIDTH, Math.max(MIN_COLUMN_WIDTH, Math.round(width)));
}

function suggestedColumnWidth(column) {
  const directHint = LONG_TEXT_COLUMN_HINTS.get(column);
  if (directHint) {
    return directHint;
  }
  if (column.includes('路径') || column.includes('依据') || column.includes('说明')) {
    return 320;
  }
  if (column.includes('引脚') || column.includes('网络')) {
    return 220;
  }
  return clampColumnWidth(Math.max(DEFAULT_COLUMN_WIDTH, column.length * 18 + 56));
}

function loadVisibleColumns(tableData) {
  const fallback = getDefaultVisibleColumns(tableData);
  try {
    const raw = window.localStorage.getItem(tableStorageKey(tableData));
    if (!raw) {
      return new Set(fallback);
    }
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return new Set(fallback);
    }
    const visible = tableData.columns.filter((column) => parsed.includes(column));
    return new Set(visible.length ? visible : fallback);
  } catch {
    return new Set(fallback);
  }
}

function persistVisibleColumns(state) {
  try {
    window.localStorage.setItem(state.storageKey, JSON.stringify(Array.from(state.visibleColumns)));
  } catch {
    // Ignore storage write failures and keep the in-memory state.
  }
}

function loadColumnWidths(tableData) {
  try {
    const raw = window.localStorage.getItem(tableWidthStorageKey(tableData));
    if (!raw) {
      return new Map();
    }
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      return new Map();
    }
    return new Map(
      Object.entries(parsed)
        .map(([column, width]) => [column, clampColumnWidth(Number(width) || suggestedColumnWidth(column))]),
    );
  } catch {
    return new Map();
  }
}

function persistColumnWidths(state) {
  try {
    const payload = Object.fromEntries(state.columnWidths.entries());
    window.localStorage.setItem(state.widthStorageKey, JSON.stringify(payload));
  } catch {
    // Ignore storage write failures and keep the in-memory state.
  }
}

function getColumnWidth(state, column) {
  return state.columnWidths.get(column) || suggestedColumnWidth(column);
}

function setColumnWidth(state, column, width) {
  state.columnWidths.set(column, clampColumnWidth(width));
}

function extractRefdesSuffixGroup(refdes) {
  const match = normalizeText(refdes).toUpperCase().match(/\d([A-Z]+\d+)$/);
  return match ? match[1] : '';
}

function splitRefdesTokens(value) {
  return normalizeText(value)
    .split(/[\s,，;/]+/)
    .map((token) => token.trim())
    .filter(Boolean);
}

function extractHierarchyLabelsFromPinName(pinName) {
  const text = normalizeText(pinName);
  if (!text.includes('@')) {
    return [];
  }
  return text
    .split('@')
    .map((segment) => segment.trim())
    .filter(Boolean)
    .map((segment) => {
      const head = segment.split(':', 1)[0].trim();
      const moduleHead = head.includes('(') ? head.split('(', 1)[0].trim() : head;
      return moduleHead.split('.').pop().trim();
    })
    .filter(Boolean);
}

function inferRowSortMeta(row, columns) {
  const prioritizedColumns = [
    ...REFDES_PRIORITY_COLUMNS.filter((column) => columns.includes(column)),
    ...columns.filter((column) => column.includes('位号') && !REFDES_PRIORITY_COLUMNS.includes(column)),
  ];
  const refTokens = [];
  prioritizedColumns.forEach((column) => {
    splitRefdesTokens(row[column]).forEach((token) => refTokens.push(token));
  });

  const suffixCounts = new Map();
  const suffixOrder = [];
  refTokens.forEach((token) => {
    const suffixGroup = extractRefdesSuffixGroup(token);
    if (!suffixGroup) {
      return;
    }
    if (!suffixCounts.has(suffixGroup)) {
      suffixOrder.push(suffixGroup);
    }
    suffixCounts.set(suffixGroup, (suffixCounts.get(suffixGroup) || 0) + 1);
  });
  suffixOrder.sort((left, right) => {
    const countGap = (suffixCounts.get(right) || 0) - (suffixCounts.get(left) || 0);
    return countGap || compareTextNatural(left, right);
  });

  const hierarchyLabels = extractHierarchyLabelsFromPinName(row['引脚名']);
  const explicitSubmodule = normalizeText(row['子模块']);
  const explicitSubmodulePath = normalizeText(row['子模块路径']);
  const inferredSubmodule = hierarchyLabels.length >= 2 ? hierarchyLabels[hierarchyLabels.length - 2] : '';
  const inferredSubmodulePath = hierarchyLabels.length >= 2 ? hierarchyLabels.slice(0, -1).join(' / ') : '';

  return {
    primaryRefdes: refTokens[0] || '',
    suffixGroup: suffixOrder[0] || '',
    contextGroup: suffixOrder[0] || explicitSubmodule || explicitSubmodulePath || inferredSubmodule || inferredSubmodulePath,
    submodule: explicitSubmodule || inferredSubmodule,
    submodulePath: explicitSubmodulePath || inferredSubmodulePath || explicitSubmodule || inferredSubmodule,
    pin: normalizeText(row['引脚']),
    network: normalizeText(row['网络名']),
  };
}

function compareRecords(left, right, state) {
  const sortMode = state.sortModeSelect?.value || 'column';
  const sortDirection = state.sortDirSelect.value;
  let result = 0;

  if (sortMode === 'suffix_group') {
    result =
      Number(Boolean(right.sortMeta.suffixGroup)) - Number(Boolean(left.sortMeta.suffixGroup)) ||
      compareTextNatural(left.sortMeta.suffixGroup, right.sortMeta.suffixGroup) ||
      compareTextNatural(left.sortMeta.primaryRefdes, right.sortMeta.primaryRefdes) ||
      compareParsedValues(parseSortableValue(left.sortMeta.pin), parseSortableValue(right.sortMeta.pin));
  } else if (sortMode === 'submodule') {
    result =
      Number(Boolean(right.sortMeta.contextGroup)) - Number(Boolean(left.sortMeta.contextGroup)) ||
      compareTextNatural(left.sortMeta.contextGroup, right.sortMeta.contextGroup) ||
      Number(Boolean(right.sortMeta.submodulePath)) - Number(Boolean(left.sortMeta.submodulePath)) ||
      compareTextNatural(left.sortMeta.submodulePath, right.sortMeta.submodulePath) ||
      compareTextNatural(left.sortMeta.suffixGroup, right.sortMeta.suffixGroup) ||
      compareTextNatural(left.sortMeta.primaryRefdes, right.sortMeta.primaryRefdes) ||
      compareParsedValues(parseSortableValue(left.sortMeta.pin), parseSortableValue(right.sortMeta.pin));
  } else {
    const sortColumn = state.sortSelect.value;
    if (!sortColumn) {
      result = left.index - right.index;
    } else {
      result = compareParsedValues(
        parseSortableValue(left.row[sortColumn]),
        parseSortableValue(right.row[sortColumn]),
      );
      if (result === 0) {
        result = left.index - right.index;
      }
    }
  }

  if (result === 0) {
    result = left.index - right.index;
  }
  return sortDirection === 'desc' ? -result : result;
}

function updateSortHeaders(state) {
  state.headerButtons.forEach((button, column) => {
    const indicator = button.querySelector('.sort-indicator');
    const active = state.sortModeSelect.value === 'column' && column === state.sortSelect.value;
    button.classList.toggle('is-sorted', active);
    indicator.textContent = !active ? '⇅' : (state.sortDirSelect.value === 'desc' ? '↓' : '↑');
  });
}

function getVisibleColumns(state) {
  const visible = state.columns.filter((column) => state.visibleColumns.has(column));
  return visible.length ? visible : state.columns.slice(0, 1);
}

function syncSortColumnOptions(state, visibleColumns) {
  const previous = state.sortSelect.value;
  state.sortSelect.replaceChildren();

  const defaultOption = document.createElement('option');
  defaultOption.value = '';
  defaultOption.textContent = '原始顺序';
  state.sortSelect.appendChild(defaultOption);

  visibleColumns.forEach((column) => {
    const option = document.createElement('option');
    option.value = column;
    option.textContent = column;
    state.sortSelect.appendChild(option);
  });

  state.sortSelect.value = visibleColumns.includes(previous) ? previous : '';
}

function syncColumnWidths(state, visibleColumns) {
  const fragment = document.createDocumentFragment();
  state.colElements = new Map();
  let totalWidth = 0;

  visibleColumns.forEach((column) => {
    const col = document.createElement('col');
    const width = getColumnWidth(state, column);
    col.style.width = `${width}px`;
    col.style.minWidth = `${MIN_COLUMN_WIDTH}px`;
    fragment.appendChild(col);
    state.colElements.set(column, col);
    totalWidth += width;
  });

  state.colgroup.replaceChildren(fragment);
  const baseWidth = state.scroll.clientWidth || 720;
  state.table.style.width = `${Math.max(totalWidth, baseWidth)}px`;
  updateScrollShadows(state.scroll);
}

function updateScrollShadows(scroll) {
  if (!scroll) return;
  const canScroll = scroll.scrollWidth > scroll.clientWidth + 2;
  const atStart = scroll.scrollLeft <= 2;
  const atEnd = scroll.scrollLeft + scroll.clientWidth >= scroll.scrollWidth - 2;
  scroll.classList.toggle('can-scroll-left', canScroll && !atStart);
  scroll.classList.toggle('at-scroll-end', canScroll && atEnd);
}

function scheduleScrollShadowUpdate(scroll) {
  if (!scroll || scroll.dataset.shadowFrame === 'true') return;
  scroll.dataset.shadowFrame = 'true';
  requestAnimationFrame(() => {
    scroll.dataset.shadowFrame = '';
    updateScrollShadows(scroll);
  });
}

function scheduleTableMount(container, tableData) {
  if (container.dataset.mounted === 'true') return;
  container.pstxTableData = tableData;
  container.dataset.pendingMount = 'true';
  if ('IntersectionObserver' in window) {
    tableMountObserver.observe(container);
  } else {
    requestAnimationFrame(() => mountTable(container, tableData));
  }
  requestAnimationFrame(() => {
    if (container.dataset.mounted === 'true' || !container.pstxTableData) return;
    const rect = container.getBoundingClientRect();
    const margin = 720;
    if (rect.top < window.innerHeight + margin && rect.bottom > -margin) {
      tableMountObserver.unobserve(container);
      container.dataset.pendingMount = '';
      mountTable(container, container.pstxTableData);
      container.pstxTableData = null;
    }
  });
}

function resetAndApplyTableState(state) {
  state.renderLimit = TABLE_INITIAL_RENDER_LIMIT;
  applyTableState(state);
}

function scheduleFilterApply(state) {
  state.renderLimit = TABLE_INITIAL_RENDER_LIMIT;
  window.clearTimeout(state.filterTimer);
  state.filterTimer = window.setTimeout(() => applyTableState(state), TABLE_FILTER_DEBOUNCE_MS);
}

function columnFilterOptions() {
  return [
    ['contains', '包含'],
    ['not_contains', '不包含'],
    ['equals', '等于'],
    ['not_equals', '不等于'],
    ['starts_with', '开头是'],
    ['ends_with', '结尾是'],
    ['empty', '为空'],
    ['not_empty', '非空'],
  ];
}

function activeColumnFilters(state) {
  return (state.columnFilters || []).filter((filter) => {
    if (!filter.column) return false;
    if (['empty', 'not_empty'].includes(filter.operator)) return true;
    return normalizeText(filter.value);
  });
}

function columnFilterMatches(row, filter) {
  const raw = normalizeText(row[filter.column]);
  const text = raw.toLowerCase();
  const query = normalizeText(filter.value).toLowerCase();
  if (filter.operator === 'empty') return !raw;
  if (filter.operator === 'not_empty') return Boolean(raw);
  if (!query) return true;
  if (filter.operator === 'not_contains') return !text.includes(query);
  if (filter.operator === 'equals') return text === query;
  if (filter.operator === 'not_equals') return text !== query;
  if (filter.operator === 'starts_with') return text.startsWith(query);
  if (filter.operator === 'ends_with') return text.endsWith(query);
  return text.includes(query);
}

function updateColumnFilterSummary(state) {
  if (!state.columnFilterSummary) return;
  const count = activeColumnFilters(state).length;
  state.columnFilterSummary.textContent = count ? ` · ${count}` : '';
}

function renderColumnFilters(state) {
  const list = state.columnFilterList;
  if (!list) return;
  list.replaceChildren();
  if (!state.columnFilters.length) {
    const empty = document.createElement('p');
    empty.className = 'column-filter-empty';
    empty.textContent = '尚未添加列条件。可按页码、状态、位号、网络名等字段组合筛选。';
    list.appendChild(empty);
  }
  state.columnFilters.forEach((filter) => {
    const row = document.createElement('div');
    row.className = 'column-filter-row';

    const columnSelect = document.createElement('select');
    columnSelect.className = 'column-filter-column';
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = '选择列';
    columnSelect.appendChild(placeholder);
    state.columns.forEach((column) => {
      const option = document.createElement('option');
      option.value = column;
      option.textContent = column;
      columnSelect.appendChild(option);
    });
    columnSelect.value = filter.column || '';
    columnSelect.addEventListener('change', () => {
      filter.column = columnSelect.value;
      scheduleFilterApply(state);
      renderColumnFilters(state);
    });

    const operatorSelect = document.createElement('select');
    operatorSelect.className = 'column-filter-operator';
    columnFilterOptions().forEach(([value, label]) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = label;
      operatorSelect.appendChild(option);
    });
    operatorSelect.value = filter.operator || 'contains';
    operatorSelect.addEventListener('change', () => {
      filter.operator = operatorSelect.value;
      renderColumnFilters(state);
      scheduleFilterApply(state);
    });

    const valueInput = document.createElement('input');
    valueInput.className = 'column-filter-value';
    valueInput.type = 'text';
    valueInput.placeholder = ['empty', 'not_empty'].includes(filter.operator) ? '无需输入' : '筛选值';
    valueInput.value = filter.value || '';
    valueInput.disabled = ['empty', 'not_empty'].includes(filter.operator);
    valueInput.addEventListener('input', () => {
      filter.value = valueInput.value;
      scheduleFilterApply(state);
    });

    const removeButton = document.createElement('button');
    removeButton.type = 'button';
    removeButton.className = 'column-action column-filter-remove';
    removeButton.textContent = '移除';
    removeButton.addEventListener('click', () => {
      state.columnFilters = state.columnFilters.filter((item) => item.id !== filter.id);
      renderColumnFilters(state);
      resetAndApplyTableState(state);
    });

    row.appendChild(columnSelect);
    row.appendChild(operatorSelect);
    row.appendChild(valueInput);
    row.appendChild(removeButton);
    list.appendChild(row);
  });
  updateColumnFilterSummary(state);
}

function startColumnResize(event, state, column) {
  event.preventDefault();
  event.stopPropagation();

  const startX = event.clientX;
  const startWidth = getColumnWidth(state, column);
  const handle = event.currentTarget;
  let pendingWidth = startWidth;
  let resizeFrame = 0;
  document.body.classList.add('is-column-resizing');
  state.table.classList.add('is-resizing');

  const onPointerMove = (moveEvent) => {
    const delta = moveEvent.clientX - startX;
    pendingWidth = startWidth + delta;
    if (resizeFrame) return;
    resizeFrame = requestAnimationFrame(() => {
      resizeFrame = 0;
      setColumnWidth(state, column, pendingWidth);
      syncColumnWidths(state, getVisibleColumns(state));
    });
  };

  const finishResize = () => {
    if (resizeFrame) {
      cancelAnimationFrame(resizeFrame);
      resizeFrame = 0;
    }
    setColumnWidth(state, column, pendingWidth);
    syncColumnWidths(state, getVisibleColumns(state));
    document.body.classList.remove('is-column-resizing');
    state.table.classList.remove('is-resizing');
    persistColumnWidths(state);
    window.removeEventListener('pointermove', onPointerMove);
    window.removeEventListener('pointerup', finishResize);
    window.removeEventListener('pointercancel', finishResize);
  };

  if (handle.setPointerCapture && event.pointerId !== undefined) {
    handle.setPointerCapture(event.pointerId);
  }

  window.addEventListener('pointermove', onPointerMove);
  window.addEventListener('pointerup', finishResize);
  window.addEventListener('pointercancel', finishResize);
}

function renderTableHead(state, visibleColumns) {
  const headRow = document.createElement('tr');
  state.headerButtons = new Map();

  visibleColumns.forEach((column) => {
    const th = document.createElement('th');
    th.className = 'table-column-head';
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'sort-header';
    button.innerHTML = `
      <span class="sort-header-label">${column}</span>
      <span class="sort-indicator">⇅</span>
    `;
    button.addEventListener('click', () => {
      state.sortModeSelect.value = 'column';
      if (state.sortSelect.value === column) {
        state.sortDirSelect.value = state.sortDirSelect.value === 'asc' ? 'desc' : 'asc';
      } else {
        state.sortSelect.value = column;
        state.sortDirSelect.value = 'asc';
      }
      resetAndApplyTableState(state);
    });
    const handle = document.createElement('span');
    handle.className = 'column-resize-handle';
    handle.title = `${column} 列宽调整`;
    handle.addEventListener('pointerdown', (event) => startColumnResize(event, state, column));
    handle.addEventListener('dblclick', (event) => {
      event.preventDefault();
      event.stopPropagation();
      setColumnWidth(state, column, suggestedColumnWidth(column));
      persistColumnWidths(state);
      syncColumnWidths(state, getVisibleColumns(state));
    });
    th.appendChild(button);
    th.appendChild(handle);
    state.headerButtons.set(column, button);
    headRow.appendChild(th);
  });

  state.thead.replaceChildren(headRow);
}

function syncColumnPicker(state) {
  const visibleColumns = getVisibleColumns(state);
  const fragment = document.createDocumentFragment();
  state.columns.forEach((column) => {
    const label = document.createElement('label');
    label.className = 'column-option';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = state.visibleColumns.has(column);
    checkbox.disabled = checkbox.checked && visibleColumns.length === 1;
    checkbox.addEventListener('change', () => {
      if (checkbox.checked) {
        state.visibleColumns.add(column);
      } else if (visibleColumns.length > 1) {
        state.visibleColumns.delete(column);
      }
      persistVisibleColumns(state);
      resetAndApplyTableState(state);
    });

    const text = document.createElement('span');
    text.textContent = column;
    label.appendChild(checkbox);
    label.appendChild(text);
    fragment.appendChild(label);
  });
  state.columnPickerList.replaceChildren(fragment);
}

function applyTableState(state) {
  const visibleColumns = getVisibleColumns(state);
  syncSortColumnOptions(state, visibleColumns);
  renderTableHead(state, visibleColumns);
  syncColumnWidths(state, visibleColumns);
  syncColumnPicker(state);
  updateColumnFilterSummary(state);

  const keyword = state.filterInput.value.trim().toLowerCase();
  const categoryValue = state.categorySelect ? state.categorySelect.value : '';
  const columnFilters = activeColumnFilters(state);
  let visibleRecords = state.records.filter((record) => {
    if (keyword && !record.search.includes(keyword)) {
      return false;
    }
    if (state.categoryColumn && categoryValue && normalizeText(record.row[state.categoryColumn]) !== categoryValue) {
      return false;
    }
    if (columnFilters.length && !columnFilters.every((filter) => columnFilterMatches(record.row, filter))) {
      return false;
    }
    return true;
  });
  const needsSort = state.sortModeSelect.value !== 'column' || state.sortSelect.value || state.sortDirSelect.value === 'desc';
  if (needsSort) {
    visibleRecords = visibleRecords.sort((left, right) => compareRecords(left, right, state));
  }

  const renderLimit = Math.min(state.renderLimit || TABLE_INITIAL_RENDER_LIMIT, visibleRecords.length);
  const renderedRecords = visibleRecords.slice(0, renderLimit);
  const fragment = document.createDocumentFragment();
  renderedRecords.forEach((record) => {
    const tr = document.createElement('tr');
    visibleColumns.forEach((column) => {
      tr.appendChild(cellNode(column, record.row[column]));
    });
    fragment.appendChild(tr);
  });
  state.tbody.replaceChildren(fragment);

  state.countNode.textContent = visibleRecords.length > renderLimit
    ? `渲染 ${renderLimit} / 筛选 ${visibleRecords.length} / 总 ${state.records.length}`
    : `显示 ${visibleRecords.length} / ${state.records.length}`;
  state.visibleColumnNode.textContent = `显示列 ${visibleColumns.length} / ${state.columns.length}`;
  if (state.renderStatusNode && state.renderMoreButton) {
    const remaining = Math.max(visibleRecords.length - renderLimit, 0);
    state.renderStatusNode.textContent = remaining
      ? `为了保持滚动流畅，当前先渲染 ${renderLimit} 行，剩余 ${remaining} 行可继续追加。`
      : `当前筛选结果已全部渲染。`;
    state.renderMoreButton.hidden = remaining === 0;
    state.renderMoreButton.textContent = `继续渲染 ${Math.min(TABLE_RENDER_STEP, remaining)} 行`;
  }
  state.sortSelect.disabled = state.sortModeSelect.value !== 'column';
  state.sortDirSelect.disabled = state.sortModeSelect.value === 'column' && !state.sortSelect.value;
  state.table.classList.toggle('density-compact', state.density === 'compact');
  state.table.classList.toggle('density-comfortable', state.density === 'comfortable');
  state.densityButton.textContent = state.density === 'compact' ? '紧凑行距' : '舒展行距';
  updateSortHeaders(state);
  updateScrollShadows(state.scroll);
}

function mountTable(container, tableData) {
  if (container.dataset.mounted === 'true') return;
  container.dataset.mounted = 'true';
  container.replaceChildren();

  if (!tableData.rows.length) {
    const empty = document.createElement('p');
    empty.className = 'empty-state';
    empty.textContent = '当前分区暂无数据。';
    container.appendChild(empty);
    return;
  }

  const categoryColumn = ['状态', '结论类型', '严重级别'].find((column) => tableData.columns.includes(column)) || '';
  const categoryOptions = categoryColumn
    ? Array.from(new Set(tableData.rows.map((row) => normalizeText(row[categoryColumn])).filter(Boolean)))
    : [];

  const toolbar = document.createElement('div');
  toolbar.className = 'table-toolbar';
  const toolbarMeta = document.createElement('div');
  toolbarMeta.className = 'table-toolbar-meta';
  toolbarMeta.innerHTML = `
    <span class="pill">列数 ${tableData.columns.length}</span>
    <span class="pill table-visible-columns-pill">显示列 ${tableData.columns.length} / ${tableData.columns.length}</span>
    <span class="pill table-count-pill">显示 ${tableData.rows.length} / ${tableData.rows.length}</span>
  `;

  const toolbarControls = document.createElement('div');
  toolbarControls.className = 'table-toolbar-controls';
  toolbarControls.innerHTML = `
    <label class="toolbar-field toolbar-field-grow">
      <span>关键字</span>
      <input type="text" class="table-filter" placeholder="筛选当前结果…">
    </label>
    <label class="toolbar-field table-category-field" hidden>
      <span class="table-category-label"></span>
      <select class="table-category-filter">
        <option value="">全部</option>
      </select>
    </label>
    <label class="toolbar-field">
      <span>排序模式</span>
      <select class="table-sort-mode"></select>
    </label>
    <label class="toolbar-field">
      <span>排序字段</span>
      <select class="table-sort-column"></select>
    </label>
    <label class="toolbar-field">
      <span>排序方式</span>
      <select class="table-sort-direction" disabled>
        <option value="asc">升序</option>
        <option value="desc">降序</option>
      </select>
    </label>
    <details class="column-picker">
      <summary class="ghost-btn toolbar-ghost">列显示</summary>
      <div class="column-picker-panel">
        <div class="column-picker-actions">
          <button type="button" class="column-action column-show-all">全部显示</button>
          <button type="button" class="column-action column-reset-default">恢复默认</button>
        </div>
        <div class="column-picker-list"></div>
      </div>
    </details>
    <details class="column-filter-builder">
      <summary class="ghost-btn toolbar-ghost">多列筛选<span class="column-filter-summary"></span></summary>
      <div class="column-filter-panel">
        <div class="column-filter-list"></div>
        <div class="column-filter-actions">
          <button type="button" class="column-action column-filter-add">添加条件</button>
          <button type="button" class="column-action column-filter-clear">清空条件</button>
        </div>
      </div>
    </details>
    <button type="button" class="ghost-btn toolbar-density">紧凑行距</button>
    <button type="button" class="ghost-btn toolbar-reset">重置</button>
  `;
  const categoryField = toolbarControls.querySelector('.table-category-field');
  const categoryLabel = toolbarControls.querySelector('.table-category-label');
  const categorySelect = toolbarControls.querySelector('.table-category-filter');
  if (categoryColumn) {
    categoryField.hidden = false;
    categoryLabel.textContent = categoryColumn;
    categoryOptions.forEach((value) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value;
      categorySelect.appendChild(option);
    });
  }

  const sortModeSelect = toolbarControls.querySelector('.table-sort-mode');
  const sortModeField = sortModeSelect.closest('.toolbar-field');
  const sortProfiles = (tableData.sort_profiles || []).length
    ? tableData.sort_profiles
    : [{ id: 'column', label: '字段排序' }];
  sortProfiles.forEach((profile) => {
    const option = document.createElement('option');
    option.value = profile.id;
    option.textContent = profile.label;
    sortModeSelect.appendChild(option);
  });
  sortModeField.hidden = sortProfiles.length <= 1;
  sortModeSelect.value = tableData.default_sort_mode || sortProfiles[0].id || 'column';
  const sortSelect = toolbarControls.querySelector('.table-sort-column');

  const scroll = document.createElement('div');
  scroll.className = 'table-scroll';
  const table = document.createElement('table');
  table.className = 'report-data-table';
  const colgroup = document.createElement('colgroup');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  const renderFooter = document.createElement('div');
  renderFooter.className = 'table-render-footer';
  renderFooter.innerHTML = `
    <span class="table-render-status"></span>
    <button type="button" class="ghost-btn table-render-more" hidden>继续渲染更多</button>
  `;
  const records = tableData.rows.map((row, index) => ({
    row,
    index,
    search: tableData.columns.map((column) => normalizeText(row[column])).join(' ').toLowerCase(),
    sortMeta: inferRowSortMeta(row, tableData.columns),
  }));

  table.appendChild(colgroup);
  table.appendChild(thead);
  table.appendChild(tbody);
  scroll.appendChild(table);
  toolbar.appendChild(toolbarMeta);
  toolbar.appendChild(toolbarControls);
  container.appendChild(toolbar);
  container.appendChild(scroll);
  container.appendChild(renderFooter);

  const state = {
    tableId: tableData.id || tableData.title,
    storageKey: tableStorageKey(tableData),
    columns: tableData.columns,
    records,
    table,
    scroll,
    colgroup,
    thead,
    tbody,
    visibleColumnNode: toolbarMeta.querySelector('.table-visible-columns-pill'),
    countNode: toolbarMeta.querySelector('.table-count-pill'),
    filterInput: toolbarControls.querySelector('.table-filter'),
    categoryColumn,
    categorySelect: categoryColumn ? categorySelect : null,
    sortModeSelect,
    sortSelect,
    sortDirSelect: toolbarControls.querySelector('.table-sort-direction'),
    densityButton: toolbarControls.querySelector('.toolbar-density'),
    density: 'compact',
    renderLimit: TABLE_INITIAL_RENDER_LIMIT,
    renderStatusNode: renderFooter.querySelector('.table-render-status'),
    renderMoreButton: renderFooter.querySelector('.table-render-more'),
    columnPickerList: toolbarControls.querySelector('.column-picker-list'),
    columnFilterList: toolbarControls.querySelector('.column-filter-list'),
    columnFilterSummary: toolbarControls.querySelector('.column-filter-summary'),
    columnFilterAddButton: toolbarControls.querySelector('.column-filter-add'),
    columnFilterClearButton: toolbarControls.querySelector('.column-filter-clear'),
    columnFilters: [],
    nextColumnFilterId: 1,
    visibleColumns: loadVisibleColumns(tableData),
    columnWidths: loadColumnWidths(tableData),
    headerButtons: new Map(),
    colElements: new Map(),
  };

  state.filterInput.addEventListener('input', () => scheduleFilterApply(state));
  if (state.categorySelect) {
    state.categorySelect.addEventListener('change', () => resetAndApplyTableState(state));
  }
  state.sortModeSelect.addEventListener('change', () => resetAndApplyTableState(state));
  state.sortSelect.addEventListener('change', () => resetAndApplyTableState(state));
  state.sortDirSelect.addEventListener('change', () => resetAndApplyTableState(state));
  state.densityButton.addEventListener('click', () => {
    state.density = state.density === 'compact' ? 'comfortable' : 'compact';
    applyTableState(state);
  });
  state.renderMoreButton.addEventListener('click', () => {
    state.renderLimit += TABLE_RENDER_STEP;
    applyTableState(state);
  });
  scroll.addEventListener('scroll', () => scheduleScrollShadowUpdate(scroll), { passive: true });
  toolbarControls.querySelector('.column-show-all').addEventListener('click', () => {
    state.visibleColumns = new Set(state.columns);
    persistVisibleColumns(state);
    resetAndApplyTableState(state);
  });
  toolbarControls.querySelector('.column-reset-default').addEventListener('click', () => {
    state.visibleColumns = new Set(getDefaultVisibleColumns(tableData));
    persistVisibleColumns(state);
    resetAndApplyTableState(state);
  });
  state.columnFilterAddButton.addEventListener('click', () => {
    state.columnFilters.push({
      id: state.nextColumnFilterId,
      column: '',
      operator: 'contains',
      value: '',
    });
    state.nextColumnFilterId += 1;
    renderColumnFilters(state);
  });
  state.columnFilterClearButton.addEventListener('click', () => {
    state.columnFilters = [];
    renderColumnFilters(state);
    resetAndApplyTableState(state);
  });
  toolbarControls.querySelector('.toolbar-reset').addEventListener('click', () => {
    state.filterInput.value = '';
    if (state.categorySelect) {
      state.categorySelect.value = '';
    }
    state.sortModeSelect.value = tableData.default_sort_mode || 'column';
    state.sortSelect.value = '';
    state.sortDirSelect.value = 'asc';
    state.columnFilters = [];
    renderColumnFilters(state);
    resetAndApplyTableState(state);
  });

  renderColumnFilters(state);
  applyTableState(state);
}

function tableBlock(tableData, initialOpen = false) {
  const block = document.createElement('article');
  block.className = 'table-block';

  const kindPills = Object.entries(tableData.kind_counts || {})
    .map(([label, value]) => `<span class="pill">${label} ${value}</span>`)
    .join('');

  block.innerHTML = `
    <div class="table-header">
      <div>
        <h3 class="table-title">${tableData.title}</h3>
      </div>
      <div class="table-meta">
        <span class="pill">记录 ${tableData.count}</span>
        ${kindPills}
      </div>
      <button type="button" class="toggle-btn">查看详情</button>
    </div>
    <div class="table-body"></div>
  `;

  const button = block.querySelector('.toggle-btn');
  const body = block.querySelector('.table-body');
  const setOpen = (open) => {
    block.classList.toggle('is-open', open);
    button.textContent = open ? '收起' : '查看详情';
    if (open) mountTable(body, tableData);
  };
  button.addEventListener('click', () => {
    const open = block.classList.toggle('is-open');
    setOpen(open);
  });
  if (initialOpen) {
    block.classList.add('is-open');
    button.textContent = '收起';
    body.innerHTML = '<p class="empty-state table-lazy-state">表格将在进入视口附近时加载。</p>';
    scheduleTableMount(body, tableData);
  }

  return block;
}

function sectionNode(section) {
  const wrapper = document.createElement('section');
  wrapper.id = section.id;
  wrapper.className = 'report-section';
  wrapper.setAttribute('data-reveal', '');
  wrapper.innerHTML = `
    <div class="section-heading">
      <p class="eyebrow">${section.id.toUpperCase()}</p>
      <h2>${section.title}</h2>
      <p>${section.lead}</p>
    </div>
  `;

  const stack = document.createElement('div');
  stack.className = 'table-stack';
  const firstOpenTable = section.tables.find((table) => table.count > 0);
  section.tables.forEach((table) => stack.appendChild(tableBlock(table, table === firstOpenTable)));
  wrapper.appendChild(stack);
  return wrapper;
}

function insightNode(insight) {
  const node = document.createElement(insight.target ? 'button' : 'article');
  node.className = `insight-card tone-${insight.tone || 'neutral'}${insight.target ? ' is-clickable' : ''}`;
  if (insight.target) {
    node.type = 'button';
    node.addEventListener('click', () => scrollToSection(insight.target));
  }
  node.innerHTML = `
    <div class="insight-title">${insight.title}</div>
    <div class="insight-body">${insight.body}</div>
  `;
  return node;
}

function sectionCardNode(section) {
  const node = document.createElement('button');
  node.type = 'button';
  node.className = `section-card tone-${section.tone || 'neutral'}`;
  node.addEventListener('click', () => scrollToSection(section.id));
  node.innerHTML = `
    <div class="section-card-header">
      <div>
        <div class="section-card-title">${section.title}</div>
        <div class="section-card-lead">${section.lead}</div>
      </div>
      <div class="section-card-count">${section.rows}</div>
    </div>
    <div class="section-card-footer">
      <span>${section.active_tables} 个子表有结果</span>
      <span>${section.top_label} · ${section.top_value}</span>
    </div>
  `;
  return node;
}

function renderSummary(report) {
  const depopMode = report.include_depop
    ? `DEPOP 排查：开启（${report.depop_count || 0} 个器件参与分析）`
    : `DEPOP 排查：关闭（已忽略 ${report.excluded_depop_count || 0} 个器件）`;
  document.getElementById('generated-at').textContent =
    `生成时间：${report.generated_at} · 降额阈值：${report.ratio_limit}% · ${depopMode}`;
  const topbarGeneratedAt = document.getElementById('topbar-generated-at');
  if (topbarGeneratedAt) {
    topbarGeneratedAt.textContent = `生成 ${report.generated_at}`;
  }

  const metricStrip = document.getElementById('metric-strip');
  metricStrip.replaceChildren();
  report.metrics.forEach((metric) => metricStrip.appendChild(metricNode(metric)));

  const topInsights = document.getElementById('top-insights');
  topInsights.replaceChildren();
  (report.top_insights || []).forEach((insight) => topInsights.appendChild(insightNode(insight)));

  const sectionCards = document.getElementById('section-cards');
  sectionCards.replaceChildren();
  (report.section_cards || []).forEach((section) => sectionCards.appendChild(sectionCardNode(section)));

  const fileMeta = document.getElementById('file-meta');
  fileMeta.replaceChildren();
  report.input_files.forEach((file) => {
    const li = document.createElement('li');
    li.textContent = `${file.label} · ${file.filename || '未上传'} · ${file.size} 字节`;
    fileMeta.appendChild(li);
  });

  const summaryLines = document.getElementById('summary-lines');
  summaryLines.replaceChildren();
  report.summary_lines.forEach((line) => {
    const li = document.createElement('li');
    li.textContent = line;
    summaryLines.appendChild(li);
  });

  if (report.warnings && report.warnings.length) {
    const warningNode = document.getElementById('warning-list');
    warningNode.hidden = false;
    warningNode.innerHTML = `<strong>补充说明</strong><ul>${report.warnings.map((item) => `<li>${item}</li>`).join('')}</ul>`;
  } else {
    document.getElementById('warning-list').hidden = true;
  }
}

function renderSectionNav(sections) {
  const nav = document.getElementById('section-nav');
  const links = [];
  const navOrder = ['summary', ...sections.map((section) => section.id), 'query'];
  navOrder.forEach((id) => {
    const labelMap = {
      summary: '概览',
      query: '查询',
    };
    const label = labelMap[id] || sections.find((section) => section.id === id)?.title || id;
    const shortLabel = shortLabelForNav(label);
    const link = document.createElement('a');
    link.href = `#${id}`;
    link.setAttribute('aria-label', label);
    if (id === 'summary') {
      link.classList.add('is-active');
    }
    link.innerHTML = `
      <span class="nav-mark" aria-hidden="true">${navIconForSection(id, label)}</span>
      <span class="nav-full">${label}</span>
      <span class="nav-short" aria-hidden="true">${shortLabel}</span>
    `;
    nav.appendChild(link);
    links.push(link);
  });

  const observer = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      const link = links.find((item) => item.getAttribute('href') === `#${entry.target.id}`);
      if (!link) return;
      if (entry.isIntersecting) {
        links.forEach((item) => item.classList.remove('is-active'));
        link.classList.add('is-active');
        const inspector = document.getElementById('inspector-current');
        const title = link.getAttribute('aria-label') || link.textContent || entry.target.id;
        if (inspector) {
          inspector.textContent = title;
        }
      }
    });
  }, { rootMargin: '-35% 0px -45% 0px', threshold: 0 });

  navOrder.forEach((id) => {
    const node = document.getElementById(id);
    if (node) observer.observe(node);
  });
}

function navIconForSection(id, label) {
  const key = String(id || '').toLowerCase();
  const text = String(label || '');
  if (key === 'summary') return 'Σ';
  if (key === 'query') return '⌕';
  if (key.includes('bom')) return 'B';
  if (key.includes('network')) return 'N';
  if (key.includes('drc')) return 'D';
  if (key.includes('resistor')) return 'R';
  if (key.includes('derating')) return 'C';
  return text.slice(0, 1).toUpperCase() || '·';
}

function setSidebarCollapsed(collapsed) {
  const layout = document.querySelector('.report-layout');
  const button = document.getElementById('sidebar-toggle');
  if (!layout || !button) return;
  layout.classList.toggle('is-sidebar-collapsed', collapsed);
  button.textContent = collapsed ? '展开导航' : '收起导航';
  button.setAttribute('aria-expanded', String(!collapsed));
  try {
    window.localStorage.setItem(SIDEBAR_STORAGE_KEY, collapsed ? '1' : '0');
  } catch {
    // Ignore storage write failures and keep the current in-memory state.
  }
}

function detailRowsNode(items) {
  const list = document.createElement('div');
  list.className = 'detail-rows';
  items.forEach((item) => {
    const row = document.createElement('div');
    row.className = 'detail-row';
    row.innerHTML = `
      <span class="detail-label">${item.label}</span>
      <span class="detail-value">${item.value || '—'}</span>
    `;
    list.appendChild(row);
  });
  return list;
}

function dataTableNode(columns, rows) {
  const wrapper = document.createElement('div');
  wrapper.className = 'query-data-table';
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  const headRow = document.createElement('tr');
  columns.forEach((column) => {
    const th = document.createElement('th');
    th.textContent = column;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  rows.forEach((row) => {
    const tr = document.createElement('tr');
    columns.forEach((column) => {
      const td = document.createElement('td');
      td.textContent = row[column] || '—';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(thead);
  table.appendChild(tbody);
  wrapper.appendChild(table);
  return wrapper;
}

function querySummaryNode(payload) {
  const node = document.createElement('div');
  node.className = 'query-summary-card';
  const meta = (payload.summary?.meta || [])
    .filter((item) => item.value)
    .map((item) => `<span class="pill">${item.label} ${item.value}</span>`)
    .join('');
  node.innerHTML = `
    <div class="query-summary-eyebrow">${payload.entity_type === 'network' ? 'NETWORK' : 'COMPONENT'}</div>
    <h3>${payload.summary?.title || payload.title}</h3>
    <p>${payload.summary?.subtitle || ''}</p>
    <div class="query-summary-meta">${meta}</div>
  `;
  return node;
}

function queryListNode(payload, runQuery) {
  const wrapper = document.createElement('div');
  wrapper.className = 'query-list';
  const header = document.createElement('div');
  header.className = 'query-list-header';
  header.innerHTML = `
    <h3>${payload.summary?.subtitle || '匹配结果'}</h3>
    <p>点击结果可直接跳转到对应对象。</p>
  `;
  wrapper.appendChild(header);

  if (!(payload.items || []).length) {
    const empty = document.createElement('div');
    empty.className = 'query-empty';
    empty.textContent = payload.lines?.[0] || '未找到匹配结果。';
    wrapper.appendChild(empty);
    return wrapper;
  }

  const list = document.createElement('div');
  list.className = 'query-result-list';
  payload.items.forEach((item) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'query-result-item';
    button.addEventListener('click', () => runQuery(payload.mode, item.keyword || item.title));
    const meta = (item.meta || [])
      .filter((entry) => entry.value)
      .map((entry) => `<span class="pill">${entry.label} ${entry.value}</span>`)
      .join('');
    button.innerHTML = `
      <div class="query-result-title">${item.title}</div>
      <div class="query-result-subtitle">${item.subtitle || ''}</div>
      <div class="query-result-meta">${meta}</div>
    `;
    list.appendChild(button);
  });
  wrapper.appendChild(list);
  return wrapper;
}

function renderQueryResults(payload, runQuery) {
  const host = document.getElementById('query-results');
  host.replaceChildren();

  if (!payload || payload.view === 'empty') {
    const empty = document.createElement('div');
    empty.className = 'query-empty';
    empty.textContent = payload?.lines?.[0] || '请输入条件后执行查询。';
    host.appendChild(empty);
    return;
  }

  if (payload.view === 'list') {
    host.appendChild(queryListNode(payload, runQuery));
    return;
  }

  host.appendChild(querySummaryNode(payload));
  (payload.cards || []).forEach((card) => {
    const block = document.createElement('section');
    block.className = 'query-card';
    const title = document.createElement('h4');
    title.textContent = card.title;
    block.appendChild(title);
    if (card.kind === 'properties') {
      block.appendChild(detailRowsNode(card.items || []));
    } else if (card.kind === 'pins') {
      block.appendChild(dataTableNode(['pin', 'net'], card.items || []));
    } else if (card.kind === 'nodes') {
      block.appendChild(dataTableNode(['refdes', 'pin', 'pin_name', 'desc', '页面', '页码一一对应'], card.items || []));
    }
    host.appendChild(block);
  });
}

function bootSidebarToggle() {
  const button = document.getElementById('sidebar-toggle');
  if (!button) return;

  let collapsed = false;
  try {
    collapsed = window.localStorage.getItem(SIDEBAR_STORAGE_KEY) === '1';
  } catch {
    collapsed = false;
  }
  setSidebarCollapsed(collapsed);
  button.addEventListener('click', () => {
    const layout = document.querySelector('.report-layout');
    const nextCollapsed = !layout?.classList.contains('is-sidebar-collapsed');
    setSidebarCollapsed(nextCollapsed);
  });
}

async function handleQuery(runId) {
  const form = document.getElementById('query-form');
  const modeInput = form?.querySelector('[name="mode"]');
  const keywordInput = form?.querySelector('[name="keyword"]');
  if (!form) return;

  const runQuery = async (mode, keyword) => {
    const results = document.getElementById('query-results');
    results.innerHTML = '<div class="query-empty">正在查询…</div>';
    const response = await fetch(`/api/report/${runId}/query`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        mode,
        keyword,
      }),
    });
    const payload = await response.json();
    renderQueryResults(payload, runQuery);
  };

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    await runQuery(modeInput.value, keywordInput.value);
  });
}

async function handleReportPage() {
  const context = window.PSTX_REPORT_CONTEXT;
  if (!context?.runId) return;
  const response = await fetch(`/api/report/${context.runId}`);
  const report = await response.json();

  renderSummary(report);
  const host = document.getElementById('report-sections');
  report.sections.forEach((section) => host.appendChild(sectionNode(section)));
  renderSectionNav(report.sections);
  bootSidebarToggle();
  bootReveals();
  handleQuery(context.runId);
}

document.addEventListener('DOMContentLoaded', () => {
  bootReveals();
  if (document.body.dataset.page === 'home') {
    handleHomePage();
  }
  if (document.body.dataset.page === 'report') {
    handleReportPage();
  }
});
