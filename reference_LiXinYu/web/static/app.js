const revealObserver = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) {
      entry.target.classList.add('is-visible');
      revealObserver.unobserve(entry.target);
    }
  });
}, { threshold: 0.16 });

const SIDEBAR_STORAGE_KEY = 'pstx-report-sidebar-collapsed';
const INSPECTOR_STORAGE_KEY = 'pstx-report-inspector-collapsed';
const TABLE_COLUMN_STORAGE_PREFIX = 'pstx-report-table-columns:';
const TABLE_WIDTH_STORAGE_PREFIX = 'pstx-report-table-widths:';
const DEFAULT_COLUMN_WIDTH = 168;
const MIN_COLUMN_WIDTH = 96;
const MAX_COLUMN_WIDTH = 640;
const TABLE_INITIAL_RENDER_LIMIT = 220;
const TABLE_RENDER_STEP = 220;
const TABLE_FILTER_DEBOUNCE_MS = 90;
const MAX_STAGGERED_NODES_PER_SELECTOR = 80;
const LONG_TEXT_COLUMN_HINTS = new Map([
  ['左侧', 380],
  ['右侧', 380],
  ['左侧引脚名', 280],
  ['右侧引脚名', 280],
  ['左侧网络', 240],
  ['右侧网络', 240],
  ['变化字段', 260],
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

const {
  prefersReducedMotion,
  restartMotion,
  staggerChildren,
  runWhenBrowserIsIdle,
  showLoadingMask,
  hideLoadingMask,
  setStatus,
  scrollToSection,
  detailRowsNode,
  dataTableNode,
  normalizeText,
} = window.PSTXUI || {};

const GLOBAL_STAGGER_SELECTORS = [
  '.note-row',
  '.field',
  '.bullet-list li',
  '.metric',
  '.insight-card',
  '.section-card',
  '.table-block',
  '.section-nav a',
  '.inspector-block',
  '.query-card',
  '.query-result-item',
  '.detail-row',
  '.project-list-item',
  '.compare-stat',
  '.compare-block',
  '.debug-stage',
  '.aster-assist-panel',
  '.aster-summary-list li',
  '.aster-focus-grid article',
  '.aster-auth-item',
  '.sync-panel',
  '.feishu-sheet-item',
];

function spreadsheetColumnName(index) {
  let value = Math.max(1, Number(index) || 1);
  let name = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function defaultFeishuColumnRange(columnCount) {
  const count = Math.max(0, Number(columnCount) || 0);
  if (!count) return 'A:Z';
  return `A:${spreadsheetColumnName(Math.min(count, 702))}`;
}

const tableMountObserver = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (!entry.isIntersecting) return;
    const body = entry.target;
    tableMountObserver.unobserve(body);
    body.dataset.pendingMount = '';
    if (body.pstxTableData) {
      runWhenBrowserIsIdle(() => {
        if (body.dataset.mounted === 'true' || !body.pstxTableData) return;
        const tableData = body.pstxTableData;
        body.pstxTableData = null;
        mountTable(body, tableData);
      });
    }
  });
}, { rootMargin: '720px 0px', threshold: 0.01 });

function applyGlobalStaggers(root = document) {
  GLOBAL_STAGGER_SELECTORS.forEach((selector) => {
    Array.from(root.querySelectorAll(selector)).slice(0, MAX_STAGGERED_NODES_PER_SELECTOR).forEach((node, index) => {
      node.style.setProperty('--stagger-index', String(Math.min(index, 10)));
    });
  });
}

function bootRuntimeHints() {
  document.documentElement.classList.toggle('is-windows', /Windows/i.test(window.navigator.userAgent || ''));
}

function bootUiDebugMode() {
  const params = new URLSearchParams(window.location.search || '');
  const enabled = params.get('debug_ui') === '1' || document.body?.dataset.debugUi === 'true';
  if (!enabled || document.querySelector('.ui-debug-panel')) return;
  document.documentElement.classList.add('ui-debug-mode');
  const panel = document.createElement('aside');
  panel.className = 'ui-debug-panel';
  panel.setAttribute('aria-label', 'UI Debug 信息');
  const render = () => {
    const page = document.body?.dataset.page || 'unknown';
    const scrollHeight = Math.max(document.documentElement.scrollHeight, document.body?.scrollHeight || 0);
    panel.innerHTML = `
      <strong>UI Debug</strong>
      <span>page: ${page}</span>
      <span>viewport: ${window.innerWidth} × ${window.innerHeight}</span>
      <span>scroll: ${scrollHeight}px</span>
      <span>mode: ${document.body?.dataset.debugFixture === 'true' ? 'fixture' : 'live'}</span>
    `;
  };
  render();
  window.addEventListener('resize', render, { passive: true });
  document.body.appendChild(panel);
}

function animateTableOpen(body) {
  const block = body?.closest('.table-block');
  restartMotion(block, 'table-open-pulse', 360);
}

function bootPageMotion() {
  bootRuntimeHints();
  document.documentElement.classList.add('motion-ready');
  applyGlobalStaggers(document);
  requestAnimationFrame(() => {
    document.body.classList.add('page-motion-ready');
  });
}

function bootReveals(root = document) {
  const nodes = [
    ...(root.matches?.('[data-reveal]') ? [root] : []),
    ...Array.from(root.querySelectorAll?.('[data-reveal]') || []),
  ];
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

async function handleHomePage() {
  const form = document.getElementById('analyze-form');
  const mask = document.getElementById('loading-mask');
  if (!form) return;
  renderProjectManager();

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    setStatus('正在生成分析报告，请稍候…');
    showLoadingMask(mask);
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
      hideLoadingMask(mask);
      setStatus(error.message || String(error), true);
    }
  });
}

function feishuBasePayload(form) {
  const payload = Object.fromEntries(new FormData(form).entries());
  return {
    base_url: payload.base_url || '',
    origin: payload.origin || '',
    user_id: payload.user_id || '',
    spreadsheet_token_or_url: payload.spreadsheet_token_or_url || '',
    library_name: payload.library_name || '',
    library_id: payload.library_id || '',
  };
}

function setFeishuStatus(message, tone = 'neutral') {
  const node = document.getElementById('feishu-sync-status');
  if (!node) return;
  node.textContent = message;
  node.style.color = tone === 'error' ? 'var(--warn)' : tone === 'ok' ? 'var(--ok)' : 'var(--muted)';
}

function sheetField(sheetNode, name) {
  return sheetNode.querySelector(`[data-feishu-field="${name}"]`);
}

function parseFeishuOptionalFields(value) {
  return String(value || '')
    .split(/[，,;；\n]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((title) => ({ label: title, column: title, source: 'manual' }));
}

function formatFeishuOptionalFields(fields) {
  return (fields || [])
    .map((field) => field?.column || field?.label || field?.title || '')
    .filter(Boolean)
    .join('，');
}

function collectSheetConfig(sheetNode) {
  const specModelCol = sheetField(sheetNode, 'spec_model_col')?.value || '';
  const hqCodeCol = sheetField(sheetNode, 'hq_code_col')?.value || '';
  return {
    sheet_id: sheetNode.dataset.sheetId || '',
    title: sheetNode.dataset.sheetTitle || '',
    enabled: Boolean(sheetField(sheetNode, 'enabled')?.checked),
    header_row: Number(sheetField(sheetNode, 'header_row')?.value || 1),
    row_count: Number(sheetField(sheetNode, 'row_count')?.value || 5000),
    column_range: sheetField(sheetNode, 'column_range')?.value || 'A:Z',
    hq_code_col: hqCodeCol,
    spec_model_col: specModelCol,
    pi_col: sheetField(sheetNode, 'pi_col')?.value || '',
    selection_order_col: sheetField(sheetNode, 'selection_order_col')?.value || '',
    optional_fields: parseFeishuOptionalFields(sheetField(sheetNode, 'optional_fields_text')?.value || ''),
    key_col: specModelCol,
    hq_no_col: hqCodeCol,
    brand_col: sheetField(sheetNode, 'brand_col')?.value || '',
    spec_col: specModelCol,
    desc_col: sheetField(sheetNode, 'desc_col')?.value || '',
  };
}

function applyFeishuSuggestion(sheetNode, suggestion) {
  const mapping = suggestion?.mapping || {};
  const values = {
    header_row: suggestion?.header_row || 1,
    hq_code_col: mapping.hq_code_col || mapping.hq_no_col || '',
    spec_model_col: mapping.spec_model_col || mapping.key_col || mapping.spec_col || '',
    pi_col: mapping.pi_col || '',
    selection_order_col: mapping.selection_order_col || '',
    optional_fields_text: formatFeishuOptionalFields(mapping.optional_fields || []),
    brand_col: mapping.brand_col || '',
    desc_col: mapping.desc_col || '',
  };
  Object.entries(values).forEach(([key, value]) => {
    const input = sheetField(sheetNode, key);
    if (input && value) input.value = value;
  });
}

function renderFeishuPreview(host, payload, suggestionPayload = null) {
  if (!host) return;
  host.replaceChildren();
  const meta = document.createElement('p');
  meta.className = 'feishu-preview-meta';
  const logFile = payload.online_debug_log_file ? ` · 日志 ${payload.online_debug_log_file}` : '';
  meta.textContent = `Sheet ${payload.sheet_id || ''} · ${payload.row_count || 0} 行预览 · 表头行 ${payload.header_row || 1}${logFile}`;
  host.appendChild(meta);

  const suggestion = suggestionPayload?.suggestion || payload.mapping_suggestion;
  if (suggestion) {
    const note = document.createElement('div');
    note.className = 'feishu-suggestion-note';
    const mapping = suggestion.mapping || {};
    const headerRole = suggestion.header_detection ? 'Agent 已识别表头与扩展字段；' : '';
    const optionalCount = (mapping.optional_fields || []).length;
    note.textContent = `${headerRole}建议：表头行 ${suggestion.header_row || 1}，规格型号 ${mapping.spec_model_col || mapping.key_col || '未识别'}，HQ料号 ${mapping.hq_code_col || mapping.hq_no_col || '未识别'}，PI ${mapping.pi_col || '未识别'}，选型顺序 ${mapping.selection_order_col || '未识别'}，扩展字段 ${optionalCount} 个。${(suggestion.notes || []).join(' ')}`;
    host.appendChild(note);
  }

  const rows = payload.rows || [];
  if (!rows.length) {
    const empty = document.createElement('p');
    empty.className = 'query-empty';
    empty.textContent = '没有可预览的行。';
    host.appendChild(empty);
    return;
  }

  const width = Math.max(...rows.map((row) => row.length), 0);
  const wrap = document.createElement('div');
  wrap.className = 'feishu-preview-table-wrap';
  const table = document.createElement('table');
  table.className = 'feishu-preview-table';
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  for (let index = 0; index < width; index += 1) {
    const th = document.createElement('th');
    th.textContent = spreadsheetColumnName(index + 1);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);
  const tbody = document.createElement('tbody');
  rows.slice(0, 16).forEach((row, rowIndex) => {
    const tr = document.createElement('tr');
    tr.dataset.rowNumber = String(rowIndex + 1);
    for (let index = 0; index < width; index += 1) {
      const td = document.createElement('td');
      td.textContent = row[index] || '';
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  table.append(thead, tbody);
  wrap.appendChild(table);
  host.appendChild(wrap);
  applyGlobalStaggers(host);
}

function renderFeishuSheetItem(sheet, basePayload, previewHost) {
  const item = document.createElement('article');
  item.className = 'feishu-sheet-item';
  item.dataset.sheetId = sheet.sheet_id || sheet.sheetId || '';
  item.dataset.sheetTitle = sheet.title || item.dataset.sheetId;

  const head = document.createElement('div');
  head.className = 'feishu-sheet-head';
  const title = document.createElement('label');
  title.className = 'feishu-sheet-title';
  const enabled = document.createElement('input');
  enabled.type = 'checkbox';
  enabled.checked = true;
  enabled.dataset.feishuField = 'enabled';
  const titleCopy = document.createElement('span');
  const strong = document.createElement('strong');
  strong.textContent = sheet.title || '未命名 Sheet';
  const small = document.createElement('small');
  small.textContent = `${item.dataset.sheetId} · ${sheet.row_count || 0} 行`;
  titleCopy.append(strong, small);
  title.append(enabled, titleCopy);

  const actions = document.createElement('div');
  actions.className = 'feishu-sheet-actions';
  const previewButton = document.createElement('button');
  previewButton.type = 'button';
  previewButton.className = 'ghost-btn';
  previewButton.textContent = '预览';
  const suggestButton = document.createElement('button');
  suggestButton.type = 'button';
  suggestButton.className = 'ghost-btn';
  suggestButton.textContent = '预览/建议';
  actions.append(previewButton, suggestButton);
  head.append(title, actions);

  const grid = document.createElement('div');
  grid.className = 'feishu-mapping-grid';
  const defaultColumnRange = sheet.column_range || sheet.columnRange || defaultFeishuColumnRange(sheet.column_count || sheet.columnCount);
  const fields = [
    ['header_row', '表头行', '1', 'number'],
    ['row_count', '同步行数', String(Math.max(sheet.row_count || 5000, 50)), 'number'],
    ['column_range', '读取列范围', defaultColumnRange, 'text'],
    ['hq_code_col', 'HQ料号 / HQ编码 / 物料编码', '', 'text'],
    ['spec_model_col', '规格型号 / Part Number', '', 'text'],
    ['pi_col', 'PI', '', 'text'],
    ['selection_order_col', '选型顺序', '', 'text'],
    ['optional_fields_text', '扩展字段', '', 'text'],
  ];
  fields.forEach(([name, labelText, value, type]) => {
    const label = document.createElement('label');
    label.className = 'field';
    const span = document.createElement('span');
    span.textContent = labelText;
    const input = document.createElement('input');
    input.type = type;
    input.value = value;
    input.dataset.feishuField = name;
    if (name === 'header_row') input.min = '1';
    if (name === 'row_count') {
      input.min = '50';
      input.max = '10000';
    }
    if (name === 'optional_fields_text') {
      input.placeholder = '例如：封装，耐压，容值，温度等级';
    }
    label.append(span, input);
    grid.appendChild(label);
  });

  const runPreview = async (applySuggestion = false) => {
    previewButton.disabled = true;
    suggestButton.disabled = true;
    setFeishuStatus(`正在读取 ${sheet.title || item.dataset.sheetId} 预览…`);
    try {
      const config = collectSheetConfig(item);
      const response = await fetch('/api/feishu-bom/preview-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...basePayload(),
          sheet_id: config.sheet_id,
          header_row: config.header_row,
          row_count: 50,
          column_range: config.column_range,
        }),
      });
      const payload = await response.json();
      if (!response.ok || payload.ok === false) throw new Error(payload.error || '预览失败。');
      item.pstxPreviewRows = payload.rows || [];
      let suggestionPayload = null;
      if (applySuggestion) {
        const suggestResponse = await fetch('/api/feishu-bom/suggest-mapping', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            rows: item.pstxPreviewRows,
            sheet_title: sheet.title || '',
            use_agent: Boolean(document.getElementById('feishu-use-agent-mapping')?.checked),
          }),
        });
        suggestionPayload = await suggestResponse.json();
        if (!suggestResponse.ok || suggestionPayload.ok === false) {
          throw new Error(suggestionPayload.error || '字段建议失败。');
        }
        applyFeishuSuggestion(item, suggestionPayload.suggestion);
      } else if (payload.mapping_suggestion) {
        applyFeishuSuggestion(item, payload.mapping_suggestion);
      }
      renderFeishuPreview(previewHost, payload, suggestionPayload);
      setFeishuStatus(`${sheet.title || item.dataset.sheetId} 预览完成。`, 'ok');
    } catch (error) {
      setFeishuStatus(error.message || String(error), 'error');
    } finally {
      previewButton.disabled = false;
      suggestButton.disabled = false;
    }
  };
  previewButton.addEventListener('click', () => runPreview(false));
  suggestButton.addEventListener('click', () => runPreview(true));

  item.append(head, grid);
  return item;
}

async function refreshFeishuCacheStatus() {
  const host = document.getElementById('feishu-cache-status');
  if (!host) return;
  host.textContent = '正在读取缓存状态…';
  try {
    const response = await fetch('/api/feishu-bom/status');
    const payload = await response.json();
    const stats = (payload.cache_stats || []).map((item) => `${item.lib_name || item.lib_id}: ${item.count}`).join('；') || '暂无库缓存';
    const logFile = payload.online_debug_log_file ? ` · 日志 ${payload.online_debug_log_file}` : '';
    host.textContent = `${payload.available ? '缓存目录可用' : '未找到缓存目录'} · 库 ${payload.library_count || 0} · 记录 ${payload.cache_count || 0} · ${stats}${logFile}`;
  } catch (error) {
    host.textContent = `缓存状态读取失败：${error.message || error}`;
  }
}

function renderFeishuSyncResult(host, payload, sheets) {
  if (!host) return;
  host.replaceChildren();
  const summary = document.createElement('div');
  summary.className = 'feishu-sync-summary';
  summary.innerHTML = `
    <strong>同步完成</strong>
    <span>库 ${payload.library_name || payload.library_id || ''} · 写入 ${payload.synced_rows || 0} 行 · 跳过 ${payload.skipped_sheets || 0} 个 Sheet</span>
    ${payload.online_debug_log_file ? `<small>飞书在线解析日志：${payload.online_debug_log_file}</small>` : ''}
  `;
  host.appendChild(summary);

  const list = document.createElement('div');
  list.className = 'feishu-sync-sheet-results';
  const sheetMap = new Map((sheets || []).map((sheet) => [sheet.sheet_id, sheet]));
  (payload.per_sheet || []).forEach((sheet) => {
    const config = sheetMap.get(sheet.sheet_id) || {};
    const article = document.createElement('article');
    article.className = `feishu-sync-sheet-result is-${sheet.status || 'unknown'}`;
    article.innerHTML = `
      <strong>${sheet.title || sheet.sheet_id || 'Sheet'}</strong>
      <p>${sheet.status === 'synced' ? '已同步' : '已跳过'} · ${sheet.row_count || 0} 行 · 表头行 ${sheet.header_row || config.header_row || 1}</p>
      <small>规格型号：${config.spec_model_col || '未配置'}；HQ料号：${config.hq_code_col || '未配置'}；PI：${config.pi_col || '未配置'}；选型顺序：${config.selection_order_col || '未配置'}；扩展字段：${(config.optional_fields || []).map((field) => field.column || field.label).filter(Boolean).join('，') || '无'}；原因：${sheet.reason || '无'}</small>
    `;
    list.appendChild(article);
  });
  host.appendChild(list);
}

async function handleFeishuSyncPage() {
  const form = document.getElementById('feishu-sync-form');
  const listHost = document.getElementById('feishu-sheet-list');
  const previewHost = document.getElementById('feishu-preview-host');
  const syncButton = document.getElementById('feishu-sync-selected');
  const resultHost = document.getElementById('feishu-sync-result');
  const refreshButton = document.getElementById('feishu-refresh-status');
  if (!form || !listHost || !previewHost || !syncButton) return;

  const getBasePayload = () => feishuBasePayload(form);
  await refreshFeishuCacheStatus();
  refreshButton?.addEventListener('click', refreshFeishuCacheStatus);

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    setFeishuStatus('正在获取 Sheet 列表…');
    listHost.innerHTML = '<p class="query-empty">正在连接飞书网关…</p>';
    try {
      const response = await fetch('/api/feishu-bom/sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(getBasePayload()),
      });
      const payload = await response.json();
      if (!response.ok || payload.ok === false) throw new Error(payload.error || 'Sheet 列表获取失败。');
      listHost.replaceChildren();
      (payload.sheets || []).forEach((sheet) => {
        listHost.appendChild(renderFeishuSheetItem(sheet, getBasePayload, previewHost));
      });
      if (!(payload.sheets || []).length) {
        listHost.innerHTML = '<p class="query-empty">该表格没有返回可用 Sheet。</p>';
      }
      applyGlobalStaggers(listHost);
      setFeishuStatus(`获取到 ${payload.sheet_count || 0} 个 Sheet。`, 'ok');
    } catch (error) {
      listHost.innerHTML = '<p class="query-empty">Sheet 获取失败，请检查连接参数。</p>';
      setFeishuStatus(error.message || String(error), 'error');
    }
  });

  syncButton.addEventListener('click', async () => {
    const sheets = Array.from(listHost.querySelectorAll('.feishu-sheet-item'))
      .map(collectSheetConfig)
      .filter((sheet) => sheet.enabled);
    if (!sheets.length) {
      setFeishuStatus('请至少勾选一个 Sheet。', 'error');
      return;
    }
    syncButton.disabled = true;
    resultHost.textContent = '正在同步选中 Sheet 到本地缓存…';
    try {
      const response = await fetch('/api/feishu-bom/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...getBasePayload(),
          sheets,
        }),
      });
      const payload = await response.json();
      if (!response.ok || payload.ok === false) throw new Error(payload.error || '同步失败。');
      renderFeishuSyncResult(resultHost, payload, sheets);
      setFeishuStatus('本地缓存已更新，可以继续进入项目分析。', 'ok');
      await refreshFeishuCacheStatus();
    } catch (error) {
      resultHost.textContent = error.message || String(error);
      setFeishuStatus(error.message || String(error), 'error');
    } finally {
      syncButton.disabled = false;
    }
  });

  bootAsterStatus();
  bootAsterCredentialForm();
}

function renderFeishuDbLibraries(host, payload, selectLibrary) {
  host.replaceChildren();
  const libraries = payload.libraries || [];
  if (!libraries.length) {
    host.innerHTML = '<p class="query-empty">还没有保存的飞书缓存库。请先去飞书同步页同步一个表格。</p>';
    return;
  }
  libraries.forEach((library) => {
    const article = document.createElement('article');
    article.className = 'feishu-db-library';
    const sheetStats = (library.sheet_stats || []).map((sheet) => `${sheet.sheet_name}: ${sheet.count}`).join('；') || '暂无 Sheet 统计';
    article.innerHTML = `
      <div class="feishu-db-library-head">
        <div>
          <strong>${library.lib_name || library.lib_id}</strong>
          <p>${library.lib_id} · ${library.cache_count || 0} 行 · ${library.last_synced_at || '未同步'}</p>
        </div>
        <div class="feishu-sheet-actions">
          <button type="button" class="ghost-btn" data-action="view">查看行</button>
          <button type="button" class="ghost-btn" data-action="delete">删除库</button>
        </div>
      </div>
      <p class="feishu-sheet-meta">${sheetStats}</p>
    `;
    article.querySelector('[data-action="view"]')?.addEventListener('click', () => selectLibrary(library.lib_id));
    article.querySelector('[data-action="delete"]')?.addEventListener('click', async () => {
      const confirmed = window.confirm(`确认删除本地缓存库 ${library.lib_name || library.lib_id}？不会修改飞书源表。`);
      if (!confirmed) return;
      const response = await fetch(`/api/feishu-bom/database/libraries/${encodeURIComponent(library.lib_id)}`, { method: 'DELETE' });
      const result = await response.json();
      if (!response.ok || result.ok === false) {
        window.alert(result.error || '删除失败。');
        return;
      }
      window.location.reload();
    });
    host.appendChild(article);
  });
  applyGlobalStaggers(host);
}

function renderFeishuDbRows(host, payload, options = {}) {
  if (!options.append) host.replaceChildren();
  if (!payload.ok) {
    host.innerHTML = `<p class="query-empty">${payload.error || '缓存行读取失败。'}</p>`;
    return;
  }
  const rows = payload.rows || [];
  if (!rows.length && !options.append) {
    host.innerHTML = '<p class="query-empty">没有匹配的缓存行。</p>';
    return;
  }
  if (!rows.length) return;
  const meta = document.createElement('p');
  meta.className = 'feishu-preview-meta';
  const shown = (payload.offset || 0) + rows.length;
  meta.textContent = `共 ${payload.total || 0} 行，当前显示到第 ${shown} 行。${payload.has_more ? ' 还有更多，可继续加载。' : ''}`;
  host.appendChild(meta);
  const wrap = document.createElement('div');
  wrap.className = 'feishu-preview-table-wrap';
  const table = document.createElement('table');
  table.className = 'feishu-preview-table';
  const columns = ['id', 'lib_name', 'sheet_name', 'key_value', 'hq_no', 'pi', 'selection_order', 'brand', 'spec', 'description', 'extra_fields', 'synced_at'];
  const labels = ['ID', '库', 'Sheet', '规格型号', 'HQ料号', 'PI', '选型顺序', '制造商', '规格', '描述', '扩展字段', '同步时间'];
  const thead = document.createElement('thead');
  const trHead = document.createElement('tr');
  labels.forEach((label) => {
    const th = document.createElement('th');
    th.textContent = label;
    trHead.appendChild(th);
  });
  const actionHead = document.createElement('th');
  actionHead.className = 'feishu-row-action-head';
  actionHead.textContent = '操作';
  trHead.appendChild(actionHead);
  thead.appendChild(trHead);
  const tbody = document.createElement('tbody');
  rows.forEach((row) => {
    const tr = document.createElement('tr');
    columns.forEach((column) => {
      const td = document.createElement('td');
      td.textContent = column === 'extra_fields'
        ? formatFeishuOptionalFields(Object.entries(row.extra_field_values || {}).map(([label, value]) => ({ label, column: `${label}:${value}` })))
        : row[column] || '';
      tr.appendChild(td);
    });
    const actionCell = document.createElement('td');
    actionCell.className = 'feishu-row-actions';
    const editButton = document.createElement('button');
    editButton.type = 'button';
    editButton.className = 'ghost-btn';
    editButton.textContent = '编辑';
    editButton.title = '编辑这条本地缓存行，不修改飞书源表。';
    editButton.addEventListener('click', () => options.onEditRow?.(row));
    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'ghost-btn danger-ghost-btn';
    deleteButton.textContent = '剔除';
    deleteButton.title = '仅删除这条本地缓存行，不修改飞书源表。';
    deleteButton.addEventListener('click', async () => {
      if (!row.id) return;
      const label = row.key_value || row.hq_no || `ID ${row.id}`;
      const confirmed = window.confirm(`确认从本地缓存剔除 ${label}？不会修改飞书源表。`);
      if (!confirmed) return;
      deleteButton.disabled = true;
      deleteButton.textContent = '剔除中…';
      try {
        await options.onDeleteRow?.(row);
      } catch (error) {
        window.alert(error.message || String(error));
        deleteButton.disabled = false;
        deleteButton.textContent = '剔除';
      }
    });
    actionCell.append(editButton, deleteButton);
    tr.appendChild(actionCell);
    tbody.appendChild(tr);
  });
  table.append(thead, tbody);
  wrap.appendChild(table);
  host.appendChild(wrap);
  if (payload.has_more && options.onLoadMore) {
    const loadMore = document.createElement('button');
    loadMore.type = 'button';
    loadMore.className = 'ghost-btn inline-btn';
    loadMore.textContent = `加载更多（从第 ${payload.next_offset + 1} 行）`;
    loadMore.addEventListener('click', () => options.onLoadMore(payload.next_offset || shown));
    host.appendChild(loadMore);
  }
}

async function bootFeishuDbPage() {
  const libraryHost = document.getElementById('feishu-db-libraries');
  const summaryHost = document.getElementById('feishu-db-summary');
  const rowsHost = document.getElementById('feishu-db-rows');
  const pathHost = document.getElementById('feishu-db-path');
  const filterForm = document.getElementById('feishu-db-row-filter');
  const editorForm = document.getElementById('feishu-db-row-editor');
  const addRowButton = document.getElementById('feishu-db-add-row');
  const loadAllButton = document.getElementById('feishu-db-load-all');
  const cancelEditButton = document.getElementById('feishu-db-cancel-edit');
  if (!libraryHost || !rowsHost) return;
  let currentLibId = '';
  let currentQuery = '';
  let currentOffset = 0;

  const editorFields = ['id', 'lib_id', 'lib_name', 'sheet_name', 'key_value', 'hq_no', 'pi', 'selection_order', 'brand', 'spec', 'description', 'extra_fields'];

  const fillEditor = (row = {}) => {
    if (!editorForm) return;
    editorFields.forEach((field) => {
      const input = editorForm.elements[field];
      if (!input) return;
      if (field === 'extra_fields') {
        input.value = JSON.stringify(row.extra_field_values || {}, null, 0);
      } else {
        input.value = row[field] || '';
      }
    });
    if (!editorForm.elements.lib_id?.value && currentLibId) {
      editorForm.elements.lib_id.value = currentLibId;
    }
    if (!editorForm.elements.lib_name?.value) {
      const selectedLib = [...libraryHost.querySelectorAll('.feishu-db-library strong')]
        .map((node) => node.textContent || '')
        .find(Boolean);
      editorForm.elements.lib_name.value = selectedLib || currentLibId || '手工维护';
    }
    if (!editorForm.elements.sheet_name?.value) {
      editorForm.elements.sheet_name.value = '手工维护';
    }
    editorForm.hidden = false;
    editorForm.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  };

  const editorPayload = () => {
    const payload = {};
    editorFields.forEach((field) => {
      if (field === 'id') return;
      const input = editorForm?.elements[field];
      if (!input) return;
      payload[field] = input.value || '';
    });
    return payload;
  };

  const loadOverview = async () => {
    const response = await fetch('/api/feishu-bom/database');
    const payload = await response.json();
    if (pathHost) pathHost.textContent = payload.cache_file || '未找到';
    if (summaryHost) {
      summaryHost.textContent = `${payload.available ? '缓存目录可用' : '未找到缓存目录'} · 库 ${payload.library_count || 0} · 记录 ${payload.cache_count || 0}`;
    }
    renderFeishuDbLibraries(libraryHost, payload, selectLibrary);
    return payload;
  };

  const loadRows = async (libId = '', query = '', { offset = 0, append = false, limit = 250 } = {}) => {
    currentLibId = libId;
    currentQuery = query;
    currentOffset = offset;
    const params = new URLSearchParams({ lib_id: libId, query, limit: String(limit), offset: String(offset) });
    if (!append) rowsHost.innerHTML = '<p class="query-empty">正在读取缓存行…</p>';
    const response = await fetch(`/api/feishu-bom/database/rows?${params.toString()}`);
    const payload = await response.json();
    renderFeishuDbRows(rowsHost, payload, {
      append,
      onEditRow: (row) => fillEditor(row),
      onLoadMore: async (nextOffset) => loadRows(currentLibId, currentQuery, { offset: nextOffset, append: true, limit }),
      onDeleteRow: async (row) => {
        const deleteResponse = await fetch(`/api/feishu-bom/database/rows/${encodeURIComponent(row.id)}`, { method: 'DELETE' });
        const result = await deleteResponse.json();
        if (!deleteResponse.ok || result.ok === false) {
          throw new Error(result.error || '缓存行剔除失败。');
        }
        await loadOverview();
        await loadRows(currentLibId, currentQuery);
      },
    });
  };

  const loadAllRows = async () => {
    rowsHost.innerHTML = '<p class="query-empty">正在分批加载全部缓存行…</p>';
    let offset = 0;
    let first = true;
    while (true) {
      const params = new URLSearchParams({
        lib_id: currentLibId,
        query: currentQuery,
        limit: '5000',
        offset: String(offset),
      });
      const response = await fetch(`/api/feishu-bom/database/rows?${params.toString()}`);
      const payload = await response.json();
      if (!response.ok || payload.ok === false) {
        rowsHost.innerHTML = `<p class="query-empty">${payload.error || '全量加载失败。'}</p>`;
        return;
      }
      renderFeishuDbRows(rowsHost, payload, {
        append: !first,
        onEditRow: (row) => fillEditor(row),
        onLoadMore: async (nextOffset) => loadRows(currentLibId, currentQuery, { offset: nextOffset, append: true, limit: 5000 }),
        onDeleteRow: async (row) => {
          const deleteResponse = await fetch(`/api/feishu-bom/database/rows/${encodeURIComponent(row.id)}`, { method: 'DELETE' });
          const result = await deleteResponse.json();
          if (!deleteResponse.ok || result.ok === false) throw new Error(result.error || '缓存行剔除失败。');
          await loadOverview();
          await loadRows(currentLibId, currentQuery);
        },
      });
      first = false;
      if (!payload.has_more) break;
      offset = payload.next_offset || offset + (payload.rows || []).length;
    }
  };

  const selectLibrary = async (libId) => {
    if (filterForm?.elements.lib_id) filterForm.elements.lib_id.value = libId;
    await loadRows(libId, filterForm?.elements.query?.value || '');
  };

  await loadOverview();

  filterForm?.addEventListener('submit', async (event) => {
    event.preventDefault();
    await loadRows(filterForm.elements.lib_id?.value || '', filterForm.elements.query?.value || '');
  });

  addRowButton?.addEventListener('click', () => fillEditor({ lib_id: currentLibId, lib_name: currentLibId, sheet_name: '手工维护' }));
  loadAllButton?.addEventListener('click', loadAllRows);
  cancelEditButton?.addEventListener('click', () => {
    if (editorForm) editorForm.hidden = true;
  });
  editorForm?.addEventListener('submit', async (event) => {
    event.preventDefault();
    const rowId = editorForm.elements.id?.value || '';
    const method = rowId ? 'PATCH' : 'POST';
    const url = rowId
      ? `/api/feishu-bom/database/rows/${encodeURIComponent(rowId)}`
      : '/api/feishu-bom/database/rows';
    const response = await fetch(url, {
      method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(editorPayload()),
    });
    const result = await response.json();
    if (!response.ok || result.ok === false) {
      window.alert(result.error || '保存失败。');
      return;
    }
    editorForm.hidden = true;
    await loadOverview();
    await loadRows(currentLibId || result.row?.lib_id || '', currentQuery);
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

function scheduleAutoRenderMoreRows(state, options = {}) {
  if (!state || !state.scroll || state.autoRenderFrame) return;
  state.autoRenderFrame = true;
  requestAnimationFrame(() => {
    state.autoRenderFrame = false;
    const scroll = state.scroll;
    const visibleCount = Number(state.visibleRecordCount) || 0;
    const currentLimit = Number(state.renderLimit) || TABLE_INITIAL_RENDER_LIMIT;
    const remaining = Math.max(visibleCount - currentLimit, 0);
    if (!remaining) return;

    const nearBottom = scroll.scrollTop + scroll.clientHeight >= scroll.scrollHeight - 180;
    const notScrollableYet = scroll.scrollHeight <= scroll.clientHeight + 16;
    if (!nearBottom && !(options.allowWhenNotScrollable && notScrollableYet)) return;

    const previousScrollTop = scroll.scrollTop;
    state.renderLimit = Math.min(currentLimit + TABLE_RENDER_STEP, visibleCount);
    applyTableState(state);
    scroll.scrollTop = previousScrollTop;
  });
}

function scheduleTableMount(container, tableData) {
  if (container.dataset.mounted === 'true') return;
  container.pstxTableData = tableData;
  container.dataset.pendingMount = 'true';
  if ('IntersectionObserver' in window) {
    tableMountObserver.observe(container);
  } else {
    runWhenBrowserIsIdle(() => mountTable(container, tableData));
  }
  requestAnimationFrame(() => {
    if (container.dataset.mounted === 'true' || !container.pstxTableData) return;
    const rect = container.getBoundingClientRect();
    const margin = 720;
    if (rect.top < window.innerHeight + margin && rect.bottom > -margin) {
      tableMountObserver.unobserve(container);
      container.dataset.pendingMount = '';
      runWhenBrowserIsIdle(() => {
        if (container.dataset.mounted === 'true' || !container.pstxTableData) return;
        const pendingTableData = container.pstxTableData;
        container.pstxTableData = null;
        mountTable(container, pendingTableData);
      });
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
  state.visibleRecordCount = visibleRecords.length;
  state.renderedRecordCount = renderLimit;
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
  if (state.renderStatusNode) {
    const remaining = Math.max(visibleRecords.length - renderLimit, 0);
    state.renderStatusNode.textContent = remaining
      ? `已按需渲染 ${renderLimit} / ${visibleRecords.length} 行，继续向下滚动会自动追加剩余 ${remaining} 行。`
      : `当前筛选结果已全部渲染。`;
  }
  state.sortSelect.disabled = state.sortModeSelect.value !== 'column';
  state.sortDirSelect.disabled = state.sortModeSelect.value === 'column' && !state.sortSelect.value;
  state.table.classList.toggle('density-compact', state.density === 'compact');
  state.table.classList.toggle('density-comfortable', state.density === 'comfortable');
  state.densityButton.textContent = state.density === 'compact' ? '紧凑行距' : '舒展行距';
  updateSortHeaders(state);
  updateScrollShadows(state.scroll);
  scheduleAutoRenderMoreRows(state, { allowWhenNotScrollable: true });
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
    density: tableData.default_density === 'comfortable' ? 'comfortable' : 'compact',
    renderLimit: TABLE_INITIAL_RENDER_LIMIT,
    renderStatusNode: renderFooter.querySelector('.table-render-status'),
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
  scroll.addEventListener('scroll', () => {
    scheduleScrollShadowUpdate(scroll);
    scheduleAutoRenderMoreRows(state);
  }, { passive: true });
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
  animateTableOpen(container);
}

const REPORT_TABLE_LEVELS = [
  {
    id: 'focus',
    title: '重点',
    lead: '高风险、必须优先人工复核的条目。',
  },
  {
    id: 'review',
    title: '常规',
    lead: '需要人工判断的候选项和规则提示。',
  },
  {
    id: 'info',
    title: '信息概览',
    lead: '用于理解项目范围、模块、网络和物料结构的基础信息。',
  },
  {
    id: 'debug',
    title: 'Debug / 证据明细',
    lead: '页码映射、索引、覆盖证据和辅助追溯材料，按需展开。',
  },
];

function reportLevelMeta(level) {
  return REPORT_TABLE_LEVELS.find((item) => item.id === level) || REPORT_TABLE_LEVELS[2];
}

function groupReportTablesByLevel(tables) {
  return REPORT_TABLE_LEVELS.map((meta) => {
    const levelTables = tables.filter((table) => (table.display_level || 'info') === meta.id);
    return {
      meta,
      tables: levelTables,
      activeTables: levelTables.filter((table) => Number(table.count || 0) > 0),
      quietTables: levelTables.filter((table) => Number(table.count || 0) <= 0),
    };
  }).filter((group) => group.tables.length > 0);
}

function tableBlock(tableData, initialOpen = false) {
  const block = document.createElement('article');
  const levelMeta = reportLevelMeta(tableData.display_level || 'info');
  block.className = `table-block table-level-${levelMeta.id}`;
  block.dataset.tableId = tableData.id || '';
  const rowCount = Number(tableData.count || 0);
  if (rowCount <= 0) {
    block.classList.add('is-empty');
  }

  const kindPills = Object.entries(tableData.kind_counts || {})
    .map(([label, value]) => `<span class="pill">${label} ${value}</span>`)
    .join('');

  block.innerHTML = `
    <div class="table-header">
      <div>
        <h3 class="table-title">${tableData.title}</h3>
        <div class="table-badge-row">
          <span class="table-level-badge">${tableData.display_level_label || levelMeta.title}</span>
          <span class="table-trust-badge trust-${tableData.trust_tone || 'info'}">${tableData.trust_label || '信息统计'}</span>
        </div>
      </div>
      <div class="table-meta">
        <span class="pill">记录 ${tableData.count}</span>
        ${kindPills}
      </div>
      <button type="button" class="toggle-btn"${rowCount <= 0 ? ' disabled' : ''}>${rowCount <= 0 ? '无明细' : '查看详情'}</button>
    </div>
    <div class="table-body"></div>
  `;

  const button = block.querySelector('.toggle-btn');
  const body = block.querySelector('.table-body');
  if (rowCount <= 0) {
    body.innerHTML = '<p class="empty-state table-empty-state">该子表没有需要展示的记录。</p>';
    return block;
  }
  const setOpen = (open) => {
    block.classList.toggle('is-open', open);
    button.textContent = open ? '收起' : '查看详情';
    if (open) {
      mountTable(body, tableData);
      animateTableOpen(body);
    }
  };
  button.addEventListener('click', () => {
    const open = !block.classList.contains('is-open');
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

function openReportTable(sectionId, tableId) {
  if (sectionId) {
    scrollToSection(sectionId);
  }
  window.setTimeout(() => {
    const blocks = Array.from(document.querySelectorAll('.table-block'));
    const block = blocks.find((item) => item.dataset.tableId === tableId);
    if (!block) return;
    const group = block.closest('details.report-level-group');
    if (group) {
      group.open = true;
    }
    const toggle = block.querySelector('.toggle-btn');
    if (toggle && !block.classList.contains('is-open') && !toggle.disabled) {
      toggle.click();
    }
    block.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 180);
}

function reportLevelGroupNode(group, defaultOpen = false) {
  const detail = document.createElement('details');
  detail.className = `report-level-group level-${group.meta.id}`;
  if (defaultOpen) {
    detail.open = true;
  }
  const activeCount = group.activeTables.length;
  const quietCount = group.quietTables.length;
  const rowCount = group.activeTables.reduce((sum, table) => sum + Number(table.count || 0), 0);
  detail.innerHTML = `
    <summary>
      <span class="report-level-title">${group.meta.title}</span>
      <span class="report-level-copy">${group.meta.lead}</span>
      <span class="report-level-count">${activeCount} 表 · ${rowCount} 条</span>
    </summary>
    <div class="report-level-body"></div>
  `;
  const body = detail.querySelector('.report-level-body');
  if (activeCount) {
    group.activeTables.forEach((table) => body.appendChild(tableBlock(table, false)));
  } else {
    const empty = document.createElement('div');
    empty.className = 'section-empty-state';
    empty.textContent = '这一层当前没有需要展示的记录。';
    body.appendChild(empty);
  }
  if (quietCount) {
    const quiet = document.createElement('details');
    quiet.className = 'quiet-table-group';
    quiet.innerHTML = `
      <summary>查看 ${quietCount} 个无结果子表</summary>
      <div class="quiet-table-list">
        ${group.quietTables.map((table) => `<span>${table.title}</span>`).join('')}
      </div>
    `;
    body.appendChild(quiet);
  }
  return detail;
}

function sectionNode(section) {
  const wrapper = document.createElement('section');
  wrapper.id = section.id;
  wrapper.className = 'report-section';
  wrapper.setAttribute('data-reveal', '');
  const tables = Array.isArray(section.tables) ? section.tables : [];
  const groups = groupReportTablesByLevel(tables);
  const activeTables = tables.filter((table) => Number(table.count || 0) > 0);
  const quietTables = tables.filter((table) => Number(table.count || 0) <= 0);
  const scanMeta = groups
    .map((group) => `<span>${group.meta.title} ${group.activeTables.length}</span>`)
    .join('');
  wrapper.innerHTML = `
    <div class="section-heading">
      <p class="eyebrow">${section.id.toUpperCase()}</p>
      <h2>${section.title}</h2>
      <p>${section.lead}</p>
      <div class="section-scan-meta">
        <span>${section.total_rows || 0} 条记录</span>
        <span>${activeTables.length} 个子表有结果</span>
        ${scanMeta}
        ${quietTables.length ? `<span>${quietTables.length} 个子表无结果已收纳</span>` : ''}
      </div>
    </div>
  `;

  const stack = document.createElement('div');
  stack.className = 'table-stack';
  if (groups.length) {
    const defaultOpenGroup =
      groups.find((group) => group.meta.id === 'focus' && group.activeTables.length) ||
      groups.find((group) => group.meta.id === 'review' && group.activeTables.length) ||
      groups.find((group) => group.meta.id === 'info' && group.activeTables.length);
    groups.forEach((group) => {
      stack.appendChild(reportLevelGroupNode(group, Boolean(defaultOpenGroup && group === defaultOpenGroup)));
    });
  } else {
    const empty = document.createElement('div');
    empty.className = 'section-empty-state';
    empty.textContent = '本分区当前没有待展开的明细记录。';
    stack.appendChild(empty);
  }
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

function reviewPlanItemNode(item) {
  const node = document.createElement('article');
  const levelMeta = reportLevelMeta(item.level || 'info');
  node.className = `review-plan-card level-${levelMeta.id}`;

  const header = document.createElement('div');
  header.className = 'review-plan-card-head';
  const titleWrap = document.createElement('div');
  const eyebrow = document.createElement('span');
  eyebrow.className = 'review-plan-eyebrow';
  eyebrow.textContent = `${levelMeta.title} · ${item.trust_label || '信息统计'} · ${item.category || '报告'}`;
  const title = document.createElement('h4');
  title.textContent = item.title || '未命名审查项';
  titleWrap.append(eyebrow, title);
  const count = document.createElement('span');
  count.className = 'review-plan-count';
  count.textContent = `${item.count || 0}`;
  header.append(titleWrap, count);

  const summary = document.createElement('p');
  summary.className = 'review-plan-summary-text';
  summary.textContent = item.summary || '暂无摘要。';

  const trust = document.createElement('div');
  trust.className = `review-plan-trust trust-${item.trust_tone || 'info'}`;
  trust.textContent = item.trust_note || '用于报告复核和证据定位。';

  const meta = document.createElement('div');
  meta.className = 'review-plan-meta';
  const addMetaPill = (label, values) => {
    if (!Array.isArray(values) || !values.length) return;
    const pill = document.createElement('span');
    pill.textContent = `${label} ${values.slice(0, 5).join(', ')}${values.length > 5 ? '…' : ''}`;
    meta.appendChild(pill);
  };
  addMetaPill('位号', item.related_refdes);
  addMetaPill('网络', item.related_nets);
  addMetaPill('页码', item.related_pages);

  const action = document.createElement('div');
  action.className = 'review-plan-action';
  const actionText = document.createElement('p');
  actionText.textContent = item.recommended_action || '展开原始表格继续复核。';
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'ghost-btn compact-btn';
  button.textContent = '查看原表';
  button.addEventListener('click', () => openReportTable(item.target || item.section_id, item.target_table_id || item.table_id));
  action.append(actionText, button);

  node.append(header, trust, summary);
  if (meta.children.length) {
    node.appendChild(meta);
  }
  node.appendChild(action);
  return node;
}

function renderReviewPlanLayer({ host, id, title, lead, items = [], groups = [], collapsed = false }) {
  const section = document.createElement('section');
  section.id = id;
  section.className = `review-plan-layer layer-${id}`;
  const body = document.createElement(collapsed ? 'details' : 'div');
  body.className = collapsed ? 'review-plan-collapsible' : 'review-plan-body';
  if (collapsed) {
    body.innerHTML = `
      <summary>
        <span>${title}</span>
        <small>${lead}</small>
      </summary>
      <div class="review-plan-body"></div>
    `;
  }
  const bodyTarget = collapsed ? body.querySelector('.review-plan-body') : body;
  const heading = document.createElement('div');
  heading.className = 'review-plan-layer-head';
  heading.innerHTML = `
    <p class="eyebrow">${id.toUpperCase()}</p>
    <h3>${title}</h3>
    <p>${lead}</p>
  `;
  if (!collapsed) {
    section.appendChild(heading);
  }
  if (groups.length) {
    groups.forEach((group, index) => {
      const detail = document.createElement('details');
      detail.className = 'review-plan-group';
      detail.open = index === 0;
      detail.innerHTML = `
        <summary>
          <span>${group.title || '常规复核组'}</span>
          <small>${group.item_count || 0} 项 · ${group.count || 0} 条记录</small>
        </summary>
        <div class="review-plan-grid"></div>
      `;
      const grid = detail.querySelector('.review-plan-grid');
      (group.items || []).forEach((item) => grid.appendChild(reviewPlanItemNode(item)));
      bodyTarget.appendChild(detail);
    });
  } else if (items.length) {
    const grid = document.createElement('div');
    grid.className = 'review-plan-grid';
    items.forEach((item) => grid.appendChild(reviewPlanItemNode(item)));
    bodyTarget.appendChild(grid);
  } else {
    const empty = document.createElement('p');
    empty.className = 'empty-state review-plan-empty';
    empty.textContent = '这一层当前没有需要展示的记录。';
    bodyTarget.appendChild(empty);
  }
  section.appendChild(body);
  host.appendChild(section);
}

function renderReviewPlan(report) {
  const host = document.getElementById('review-plan');
  if (!host) return;
  host.replaceChildren();
  const plan = report.review_plan || {};
  const summary = plan.summary || {};
  const overview = document.createElement('div');
  overview.className = 'review-plan-overview';
  overview.innerHTML = `
    <div>
      <p class="eyebrow">REVIEW PLANNER</p>
      <h3>审查任务分层</h3>
      <p>默认只显示行动清单；统计、Debug 和完整索引按需展开。</p>
    </div>
    <div class="review-plan-overview-metrics">
      <span>重点 ${summary.focus_count || 0}</span>
      <span>常规 ${summary.review_item_count || 0}</span>
      <span>信息 ${summary.info_count || 0}</span>
      <span>Debug ${summary.debug_count || 0}</span>
    </div>
  `;
  host.appendChild(overview);
  renderReviewPlanLayer({
    host,
    id: 'focus',
    title: '重点',
    lead: '需要最先处理的明确异常和高优先级规则候选。',
    items: plan.focus_items || [],
  });
  renderReviewPlanLayer({
    host,
    id: 'review',
    title: '常规',
    lead: '按功能分区聚合的常规规则候选。',
    groups: plan.review_groups || [],
  });
  renderReviewPlanLayer({
    host,
    id: 'info',
    title: '信息',
    lead: '项目范围、页码分布、网络和 BOM 概览，默认折叠。',
    items: plan.info_items || [],
    collapsed: true,
  });
  renderReviewPlanLayer({
    host,
    id: 'debug',
    title: 'Debug',
    lead: '解析诊断、映射索引和原始证据入口，按需打开。',
    items: plan.debug_items || [],
    collapsed: true,
  });
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

function decisionToneClass(value, fallback = 'neutral') {
  return Number(value || 0) > 0 ? 'warning' : fallback;
}

function reportDecisionItemNode(item) {
  const node = document.createElement(item.target ? 'button' : 'article');
  node.className = `report-decision-item tone-${item.tone || 'neutral'}${item.target ? ' is-clickable' : ''}`;
  if (item.target) {
    node.type = 'button';
    node.addEventListener('click', () => scrollToSection(item.target));
  }
  node.innerHTML = `
    <span>${item.label}</span>
    <strong>${item.value}</strong>
    <p>${item.caption}</p>
  `;
  return node;
}

function renderReportDecisionStrip(report) {
  const host = document.getElementById('report-decision-strip');
  if (!host) return;
  const summary = report.review_plan?.summary || {};
  const trustCounts = summary.trust_counts || {};
  const insight = (report.top_insights || [])[0] || {};
  const items = [
    {
      label: '明确异常',
      value: trustCounts['明确异常'] || 0,
      caption: '字段缺失或确定状态，优先处理。',
      tone: decisionToneClass(trustCounts['明确异常'], 'ok'),
      target: 'focus',
    },
    {
      label: '规则候选',
      value: trustCounts['规则候选'] || 0,
      caption: '需要结合设计意图人工确认。',
      tone: decisionToneClass(trustCounts['规则候选']),
      target: 'review',
    },
    {
      label: '常规复核',
      value: summary.review_item_count || 0,
      caption: '按功能分区聚合的候选项。',
      tone: decisionToneClass(summary.review_item_count),
      target: 'review',
    },
    {
      label: '优先提示',
      value: (report.top_insights || []).length,
      caption: insight.title || '暂无额外高优先级提示。',
      tone: (report.top_insights || []).length ? (insight.tone || 'neutral') : 'ok',
      target: insight.target || 'summary',
    },
  ];
  host.replaceChildren();
  items.forEach((item) => host.appendChild(reportDecisionItemNode(item)));
  staggerChildren(host, '.report-decision-item');
}

function renderSummary(report) {
  const depopMode = report.include_depop
    ? `DEPOP 排查：开启（${report.depop_count || 0} 个器件参与分析）`
    : `DEPOP 排查：关闭（已忽略 ${report.excluded_depop_count || 0} 个器件）`;
  const totalBomMode = report.include_total_bom ? '总 BOM：开启' : '总 BOM：关闭';
  document.getElementById('generated-at').textContent =
    `生成时间：${report.generated_at} · 降额阈值：${report.ratio_limit}% · ${depopMode} · ${totalBomMode}`;
  const topbarGeneratedAt = document.getElementById('topbar-generated-at');
  if (topbarGeneratedAt) {
    topbarGeneratedAt.textContent = `生成 ${report.generated_at}`;
  }
  renderReportDecisionStrip(report);

  const metricStrip = document.getElementById('metric-strip');
  metricStrip.replaceChildren();
  report.metrics.forEach((metric) => metricStrip.appendChild(metricNode(metric)));
  staggerChildren(metricStrip, '.metric');

  const topInsights = document.getElementById('top-insights');
  topInsights.replaceChildren();
  (report.top_insights || []).forEach((insight, index) => {
    const node = insightNode(insight);
    if (index === 0) {
      node.classList.add('is-primary');
    }
    topInsights.appendChild(node);
  });
  staggerChildren(topInsights, '.insight-card', 1);

  const sectionCards = document.getElementById('section-cards');
  sectionCards.replaceChildren();
  (report.section_cards || []).forEach((section) => sectionCards.appendChild(sectionCardNode(section)));
  staggerChildren(sectionCards, '.section-card', 2);
  renderReviewPlan(report);

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

function renderSectionNav(sections, reviewPlan = null) {
  const nav = document.getElementById('section-nav');
  nav.replaceChildren();
  const links = [];
  const navOrder = ['summary', 'focus', 'review', 'info', 'debug', ...sections.map((section) => section.id), 'query'];
  navOrder.forEach((id) => {
    const labelMap = {
      summary: '概览',
      focus: '重点',
      review: '常规',
      info: '信息',
      debug: 'Debug',
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
  if (key === 'focus') return '!';
  if (key === 'review') return 'R';
  if (key === 'info') return 'i';
  if (key === 'debug') return 'D';
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
  const label = collapsed ? '展开导航' : '收起导航';
  button.textContent = label;
  button.setAttribute('aria-label', label);
  button.setAttribute('aria-expanded', String(!collapsed));
  button.title = label;
  try {
    window.localStorage.setItem(SIDEBAR_STORAGE_KEY, collapsed ? '1' : '0');
  } catch {
    // Ignore storage write failures and keep the current in-memory state.
  }
}

function setInspectorCollapsed(collapsed) {
  const layout = document.querySelector('.report-layout');
  const inspector = document.querySelector('.report-inspector');
  const button = document.getElementById('inspector-toggle');
  if (!layout || !inspector || !button) return;
  layout.classList.toggle('is-inspector-collapsed', collapsed);
  inspector.classList.toggle('is-collapsed', collapsed);
  button.textContent = collapsed ? '展开右栏' : '收起右栏';
  button.setAttribute('aria-expanded', String(!collapsed));
  button.setAttribute('aria-label', collapsed ? '展开右侧信息栏' : '收起右侧信息栏');
  try {
    window.localStorage.setItem(INSPECTOR_STORAGE_KEY, collapsed ? '1' : '0');
  } catch {
    // Ignore storage write failures and keep the current in-memory state.
  }
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
    restartMotion(host, 'query-result-enter', 520);
    return;
  }

  if (payload.view === 'list') {
    host.appendChild(queryListNode(payload, runQuery));
    applyGlobalStaggers(host);
    restartMotion(host, 'query-result-enter', 620);
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
      block.appendChild(dataTableNode(['refdes', 'pin', 'pin_name', 'desc', '页码', '主模块页映射一一对应'], card.items || []));
    }
    host.appendChild(block);
  });
  applyGlobalStaggers(host);
  restartMotion(host, 'query-result-enter', 620);
}

function renderAsterSummary(host, payload) {
  host.hidden = false;
  host.replaceChildren();
  const title = document.createElement('h4');
  title.textContent = `Aster 摘要 · ${payload.mode || 'mock'} · ${payload.provider || 'local'}`;
  const summary = document.createElement('p');
  summary.textContent = payload.summary || '暂无摘要。';
  host.appendChild(title);
  host.appendChild(summary);

  const list = document.createElement('ul');
  list.className = 'aster-summary-list';
  (payload.priorities || []).forEach((item) => {
    const li = document.createElement('li');
    const strong = document.createElement('strong');
    strong.textContent = item.title || '建议';
    const body = document.createElement('p');
    body.textContent = item.body || '';
    li.appendChild(strong);
    li.appendChild(body);
    list.appendChild(li);
  });
  if (list.children.length) {
    host.appendChild(list);
  }

  if ((payload.review_checklist || []).length) {
    const checklist = document.createElement('div');
    checklist.className = 'aster-checklist';
    const heading = document.createElement('h5');
    heading.textContent = 'AI 审查清单';
    checklist.appendChild(heading);
    (payload.review_checklist || []).slice(0, 8).forEach((item) => {
      const row = document.createElement('article');
      row.className = `aster-check-item is-${item.status || 'needs_review'} severity-${item.severity || 'medium'}`;
      const top = document.createElement('div');
      const name = document.createElement('strong');
      name.textContent = item.item || '审查项';
      const badge = document.createElement('span');
      badge.textContent = asterChecklistStatusLabel(item.status);
      top.append(name, badge);
      const evidence = document.createElement('p');
      evidence.textContent = item.evidence || '已纳入 AI 审查上下文。';
      row.append(top, evidence);
      checklist.appendChild(row);
    });
    host.appendChild(checklist);
  }

  if ((payload.section_focus || []).length) {
    const focusGrid = document.createElement('div');
    focusGrid.className = 'aster-focus-grid';
    (payload.section_focus || []).slice(0, 4).forEach((item) => {
      const card = document.createElement('article');
      const label = document.createElement('span');
      label.textContent = item.section || '分区';
      const count = document.createElement('strong');
      count.textContent = `${item.rows || 0} 条`;
      const reason = document.createElement('p');
      reason.textContent = item.reason || '';
      card.append(label, count, reason);
      focusGrid.appendChild(card);
    });
    host.appendChild(focusGrid);
  }

  if ((payload.manual_review || []).length) {
    const manual = document.createElement('div');
    manual.className = 'aster-manual-review';
    const heading = document.createElement('h5');
    heading.textContent = '必须人工确认';
    manual.appendChild(heading);
    (payload.manual_review || []).slice(0, 6).forEach((item) => {
      const row = document.createElement('article');
      const topic = document.createElement('strong');
      topic.textContent = item.topic || '人工复核项';
      const reason = document.createElement('p');
      reason.textContent = item.reason || '';
      row.append(topic, reason);
      manual.appendChild(row);
    });
    host.appendChild(manual);
  }

  const guard = document.createElement('p');
  guard.textContent = (payload.safeguards || []).join(' ');
  host.appendChild(guard);
  applyGlobalStaggers(host);
  restartMotion(host, 'query-result-enter', 620);
}

function asterChecklistStatusLabel(status) {
  const labels = {
    pass: '通过',
    covered_no_findings: '已覆盖',
    covered_with_findings: '有发现',
    needs_review: '需复核',
    manual_only: '人工判断',
  };
  return labels[status] || '需复核';
}

function renderAsterError(host, error, payload = {}) {
  host.hidden = false;
  host.replaceChildren();
  const title = document.createElement('h4');
  title.textContent = 'Aster 调用失败';
  const message = document.createElement('p');
  message.textContent = error.message || String(error);
  host.append(title, message);

  const diagnostics = payload.diagnostics || {};
  const detailItems = [
    ['请求 ID', diagnostics.request_id],
    ['操作', diagnostics.operation],
    ['后端', diagnostics.backend],
    ['HTTP 状态', diagnostics.status],
    ['日志文件', payload.log_file || diagnostics.log_file],
  ].filter(([, value]) => value !== undefined && value !== null && value !== '');
  if (detailItems.length) {
    host.appendChild(detailRowsNode(detailItems.map(([label, value]) => ({ label, value }))));
  }

  if ((payload.diagnostic_hints || []).length) {
    const list = document.createElement('ul');
    list.className = 'aster-summary-list';
    payload.diagnostic_hints.forEach((hint) => {
      const li = document.createElement('li');
      const body = document.createElement('p');
      body.textContent = hint;
      li.appendChild(body);
      list.appendChild(li);
    });
    host.appendChild(list);
  }
  restartMotion(host, 'query-result-enter', 620);
}

function bootAsterFloatingPanel() {
  const panel = document.getElementById('aster-assist');
  const launcher = document.getElementById('aster-float-launcher');
  const minimize = document.getElementById('aster-panel-minimize');
  const resetPosition = document.getElementById('aster-panel-reset-position');
  const handle = panel?.querySelector('[data-drag-handle="aster"]');
  if (!panel || !launcher) return;
  const storageKey = 'pstx_aster_panel_state';
  const positionKey = 'pstx_aster_panel_position';

  const clearPanelPosition = () => {
    panel.style.left = '';
    panel.style.top = '';
    panel.style.right = '';
    panel.style.bottom = '';
  };

  const clampPanelPosition = (left, top) => {
    const rect = panel.getBoundingClientRect();
    const margin = 12;
    const maxLeft = Math.max(margin, window.innerWidth - rect.width - margin);
    const maxTop = Math.max(margin, window.innerHeight - rect.height - margin);
    return {
      left: Math.min(Math.max(left, margin), maxLeft),
      top: Math.min(Math.max(top, margin), maxTop),
    };
  };

  const applyPanelPosition = (position) => {
    if (!position || !Number.isFinite(position.left) || !Number.isFinite(position.top)) {
      clearPanelPosition();
      return;
    }
    const clamped = clampPanelPosition(position.left, position.top);
    panel.style.left = `${clamped.left}px`;
    panel.style.top = `${clamped.top}px`;
    panel.style.right = 'auto';
    panel.style.bottom = 'auto';
  };

  const loadPanelPosition = () => {
    try {
      const saved = JSON.parse(localStorage.getItem(positionKey) || 'null');
      applyPanelPosition(saved);
    } catch (error) {
      clearPanelPosition();
    }
  };

  const savePanelPosition = () => {
    const rect = panel.getBoundingClientRect();
    try {
      localStorage.setItem(positionKey, JSON.stringify({ left: Math.round(rect.left), top: Math.round(rect.top) }));
    } catch (error) {
      // LocalStorage can be disabled in hardened browser environments.
    }
  };

  const setOpen = (open) => {
    panel.classList.toggle('is-collapsed', !open);
    launcher.hidden = open;
    launcher.setAttribute('aria-expanded', open ? 'true' : 'false');
    panel.setAttribute('aria-hidden', open ? 'false' : 'true');
    document.body.classList.toggle('aster-panel-open', open);
    if (open) {
      requestAnimationFrame(loadPanelPosition);
    }
    try {
      localStorage.setItem(storageKey, open ? 'open' : 'closed');
    } catch (error) {
      // LocalStorage can be disabled in hardened browser environments.
    }
  };

  launcher.addEventListener('click', () => setOpen(true));
  minimize?.addEventListener('click', () => setOpen(false));
  resetPosition?.addEventListener('click', () => {
    clearPanelPosition();
    try {
      localStorage.removeItem(positionKey);
    } catch (error) {
      // LocalStorage can be disabled in hardened browser environments.
    }
  });
  handle?.addEventListener('pointerdown', (event) => {
    if (event.target.closest('button, input, select, textarea, a')) return;
    event.preventDefault();
    const startX = event.clientX;
    const startY = event.clientY;
    const startRect = panel.getBoundingClientRect();
    let dragFrame = 0;
    let nextPosition = { left: startRect.left, top: startRect.top };
    panel.classList.add('is-dragging');
    handle.setPointerCapture?.(event.pointerId);

    const onPointerMove = (moveEvent) => {
      const proposed = clampPanelPosition(
        startRect.left + moveEvent.clientX - startX,
        startRect.top + moveEvent.clientY - startY,
      );
      nextPosition = proposed;
      if (dragFrame) return;
      dragFrame = requestAnimationFrame(() => {
        dragFrame = 0;
        applyPanelPosition(nextPosition);
      });
    };

    const onPointerUp = () => {
      if (dragFrame) {
        cancelAnimationFrame(dragFrame);
        dragFrame = 0;
      }
      applyPanelPosition(nextPosition);
      savePanelPosition();
      panel.classList.remove('is-dragging');
      window.removeEventListener('pointermove', onPointerMove);
      window.removeEventListener('pointerup', onPointerUp);
      window.removeEventListener('pointercancel', onPointerUp);
    };

    window.addEventListener('pointermove', onPointerMove);
    window.addEventListener('pointerup', onPointerUp);
    window.addEventListener('pointercancel', onPointerUp);
  });
  window.addEventListener('resize', () => {
    if (!panel.classList.contains('is-collapsed')) {
      const rect = panel.getBoundingClientRect();
      applyPanelPosition({ left: rect.left, top: rect.top });
      savePanelPosition();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !panel.classList.contains('is-collapsed')) {
      setOpen(false);
      launcher.focus({ preventScroll: true });
    }
  });

  let saved = 'closed';
  try {
    saved = localStorage.getItem(storageKey) || 'closed';
  } catch (error) {
    saved = 'closed';
  }
  setOpen(saved === 'open');
}

function renderAsterStatus(host, payload) {
  host.hidden = false;
  host.replaceChildren();

  const header = document.createElement('div');
  header.className = 'aster-auth-header';
  const title = document.createElement('strong');
  title.textContent = `认证状态 · ${payload.mode || 'mock'} / ${payload.backend || 'chat-flow'}`;
  const badge = document.createElement('span');
  badge.className = `aster-auth-badge is-${payload.status || 'unknown'}`;
  badge.textContent = payload.status || 'unknown';
  header.append(title, badge);
  host.appendChild(header);

  const message = document.createElement('p');
  message.textContent = payload.message || '未读取到 Aster 状态。';
  host.appendChild(message);

  const grid = document.createElement('div');
  grid.className = 'aster-auth-grid';
  (payload.items || []).forEach((item) => {
    const row = document.createElement('div');
    row.className = `aster-auth-item ${item.configured ? 'is-configured' : 'is-missing'}`;
    const label = document.createElement('span');
    label.textContent = item.label || item.name || '配置项';
    const state = document.createElement('strong');
    state.textContent = item.configured ? (item.secret ? '已配置（隐藏）' : (item.value || '已配置')) : '未配置';
    const meta = document.createElement('small');
    meta.textContent = `${item.name || ''}${item.required ? ' · 必需' : ' · 可选'}`;
    row.append(label, state, meta);
    grid.appendChild(row);
  });
  host.appendChild(grid);

  if ((payload.safeguards || []).length) {
    const guard = document.createElement('p');
    guard.textContent = payload.safeguards.join(' ');
    host.appendChild(guard);
  }
  applyGlobalStaggers(host);
}

function bootAsterStatus() {
  const host = document.getElementById('aster-auth-status');
  if (!host) return;
  host.hidden = false;
  host.textContent = '正在读取 Aster 认证状态…';
  return fetch('/api/aster/status')
    .then((response) => response.json())
    .then((payload) => renderAsterStatus(host, payload))
    .catch((error) => {
      host.replaceChildren();
      const message = document.createElement('p');
      message.textContent = `Aster 状态读取失败：${error.message || error}`;
      host.appendChild(message);
    });
}

function clearSecretInputs(form) {
  ['api_key', 'app_secret'].forEach((name) => {
    const input = form.elements[name];
    if (input) input.value = '';
  });
}

function setCredentialMessage(text, tone = 'neutral') {
  const message = document.getElementById('aster-credential-message');
  if (!message) return;
  message.textContent = text;
  message.dataset.tone = tone;
}

function bootAsterCredentialForm() {
  const form = document.getElementById('aster-credential-form');
  const clearButton = document.getElementById('aster-credential-clear');
  if (!form) return;

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const submitButton = form.querySelector('button[type="submit"]');
    if (submitButton) submitButton.disabled = true;
    setCredentialMessage('正在应用临时凭据…');
    const payload = Object.fromEntries(new FormData(form).entries());
    try {
      const response = await fetch('/api/aster/runtime-config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const result = await response.json();
      clearSecretInputs(form);
      if (!response.ok || result.ok === false) {
        throw new Error(result.error || '临时凭据设置失败。');
      }
      renderAsterStatus(document.getElementById('aster-auth-status'), result);
      setCredentialMessage('临时凭据已应用到当前进程内存。', 'ok');
    } catch (error) {
      setCredentialMessage(error.message || String(error), 'error');
    } finally {
      if (submitButton) submitButton.disabled = false;
    }
  });

  clearButton?.addEventListener('click', async () => {
    clearButton.disabled = true;
    setCredentialMessage('正在清除临时凭据…');
    try {
      const response = await fetch('/api/aster/runtime-config', { method: 'DELETE' });
      const result = await response.json();
      clearSecretInputs(form);
      renderAsterStatus(document.getElementById('aster-auth-status'), result);
      setCredentialMessage('临时凭据已清除。', response.ok ? 'ok' : 'error');
    } catch (error) {
      setCredentialMessage(error.message || String(error), 'error');
    } finally {
      clearButton.disabled = false;
    }
  });
}

function bootAsterSummary(runId) {
  const button = document.getElementById('aster-summary-button');
  const host = document.getElementById('aster-summary-result');
  const chatHost = document.getElementById('harness-agent-result');
  if (!button || !host) return;

  button.addEventListener('click', async () => {
    button.disabled = true;
    const originalText = button.textContent;
    button.textContent = '生成中…';
    host.hidden = true;
    let loadingMessage = null;
    if (chatHost) {
      appendAgentChatMessage(chatHost, 'user', '请生成当前报告的审查摘要。');
      loadingMessage = appendAgentChatLoading(chatHost, '收集报告摘要', { kind: 'summary' });
    } else {
      host.hidden = false;
      host.innerHTML = '<p>正在生成 Aster 摘要…</p>';
    }
    let errorPayload = {};
    try {
      const response = await fetch(`/api/report/${runId}/aster-summary`);
      loadingMessage?.updateAgentStage?.('读取摘要结果');
      const payload = await response.json();
      if (!response.ok || !payload.ok) {
        errorPayload = payload;
        throw new Error(payload.error || '生成失败。');
      }
      if (loadingMessage) loadingMessage.remove();
      if (chatHost) {
        const summaryNode = document.createElement('div');
        renderAsterSummary(summaryNode, payload);
        appendAgentChatMessage(chatHost, 'assistant', summaryNode, { className: 'is-result' });
      } else {
        host.hidden = false;
        renderAsterSummary(host, payload);
      }
    } catch (error) {
      if (loadingMessage) loadingMessage.remove();
      if (chatHost) {
        const errorNode = document.createElement('div');
        renderAsterError(errorNode, error, errorPayload);
        appendAgentChatMessage(chatHost, 'assistant', errorNode, { className: 'is-result' });
      } else {
        host.hidden = false;
        renderAsterError(host, error, errorPayload);
      }
    } finally {
      button.disabled = false;
      button.textContent = originalText;
    }
  });
}

function agentTraceMetaNode(payload, { compact = false } = {}) {
  const summary = payload.trace_summary || {};
  const durable = payload.durable_status || payload;
  const progress = durable.progress || {};
  const capabilityText = (payload.capability_plan || [])
    .map((item) => item.title || item.id)
    .filter(Boolean)
    .join(' / ');
  const rows = [
    { label: 'Agent Run', value: payload.agent_run_id || summary.agent_run_id || '未生成' },
    { label: 'Profile', value: payload.profile || summary.profile || 'quick_scan' },
    { label: '能力组合', value: capabilityText || (payload.capability_profiles || []).join(' / ') || '—' },
    { label: '当前阶段', value: durable.current_phase || durable.checkpoint?.phase || '—' },
    { label: 'Heartbeat', value: durable.heartbeat_at || '—' },
    { label: '终止原因', value: summary.stopped_reason || payload.model_metadata?.stopped_reason || 'unknown' },
    { label: '步骤', value: summary.steps ?? progress.step_index ?? (payload.agent_steps || []).length },
    { label: '工具调用', value: summary.tool_call_count ?? progress.tool_call_count ?? (payload.tool_calls || []).length },
    { label: '证据节点', value: summary.evidence_node_count ?? progress.evidence_count ?? (payload.final_evidence || []).length },
    { label: '任务账本 Open', value: summary.task_ledger_open_count ?? payload.runtime_state?.task_ledger?.progress?.open ?? 0 },
    { label: '建议下一步', value: summary.task_ledger_next_action_count ?? (payload.runtime_state?.task_ledger?.next_actions || []).length },
    { label: '回答质量', value: `${summary.final_quality_status || payload.final_answer_quality_gate?.status || '—'} ${summary.final_quality_score ?? payload.final_answer_quality_gate?.score ?? ''}`.trim() },
    { label: '修正动作', value: payload.final_answer_quality_gate?.repair_action_count ?? 0 },
    { label: '自动补证据', value: summary.quality_repair_attempt_count ? `${summary.quality_repair_attempt_count} 次 / ${summary.quality_repair_tool_count || 0} 工具` : '未触发' },
    { label: '运行日志', value: payload.journal_summary?.event_count ?? (payload.execution_journal || []).length ?? 0 },
    { label: 'Artifacts', value: durable.artifact_count ?? 0 },
    { label: '续跑意图', value: payload.continuation_pack?.next_intent || '—' },
    { label: 'Subagents', value: summary.subagent_count ?? (payload.subagents || []).length },
    { label: '耗时', value: `${summary.elapsed_ms ?? payload.elapsed_ms ?? 0} ms` },
  ];
  if (compact) {
    return detailRowsNode(rows.filter((row) => ['Profile', '能力组合', '步骤', '工具调用'].includes(row.label)));
  }
  return detailRowsNode(rows);
}

function agentResultStatusClass(payload) {
  if (payload.ok === false) return 'is-error';
  if (payload.status === 'waiting_for_user') return 'is-waiting';
  if (payload.status === 'limited') return 'is-limited';
  if (payload.status === 'incomplete') return 'is-incomplete';
  const quality = payload.final_answer_quality_gate?.status || payload.trace_summary?.final_quality_status || '';
  if (quality === 'warn') return 'is-warning';
  if (quality === 'fail') return 'is-error';
  return 'is-ok';
}

function agentStoppedReasonLabel(reason) {
  const labels = {
    final_answer: '模型已给出最终回答',
    needs_user_input: '等待用户补充信息',
    max_tool_calls: '工具调用预算已用尽',
    max_steps: '执行轮数预算已用尽',
    model_error: '模型服务异常',
    invalid_model_json: '模型返回格式异常',
    protocol_error: '工具协议校验未通过',
    tool_error: '工具执行异常',
    tool_error_recovery: '工具异常后已尝试恢复',
    quality_repair_continue: '已自动补证据后再次生成',
    empty_answer: '模型未生成有效回答',
  };
  return labels[String(reason || '')] || (reason ? String(reason) : '正常结束');
}

function agentStatusOverview(payload) {
  const summary = payload.trace_summary || {};
  const durable = payload.durable_status || payload;
  const progress = durable.progress || {};
  const quality = payload.final_answer_quality_gate || {};
  const stoppedReason = summary.stopped_reason || payload.model_metadata?.stopped_reason || '';
  const status = payload.status || (payload.ok === false ? 'failed' : 'completed');
  const toolCalls = summary.tool_call_count ?? progress.tool_call_count ?? (payload.tool_calls || []).length;
  const steps = summary.steps ?? progress.step_index ?? (payload.agent_steps || []).length;
  const openTasks = summary.task_ledger_open_count ?? payload.runtime_state?.task_ledger?.progress?.open ?? 0;
  const nextActions = payload.runtime_state?.task_ledger?.next_actions || [];
  const needs = payload.needs_user_input || {};
  const repairCount = summary.quality_repair_attempt_count || 0;

  let title = '已完成';
  let description = 'Agent 已完成本轮取证和回答，可打开 Trace 查看证据链。';
  let next = nextActions[0]?.title || '如需更细节，可继续追问具体位号、网络、页码或证据 id。';
  if (payload.ok === false) {
    title = '需要人工接管';
    description = payload.error || '本轮执行遇到错误，建议打开 Trace 查看失败工具或模型返回。';
    next = '可以缩小问题范围，或提高轮数/工具预算后重试。';
  } else if (status === 'waiting_for_user') {
    title = '等待补充';
    description = needs.reason || '当前证据缺口需要你补一小段信息，提交后会继续同一任务。';
    next = `需要补充 ${Array.isArray(needs.questions) ? needs.questions.length : 0} 项信息。`;
  } else if (status === 'limited') {
    title = '达到预算上限';
    description = `${agentStoppedReasonLabel(stoppedReason)}，已保留当前证据和任务账本。`;
    next = '建议继续追问同一主题，或把最大轮数/工具调用调大后再运行。';
  } else if (status === 'incomplete') {
    title = '证据仍不完整';
    description = durable.error || `${agentStoppedReasonLabel(stoppedReason)}；任务已保留 checkpoint，可继续同一 run。`;
    next = nextActions[0]?.title || '建议继续取证：优先打开 Trace 中推荐的 detail/aggregation 工具。';
  } else if (quality.status === 'warn') {
    title = '完成但建议复核';
    description = quality.summary || '回答已生成，但质量门禁提示存在证据覆盖或引用风险。';
    next = nextActions[0]?.title || '建议打开 Trace 查看质量门禁和证据引用。';
  }

  return {
    title,
    description,
    next,
    status,
    stopped_reason: stoppedReason,
    stats: [
      { label: '步骤', value: steps },
      { label: '工具', value: toolCalls },
      { label: '证据', value: summary.evidence_node_count ?? progress.evidence_count ?? (payload.final_evidence || []).length },
      { label: 'Open', value: openTasks },
      { label: '补证', value: repairCount ? `${repairCount}次` : '无' },
    ],
  };
}

function renderAgentStatusOverview(payload) {
  const overview = agentStatusOverview(payload);
  const block = document.createElement('section');
  block.className = `agent-status-overview ${agentResultStatusClass(payload)}`;
  const text = document.createElement('div');
  text.className = 'agent-status-copy';
  const title = document.createElement('strong');
  title.textContent = overview.title;
  const desc = document.createElement('p');
  desc.textContent = overview.description;
  const next = document.createElement('p');
  next.className = 'agent-status-next';
  next.textContent = `下一步：${overview.next}`;
  text.append(title, desc, next);

  const stats = document.createElement('div');
  stats.className = 'agent-status-stats';
  overview.stats.forEach((item) => {
    const chip = document.createElement('span');
    chip.textContent = `${item.label} ${item.value}`;
    stats.appendChild(chip);
  });
  block.append(text, stats);
  return block;
}

function renderAgentList(title, items, renderer, { hideEmpty = false, limit = 0 } = {}) {
  if (!items.length && hideEmpty) return null;
  const block = document.createElement('div');
  block.className = 'agent-result-block';
  const heading = document.createElement('h4');
  heading.textContent = title;
  block.appendChild(heading);
  if (!items.length) {
    const empty = document.createElement('p');
    empty.className = 'agent-empty';
    empty.textContent = '暂无。';
    block.appendChild(empty);
    return block;
  }
  const list = document.createElement('div');
  list.className = 'agent-result-list';
  const visibleItems = limit > 0 ? items.slice(0, limit) : items;
  visibleItems.forEach((item, index) => list.appendChild(renderer(item, index)));
  block.appendChild(list);
  if (limit > 0 && items.length > limit) {
    const note = document.createElement('p');
    note.className = 'agent-result-compact-note';
    note.textContent = `已收起 ${items.length - limit} 项，打开 Trace 抽屉查看全部。`;
    block.appendChild(note);
  }
  return block;
}

function traceJsonPreview(value, limit = 22000) {
  let text = '';
  try {
    text = JSON.stringify(value ?? {}, null, 2);
  } catch (error) {
    text = String(value ?? '');
  }
  if (text.length <= limit) return text;
  return `${text.slice(0, limit)}\n… 已截断 ${text.length - limit} 字符，完整内容仍在本地 agent trace/store 中。`;
}

function traceDetailsNode(title, value, { open = false, className = '' } = {}) {
  const details = document.createElement('details');
  details.className = `agent-trace-details ${className}`.trim();
  if (open) details.open = true;
  const summary = document.createElement('summary');
  summary.textContent = title;
  const pre = document.createElement('pre');
  pre.textContent = traceJsonPreview(value);
  details.append(summary, pre);
  return details;
}

function renderEvidenceLayers(layers) {
  if (!layers || typeof layers !== 'object') return null;
  const block = document.createElement('div');
  block.className = 'agent-evidence-layers';
  const heading = document.createElement('span');
  heading.className = 'agent-evidence-label';
  heading.textContent = '三层证据';
  block.appendChild(heading);

  const summaryLayer = layers.summary_layer || {};
  const summary = document.createElement('p');
  summary.textContent = [
    summaryLayer.completeness ? `完整性：${summaryLayer.completeness}` : '',
    summaryLayer.evidence_count !== undefined ? `证据卡：${summaryLayer.evidence_count}` : '',
    summaryLayer.scope_summary ? `范围：${summaryLayer.scope_summary}` : '',
  ].filter(Boolean).join(' · ') || '摘要层已生成。';
  block.appendChild(summary);

  const cards = Array.isArray(layers.evidence_card_layer) ? layers.evidence_card_layer : [];
  if (cards.length) {
    const list = document.createElement('div');
    list.className = 'agent-evidence-card-list';
    cards.slice(0, 6).forEach((card) => {
      const chip = document.createElement('span');
      chip.className = 'agent-evidence-card';
      chip.textContent = `${card.id || 'ev'}${card.refdes ? ` · ${card.refdes}` : ''}${card.page ? ` · 页${card.page}` : ''}`;
      list.appendChild(chip);
    });
    if (cards.length > 6) {
      const more = document.createElement('span');
      more.className = 'agent-evidence-card is-muted';
      more.textContent = `+${cards.length - 6}`;
      list.appendChild(more);
    }
    block.appendChild(list);
    block.appendChild(traceDetailsNode('展开证据卡层', cards, { className: 'is-compact' }));
  }

  if (layers.raw_layer) {
    block.appendChild(traceDetailsNode('展开原始层说明/预览', layers.raw_layer, { className: 'is-compact' }));
  }
  return block;
}

function findTopologyBusinessView(value) {
  if (!value || typeof value !== 'object') return null;
  if (value.business_view && typeof value.business_view === 'object') return value.business_view;
  if (value.topology_business_view && typeof value.topology_business_view === 'object') return value.topology_business_view;
  if (value.topology_netlist?.business_view && typeof value.topology_netlist.business_view === 'object') return value.topology_netlist.business_view;
  return null;
}

function renderTopologyBusinessView(view) {
  if (!view || typeof view !== 'object') return null;
  const block = document.createElement('div');
  block.className = 'agent-topology-business-view';
  const head = document.createElement('div');
  head.className = 'agent-topology-business-head';
  const title = document.createElement('strong');
  title.textContent = '拓扑业务视角';
  const counts = view.counts || {};
  const meta = document.createElement('span');
  meta.textContent = [
    counts.total_node_count !== undefined ? `节点 ${counts.total_node_count}` : '',
    counts.total_signal_edge_count !== undefined ? `信号边 ${counts.total_signal_edge_count}` : '',
    counts.total_supply_edge_count !== undefined ? `供电 ${counts.total_supply_edge_count}` : '',
  ].filter(Boolean).join(' · ');
  head.append(title, meta);
  block.appendChild(head);
  const summary = document.createElement('p');
  summary.textContent = view.summary || view.scope_note || '已生成拓扑业务视角。';
  block.appendChild(summary);

  const queue = Array.isArray(view.review_queue) ? view.review_queue.slice(0, 5) : [];
  if (queue.length) {
    const list = document.createElement('div');
    list.className = 'agent-topology-review-queue';
    queue.forEach((item) => {
      const row = document.createElement('article');
      row.className = `agent-topology-review-item is-${item.review_priority || 'low'}`;
      const strong = document.createElement('strong');
      strong.textContent = item.title || item.item_id || '审查项';
      const body = document.createElement('p');
      body.textContent = item.summary || (item.review_focus || []).join('、') || '需结合 detail tool 复核。';
      row.append(strong, body);
      list.appendChild(row);
    });
    block.appendChild(list);
  }

  const partitions = Array.isArray(view.review_partitions) ? view.review_partitions : [];
  if (partitions.length) {
    const chips = document.createElement('div');
    chips.className = 'agent-topology-partitions';
    partitions.slice(0, 8).forEach((partition) => {
      const chip = document.createElement('span');
      chip.className = `agent-topology-partition is-${partition.priority || 'low'}`;
      chip.textContent = `${partition.title || partition.partition_id} ${partition.item_count || 0}`;
      chips.appendChild(chip);
    });
    block.appendChild(chips);
  }
  block.appendChild(traceDetailsNode('展开拓扑业务视角 JSON', view, { className: 'is-compact' }));
  return block;
}

function renderTaskLedger(ledger) {
  if (!ledger || typeof ledger !== 'object') return null;
  const items = Array.isArray(ledger.items) ? ledger.items : [];
  const actions = Array.isArray(ledger.next_actions) ? ledger.next_actions : [];
  if (!items.length && !actions.length) return null;
  const block = document.createElement('div');
  block.className = 'agent-task-ledger';
  const head = document.createElement('div');
  head.className = 'agent-task-ledger-head';
  const title = document.createElement('strong');
  title.textContent = '任务账本';
  const progress = ledger.progress || {};
  const badge = document.createElement('span');
  badge.textContent = `open ${progress.open ?? 0} · blocked ${progress.blocked ?? 0}`;
  head.append(title, badge);
  block.appendChild(head);

  if (items.length) {
    const list = document.createElement('div');
    list.className = 'agent-task-ledger-list';
    items.slice(0, 8).forEach((entry) => {
      const row = document.createElement('article');
      row.className = `agent-ledger-item is-${entry.status || 'pending'}`;
      const strong = document.createElement('strong');
      strong.textContent = entry.title || entry.id || '任务项';
      const meta = document.createElement('p');
      const tools = Array.isArray(entry.recommended_tools) && entry.recommended_tools.length
        ? ` · 工具 ${entry.recommended_tools.slice(0, 3).join(', ')}`
        : '';
      meta.textContent = `${entry.status || 'pending'} · ${entry.source || 'runtime'}${tools}`;
      row.append(strong, meta);
      list.appendChild(row);
    });
    block.appendChild(list);
  }

  if (actions.length) {
    const next = document.createElement('div');
    next.className = 'agent-task-next';
    const nextTitle = document.createElement('span');
    nextTitle.textContent = '建议下一步';
    next.appendChild(nextTitle);
    actions.slice(0, 5).forEach((action) => {
      const chip = document.createElement('span');
      chip.className = 'agent-next-chip';
      chip.textContent = action.tool ? `${action.tool}: ${action.title || ''}` : (action.title || action.type || 'next');
      next.appendChild(chip);
    });
    block.appendChild(next);
  }
  block.appendChild(traceDetailsNode('展开完整任务账本 JSON', ledger, { className: 'is-compact' }));
  return block;
}

function openAgentTraceDrawer() {
  const drawer = document.getElementById('agent-trace-drawer');
  if (!drawer) return;
  drawer.classList.add('is-open');
  drawer.setAttribute('aria-hidden', 'false');
  document.body.classList.add('is-agent-trace-open');
}

function closeAgentTraceDrawer() {
  const drawer = document.getElementById('agent-trace-drawer');
  if (!drawer) return;
  drawer.classList.remove('is-open');
  drawer.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('is-agent-trace-open');
}

function bootAgentTraceDrawer() {
  const drawer = document.getElementById('agent-trace-drawer');
  if (!drawer || drawer.dataset.traceDrawerBooted) return;
  drawer.dataset.traceDrawerBooted = '1';
  const closeTargets = drawer.querySelectorAll('#agent-trace-close, [data-agent-trace-close]');
  closeTargets.forEach((target) => target.addEventListener('click', closeAgentTraceDrawer));
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && drawer.classList.contains('is-open')) {
      closeAgentTraceDrawer();
    }
  });
}

function jumpToAgentCitation(citation) {
  const locator = citation?.locator || {};
  const sectionId = locator.section_id;
  if (sectionId) {
    const section = document.getElementById(`compare-section-${sectionId}`);
    if (section) {
      section.scrollIntoView({ behavior: 'smooth', block: 'start' });
      restartMotion(section, 'query-result-enter', 620);
      return true;
    }
  }
  const tableId = locator.table_id;
  if (tableId) {
    const escaped = window.CSS?.escape ? CSS.escape(tableId) : String(tableId).replace(/"/g, '\\"');
    const table = document.querySelector(`[data-table-id="${escaped}"]`) || document.getElementById(`table-${tableId}`);
    if (table) {
      table.scrollIntoView({ behavior: 'smooth', block: 'center' });
      restartMotion(table, 'query-result-enter', 620);
      return true;
    }
  }
  return false;
}

function renderAgentTraceDrawer(payload, { title = 'Agent 执行复盘' } = {}) {
  const drawer = document.getElementById('agent-trace-drawer');
  const body = document.getElementById('agent-trace-body');
  if (!drawer || !body) return;
  body.replaceChildren();
  const heading = document.getElementById('agent-trace-title') || drawer.querySelector('.agent-trace-head h3');
  if (heading) heading.textContent = title;

  const shell = document.createElement('div');
  shell.className = 'agent-trace-shell';
  shell.appendChild(agentTraceMetaNode(payload));
  const durable = payload.durable_status || payload;
  if (durable.partial_trace) {
    shell.appendChild(traceDetailsNode('后台 Checkpoint / Partial Trace', durable.partial_trace, { open: true, className: 'is-compact' }));
  }
  if (durable.checkpoint) {
    shell.appendChild(traceDetailsNode('当前 Checkpoint', durable.checkpoint, { className: 'is-compact' }));
  }
  const ledgerNode = renderTaskLedger(payload.runtime_state?.task_ledger || payload.session_state?.task_ledger);
  if (ledgerNode) shell.appendChild(ledgerNode);
  if (payload.final_answer_quality_gate) {
    shell.appendChild(traceDetailsNode('最终回答质量门禁', payload.final_answer_quality_gate, { open: false, className: 'is-compact' }));
  }

  shell.appendChild(renderAgentList('证据引用', payload.citations || [], (citation) => {
    const item = document.createElement('article');
    item.className = `agent-citation ${citation.valid ? 'is-valid' : 'is-invalid'}`;
    const strong = document.createElement('strong');
    strong.textContent = `${citation.id || 'unknown'} · ${citation.title || citation.type || ''}`;
    const note = document.createElement('p');
    note.textContent = citation.note || (citation.valid ? '有效引用' : '引用不存在');
    const jump = document.createElement('button');
    jump.type = 'button';
    jump.className = 'ghost-btn inline-btn agent-citation-jump';
    jump.textContent = '定位证据';
    jump.addEventListener('click', () => {
      if (jumpToAgentCitation(citation)) closeAgentTraceDrawer();
    });
    item.append(strong, note, jump);
    return item;
  }));

  shell.appendChild(renderAgentList('工具调用', payload.tool_calls || [], (call) => {
    const item = document.createElement('article');
    item.className = `agent-step ${call.ok === false ? 'is-error' : ''}`;
    const strong = document.createElement('strong');
    strong.textContent = `#${call.index || '?'} ${call.tool || 'tool'} · ${call.ok === false ? '拒绝/失败' : '完成'}`;
    const bodyText = document.createElement('p');
    bodyText.textContent = call.error || call.reason || '已执行。';
    item.append(strong, bodyText);
    return item;
  }));

  shell.appendChild(renderAgentList('执行步骤', payload.agent_steps || [], (step) => {
    const item = document.createElement('article');
    item.className = `agent-step ${step.ok === false ? 'is-error' : ''}`;
    const strong = document.createElement('strong');
    strong.textContent = `#${step.index || '?'} ${step.type || 'step'} ${step.tool ? `· ${step.tool}` : ''}`;
    const bodyText = document.createElement('p');
    bodyText.textContent = step.summary || step.error || '已记录。';
    item.append(strong, bodyText);
    return item;
  }));

  shell.appendChild(renderAgentList('观察结果', payload.observations || [], (observation) => {
    const item = document.createElement('article');
    item.className = 'agent-step';
    const strong = document.createElement('strong');
    strong.textContent = observation.title || observation.tool || 'Observation';
    const bodyText = document.createElement('p');
    bodyText.textContent = observation.summary || `证据节点 ${(observation.evidence_node_ids || []).length} 个。`;
    item.append(strong, bodyText);
    const layers = renderEvidenceLayers(observation.evidence_layers);
    if (layers) item.appendChild(layers);
    const topologyView = renderTopologyBusinessView(findTopologyBusinessView(observation.raw_result || observation.result || observation));
    if (topologyView) item.appendChild(topologyView);
    return item;
  }));

  shell.appendChild(renderAgentList('原始证据层', payload.raw_observations || [], (raw) => {
    const item = document.createElement('article');
    item.className = 'agent-step';
    const strong = document.createElement('strong');
    strong.textContent = `#${raw.call_index || '?'} ${raw.tool || 'raw result'}`;
    const bodyText = document.createElement('p');
    bodyText.textContent = raw.summary || `原始 JSON ${raw.raw_result_json_chars || 0} 字符，默认不进入模型上下文。`;
    item.append(strong, bodyText);
    const layers = renderEvidenceLayers(raw.evidence_layers);
    if (layers) item.appendChild(layers);
    const topologyView = renderTopologyBusinessView(findTopologyBusinessView(raw.raw_result || raw));
    if (topologyView) item.appendChild(topologyView);
    item.appendChild(traceDetailsNode('展开完整工具结果', raw.raw_result || {}, { className: 'is-raw' }));
    return item;
  }, { hideEmpty: true }));

  shell.appendChild(renderAgentList('最终证据节点', payload.final_evidence || [], (node) => {
    const item = document.createElement('article');
    item.className = 'agent-citation is-valid';
    const strong = document.createElement('strong');
    strong.textContent = `${node.id || 'ev'} · ${node.title || node.type || ''}`;
    const bodyText = document.createElement('p');
    bodyText.textContent = node.summary || node.type || '';
    item.append(strong, bodyText);
    return item;
  }));

  shell.appendChild(renderAgentList('并行 Subagents', payload.subagents || [], (subagent) => {
    const item = document.createElement('article');
    item.className = `agent-subagent ${subagent.ok === false ? 'is-error' : ''}`;
    const strong = document.createElement('strong');
    strong.textContent = `${subagent.title || subagent.profile || 'Subagent'} · ${subagent.ok === false ? '需接管' : '完成'}`;
    const bodyText = document.createElement('p');
    const summary = subagent.trace_summary || {};
    bodyText.textContent = `${subagent.answer || '未生成回答。'} · steps ${summary.steps || 0} · evidence ${subagent.evidence_node_count || 0} · actions ${subagent.proposed_action_count || 0}`;
    item.append(strong, bodyText);
    return item;
  }));

  shell.appendChild(renderAgentList('建议动作', payload.proposed_actions || [], (action) => {
    const item = document.createElement('article');
    item.className = 'agent-action';
    const strong = document.createElement('strong');
    strong.textContent = action.title || action.id || '建议';
    const bodyText = document.createElement('p');
    bodyText.textContent = action.reason || action.priority || '需要人工复核。';
    item.append(strong, bodyText);
    return item;
  }));

  body.appendChild(shell);
  applyGlobalStaggers(shell);
}

function agentResultStatusLabel(payload) {
  if (payload.status === 'waiting_for_user') return '等待补充';
  if (payload.status === 'limited') return '达到上限';
  if (payload.status === 'incomplete') return '未完成';
  if (payload.ok === false) return '需要人工接管';
  return '完成';
}

function ensureAgentChatThread(host) {
  if (!host) return;
  host.hidden = false;
  host.classList.add('agent-chat-thread');
  if (!host.dataset.chatReady) {
    host.replaceChildren();
    host.dataset.chatReady = '1';
  }
}

function appendAgentChatMessage(host, role, content, { className = '', label = '' } = {}) {
  ensureAgentChatThread(host);
  const message = document.createElement('article');
  message.className = ['agent-chat-message', `is-${role}`, className].filter(Boolean).join(' ');
  const roleNode = document.createElement('span');
  roleNode.className = 'agent-chat-role';
  roleNode.textContent = label || (role === 'user' ? '你' : role === 'system' ? '系统' : 'Agent');
  const bubble = document.createElement('div');
  bubble.className = 'agent-chat-bubble';
  if (content instanceof Node) {
    bubble.appendChild(content);
  } else {
    const paragraph = document.createElement('p');
    paragraph.textContent = content || '';
    bubble.appendChild(paragraph);
  }
  message.append(roleNode, bubble);
  host.appendChild(message);
  host.scrollTop = host.scrollHeight;
  return message;
}

const AGENT_STATUS_STAGES = {
  report: [
    '整理你的问题',
    '规划取证路线',
    '调用本地只读工具',
    '压缩证据摘要',
    '等待模型生成',
    '校验证据引用',
    '整理回答',
  ],
  compare: [
    '确认 A/B 项目',
    '规划对比路线',
    '批量检索差异',
    '读取必要证据',
    '等待模型生成',
    '校验证据引用',
    '整理回答',
  ],
  continue: [
    '保存补充信息',
    '恢复任务上下文',
    '继续本地取证',
    '更新证据摘要',
    '整理回答',
  ],
  summary: [
    '收集报告摘要',
    '调用 Aster 服务',
    '等待摘要生成',
    '整理审查要点',
  ],
};

function normalizeAgentStages(stages, fallbackText) {
  const items = Array.isArray(stages) ? stages.map((item) => String(item || '').trim()).filter(Boolean) : [];
  if (items.length) return items;
  return [fallbackText || '规划取证路线'];
}

function updateAgentLoadingStage(message, label, index) {
  const status = message?.querySelector?.('[data-agent-stage-status]');
  if (!status) return;
  status.textContent = label;
  status.dataset.stageIndex = String(index);
}

function attachAgentStageController(message, stages, { interval = 1800, initialDelay = 900 } = {}) {
  const labels = normalizeAgentStages(stages, '规划取证路线');
  let index = 0;
  let timer = null;
  const tick = () => {
    if (!message.isConnected && timer) {
      clearInterval(timer);
      timer = null;
      return;
    }
    index = Math.min(index + 1, labels.length - 1);
    updateAgentLoadingStage(message, labels[index], index);
  };
  updateAgentLoadingStage(message, labels[0], 0);
  timer = window.setInterval(tick, interval);
  const firstTimer = window.setTimeout(tick, initialDelay);
  const stop = (finalText = '') => {
    window.clearTimeout(firstTimer);
    if (timer) {
      window.clearInterval(timer);
      timer = null;
    }
    if (finalText) updateAgentLoadingStage(message, finalText, index);
  };
  const originalRemove = message.remove.bind(message);
  message.remove = () => {
    stop();
    originalRemove();
  };
  message.updateAgentStage = (text) => {
    if (text) updateAgentLoadingStage(message, text, index);
  };
  message.stopAgentStages = stop;
  return message;
}

function appendAgentChatLoading(host, text, options = {}) {
  const loading = document.createElement('div');
  loading.className = 'agent-chat-loading';
  loading.innerHTML = '<span></span><span></span><span></span>';
  const label = document.createElement('p');
  label.dataset.agentStageStatus = '1';
  label.textContent = text || '规划取证路线';
  loading.appendChild(label);
  const message = appendAgentChatMessage(host, 'assistant', loading, { className: 'is-loading' });
  const stages = options.stages || (options.kind ? AGENT_STATUS_STAGES[options.kind] : null) || [text];
  return attachAgentStageController(message, stages, options);
}

function agentRunStageText(payload) {
  const status = String(payload?.status || '').toLowerCase();
  const phase = String(payload?.current_phase || payload?.checkpoint?.phase || '').toLowerCase();
  const progress = payload?.progress || {};
  const bits = [];
  if (progress.step_index !== undefined && progress.max_steps) bits.push(`步骤 ${progress.step_index}/${progress.max_steps}`);
  if (progress.tool_call_count !== undefined && progress.max_tool_calls) bits.push(`工具 ${progress.tool_call_count}/${progress.max_tool_calls}`);
  if (progress.evidence_count !== undefined) bits.push(`证据 ${progress.evidence_count}`);
  const suffix = bits.length ? ` · ${bits.join(' · ')}` : '';
  if (status === 'queued') return `已进入后台队列${suffix}`;
  if (status === 'running') {
    if (phase.includes('prefetch')) return `正在预取证据${suffix}`;
    if (phase.includes('batch_tool')) return `正在批量调用工具${suffix}`;
    if (phase.includes('tool')) return `正在调用工具取证${suffix}`;
    if (phase.includes('repair')) return `正在补证据/修复质量${suffix}`;
    if (phase.includes('model')) return `正在让模型综合证据${suffix}`;
    if (phase.includes('planning')) return `正在规划技能和工具${suffix}`;
    if (phase.includes('finalizing')) return `正在整理最终回答${suffix}`;
    return `后台正在取证${suffix}`;
  }
  if (status === 'waiting_for_user') return '需要补充信息';
  if (status === 'completed') return '结果已生成';
  if (status === 'cancelled') return '任务已取消';
  if (status === 'incomplete') return '任务已保存，可继续';
  if (status === 'failed') return '任务执行失败';
  return '读取后台任务状态';
}

async function pollAgentRunUntilReady(agentRunId, { loadingMessage = null, intervalMs = 1600, timeoutMs = 10 * 60 * 1000 } = {}) {
  const startedAt = Date.now();
  let lastPayload = null;
  while (Date.now() - startedAt < timeoutMs) {
    const response = await fetch(`/api/harness/agent-runs/${encodeURIComponent(agentRunId)}`);
    const payload = await response.json();
    lastPayload = payload;
    if (!response.ok || payload.ok === false) {
      throw new Error(payload.error || '读取后台 Agent 状态失败。');
    }
    loadingMessage?.updateAgentStage?.(agentRunStageText(payload));
    const status = String(payload.status || '').toLowerCase();
    if (['completed', 'waiting_for_user', 'failed', 'cancelled', 'incomplete'].includes(status)) {
      const result = payload.agent_run && Object.keys(payload.agent_run).length ? payload.agent_run : payload;
      return {
        ...result,
        status: result.status || payload.status,
        agent_run_id: result.agent_run_id || payload.agent_run_id,
        durable_status: payload,
      };
    }
    await new Promise((resolve) => window.setTimeout(resolve, intervalMs));
  }
  throw new Error(`后台 Agent 执行超过 ${Math.round(timeoutMs / 1000)} 秒仍未完成，请稍后通过复盘继续查看。${lastPayload?.agent_run_id ? ` agent_run_id=${lastPayload.agent_run_id}` : ''}`);
}

async function parseAgentResponseOrPoll(response, { loadingMessage = null } = {}) {
  const payload = await response.json();
  if (response.status === 202 || payload.async) {
    if (!payload.agent_run_id) {
      throw new Error(payload.error || '后台 Agent 未返回 agent_run_id。');
    }
    loadingMessage?.updateAgentStage?.(agentRunStageText(payload));
    return pollAgentRunUntilReady(payload.agent_run_id, { loadingMessage });
  }
  if (!response.ok || payload.ok === false) {
    throw Object.assign(new Error(payload.error || 'Agent 审查失败。'), { payload });
  }
  return payload;
}

function summarizeContextAnswers(contextAnswers) {
  return contextAnswers.map((item, index) => {
    const target = [item.applies_to?.refdes, item.applies_to?.field].filter(Boolean).join(' · ');
    return `${index + 1}. ${target ? `${target}：` : ''}${item.answer}`;
  }).join('\n');
}

function inferReportAgentProfile(question, profileMap) {
  const text = normalizeText(question).toLowerCase();
  const candidates = [
    ['agent_ref_qa', ['ref', '参考资料', '资料库', '文档', 'manual', '能力边界']],
    ['dfmea_prep', ['dfmea', 'fmea', '失效模式', '失效后果', '规格书']],
    ['feishu_bom_qa', ['飞书', 'hq', '料号', 'pi', '选型顺序', '规格型号', 'part number']],
    ['bom_depop', ['bom_option', 'depop', 'dnp', '打圈', 'bom']],
    ['page_mapping', ['页码', '主模块页', 'page', '映射']],
    ['resistor_bias', ['串阻', '上拉', '下拉', 'od', 'oc', '电阻']],
    ['derating', ['降额', '电容', '耐压', 'ac耦合']],
    ['csa_geometry', ['csa', '画圈', '几何', 'dot', 'arc', 'circle']],
    ['full_review', ['完整', '全面', '全部', '所有']],
  ];
  const matched = candidates.find(([profile, keywords]) => profileMap.has(profile) && keywords.some((keyword) => text.includes(keyword)));
  return matched?.[0] || (profileMap.has('quick_scan') ? 'quick_scan' : [...profileMap.keys()][0] || 'quick_scan');
}

function inferCompareAgentProfile(question, profileMap) {
  const text = normalizeText(question).toLowerCase();
  const candidates = [
    ['compare_cadence_pages', ['第', '页', 'page', 'cadence', 'csa', '原始文件', 'page1', 'page*.csv']],
    ['compare_bom_feishu', ['飞书', 'hq', '料号', 'pi', '选型顺序', 'bom', '规格']],
    ['compare_pin_net', ['pin', 'net', '引脚', '网络', '串阻', '连接']],
    ['compare_key_devices', ['芯片', '连接器', 'pu', 'xu', '关键器件', '新增', '删除']],
    ['compare_page_mapping', ['页码', '主模块页', '映射']],
    ['compare_full_review', ['完整', '全面', '全部', '所有']],
  ];
  const matched = candidates.find(([profile, keywords]) => profileMap.has(profile) && keywords.some((keyword) => text.includes(keyword)));
  return matched?.[0] || (profileMap.has('compare_quick_scan') ? 'compare_quick_scan' : [...profileMap.keys()][0] || 'compare_quick_scan');
}

function renderNeedsUserInputForm(host, payload, { runId = '', traceTitle = '', append = false } = {}) {
  const needs = payload.needs_user_input || {};
  const questions = needs.questions || [];
  if (!runId || !questions.length) return null;

  const form = document.createElement('form');
  form.className = 'agent-clarify-form';
  const heading = document.createElement('h4');
  heading.textContent = '需要你补充的信息';
  const reason = document.createElement('p');
  reason.textContent = needs.reason || '当前证据不足，补充后 Agent 会继续同一个审查任务。';
  form.append(heading, reason);

  questions.forEach((question, index) => {
    const label = document.createElement('label');
    label.className = 'agent-clarify-question';
    const title = document.createElement('span');
    const appliesTo = question.applies_to || {};
    const target = [appliesTo.refdes, appliesTo.field].filter(Boolean).join(' · ');
    title.textContent = `${index + 1}. ${question.question || question.question_id || '请补充信息'}${target ? `（${target}）` : ''}`;
    const textarea = document.createElement('textarea');
    textarea.name = question.question_id || `q-${index + 1}`;
    textarea.rows = 3;
    textarea.placeholder = '在这里填写补充信息，例如 HQ 料号、规格型号、芯片类别或人工待查说明。';
    textarea.dataset.appliesTo = JSON.stringify(appliesTo);
    label.append(title, textarea);
    form.appendChild(label);
  });

  const message = document.createElement('p');
  message.className = 'agent-clarify-message';
  const submit = document.createElement('button');
  submit.type = 'submit';
  submit.className = 'primary-btn inline-btn';
  submit.textContent = '提交并继续 Agent';
  form.append(submit, message);

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const contextAnswers = [...form.querySelectorAll('textarea')].map((textarea) => {
      let appliesTo = {};
      try {
        appliesTo = JSON.parse(textarea.dataset.appliesTo || '{}');
      } catch (error) {
        appliesTo = {};
      }
      return {
        question_id: textarea.name,
        answer: textarea.value.trim(),
        applies_to: appliesTo,
      };
    }).filter((item) => item.answer);
    if (!contextAnswers.length) {
      message.textContent = '请至少填写一条补充信息。';
      return;
    }
    submit.disabled = true;
    message.textContent = '正在保存补充并继续执行…';
    let loadingMessage = null;
    if (append) {
      appendAgentChatMessage(host, 'user', summarizeContextAnswers(contextAnswers), {
        className: 'is-context-answer',
        label: '补充',
      });
      loadingMessage = appendAgentChatLoading(host, '保存补充信息', { kind: 'continue' });
    }
    const request = payload.request || {};
    const limits = payload.limits || {};
    try {
      const endpoint = payload.agent_run_id
        ? `/api/harness/agent-runs/${encodeURIComponent(payload.agent_run_id)}/continue`
        : `/api/report/${runId}/harness/agent`;
      const response = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          profile: request.profile || payload.profile || 'dfmea_prep',
          question: request.question || '',
          max_steps: Number(request.max_steps || limits.max_steps || 10),
          max_tool_calls: Number(request.max_tool_calls || limits.max_tool_calls || 18),
          max_rows_per_table: Number(request.max_rows_per_table || limits.max_rows_per_table || 12),
          enable_subagents: Boolean(request.enable_subagents),
          subagent_profiles: request.subagent_profiles || [],
          max_subagents: Number(request.max_subagents || limits.max_subagents || 2),
          debug: Boolean(request.debug),
          continue_agent_run_id: payload.agent_run_id || '',
          context_answers: contextAnswers,
          async: true,
        }),
      });
      const nextPayload = await parseAgentResponseOrPoll(response, { loadingMessage });
      if (loadingMessage) loadingMessage.remove();
      if (nextPayload.agent_run_id && host?.dataset) {
        host.dataset.lastAgentRunId = nextPayload.agent_run_id;
      }
      renderHarnessAgentResult(host, nextPayload, { autoOpenTrace: true, traceTitle, runId, append: append || Boolean(host?.dataset?.chatReady) });
    } catch (error) {
      message.textContent = error.message || String(error);
      if (loadingMessage) loadingMessage.remove();
    } finally {
      submit.disabled = false;
    }
  });

  return form;
}

function renderHarnessAgentResult(host, payload, { replay = false, autoOpenTrace = false, traceTitle = '', runId = '', append = false } = {}) {
  if (append) {
    ensureAgentChatThread(host);
  } else {
    host.replaceChildren();
    host.hidden = false;
  }
  const shell = document.createElement('div');
  shell.className = `agent-result-shell ${agentResultStatusClass(payload)}`;
  if (append) shell.classList.add('is-compact');

  const title = document.createElement('div');
  title.className = 'agent-result-title';
  const titleMain = document.createElement('div');
  titleMain.className = 'agent-result-title-main';
  const heading = document.createElement('h4');
  heading.textContent = replay ? 'Agent Run 复盘' : 'Agent 审查结果';
  const badge = document.createElement('span');
  badge.className = `agent-result-badge ${agentResultStatusClass(payload)}`;
  badge.textContent = agentResultStatusLabel(payload);
  titleMain.append(heading, badge);
  const titleActions = document.createElement('div');
  titleActions.className = 'agent-result-title-actions';
  const collapseButton = document.createElement('button');
  collapseButton.type = 'button';
  collapseButton.className = 'ghost-btn inline-btn agent-result-collapse';
  collapseButton.textContent = '收起';
  collapseButton.setAttribute('aria-expanded', 'true');
  collapseButton.addEventListener('click', () => {
    const collapsed = shell.classList.toggle('is-body-collapsed');
    collapseButton.textContent = collapsed ? '展开' : '收起';
    collapseButton.setAttribute('aria-expanded', String(!collapsed));
  });
  titleActions.appendChild(collapseButton);
  if (append) {
    const dismissButton = document.createElement('button');
    dismissButton.type = 'button';
    dismissButton.className = 'ghost-btn inline-btn agent-result-dismiss';
    dismissButton.textContent = '关闭';
    dismissButton.setAttribute('aria-label', '关闭这张 Agent 结果卡片');
    dismissButton.addEventListener('click', () => {
      const message = shell.closest('.agent-chat-message');
      (message || shell).remove();
    });
    titleActions.appendChild(dismissButton);
  }
  title.append(titleMain, titleActions);
  shell.appendChild(title);

  const body = document.createElement('div');
  body.className = 'agent-result-body';
  shell.appendChild(body);
  body.appendChild(renderAgentStatusOverview(payload));

  const answer = document.createElement('p');
  answer.className = 'agent-answer';
  answer.textContent = payload.answer || payload.error || '未生成回答。';
  body.appendChild(answer);
  body.appendChild(agentTraceMetaNode(payload, { compact: append }));

  const clarifyForm = renderNeedsUserInputForm(host, payload, { runId, traceTitle: traceTitle || '报告 Agent 执行复盘', append });
  if (clarifyForm) {
    body.appendChild(clarifyForm);
  }

  const traceButton = document.createElement('button');
  traceButton.type = 'button';
  traceButton.className = 'primary-btn inline-btn';
  traceButton.textContent = '打开 Trace 抽屉';
  traceButton.addEventListener('click', () => {
    renderAgentTraceDrawer(payload, { title: traceTitle || (replay ? 'Agent Run 复盘' : 'Agent 执行复盘') });
    openAgentTraceDrawer();
  });
  body.appendChild(traceButton);
  renderAgentTraceDrawer(payload, { title: traceTitle || (replay ? 'Agent Run 复盘' : 'Agent 执行复盘') });

  if (payload.agent_run_id && !replay) {
    const replayButton = document.createElement('button');
    replayButton.type = 'button';
    replayButton.className = 'ghost-btn inline-btn';
    replayButton.textContent = '读取本次复盘';
    replayButton.addEventListener('click', async () => {
      replayButton.disabled = true;
      try {
        const response = await fetch(`/api/harness/agent-runs/${payload.agent_run_id}`);
        const replayPayload = await response.json();
        if (!response.ok || replayPayload.ok === false) {
          throw new Error(replayPayload.error || '读取复盘失败。');
        }
        renderHarnessAgentResult(host, replayPayload.agent_run, { replay: true, runId, append });
      } catch (error) {
        replayButton.textContent = error.message || String(error);
      } finally {
        replayButton.disabled = false;
      }
    });
    body.appendChild(replayButton);
  }

  const listOptions = append ? { hideEmpty: true, limit: 3 } : {};
  const citationsBlock = renderAgentList('证据引用', payload.citations || [], (citation) => {
    const item = document.createElement('article');
    item.className = `agent-citation ${citation.valid ? 'is-valid' : 'is-invalid'}`;
    const strong = document.createElement('strong');
    strong.textContent = `${citation.id || 'unknown'} · ${citation.title || citation.type || ''}`;
    const note = document.createElement('p');
    note.textContent = citation.note || (citation.valid ? '有效引用' : '引用不存在');
    item.append(strong, note);
    return item;
  }, listOptions);
  if (citationsBlock) body.appendChild(citationsBlock);

  const stepsBlock = renderAgentList('执行步骤', payload.agent_steps || [], (step) => {
    const item = document.createElement('article');
    item.className = `agent-step ${step.ok === false ? 'is-error' : ''}`;
    const strong = document.createElement('strong');
    strong.textContent = `#${step.index || '?'} ${step.type || 'step'} ${step.tool ? `· ${step.tool}` : ''}`;
    const body = document.createElement('p');
    body.textContent = step.summary || step.error || '已记录。';
    item.append(strong, body);
    return item;
  }, listOptions);
  if (stepsBlock) body.appendChild(stepsBlock);

  const subagentsBlock = renderAgentList('并行 Subagents', payload.subagents || [], (subagent) => {
    const item = document.createElement('article');
    item.className = `agent-subagent ${subagent.ok === false ? 'is-error' : ''}`;
    const strong = document.createElement('strong');
    strong.textContent = `${subagent.title || subagent.profile || 'Subagent'} · ${subagent.ok === false ? '需接管' : '完成'}`;
    const body = document.createElement('p');
    const summary = subagent.trace_summary || {};
    body.textContent = `${subagent.answer || '未生成回答。'} · steps ${summary.steps || 0} · evidence ${subagent.evidence_node_count || 0} · actions ${subagent.proposed_action_count || 0}`;
    item.append(strong, body);
    return item;
  }, append ? { hideEmpty: true, limit: 2 } : {});
  if (subagentsBlock) body.appendChild(subagentsBlock);

  const actionsBlock = renderAgentList('建议动作', payload.proposed_actions || [], (action) => {
    const item = document.createElement('article');
    item.className = 'agent-action';
    const strong = document.createElement('strong');
    strong.textContent = action.title || action.id || '建议';
    const body = document.createElement('p');
    body.textContent = action.reason || action.priority || '需要人工复核。';
    item.append(strong, body);
    return item;
  }, listOptions);
  if (actionsBlock) body.appendChild(actionsBlock);

  if (append) {
    appendAgentChatMessage(host, 'assistant', shell, { className: 'is-result' });
  } else {
    host.appendChild(shell);
  }
  applyGlobalStaggers(shell);
  restartMotion(shell, 'query-result-enter', 620);
  if (autoOpenTrace) {
    openAgentTraceDrawer();
  }
}

async function bootAsterChatAgent(runId) {
  const form = document.getElementById('harness-agent-form');
  const profileSelect = document.getElementById('harness-agent-profile');
  const questionInput = document.getElementById('harness-agent-question');
  const maxStepsInput = document.getElementById('harness-agent-max-steps');
  const maxToolCallsInput = document.getElementById('harness-agent-max-tool-calls');
  const enableSubagentsInput = document.getElementById('harness-agent-enable-subagents');
  const maxSubagentsInput = document.getElementById('harness-agent-max-subagents');
  const debugInput = document.getElementById('harness-agent-debug');
  const resultHost = document.getElementById('harness-agent-result');
  if (!form || !profileSelect || !resultHost) return;
  if (form.dataset.agentChatBooted) return;
  form.dataset.agentChatBooted = '1';
  ensureAgentChatThread(resultHost);
  appendAgentChatMessage(resultHost, 'system', '浮窗已进入连续对话模式。你可以追问、补充 DFMEA 缺失信息，Agent 会带着当前项目上下文继续。');

  if (!form.dataset.chatToolbarReady) {
    const toolbar = document.createElement('div');
    toolbar.className = 'agent-chat-toolbar';
    const clearButton = document.createElement('button');
    clearButton.type = 'button';
    clearButton.className = 'ghost-btn inline-btn';
    clearButton.textContent = '清空本地对话';
    clearButton.addEventListener('click', () => {
      resultHost.replaceChildren();
      resultHost.dataset.chatReady = '1';
      delete resultHost.dataset.lastAgentRunId;
      appendAgentChatMessage(resultHost, 'system', '本地浮窗对话已清空；项目级补充上下文仍保留，必要时可通过 API 清空。');
    });
    const clearContextButton = document.createElement('button');
    clearContextButton.type = 'button';
    clearContextButton.className = 'ghost-btn inline-btn';
    clearContextButton.textContent = '清空项目补充';
    clearContextButton.addEventListener('click', async () => {
      clearContextButton.disabled = true;
      try {
        const response = await fetch(`/api/report/${runId}/harness/context/clear`, { method: 'POST' });
        const payload = await response.json();
        if (!response.ok || payload.ok === false) {
          throw new Error(payload.error || '清空项目补充失败。');
        }
        delete resultHost.dataset.lastAgentRunId;
        appendAgentChatMessage(resultHost, 'system', '项目级补充上下文已清空，后续对话会重新收集缺失信息。');
      } catch (error) {
        appendAgentChatMessage(resultHost, 'system', error.message || String(error));
      } finally {
        clearContextButton.disabled = false;
      }
    });
    toolbar.append(clearButton, clearContextButton);
    form.insertAdjacentElement('afterend', toolbar);
    form.dataset.chatToolbarReady = '1';
  }

  const profileMap = new Map();
  try {
    const response = await fetch('/api/harness/profiles');
    const payload = await response.json();
    (payload.profiles || []).forEach((profile) => {
      profileMap.set(profile.id, profile);
      const option = document.createElement('option');
      option.value = profile.id;
      option.textContent = profile.title || profile.id;
      option.dataset.defaultQuestion = profile.default_question || '';
      option.dataset.maxSteps = profile.max_steps || '';
      option.dataset.maxToolCalls = profile.max_tool_calls || '';
      option.dataset.subagentProfiles = (profile.subagent_profiles || []).join(',');
      if (![...profileSelect.options].some((item) => item.value === profile.id)) {
        profileSelect.appendChild(option);
      }
    });
  } catch (error) {
    appendAgentChatMessage(resultHost, 'system', `Profile 读取失败：${error.message || error}`);
  }

  const applyProfileDefaults = () => {
    const profile = profileMap.get(profileSelect.value);
    if (!profile) return;
    if (questionInput && !questionInput.value.trim()) {
      questionInput.placeholder = profile.default_question || questionInput.placeholder;
    }
    if (maxStepsInput && profile.max_steps) {
      maxStepsInput.value = profile.max_steps;
    }
    if (maxToolCallsInput && profile.max_tool_calls) {
      maxToolCallsInput.value = profile.max_tool_calls;
    }
  };
  profileSelect.addEventListener('change', applyProfileDefaults);
  applyProfileDefaults();

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const submit = form.querySelector('button[type="submit"]');
    if (submit) submit.disabled = true;
    const rawQuestion = (questionInput?.value || '').trim();
    const profile = profileSelect.value || 'auto';
    const fallbackQuestion = profileMap.get(profile)?.default_question || questionInput?.placeholder || '';
    const question = rawQuestion || fallbackQuestion || '请根据当前项目继续审查。';
    appendAgentChatMessage(resultHost, 'user', question);
    const loadingMessage = appendAgentChatLoading(resultHost, '整理你的问题', { kind: 'report' });
    restartMotion(loadingMessage, 'query-loading-pulse', 480);
    try {
      const response = await fetch(`/api/report/${runId}/harness/agent`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          profile,
          question,
          max_steps: Number(maxStepsInput?.value || 6),
          max_tool_calls: Number(maxToolCallsInput?.value || 10),
          enable_subagents: Boolean(enableSubagentsInput?.checked),
          max_subagents: Number(maxSubagentsInput?.value || 2),
          debug: Boolean(debugInput?.checked),
          continue_agent_run_id: resultHost.dataset.lastAgentRunId || '',
          async: true,
        }),
      });
      const payload = await parseAgentResponseOrPoll(response, { loadingMessage });
      loadingMessage.remove();
      if (payload.agent_run_id) {
        resultHost.dataset.lastAgentRunId = payload.agent_run_id;
      }
      renderHarnessAgentResult(resultHost, payload, { autoOpenTrace: true, traceTitle: '报告 Agent 执行复盘', runId, append: true });
      if (questionInput) questionInput.value = '';
    } catch (error) {
      loadingMessage.remove();
      renderHarnessAgentResult(resultHost, { ok: false, error: error.message || String(error), answer: '' }, { append: true });
    } finally {
      if (submit) submit.disabled = false;
    }
  });
}

function evalCaseItem(caseItem) {
  const label = document.createElement('label');
  label.className = 'agent-eval-case-item';
  const checkbox = document.createElement('input');
  checkbox.type = 'checkbox';
  checkbox.value = caseItem.case_id;
  const body = document.createElement('span');
  const strong = document.createElement('strong');
  strong.textContent = caseItem.title || caseItem.case_id;
  const meta = document.createElement('small');
  meta.textContent = `${caseItem.case_id} · ${caseItem.profile || 'quick_scan'} · ${caseItem.expected_stopped_reason || 'any'}`;
  const desc = document.createElement('p');
  desc.textContent = caseItem.description || '';
  body.append(strong, meta, desc);
  label.append(checkbox, body);
  return label;
}

function renderAgentEvalResult(host, payload) {
  host.replaceChildren();
  const shell = document.createElement('div');
  shell.className = `agent-eval-result-shell ${payload.ok ? 'is-ok' : 'is-error'}`;
  const score = document.createElement('div');
  score.className = 'agent-eval-score';
  const number = document.createElement('strong');
  number.textContent = `${payload.score ?? 0}%`;
  const caption = document.createElement('span');
  caption.textContent = `通过 ${payload.passed_count || 0}/${payload.case_count || 0} · 失败 ${payload.failed_count || 0}`;
  score.append(number, caption);
  shell.appendChild(score);

  const list = document.createElement('div');
  list.className = 'agent-eval-result-list';
  (payload.cases || []).forEach((caseResult) => {
    const item = document.createElement('article');
    item.className = `agent-eval-result-item ${caseResult.passed ? 'is-pass' : 'is-fail'}`;
    const title = document.createElement('strong');
    title.textContent = `${caseResult.passed ? 'PASS' : 'FAIL'} · ${caseResult.title || caseResult.case_id}`;
    const meta = document.createElement('p');
    meta.textContent = `工具：${(caseResult.tool_calls || []).join(', ') || '无'} · stopped=${caseResult.metrics?.stopped_reason || ''}`;
    item.append(title, meta);
    if ((caseResult.failures || []).length) {
      const failures = document.createElement('ul');
      (caseResult.failures || []).forEach((failure) => {
        const li = document.createElement('li');
        li.textContent = failure;
        failures.appendChild(li);
      });
      item.appendChild(failures);
    }
    list.appendChild(item);
  });
  shell.appendChild(list);
  host.appendChild(shell);
  restartMotion(shell, 'query-result-enter', 620);
}

async function bootAgentEvalPage() {
  const statusHost = document.getElementById('agent-eval-status');
  const listHost = document.getElementById('agent-eval-case-list');
  const resultHost = document.getElementById('agent-eval-result');
  const runAll = document.getElementById('agent-eval-run-all');
  const runSelected = document.getElementById('agent-eval-run-selected');
  if (!statusHost || !listHost || !resultHost || !runAll || !runSelected) return;

  try {
    const response = await fetch('/api/agent-eval/status');
    const payload = await response.json();
    statusHost.textContent = `共 ${payload.case_count || 0} 个 eval case，本地 deterministic provider，不调用真实 Aster。`;
    listHost.replaceChildren();
    (payload.cases || []).forEach((caseItem) => listHost.appendChild(evalCaseItem(caseItem)));
  } catch (error) {
    statusHost.textContent = `Eval 状态读取失败：${error.message || error}`;
  }

  const runEval = async (selectedOnly) => {
    const selected = [...listHost.querySelectorAll('input[type="checkbox"]:checked')].map((item) => item.value);
    const body = selectedOnly ? { case_ids: selected } : {};
    resultHost.innerHTML = '<p class="agent-empty">正在运行 Agent Eval…</p>';
    try {
      const response = await fetch('/api/agent-eval/run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      const payload = await response.json();
      if (!response.ok || payload.error) {
        throw new Error(payload.error || 'Agent Eval 执行失败。');
      }
      renderAgentEvalResult(resultHost, payload);
    } catch (error) {
      renderAgentEvalResult(resultHost, {
        ok: false,
        score: 0,
        passed_count: 0,
        case_count: 1,
        failed_count: 1,
        cases: [{ title: 'Eval 调用失败', passed: false, failures: [error.message || String(error)] }],
      });
    }
  };
  runAll.addEventListener('click', () => runEval(false));
  runSelected.addEventListener('click', () => runEval(true));
}

function renderAgentLabDocList(host, source, options = {}) {
  if (!host) return;
  const label = options.label || '资料';
  const countKey = options.countKey || 'file_count';
  host.replaceChildren();
  const docs = source.documents || [];
  if (!docs.length) {
    const empty = document.createElement('p');
    empty.className = 'agent-empty';
    empty.textContent = source[countKey] ? `已发现 ${label}，但还没有可展示索引；请点击重建索引。` : `${label} 目录还没有文件。把资料放进去后点击重建索引。`;
    host.appendChild(empty);
    return;
  }
  docs.forEach((doc) => {
    const item = document.createElement('article');
    item.className = `agent-lab-doc-item ${doc.status === 'indexed' ? 'is-indexed' : 'is-warning'}`;
    const title = document.createElement('strong');
    title.textContent = doc.title || doc.rel_path || `doc ${doc.doc_id}`;
    const meta = document.createElement('small');
    meta.textContent = `${doc.status || 'unknown'} · ${doc.page_count || 0} 页 · ${doc.rel_path || ''}`;
    item.append(title, meta);
    if (doc.error) {
      const error = document.createElement('p');
      error.textContent = doc.error;
      item.appendChild(error);
    }
    host.appendChild(item);
  });
}

function renderAgentLabStatus(statusHost, docsHost, payload, checklistHost = null) {
  const ref = payload.ref || payload || {};
  const checklist = payload.checklist || {};
  statusHost.textContent = `${checklist.summary || 'ref_checklist 状态未知'} · ${ref.summary || 'ref/ 状态未知'}`;
  renderAgentLabDocList(checklistHost, checklist, { label: 'ref_checklist', countKey: 'file_count' });
  renderAgentLabDocList(docsHost, ref, { label: 'ref PDF', countKey: 'pdf_count' });
}

async function bootAgentLabPage() {
  const defaultAgentLabProfile = 'review_checklist_qa';
  const statusHost = document.getElementById('agent-lab-status');
  const docsHost = document.getElementById('agent-lab-ref-docs');
  const checklistDocsHost = document.getElementById('agent-lab-checklist-docs');
  const refreshButton = document.getElementById('agent-lab-refresh');
  const reindexButton = document.getElementById('agent-lab-reindex');
  const checklistReindexButton = document.getElementById('agent-lab-checklist-reindex');
  const form = document.getElementById('agent-lab-form');
  const profileSelect = document.getElementById('agent-lab-profile');
  const questionInput = document.getElementById('agent-lab-question');
  const maxStepsInput = document.getElementById('agent-lab-max-steps');
  const maxToolCallsInput = document.getElementById('agent-lab-max-tool-calls');
  const debugInput = document.getElementById('agent-lab-debug');
  const resultHost = document.getElementById('agent-lab-result');
  if (!statusHost || !docsHost || !form || !resultHost) return;

  const loadStatus = async () => {
    statusHost.textContent = '正在读取 ref/ 与 ref_checklist/ 状态…';
    try {
      const response = await fetch('/api/agent-lab/status');
      const payload = await response.json();
      if (!response.ok || payload.error) throw new Error(payload.error || 'Agent Lab 状态读取失败。');
      renderAgentLabStatus(statusHost, docsHost, payload, checklistDocsHost);
      (payload.profiles || []).forEach((profile) => {
        if (!profileSelect || [...profileSelect.options].some((option) => option.value === profile.id)) return;
        const option = document.createElement('option');
        option.value = profile.id;
        option.textContent = profile.title || profile.id;
        option.dataset.defaultQuestion = profile.default_question || '';
        profileSelect.appendChild(option);
      });
    } catch (error) {
      statusHost.textContent = `Agent Lab 状态读取失败：${error.message || error}`;
    }
  };

  refreshButton?.addEventListener('click', loadStatus);
  reindexButton?.addEventListener('click', async () => {
    reindexButton.disabled = true;
    statusHost.textContent = '正在重建 ref PDF 索引…';
    try {
      const response = await fetch('/api/agent-lab/ref/reindex', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ force: true, max_files: 1000 }),
      });
      const payload = await response.json();
      if (!response.ok || payload.error) throw new Error(payload.error || 'ref PDF 索引重建失败。');
      statusHost.textContent = payload.summary || 'ref PDF 索引已重建。';
      await loadStatus();
    } catch (error) {
      statusHost.textContent = `ref PDF 索引重建失败：${error.message || error}`;
    } finally {
      reindexButton.disabled = false;
    }
  });

  checklistReindexButton?.addEventListener('click', async () => {
    checklistReindexButton.disabled = true;
    statusHost.textContent = '正在重建 review checklist 索引…';
    try {
      const response = await fetch('/api/agent-lab/checklist/reindex', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ force: true, max_files: 1000 }),
      });
      const payload = await response.json();
      if (!response.ok || payload.error) throw new Error(payload.error || 'review checklist 索引重建失败。');
      statusHost.textContent = payload.summary || 'review checklist 索引已重建。';
      await loadStatus();
    } catch (error) {
      statusHost.textContent = `review checklist 索引重建失败：${error.message || error}`;
    } finally {
      checklistReindexButton.disabled = false;
    }
  });

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const submit = form.querySelector('button[type="submit"]');
    if (submit) submit.disabled = true;
    const question = (questionInput?.value || '').trim() || '请参考 ref_checklist 中真实 review 问题模式，帮我检查当前项目有哪些需要优先复核的原理图风险。';
    appendAgentChatMessage(resultHost, 'user', question);
    const loadingMessage = appendAgentChatLoading(resultHost, '规划参考资料检索', {
      stages: ['规划参考资料检索', '读取本地索引', '生成证据引用', '整理实验结论'],
      interval: 1500,
    });
    try {
      const response = await fetch('/api/agent-lab/ask', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          profile: profileSelect?.value || defaultAgentLabProfile,
          question,
          max_steps: Number(maxStepsInput?.value || 8),
          max_tool_calls: Number(maxToolCallsInput?.value || 14),
          debug: Boolean(debugInput?.checked),
        }),
      });
      loadingMessage.updateAgentStage?.('读取 Agent Lab 结果');
      const payload = await response.json();
      loadingMessage.remove();
      if (!response.ok && payload.error) {
        payload.ok = false;
      }
      renderHarnessAgentResult(resultHost, payload, {
        autoOpenTrace: true,
        traceTitle: 'Agent Lab 执行复盘',
        append: true,
      });
      if (questionInput) questionInput.value = '';
      await loadStatus();
    } catch (error) {
      loadingMessage.remove();
      renderHarnessAgentResult(resultHost, { ok: false, error: error.message || String(error), answer: '' }, { append: true });
    } finally {
      if (submit) submit.disabled = false;
    }
  });

  await loadStatus();
}

function projectOptionLabel(project) {
  const name = project.project_name || project.run_id;
  const time = project.generated_at ? ` · ${project.generated_at}` : '';
  return `${name}${time}`;
}

function diffCountText(diff) {
  if (!diff) return '0';
  return `+${diff.added_count || 0} / -${diff.removed_count || 0} / Δ${diff.changed_count || 0}`;
}

function createProjectSelect(projects, selectedRunId) {
  const select = document.createElement('select');
  select.className = 'project-compare-select';
  projects.forEach((project) => {
    const option = document.createElement('option');
    option.value = project.run_id;
    option.textContent = projectOptionLabel(project);
    select.appendChild(option);
  });
  if (selectedRunId && projects.some((project) => project.run_id === selectedRunId)) {
    select.value = selectedRunId;
  }
  return select;
}

function populateProjectSelect(select, projects, selectedRunId) {
  select.replaceChildren();
  projects.forEach((project) => {
    const option = document.createElement('option');
    option.value = project.run_id;
    option.textContent = projectOptionLabel(project);
    select.appendChild(option);
  });
  if (selectedRunId && projects.some((project) => project.run_id === selectedRunId)) {
    select.value = selectedRunId;
  }
}

function projectListNode(projects) {
  const list = document.createElement('div');
  list.className = 'project-list';
  projects.slice(0, 6).forEach((project) => {
    const item = document.createElement('a');
    item.className = 'project-list-item';
    item.href = `/report/${project.run_id}`;
    const title = document.createElement('strong');
    title.textContent = project.project_name || project.run_id;
    const meta = document.createElement('span');
    meta.textContent = `${project.generated_at || '未记录时间'} · 元件 ${project.component_count || 0} · 网络 ${project.net_count || 0}`;
    item.appendChild(title);
    item.appendChild(meta);
    list.appendChild(item);
  });
  return list;
}

function compareProjectSummaryCard(project, sideLabel) {
  const card = document.createElement('article');
  card.className = 'compare-project-card';
  const metrics = [
    ['元件', project.component_count || 0],
    ['网络', project.net_count || 0],
    ['DRC', project.drc_count || 0],
  ];
  const side = document.createElement('span');
  side.textContent = sideLabel;
  const title = document.createElement('h3');
  title.textContent = project.project_name || project.run_id || '未命名项目';
  const time = document.createElement('p');
  time.textContent = project.generated_at || '未记录生成时间';
  const metricWrap = document.createElement('div');
  metricWrap.className = 'compare-project-metrics';
  metrics.forEach(([label, value]) => {
    const strong = document.createElement('strong');
    strong.textContent = value;
    const small = document.createElement('small');
    small.textContent = label;
    strong.appendChild(small);
    metricWrap.appendChild(strong);
  });
  card.append(side, title, time, metricWrap);
  return card;
}

function compareSectionsFromPayload(payload) {
  if (Array.isArray(payload.compare_sections) && payload.compare_sections.length) {
    return payload.compare_sections;
  }
  const overviewRows = payload.overview || [];
  const sections = [
    ['overview', '指标差异', '项目级指标变化。', {
      added_count: 0,
      removed_count: 0,
      changed_count: overviewRows.length,
      rows: overviewRows,
      total_rows: overviewRows.length,
    }, 'overview'],
    ['net_view', 'Net 视角变化', '按左侧网络到右侧网络聚合 Pin/Net 证据。', payload.net_view_diff, 'net'],
    ['key_components', '关键器件增删', '芯片、连接器和其他关键器件增删。', payload.key_component_diff],
    ['key_pin_nets', '关键器件 Pin/Net 连接差异', '关键器件逐 pin 连接差异。', payload.key_pin_net_diff, 'net'],
    ['passive_pin_nets', 'R/C/L Pin/Net 连接差异', '无源件逐 pin 连接差异。', payload.passive_pin_net_diff, 'net'],
    ['components', '元件属性差异', '全量元件属性差异。', payload.component_diff, 'parts'],
    ['nets', '网络节点明细', '网络连接变化。', payload.net_diff, 'net'],
  ];
  return sections.map(([id, title, lead, diff, group]) => ({
    id,
    title,
    lead,
    group: group || (id === 'key_components' ? 'device' : 'detail'),
    diff,
    table: {
      id: `compare_${id}`,
      title,
      count: diff?.total_rows || diff?.rows?.length || 0,
      columns: Object.keys((diff?.rows || [])[0] || {}),
      rows: diff?.rows || [],
      kind_counts: {},
      default_hidden_columns: [],
      sort_profiles: [{ id: 'column', label: '字段排序' }],
    },
  }));
}

function compareRowPrimaryText(row) {
  const fields = ['网络迁移', '位号', '芯片位号', '网络名', '对象', '指标', '引脚', '料号'];
  for (const field of fields) {
    if (row?.[field]) {
      return `${field} ${row[field]}`;
    }
  }
  return row?.类型 || '差异项';
}

function compareSectionTotal(section) {
  const diff = section?.diff || {};
  return (diff.added_count || 0) + (diff.removed_count || 0) + (diff.changed_count || 0);
}

function compareSectionGroup(section) {
  if (section?.group) return section.group;
  const id = section?.id || '';
  if (['net_view', 'key_pin_nets', 'passive_pin_nets', 'nets'].includes(id)) return 'net';
  if (['key_components'].includes(id)) return 'device';
  if (['components'].includes(id)) return 'parts';
  if (String(id).startsWith('report_table_')) return 'report';
  return 'overview';
}

function compareGroupDefinitions(sections) {
  const groups = [
    ['net', '网络视角'],
    ['device', '器件视角'],
    ['parts', '料号属性'],
    ['report', '检查表'],
    ['all', '全部'],
  ];
  return groups.map(([id, label]) => {
    const count = id === 'all'
      ? sections.reduce((sum, section) => sum + compareSectionTotal(section), 0)
      : sections.filter((section) => compareSectionGroup(section) === id)
        .reduce((sum, section) => sum + compareSectionTotal(section), 0);
    return { id, label, count };
  });
}

function applyComparePerspective(root, groupId) {
  if (!root) return;
  const selected = groupId || 'net';
  root.querySelectorAll('[data-compare-group-button]').forEach((button) => {
    const active = button.dataset.compareGroupButton === selected;
    button.classList.toggle('is-active', active);
    button.setAttribute('aria-pressed', active ? 'true' : 'false');
  });
  root.querySelectorAll('[data-compare-section-group]').forEach((node) => {
    const visible = selected === 'all' || node.dataset.compareSectionGroup === selected;
    node.hidden = !visible;
  });
  root.querySelectorAll('[data-compare-nav-group]').forEach((node) => {
    const visible = selected === 'all' || node.dataset.compareNavGroup === selected;
    node.hidden = !visible;
  });
  const focus = root.querySelector('[data-compare-net-focus]');
  if (focus) {
    focus.hidden = selected !== 'net' && selected !== 'all';
  }
}

function comparePerspectiveControls(sections) {
  const wrapper = document.createElement('div');
  wrapper.className = 'compare-perspective-controls';
  const label = document.createElement('span');
  label.textContent = '观察视角';
  wrapper.appendChild(label);
  compareGroupDefinitions(sections).forEach((group) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'compare-perspective-btn';
    button.dataset.compareGroupButton = group.id;
    button.setAttribute('aria-pressed', group.id === 'net' ? 'true' : 'false');
    button.innerHTML = `<strong>${group.label}</strong><small>${group.count}</small>`;
    button.addEventListener('click', () => applyComparePerspective(wrapper.closest('.compare-page-result-shell'), group.id));
    wrapper.appendChild(button);
  });
  return wrapper;
}

function compareNetFocusNode(payload) {
  const diff = payload.net_view_diff || {};
  const rows = diff.rows || [];
  const summary = diff.summary || {};
  const section = document.createElement('section');
  section.className = 'compare-net-focus';
  section.dataset.compareNetFocus = '1';
  const head = document.createElement('div');
  head.className = 'compare-net-focus-head';
  const copy = document.createElement('div');
  const eyebrow = document.createElement('p');
  eyebrow.className = 'eyebrow';
  eyebrow.textContent = 'NET FIRST';
  const title = document.createElement('h2');
  title.textContent = '先从网络变化看影响面';
  const lead = document.createElement('p');
  lead.textContent = '按左侧网络到右侧网络聚合关键器件与 R/C/L pin 证据；这里只做证据归组，不猜测电气等价。';
  copy.append(eyebrow, title, lead);
  const metrics = document.createElement('div');
  metrics.className = 'compare-net-focus-metrics';
  [
    ['网络迁移', summary.transition_count || rows.filter((row) => row?.类型 === '网络迁移').length],
    ['新增网络', summary.net_added_count || 0],
    ['删除网络', summary.net_removed_count || 0],
    ['节点变化', summary.net_changed_count || 0],
  ].forEach(([labelText, value]) => {
    const item = document.createElement('span');
    const strong = document.createElement('strong');
    strong.textContent = value;
    const small = document.createElement('small');
    small.textContent = labelText;
    item.append(strong, small);
    metrics.appendChild(item);
  });
  head.append(copy, metrics);
  section.appendChild(head);

  const list = document.createElement('div');
  list.className = 'compare-net-focus-list';
  if (!rows.length) {
    const empty = document.createElement('p');
    empty.className = 'compare-empty';
    empty.textContent = '未发现网络视角差异。';
    list.appendChild(empty);
  } else {
    rows.slice(0, 8).forEach((row) => {
      const item = document.createElement('article');
      item.className = 'compare-net-focus-row';
      const tag = document.createElement('span');
      tag.className = 'compare-diff-type';
      tag.textContent = row?.类型 || '变化';
      const body = document.createElement('div');
      const titleLine = document.createElement('strong');
      titleLine.textContent = row?.网络迁移 || `${row?.左侧网络 || '未连接'} -> ${row?.右侧网络 || '未连接'}`;
      const meta = document.createElement('p');
      const sample = row?.样例引脚 ? ` · ${row.样例引脚}` : '';
      meta.textContent = `影响 ${row?.影响位号数 || 0} 个位号 / ${row?.影响引脚数 || 0} 个引脚，关键器件 ${row?.关键器件数 || 0}，R/C/L ${row?.['R/C/L数'] || 0}${sample}`;
      body.append(titleLine, meta);
      const nodes = document.createElement('span');
      nodes.className = 'compare-net-node-count';
      nodes.textContent = `${row?.左侧节点数 || 0} -> ${row?.右侧节点数 || 0} 节点`;
      item.append(tag, body, nodes);
      list.appendChild(item);
    });
    if (diff.total_rows > rows.length || diff.truncated) {
      const note = document.createElement('p');
      note.className = 'compare-note';
      note.textContent = diff.total_rows > rows.length
        ? `Net 视角先展示 ${rows.length} / ${diff.total_rows} 行，可调大明细上限后重算。`
        : '上游 Pin/Net 或网络明细已截断，Net 视角摘要以当前明细为准。';
      list.appendChild(note);
    }
  }
  section.appendChild(list);
  return section;
}

function compactCompareCell(value, limit = 150) {
  const text = normalizeText(value);
  if (!text) return '—';
  return text.length > limit ? `${text.slice(0, limit)}…` : text;
}

function compareRowDeltaText(row) {
  const networkPair = row?.左侧网络 || row?.右侧网络
    ? `${compactCompareCell(row?.左侧网络, 80)} → ${compactCompareCell(row?.右侧网络, 80)}`
    : '';
  const valuePair = row?.左侧 || row?.右侧
    ? `${compactCompareCell(row?.左侧)} → ${compactCompareCell(row?.右侧)}`
    : '';
  return networkPair || valuePair || compactCompareCell(row?.变化字段 || row?.类型 || '内容变化');
}

function comparePreviewNode(section) {
  const rows = section.table?.rows || [];
  const previewLimit = 4;
  const wrapper = document.createElement('div');
  wrapper.className = 'compare-diff-preview';
  const header = document.createElement('div');
  header.className = 'compare-diff-preview-head';
  const title = document.createElement('strong');
  title.textContent = '关键差异预览';
  const meta = document.createElement('span');
  const displayed = rows.length;
  const total = section.diff?.total_rows || displayed;
  const previewCount = Math.min(displayed, previewLimit);
  meta.textContent = total > displayed ? `先展示 ${previewCount} / ${total}` : `展示 ${previewCount} / ${total}`;
  header.append(title, meta);
  wrapper.appendChild(header);

  rows.slice(0, previewLimit).forEach((row) => {
    const item = document.createElement('article');
    item.className = 'compare-diff-preview-row';
    const type = document.createElement('span');
    type.className = 'compare-diff-type';
    type.textContent = row?.类型 || '变化';
    const body = document.createElement('div');
    const primary = document.createElement('strong');
    primary.textContent = compareRowPrimaryText(row);
    const delta = document.createElement('p');
    delta.textContent = compareRowDeltaText(row);
    body.append(primary, delta);
    item.append(type, body);
    wrapper.appendChild(item);
  });

  if (section.diff?.truncated) {
    const note = document.createElement('p');
    note.className = 'compare-note';
    note.textContent = `当前仅拉取前 ${displayed} 行差异，总差异 ${section.diff.total_rows} 行；可调大“明细上限”后重新生成。`;
    wrapper.appendChild(note);
  }
  return wrapper;
}

function compareSectionNode(section, index) {
  const wrapper = document.createElement('section');
  wrapper.className = `compare-domain-section priority-${section.priority || 'normal'}`;
  wrapper.id = `compare-section-${section.id}`;
  wrapper.dataset.compareSectionGroup = compareSectionGroup(section);
  wrapper.setAttribute('data-reveal', '');
  const diff = section.diff || {};
  const total = compareSectionTotal(section);
  const head = document.createElement('div');
  head.className = 'compare-domain-head';
  const copy = document.createElement('div');
  const eyebrow = document.createElement('p');
  eyebrow.className = 'eyebrow';
  eyebrow.textContent = String(section.id || 'compare').toUpperCase();
  const title = document.createElement('h2');
  title.textContent = section.title || '对比分区';
  const lead = document.createElement('p');
  lead.textContent = section.lead || '';
  copy.append(eyebrow, title, lead);
  const stat = document.createElement('div');
  stat.className = 'compare-domain-stat';
  const totalNode = document.createElement('strong');
  totalNode.textContent = total;
  const breakdown = document.createElement('span');
  breakdown.textContent = diffCountText(diff);
  stat.append(totalNode, breakdown);
  head.append(copy, stat);
  wrapper.appendChild(head);
  if (section.table?.rows?.length) {
    wrapper.appendChild(comparePreviewNode(section));
    wrapper.appendChild(tableBlock(section.table, false));
  } else {
    const empty = document.createElement('p');
    empty.className = 'compare-empty';
    empty.textContent = '未发现差异。';
    wrapper.appendChild(empty);
  }
  return wrapper;
}

function renderComparePageResult(host, payload) {
  host.replaceChildren();
  const result = document.createElement('div');
  result.className = 'compare-page-result-shell compare-result-enter';
  const sections = compareSectionsFromPayload(payload);
  const title = document.createElement('div');
  title.className = 'compare-result-title compare-page-title';
  title.append(
    compareProjectSummaryCard(payload.left || {}, 'LEFT'),
    (() => {
      const versus = document.createElement('strong');
      versus.textContent = 'vs';
      return versus;
    })(),
    compareProjectSummaryCard(payload.right || {}, 'RIGHT'),
  );
  result.appendChild(title);

  const cards = document.createElement('div');
  cards.className = 'compare-stat-grid compare-page-stat-grid';
  [
    ['指标变化', payload.diff_totals?.overview || 0],
    ['Net 视角', payload.diff_totals?.net_view || 0],
    ['关键器件', payload.diff_totals?.key_components || 0],
    ['关键 Pin/Net', payload.diff_totals?.key_pin_nets || 0],
    ['R/C/L Pin/Net', payload.diff_totals?.passive_pin_nets || 0],
    ['元件属性', payload.diff_totals?.components || 0],
    ['网络连接', payload.diff_totals?.nets || 0],
    ['检查表', payload.diff_totals?.tables || 0],
  ].forEach(([label, value]) => {
    const card = document.createElement('article');
    card.className = 'compare-stat';
    card.innerHTML = `<span>${label}</span><strong>${value}</strong>`;
    cards.appendChild(card);
  });
  result.appendChild(cards);
  result.appendChild(compareNetFocusNode(payload));
  result.appendChild(comparePerspectiveControls(sections));

  const nav = document.createElement('nav');
  nav.className = 'compare-domain-nav';
  sections.forEach((section) => {
    const link = document.createElement('a');
    link.href = `#compare-section-${section.id}`;
    link.textContent = section.title;
    link.dataset.compareNavGroup = compareSectionGroup(section);
    nav.appendChild(link);
  });
  result.appendChild(nav);

  const stack = document.createElement('div');
  stack.className = 'compare-domain-stack';
  sections.forEach((section, index) => stack.appendChild(compareSectionNode(section, index)));
  result.appendChild(stack);
  host.appendChild(result);
  applyComparePerspective(result, 'net');
  bootReveals(result);
  staggerChildren(cards, '.compare-stat');
  staggerChildren(stack, '.compare-domain-section', 2);
  restartMotion(result, 'compare-result-enter', 920);
}

async function renderProjectManager(currentRunId = '') {
  const host = document.getElementById('project-manager');
  if (!host) return;
  const body = host.querySelector('.project-manager-body') || host;
  body.textContent = '正在读取会话项目…';
  try {
    const response = await fetch('/api/projects');
    const payload = await response.json();
    const projects = payload.projects || [];
    body.replaceChildren();
    if (!projects.length) {
      const empty = document.createElement('p');
      empty.className = 'compare-empty';
      empty.textContent = '当前会话还没有已分析项目。生成报告后，这里会自动出现项目列表。';
      body.appendChild(empty);
      return;
    }
    const projectList = projectListNode(projects);
    body.appendChild(projectList);
    staggerChildren(projectList, '.project-list-item');
    const form = document.createElement('form');
    form.className = 'project-compare-form';
    const leftDefault = currentRunId || projects[0]?.run_id || '';
    const rightDefault = projects.find((project) => project.run_id !== leftDefault)?.run_id || projects[1]?.run_id || leftDefault;
    const leftSelect = createProjectSelect(projects, leftDefault);
    const rightSelect = createProjectSelect(projects, rightDefault);
    const submit = document.createElement('button');
    submit.type = 'submit';
    submit.className = 'primary-btn inline-btn';
    submit.textContent = '进入项目对比';
    form.appendChild(leftSelect);
    form.appendChild(rightSelect);
    form.appendChild(submit);
    body.appendChild(form);

    const compareLink = document.createElement('a');
    compareLink.className = 'ghost-btn inline-btn';
    compareLink.href = '/compare';
    compareLink.textContent = '打开独立对比工作台';
    body.appendChild(compareLink);
    restartMotion(host, 'project-manager-refresh', 900);

    const syncCompareHref = () => {
      const params = new URLSearchParams();
      if (leftSelect.value) params.set('left_run_id', leftSelect.value);
      if (rightSelect.value) params.set('right_run_id', rightSelect.value);
      const query = params.toString();
      compareLink.href = `/compare${query ? `?${query}` : ''}`;
    };
    leftSelect.addEventListener('change', syncCompareHref);
    rightSelect.addEventListener('change', syncCompareHref);
    syncCompareHref();

    form.addEventListener('submit', (event) => {
      event.preventDefault();
      syncCompareHref();
      window.location.href = compareLink.href;
    });
  } catch (error) {
    body.textContent = error.message || String(error);
  }
}

async function bootComparePage() {
  const form = document.getElementById('compare-page-form');
  const leftSelect = document.getElementById('compare-left-run');
  const rightSelect = document.getElementById('compare-right-run');
  const detailLimitInput = document.getElementById('compare-detail-limit');
  const status = document.getElementById('compare-projects-status');
  const resultHost = document.getElementById('compare-result-host');
  if (!form || !leftSelect || !rightSelect || !resultHost) return;

  const state = {
    payload: null,
  };
  window.PSTX_COMPARE_CONTEXT = state;
  bootAsterFloatingPanel();
  bootAsterStatus();
  bootAsterCredentialForm();
  bootCompareHarnessAgent(state);

  const params = new URLSearchParams(window.location.search);
  try {
    const response = await fetch('/api/projects');
    const payload = await response.json();
    const projects = payload.projects || [];
    if (!projects.length) {
      status.textContent = '当前会话还没有已分析项目。请先返回首页分析至少两个项目。';
      return;
    }
    const leftDefault = params.get('left_run_id') || projects[0]?.run_id || '';
    const rightDefault = params.get('right_run_id')
      || projects.find((project) => project.run_id !== leftDefault)?.run_id
      || projects[1]?.run_id
      || '';
    populateProjectSelect(leftSelect, projects, leftDefault);
    populateProjectSelect(rightSelect, projects, rightDefault);
    status.textContent = `已读取 ${projects.length} 个会话项目。`;
    if (leftSelect.value && rightSelect.value && leftSelect.value !== rightSelect.value && params.has('left_run_id')) {
      window.setTimeout(() => form.requestSubmit(), 0);
    }
  } catch (error) {
    status.textContent = `项目列表读取失败：${error.message || error}`;
  }

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    resultHost.innerHTML = '<p class="compare-empty">正在生成完整项目差异…</p>';
    restartMotion(resultHost, 'compare-loading-pulse', 480);
    const submit = form.querySelector('button[type="submit"]');
    if (submit) submit.disabled = true;
    try {
      const response = await fetch('/api/compare', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          left_run_id: leftSelect.value,
          right_run_id: rightSelect.value,
          detail_limit: Number(detailLimitInput?.value || 1000),
        }),
      });
      const payload = await response.json();
      if (!response.ok || !payload.ok) {
        throw new Error(payload.error || '对比失败。');
      }
      state.payload = payload;
      const query = new URLSearchParams({
        left_run_id: leftSelect.value,
        right_run_id: rightSelect.value,
      });
      window.history.replaceState({}, '', `/compare?${query.toString()}`);
      renderComparePageResult(resultHost, payload);
      status.textContent = `对比完成：${payload.left.project_name} vs ${payload.right.project_name}`;
    } catch (error) {
      state.payload = null;
      resultHost.innerHTML = `<p class="compare-empty compare-error">${error.message || error}</p>`;
      status.textContent = '对比失败，请检查两个项目是否仍在当前会话中。';
    } finally {
      if (submit) submit.disabled = false;
    }
  });

}

function bootCompareHarnessAgent(state) {
  const form = document.getElementById('harness-agent-form');
  const profileSelect = document.getElementById('harness-agent-profile');
  const questionInput = document.getElementById('harness-agent-question');
  const maxStepsInput = document.getElementById('harness-agent-max-steps');
  const maxToolCallsInput = document.getElementById('harness-agent-max-tool-calls');
  const debugInput = document.getElementById('harness-agent-debug');
  const detailLimitInput = document.getElementById('compare-detail-limit');
  const resultHost = document.getElementById('harness-agent-result');
  const summaryButton = document.getElementById('aster-summary-button');
  if (!form || !profileSelect || !resultHost) return;
  if (form.dataset.compareAgentChatBooted) return;
  form.dataset.compareAgentChatBooted = '1';
  ensureAgentChatThread(resultHost);
  appendAgentChatMessage(resultHost, 'system', '对比页浮窗已进入连续对话模式。先生成 A/B 项目对比，然后可以连续追问差异、证据和风险。');

  if (!form.dataset.chatToolbarReady) {
    const toolbar = document.createElement('div');
    toolbar.className = 'agent-chat-toolbar';
    const clearButton = document.createElement('button');
    clearButton.type = 'button';
    clearButton.className = 'ghost-btn inline-btn';
    clearButton.textContent = '清空本地对话';
    clearButton.addEventListener('click', () => {
      resultHost.replaceChildren();
      resultHost.dataset.chatReady = '1';
      appendAgentChatMessage(resultHost, 'system', '本地对比浮窗对话已清空；当前 A/B 对比结果仍保留。');
    });
    toolbar.appendChild(clearButton);
    form.insertAdjacentElement('afterend', toolbar);
    form.dataset.chatToolbarReady = '1';
  }

  const profileMap = new Map();
  fetch('/api/compare/harness/profiles')
    .then((response) => response.json())
    .then((payload) => {
      profileSelect.replaceChildren();
      (payload.profiles || []).forEach((profile) => {
        profileMap.set(profile.id, profile);
        const option = document.createElement('option');
        option.value = profile.id;
        option.textContent = profile.title || profile.id;
        option.dataset.defaultQuestion = profile.default_question || '';
        option.dataset.maxSteps = profile.max_steps || '';
        option.dataset.maxToolCalls = profile.max_tool_calls || '';
        profileSelect.appendChild(option);
      });
      const defaultProfile = payload.default_profile || 'compare_quick_scan';
      if ([...profileSelect.options].some((item) => item.value === 'auto')) {
        profileSelect.value = 'auto';
      } else if (profileMap.has(defaultProfile)) {
        profileSelect.value = defaultProfile;
      }
      applyCompareProfileDefaults();
    })
    .catch((error) => {
      appendAgentChatMessage(resultHost, 'system', `Compare Agent Profile 读取失败：${error.message || error}`);
    });

  const applyCompareProfileDefaults = () => {
    const profile = profileMap.get(profileSelect.value);
    if (!profile) return;
    if (questionInput && !questionInput.value.trim()) {
      questionInput.placeholder = profile.default_question || questionInput.placeholder;
    }
    if (maxStepsInput && profile.max_steps) {
      maxStepsInput.value = profile.max_steps;
    }
    if (maxToolCallsInput && profile.max_tool_calls) {
      maxToolCallsInput.value = profile.max_tool_calls;
    }
  };
  profileSelect.addEventListener('change', applyCompareProfileDefaults);

  const run = async (question, button) => {
    if (button) button.disabled = true;
    const finalQuestion = (question || '').trim() || '请深度审查当前项目对比差异。';
    appendAgentChatMessage(resultHost, 'user', finalQuestion);
    const loadingMessage = appendAgentChatLoading(resultHost, '确认 A/B 项目', { kind: 'compare' });
    resultHost.dataset.activeLoadingMessage = '1';
    restartMotion(loadingMessage, 'query-loading-pulse', 480);
    try {
      const responsePayload = state.payload;
      if (!responsePayload) {
        throw new Error('请先生成项目对比结果。');
      }
      const profile = profileSelect.value || 'auto';
      const response = await fetch('/api/compare/harness-agent', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          left_run_id: responsePayload.left?.run_id || '',
          right_run_id: responsePayload.right?.run_id || '',
          profile,
          question: finalQuestion,
          max_steps: Number(maxStepsInput?.value || 8),
          max_tool_calls: Number(maxToolCallsInput?.value || 14),
          detail_limit: Number(detailLimitInput?.value || responsePayload.detail_limit || 500),
          debug: Boolean(debugInput?.checked),
          async: true,
        }),
      });
      const payload = await parseAgentResponseOrPoll(response, { loadingMessage });
      loadingMessage.updateAgentStage?.('整理对比回答');
      loadingMessage.remove();
      renderHarnessAgentResult(resultHost, payload, { autoOpenTrace: true, traceTitle: '对比 Agent 执行复盘', append: true });
    } catch (error) {
      loadingMessage.remove();
      renderHarnessAgentResult(resultHost, { ok: false, error: error.message || String(error), answer: '', ...(error.payload || {}) }, { append: true });
    } finally {
      delete resultHost.dataset.activeLoadingMessage;
      if (button) button.disabled = false;
    }
  };

  summaryButton?.addEventListener('click', () => run('请从工程审查角度深度审查当前 A/B 项目的最高优先级差异，重点关注关键器件、芯片/连接器 Pin-Net、R/C/L、网络和飞书 PI/选型顺序。', summaryButton));
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const submit = form.querySelector('button[type="submit"]');
    run(questionInput?.value || '请深度审查当前项目对比差异。', submit);
    if (questionInput) questionInput.value = '';
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

function bootInspectorToggle() {
  const button = document.getElementById('inspector-toggle');
  if (!button) return;

  let collapsed = false;
  try {
    collapsed = window.localStorage.getItem(INSPECTOR_STORAGE_KEY) === '1';
  } catch {
    collapsed = false;
  }
  setInspectorCollapsed(collapsed);
  button.addEventListener('click', () => {
    const layout = document.querySelector('.report-layout');
    const nextCollapsed = !layout?.classList.contains('is-inspector-collapsed');
    setInspectorCollapsed(nextCollapsed);
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
    restartMotion(results, 'query-loading-pulse', 480);
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
  const debugFixture = Boolean(context.debugFixture || document.body?.dataset.debugFixture === 'true');
  let report = window.PSTX_DEBUG_REPORT;
  if (!debugFixture || !report) {
    const response = await fetch(`/api/report/${context.runId}`);
    report = await response.json();
  }

  renderSummary(report);
  const host = document.getElementById('report-sections');
  host.replaceChildren();
  report.sections.forEach((section) => host.appendChild(sectionNode(section)));
  renderSectionNav(report.sections, report.review_plan);
  applyGlobalStaggers(document);
  document.body.classList.add('report-data-ready');
  bootSidebarToggle();
  bootInspectorToggle();
  bootReveals();
  if (debugFixture) {
    const manager = document.querySelector('#project-manager .project-manager-body');
    if (manager) {
      manager.innerHTML = '<p class="empty-state">Debug fixture 不绑定真实 run，只用于观察报告页布局、卡片密度和表格交互。</p>';
    }
    const queryResults = document.getElementById('query-results');
    if (queryResults) {
      queryResults.innerHTML = '<div class="query-empty">Debug fixture 不执行真实查询；请在真实报告页使用查询。</div>';
    }
    const summaryButton = document.getElementById('aster-summary-button');
    if (summaryButton) {
      summaryButton.hidden = true;
    }
  } else {
    renderProjectManager(context.runId);
    handleQuery(context.runId);
    bootAsterSummary(context.runId);
    bootAsterChatAgent(context.runId);
  }
  bootAsterFloatingPanel();
  bootAsterStatus();
  bootAsterCredentialForm();
}

window.PSTXApp = Object.assign(window.PSTXApp || {}, {
  bootCommon() {
    bootPageMotion();
    bootUiDebugMode();
    bootReveals();
    bootAgentTraceDrawer();
  },
  handleHomePage,
  handleFeishuSyncPage,
  bootFeishuDbPage,
  bootAsterStatus,
  bootAsterCredentialForm,
  bootComparePage,
  handleReportPage,
  bootAgentEvalPage,
  bootAgentLabPage,
  renderProjectManager,
});
