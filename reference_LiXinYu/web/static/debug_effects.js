(function () {
  function reducedMotion() {
    const query = window.matchMedia?.('(prefers-reduced-motion: reduce)');
    return Boolean(query?.matches);
  }

  function replayClass(node, className, duration = 760) {
    if (!node || reducedMotion()) return;
    node.classList.remove(className);
    void node.offsetWidth;
    node.classList.add(className);
    window.setTimeout(() => node.classList.remove(className), duration);
  }

  function setStagger(parent, selector) {
    parent?.querySelectorAll(selector).forEach((node, index) => {
      node.style.setProperty('--stagger-index', String(Math.min(index, 10)));
    });
  }

  function replayOverview() {
    const stage = document.querySelector('[data-debug-stage="overview"]');
    setStagger(stage, '.metric, .insight-card');
    replayClass(stage, 'debug-replay', 900);
  }

  function setTableOpen(open) {
    const table = document.querySelector('.debug-table-demo');
    if (!table) return;
    const button = table.querySelector('.toggle-btn');
    table.classList.toggle('is-open', open);
    if (button) {
      button.textContent = open ? '收起' : '查看详情';
    }
    if (open) {
      replayClass(table, 'table-open-pulse', 420);
    }
  }

  function toggleTable() {
    const table = document.querySelector('.debug-table-demo');
    setTableOpen(!table?.classList.contains('is-open'));
  }

  function compareBlock(title, status, count) {
    return `
      <details class="compare-block" open>
        <summary>${title} · ${status}</summary>
        <div class="query-data-table">
          <table>
            <thead>
              <tr><th>对象</th><th>类型</th><th>说明</th></tr>
            </thead>
            <tbody>
              <tr><td>${title}-A</td><td>新增</td><td>模拟新增 ${count} 项</td></tr>
              <tr><td>${title}-B</td><td>变化</td><td>模拟字段变化，用于观察展开块动画</td></tr>
            </tbody>
          </table>
        </div>
      </details>
    `;
  }

  function renderCompare() {
    const host = document.querySelector('[data-debug-compare-host]');
    if (!host) return;
    host.innerHTML = `
      <div class="compare-result compare-result-enter">
        <div class="compare-result-title">
          <span>GPU_2SW_BOARD_A00</span>
          <strong>vs</strong>
          <span>GPU_2SW_BOARD_A01</span>
        </div>
        <div class="compare-stat-grid">
          <div class="compare-stat"><span>指标变化</span><strong>6</strong></div>
          <div class="compare-stat"><span>元件差异</span><strong>18</strong></div>
          <div class="compare-stat"><span>网络差异</span><strong>9</strong></div>
          <div class="compare-stat"><span>结果表差异</span><strong>4</strong></div>
        </div>
        ${compareBlock('元件差异', '+12 / -4 / Δ2', 12)}
        ${compareBlock('网络差异', '+5 / -1 / Δ3', 5)}
      </div>
    `;
    const result = host.querySelector('.compare-result');
    setStagger(result, '.compare-stat, .compare-block');
    replayClass(result, 'compare-result-enter', 920);
  }

  function replayAll() {
    replayOverview();
    setTableOpen(true);
    renderCompare();
    document.querySelectorAll('[data-reveal]').forEach((node) => {
      replayClass(node, 'debug-reveal-pulse', 640);
    });
  }

  document.addEventListener('DOMContentLoaded', () => {
    setStagger(document, '.metric, .insight-card, .compare-stat, .compare-block');
    setTableOpen(true);
    renderCompare();

    document.querySelectorAll('[data-debug-action]').forEach((control) => {
      control.addEventListener('click', () => {
        const action = control.dataset.debugAction;
        if (action === 'replay-all') replayAll();
        if (action === 'replay-overview') replayOverview();
        if (action === 'toggle-table') toggleTable();
        if (action === 'replay-compare') renderCompare();
      });
    });
  });
}());
