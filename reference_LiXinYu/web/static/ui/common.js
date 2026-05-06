(function () {
  function prefersReducedMotion() {
    const query = window.matchMedia?.('(prefers-reduced-motion: reduce)');
    return Boolean(query?.matches);
  }

  function restartMotion(node, className, duration = 420) {
    if (!node || prefersReducedMotion()) return;
    node.classList.remove(className);
    void node.offsetWidth;
    node.classList.add(className);
    window.setTimeout(() => node.classList.remove(className), duration);
  }

  function staggerChildren(parent, selector, baseDelay = 0) {
    if (!parent) return;
    parent.querySelectorAll(selector).forEach((node, index) => {
      node.style.setProperty('--stagger-index', String(Math.min(index + baseDelay, 10)));
    });
  }

  function runWhenBrowserIsIdle(callback) {
    if ('requestIdleCallback' in window) {
      window.requestIdleCallback(callback, { timeout: 180 });
      return;
    }
    requestAnimationFrame(callback);
  }

  function showLoadingMask(mask) {
    if (!mask) return;
    mask.hidden = false;
    mask.classList.add('is-active');
    requestAnimationFrame(() => mask.classList.add('is-visible'));
  }

  function hideLoadingMask(mask) {
    if (!mask) return;
    mask.classList.remove('is-visible');
    window.setTimeout(() => {
      mask.classList.remove('is-active');
      mask.hidden = true;
    }, prefersReducedMotion() ? 0 : 220);
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
    target.scrollIntoView({ behavior: prefersReducedMotion() ? 'auto' : 'smooth', block: 'start' });
    restartMotion(target, 'section-flash', 1400);
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

  function normalizeText(value) {
    return String(value ?? '').trim();
  }

  window.PSTXUI = Object.assign(window.PSTXUI || {}, {
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
  });
}());
