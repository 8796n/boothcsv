(function(global){
  'use strict';

  const state = {
    sortKey: 'paymentDate',
    sortDirection: 'desc',
    currentPage: 1,
    pageSize: 20,
    unprintedOnly: false,
    totalItems: 0,
    totalPages: 1,
    pageItems: []
  };

  const selection = new Set();

  const ui = {
    panel: null,
    body: null,
    selectAll: null,
    deleteButton: null,
    previewButton: null,
    prev: null,
    next: null,
    pageInfo: null,
    empty: null,
    filter: null,
    headerCells: []
  };

  let unsubscribe = null;
  const dependencies = {
    ensureOrderRepository: null,
    clearPreviousResults: null,
    hideQuickGuide: null,
    renderPreview: null
  };

  function assignDependencies(options){
    dependencies.ensureOrderRepository = options.ensureOrderRepository || dependencies.ensureOrderRepository;
    dependencies.clearPreviousResults = options.clearPreviousResults || dependencies.clearPreviousResults;
    dependencies.hideQuickGuide = options.hideQuickGuide || dependencies.hideQuickGuide;
    dependencies.renderPreview = options.renderPreview || dependencies.renderPreview;
  }

  async function ensureRepository(){
    if (typeof dependencies.ensureOrderRepository !== 'function') return null;
    return await dependencies.ensureOrderRepository();
  }

  function parseDateToTimestamp(value){
    if (!value) return Number.NaN;
    const direct = Date.parse(value);
    if (!Number.isNaN(direct)) return direct;
    const normalized = value.replace(/T/, ' ').replace(/-/g, '/');
    const fallback = Date.parse(normalized);
    return Number.isNaN(fallback) ? Number.NaN : fallback;
  }

  function formatDateTime(value){
    if (!value) return '';
    const date = new Date(value);
    if (!Number.isNaN(date.getTime())) return date.toLocaleString();
    const normalized = value.replace(/T/, ' ').replace(/-/g, '/');
    const fallback = new Date(normalized);
    if (!Number.isNaN(fallback.getTime())) return fallback.toLocaleString();
    return value;
  }

  function buildPreviewConfig(){
    const cache = global.settingsCache || {};
    const enabledCustom = !!cache.customLabelEnable;
    let customLabels = [];
    if (enabledCustom && global.CustomLabels && typeof global.CustomLabels.getFromUI === 'function'){
      customLabels = global.CustomLabels.getFromUI().filter(label => label && label.enabled);
    }
    return {
      labelyn: !!cache.labelyn,
      labelskip: cache.labelskip,
      sortByPaymentDate: !!cache.sortByPaymentDate,
      customLabelEnable: enabledCustom,
      customLabels
    };
  }

  function updateSortIndicators(){
    if (!Array.isArray(ui.headerCells)) return;
    ui.headerCells.forEach(cell => {
      cell.classList.remove('sort-asc', 'sort-desc');
      if (cell.dataset.sort === state.sortKey){
        cell.classList.add(state.sortDirection === 'asc' ? 'sort-asc' : 'sort-desc');
      }
    });
  }

  function syncSelectAll(pageItems){
    if (!ui.selectAll) return;
    if (pageItems.length === 0){
      ui.selectAll.checked = false;
      ui.selectAll.indeterminate = false;
      ui.selectAll.disabled = state.totalItems === 0;
      return;
    }
    const selectedCount = pageItems.filter(item => selection.has(item.orderNumber)).length;
    ui.selectAll.disabled = false;
    ui.selectAll.checked = selectedCount === pageItems.length;
    ui.selectAll.indeterminate = selectedCount > 0 && selectedCount < pageItems.length;
  }

  function syncActionButtons(){
    const count = selection.size;
    if (ui.deleteButton){
      ui.deleteButton.disabled = count === 0;
      ui.deleteButton.textContent = count > 0 ? `選択した注文を削除 (${count})` : '選択した注文を削除';
    }
    if (ui.previewButton){
      ui.previewButton.disabled = count === 0;
      ui.previewButton.textContent = count > 0 ? `選択した注文を印刷プレビュー (${count})` : '選択した注文を印刷プレビュー';
    }
  }

  function updatePagination(totalItems){
    if (ui.prev){
      ui.prev.disabled = state.currentPage <= 1 || totalItems === 0;
    }
    if (ui.next){
      ui.next.disabled = state.currentPage >= state.totalPages || totalItems === 0;
    }
    if (ui.pageInfo){
      ui.pageInfo.textContent = totalItems === 0
        ? '0件'
        : `${state.currentPage} / ${state.totalPages} ページ（全${totalItems}件）`;
    }
  }

  function compareEntries(a, b){
    const dir = state.sortDirection === 'asc' ? 1 : -1;
    let result = 0;
    if (state.sortKey === 'paymentDate'){
      const av = Number.isNaN(a.paymentDateValue) ? -Infinity : a.paymentDateValue;
      const bv = Number.isNaN(b.paymentDateValue) ? -Infinity : b.paymentDateValue;
      if (av < bv) result = -1;
      else if (av > bv) result = 1;
    } else if (state.sortKey === 'printedAt'){
      const av = Number.isNaN(a.printedAtValue) ? -Infinity : a.printedAtValue;
      const bv = Number.isNaN(b.printedAtValue) ? -Infinity : b.printedAtValue;
      if (av < bv) result = -1;
      else if (av > bv) result = 1;
    } else {
      result = a.orderNumber.localeCompare(b.orderNumber, 'ja', { numeric: true, sensitivity: 'base' });
    }
    if (result === 0){
      result = a.orderNumber.localeCompare(b.orderNumber, 'ja', { numeric: true, sensitivity: 'base' });
    }
    return result * dir;
  }

  async function refresh(){
    if (!ui.panel) return;
    const repo = await ensureRepository();
    if (!repo) return;
    const constants = global.CONSTANTS;
    if (!constants || !constants.CSV){
      console.error('ProcessedOrdersPanel: CONSTANTS.CSV が利用できません');
      return;
    }

    const records = repo.getAll();
    const data = records.map(rec => {
      const paymentRaw = rec && rec.row ? rec.row[constants.CSV.PAYMENT_DATE_COLUMN] || '' : '';
      const printedRaw = rec && rec.printedAt ? rec.printedAt : '';
      return {
        orderNumber: rec.orderNumber,
        paymentDateRaw: paymentRaw,
        paymentDateValue: parseDateToTimestamp(paymentRaw),
        printedAtRaw: printedRaw,
        printedAtValue: parseDateToTimestamp(printedRaw),
        printed: !!printedRaw
      };
    }).filter(item => !!item.orderNumber);

    const sorted = [...data].sort(compareEntries);
    const filtered = state.unprintedOnly ? sorted.filter(item => !item.printed) : sorted;

    state.totalItems = filtered.length;
    state.totalPages = filtered.length === 0 ? 1 : Math.ceil(filtered.length / state.pageSize);
    if (state.currentPage > state.totalPages){
      state.currentPage = state.totalPages;
    }
    if (filtered.length === 0){
      state.currentPage = 1;
    }

    const startIndex = (state.currentPage - 1) * state.pageSize;
    const pageItems = filtered.slice(startIndex, startIndex + state.pageSize);
    state.pageItems = pageItems;

    const validKeys = new Set(data.map(item => item.orderNumber));
    selection.forEach(key => {
      if (!validKeys.has(key)) selection.delete(key);
    });

    if (ui.body){
      ui.body.textContent = '';
      pageItems.forEach(item => {
        const tr = global.document.createElement('tr');

        const selectTd = global.document.createElement('td');
        selectTd.className = 'col-select';
        const checkbox = global.document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.dataset.orderNumber = item.orderNumber;
        checkbox.checked = selection.has(item.orderNumber);
        checkbox.setAttribute('aria-label', `注文 ${item.orderNumber} を選択`);
        checkbox.addEventListener('change', (ev) => handleRowSelection(item.orderNumber, ev.target.checked));
        selectTd.appendChild(checkbox);
        tr.appendChild(selectTd);

        const orderTd = global.document.createElement('td');
        const link = global.document.createElement('a');
        link.href = `https://manage.booth.pm/orders/${encodeURIComponent(item.orderNumber)}`;
        link.target = '_blank';
        link.rel = 'noopener';
        link.className = 'order-link';
        link.textContent = item.orderNumber;
        orderTd.appendChild(link);
        tr.appendChild(orderTd);

        const paymentTd = global.document.createElement('td');
        paymentTd.textContent = item.paymentDateRaw || '';
        tr.appendChild(paymentTd);

        const printedTd = global.document.createElement('td');
        printedTd.textContent = item.printedAtRaw ? formatDateTime(item.printedAtRaw) : '';
        tr.appendChild(printedTd);

        ui.body.appendChild(tr);
      });
    }

    if (ui.empty){
      if (filtered.length === 0){
        ui.empty.hidden = false;
        ui.empty.textContent = state.unprintedOnly ? '未印刷の注文はありません。' : '保存されている注文はありません。';
      } else {
        ui.empty.hidden = true;
      }
    }

    syncSelectAll(pageItems);
    syncActionButtons();
    updatePagination(filtered.length);
    updateSortIndicators();
    updateVisibility();
  }

  function handleSort(key){
    if (!key) return;
    if (state.sortKey === key){
      state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
      state.sortKey = key;
      state.sortDirection = key === 'orderNumber' ? 'asc' : 'desc';
    }
    state.currentPage = 1;
    refresh().catch(err => console.error('processed orders refresh error', err));
  }

  function handleSelectAll(event){
    const checked = event.target.checked;
    state.pageItems.forEach(item => {
      if (checked) selection.add(item.orderNumber);
      else selection.delete(item.orderNumber);
    });
    if (ui.body){
      ui.body.querySelectorAll('input[type="checkbox"][data-order-number]').forEach(box => {
        box.checked = selection.has(box.dataset.orderNumber);
      });
    }
    syncSelectAll(state.pageItems);
    syncActionButtons();
  }

  function handleRowSelection(orderNumber, checked){
    if (!orderNumber) return;
    if (checked) selection.add(orderNumber);
    else selection.delete(orderNumber);
    syncSelectAll(state.pageItems);
    syncActionButtons();
  }

  function handleFilterChange(event){
    state.unprintedOnly = !!event.target.checked;
    state.currentPage = 1;
    refresh().catch(err => console.error('processed orders refresh error', err));
  }

  function handlePrevPage(){
    if (state.currentPage <= 1) return;
    state.currentPage -= 1;
    refresh().catch(err => console.error('processed orders refresh error', err));
  }

  function handleNextPage(){
    if (state.currentPage >= state.totalPages) return;
    state.currentPage += 1;
    refresh().catch(err => console.error('processed orders refresh error', err));
  }

  async function handlePreview(){
    if (selection.size === 0) return;
    if (typeof dependencies.clearPreviousResults === 'function'){
      dependencies.clearPreviousResults();
    }
    if (typeof dependencies.hideQuickGuide === 'function'){
      dependencies.hideQuickGuide();
    }
    if (typeof dependencies.renderPreview !== 'function') return;
    const orderNumbers = Array.from(selection);
    const config = buildPreviewConfig();
    try {
      await dependencies.renderPreview(orderNumbers, config);
    } catch (e){
      console.error('order preview error', e);
      global.alert('プレビューの生成に失敗しました: ' + (e && e.message ? e.message : e));
    }
  }

  async function handleDelete(){
    if (selection.size === 0) return;
    const orderNumbers = Array.from(selection);
    if (!global.confirm(`選択した${orderNumbers.length}件の注文を削除しますか？`)) return;
    const repo = await ensureRepository();
    if (!repo) return;
    try {
      await repo.deleteMany(orderNumbers);
      selection.clear();
      global.alert('選択した注文を削除しました');
    } catch (e){
      console.error('order delete error', e);
      global.alert('注文の削除に失敗しました: ' + (e && e.message ? e.message : e));
    }
    await refresh();
  }

  function updateVisibility(){
    if (!ui.panel) return;
    const fileInput = global.document.getElementById('file');
    const hasFile = !!(fileInput && fileInput.files && fileInput.files.length > 0);
    ui.panel.hidden = hasFile;
  }

  function attachEventHandlers(){
    if (ui.selectAll){
      ui.selectAll.addEventListener('change', handleSelectAll);
    }
    if (ui.filter){
      ui.filter.addEventListener('change', handleFilterChange);
    }
    if (ui.prev){
      ui.prev.addEventListener('click', handlePrevPage);
    }
    if (ui.next){
      ui.next.addEventListener('click', handleNextPage);
    }
    if (ui.deleteButton){
      ui.deleteButton.addEventListener('click', handleDelete);
    }
    if (ui.previewButton){
      ui.previewButton.addEventListener('click', handlePreview);
    }
    if (Array.isArray(ui.headerCells)){
      ui.headerCells.forEach(cell => {
        const key = cell.dataset.sort;
        if (!key) return;
        cell.addEventListener('click', () => handleSort(key));
      });
    }
  }

  function cacheUIElements(){
    ui.panel = global.document.getElementById('processedOrdersPanel');
    ui.body = global.document.getElementById('processedOrdersBody');
    ui.selectAll = global.document.getElementById('processedOrdersSelectAll');
    ui.deleteButton = global.document.getElementById('processedOrdersDeleteButton');
    ui.previewButton = global.document.getElementById('processedOrdersPreviewButton');
    ui.prev = global.document.getElementById('processedOrdersPrev');
    ui.next = global.document.getElementById('processedOrdersNext');
    ui.pageInfo = global.document.getElementById('processedOrdersPageInfo');
    ui.empty = global.document.getElementById('processedOrdersEmpty');
    ui.filter = global.document.getElementById('processedOrdersUnprintedOnly');
    ui.headerCells = ui.panel ? Array.from(ui.panel.querySelectorAll('th.sortable')) : [];
  }

  async function init(options = {}){
    assignDependencies(options);
    const template = global.document.getElementById('processedOrdersPanelTemplate');
    if (!template) return;

    const repo = await ensureRepository();
    if (!repo) return;

    if (!ui.panel){
      const fragment = template.content.cloneNode(true);
      const guide = global.document.getElementById('initialQuickGuide');
      if (guide && guide.parentNode){
        guide.parentNode.insertBefore(fragment, guide.nextSibling);
      } else {
        const main = global.document.getElementById('mainContent');
        if (main) main.appendChild(fragment);
      }
      cacheUIElements();
      attachEventHandlers();
    }

    if (unsubscribe){
      unsubscribe();
      unsubscribe = null;
    }
    if (typeof repo.onUpdate === 'function'){
      unsubscribe = repo.onUpdate(() => {
        refresh().catch(err => console.error('processed orders refresh error', err));
      });
    }

    await refresh();
  }

  function clearSelection(){
    selection.clear();
    syncActionButtons();
    syncSelectAll(state.pageItems);
  }

  global.ProcessedOrdersPanel = {
    init,
    refresh,
    updateVisibility,
    clearSelection
  };
})(window);
