(function(global){
  'use strict';

  const state = {
    sortKey: 'paymentDate',
    sortDirection: 'desc',
    currentPage: 1,
    pageSize: 20,
    unprintedOnly: false,
    unshippedOnly: false,
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
    shipButton: null,
    prev: null,
    next: null,
    pageInfo: null,
    empty: null,
    filter: null,
    shippedFilter: null,
    headerCells: []
  };

  let unsubscribe = null;
  const dependencies = {
    ensureOrderRepository: null,
    clearPreviousResults: null,
    hideQuickGuide: null,
    renderPreview: null,
    setPreviewSource: null,
    notifyShipment: null
  };

  function assignDependencies(options){
    dependencies.ensureOrderRepository = options.ensureOrderRepository || dependencies.ensureOrderRepository;
    dependencies.clearPreviousResults = options.clearPreviousResults || dependencies.clearPreviousResults;
    dependencies.hideQuickGuide = options.hideQuickGuide || dependencies.hideQuickGuide;
    dependencies.renderPreview = options.renderPreview || dependencies.renderPreview;
    dependencies.setPreviewSource = options.setPreviewSource || dependencies.setPreviewSource;
    dependencies.notifyShipment = options.notifyShipment || dependencies.notifyShipment;
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

  function createSvgIcon(name){
    const namespace = 'http://www.w3.org/2000/svg';
    const svg = global.document.createElementNS(namespace, 'svg');
    svg.setAttribute('viewBox', '0 0 16 16');
    svg.setAttribute('aria-hidden', 'true');
    svg.classList.add('processed-orders-icon-svg');

    if (name === 'booth'){
      const path = global.document.createElementNS(namespace, 'path');
      path.setAttribute('d', 'M3 3h4v2H5v6h6V9h2v4H3V3zm6 0h4v4h-2V6.41L8.7 8.7 7.3 7.3 9.59 5H9V3z');
      path.setAttribute('fill', 'currentColor');
      svg.appendChild(path);
      return svg;
    }

    if (name === 'qr'){
      const path = global.document.createElementNS(namespace, 'path');
      path.setAttribute('d', 'M2 2h5v5H2V2zm1.5 1.5v2h2v-2h-2zM9 2h5v5H9V2zm1.5 1.5v2h2v-2h-2zM2 9h5v5H2V9zm1.5 1.5v2h2v-2h-2zM9 9h1.5v1.5H9V9zm1.5 1.5H12V12h-1.5v-1.5zM12 9h2v2h-2V9zm-3 3h2v2H9v-2zm3 1h2v1h-2v-1z');
      path.setAttribute('fill', 'currentColor');
      svg.appendChild(path);
      return svg;
    }

    const path = global.document.createElementNS(namespace, 'path');
    path.setAttribute('d', 'M2 3h12v10H2V3zm1.2 1.2v7.6h9.6V4.2H3.2zm1.1 6.3 2.2-2.4 1.6 1.8 1.7-2.1 2.4 2.7H4.3zm1-4.3a1.2 1.2 0 110 2.4 1.2 1.2 0 010-2.4z');
    path.setAttribute('fill', 'currentColor');
    svg.appendChild(path);
    return svg;
  }

  function createStatusIcon(name, label){
    const icon = global.document.createElement('span');
    icon.className = `processed-orders-icon processed-orders-icon-${name}`;
    icon.title = label;
    icon.setAttribute('aria-label', label);
    icon.appendChild(createSvgIcon(name));
    return icon;
  }

  function createBoothLink(orderNumber){
    const link = global.document.createElement('a');
    link.href = `https://manage.booth.pm/orders/${encodeURIComponent(orderNumber)}`;
    link.target = '_blank';
    link.rel = 'noopener';
    link.className = 'processed-orders-icon-button processed-orders-icon-button-booth';
    link.title = 'BOOTHの注文詳細を開く';
    link.setAttribute('aria-label', `注文 ${orderNumber} を BOOTH で開く`);
    link.appendChild(createSvgIcon('booth'));
    return link;
  }

  async function openOrderPreview(orderNumber){
    if (!orderNumber) return;
    if (typeof dependencies.clearPreviousResults === 'function'){
      dependencies.clearPreviousResults();
    }
    if (typeof dependencies.hideQuickGuide === 'function'){
      dependencies.hideQuickGuide();
    }
    if (typeof dependencies.renderPreview !== 'function') return;
    const config = buildPreviewConfig();
    try {
      await dependencies.renderPreview([orderNumber], config);
      if (typeof dependencies.setPreviewSource === 'function') {
        dependencies.setPreviewSource('processed-orders');
      }
    } catch (e){
      console.error('single order preview error', e);
      global.alert('プレビューの生成に失敗しました: ' + (e && e.message ? e.message : e));
    }
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
    if (ui.shipButton){
      ui.shipButton.disabled = count === 0;
      ui.shipButton.textContent = count > 0 ? `選択した注文を発送通知 (${count})` : '選択した注文を発送通知';
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
    } else if (state.sortKey === 'shippedAt'){
      const av = Number.isNaN(a.shippedAtValue) ? -Infinity : a.shippedAtValue;
      const bv = Number.isNaN(b.shippedAtValue) ? -Infinity : b.shippedAtValue;
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
      const shippedRaw = rec && rec.shippedAt ? rec.shippedAt : '';
      return {
        orderNumber: rec.orderNumber,
        paymentDateRaw: paymentRaw,
        paymentDateValue: parseDateToTimestamp(paymentRaw),
        printedAtRaw: printedRaw,
        printedAtValue: parseDateToTimestamp(printedRaw),
        shippedAtRaw: shippedRaw,
        shippedAtValue: parseDateToTimestamp(shippedRaw),
        hasQr: !!(rec && rec.qr),
        hasImage: !!(rec && rec.image),
        printed: !!printedRaw,
        shipped: !!shippedRaw
      };
    }).filter(item => !!item.orderNumber);

    const sorted = [...data].sort(compareEntries);
    const filtered = sorted.filter(item => {
      if (state.unprintedOnly && item.printed) return false;
      if (state.unshippedOnly && item.shipped) return false;
      return true;
    });

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
        const orderCell = global.document.createElement('div');
        orderCell.className = 'processed-orders-order-cell';

        const previewButton = global.document.createElement('button');
        previewButton.type = 'button';
        previewButton.className = 'order-link order-preview-link';
        previewButton.textContent = item.orderNumber;
        previewButton.title = `注文 ${item.orderNumber} の印刷プレビューを表示`;
        previewButton.setAttribute('aria-label', `注文 ${item.orderNumber} の印刷プレビューを表示`);
        previewButton.addEventListener('click', () => {
          openOrderPreview(item.orderNumber).catch(err => console.error('single order preview error', err));
        });
        orderCell.appendChild(previewButton);

        const actions = global.document.createElement('span');
        actions.className = 'processed-orders-order-actions';
        actions.appendChild(createBoothLink(item.orderNumber));
        if (item.hasQr) {
          actions.appendChild(createStatusIcon('qr', 'QRコードあり'));
        }
        if (item.hasImage) {
          actions.appendChild(createStatusIcon('image', '個別画像あり'));
        }
        orderCell.appendChild(actions);

        orderTd.appendChild(orderCell);
        tr.appendChild(orderTd);

        const paymentTd = global.document.createElement('td');
        paymentTd.textContent = item.paymentDateRaw || '';
        tr.appendChild(paymentTd);

        const printedTd = global.document.createElement('td');
        printedTd.textContent = item.printedAtRaw ? formatDateTime(item.printedAtRaw) : '';
        tr.appendChild(printedTd);

        const shippedTd = global.document.createElement('td');
        shippedTd.textContent = item.shippedAtRaw ? formatDateTime(item.shippedAtRaw) : '';
        tr.appendChild(shippedTd);

        ui.body.appendChild(tr);
      });
    }

    if (ui.empty){
      if (filtered.length === 0){
        ui.empty.hidden = false;
        if (state.unprintedOnly && state.unshippedOnly) {
          ui.empty.textContent = '未印刷かつ未発送の注文はありません。';
        } else if (state.unprintedOnly) {
          ui.empty.textContent = '未印刷の注文はありません。';
        } else if (state.unshippedOnly) {
          ui.empty.textContent = '未発送の注文はありません。';
        } else {
          ui.empty.textContent = '保存されている注文はありません。';
        }
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

  function handleShippedFilterChange(event){
    state.unshippedOnly = !!event.target.checked;
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
      if (typeof dependencies.setPreviewSource === 'function') {
        dependencies.setPreviewSource('processed-orders');
      }
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

  async function handleShipNotification(){
    if (selection.size === 0) return;
    if (typeof dependencies.notifyShipment !== 'function') return;
    try {
      await dependencies.notifyShipment(Array.from(selection));
    } catch (e){
      console.error('shipment notification error', e);
      global.alert('発送通知処理に失敗しました: ' + (e && e.message ? e.message : e));
    }
  }

  function updateVisibility(){
    if (!ui.panel) return;
    const fileInput = global.document.getElementById('file');
    const hasFile = !!(fileInput && fileInput.files && fileInput.files.length > 0);
    const allSheets = global.document.querySelectorAll('section.sheet');
    const labelSheets = global.document.querySelectorAll('section.sheet.label-sheet');
    const hasOrderSheets = allSheets.length > labelSheets.length;
    const hasCurrentPreview = Array.isArray(global.currentDisplayedOrderNumbers) && global.currentDisplayedOrderNumbers.length > 0;
    ui.panel.hidden = hasFile || hasOrderSheets || hasCurrentPreview;
  }

  function attachEventHandlers(){
    if (ui.selectAll){
      ui.selectAll.addEventListener('change', handleSelectAll);
    }
    if (ui.filter){
      ui.filter.addEventListener('change', handleFilterChange);
    }
    if (ui.shippedFilter){
      ui.shippedFilter.addEventListener('change', handleShippedFilterChange);
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
    if (ui.shipButton){
      ui.shipButton.addEventListener('click', handleShipNotification);
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
    ui.shipButton = global.document.getElementById('processedOrdersShipButton');
    ui.prev = global.document.getElementById('processedOrdersPrev');
    ui.next = global.document.getElementById('processedOrdersNext');
    ui.pageInfo = global.document.getElementById('processedOrdersPageInfo');
    ui.empty = global.document.getElementById('processedOrdersEmpty');
    ui.filter = global.document.getElementById('processedOrdersUnprintedOnly');
    ui.shippedFilter = global.document.getElementById('processedOrdersUnshippedOnly');
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
