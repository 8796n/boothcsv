(function() {
  if (window.BoothCSVDocsCapture) return;

  const SHOTS_PATH = './scripts/docs-capture/shots.web.json';
  const TARGET_ATTR = 'data-docs-capture-target';
  const INCLUDE_ATTR = 'data-docs-capture-include';
  const state = {
    definitions: null,
    currentShot: null,
    initialized: false,
    captureUi: {
      headerHidden: false,
      fixedHeaderOffset: '',
      headerDisplay: '',
      previewBackDisplay: '',
      documentOverflow: '',
      bodyOverflow: ''
    },
    captureStage: null
  };

  function log(...args) {
    if (typeof window.debugLog === 'function') {
      window.debugLog('[docsCapture]', ...args);
      return;
    }
    console.debug('[docsCapture]', ...args);
  }

  function sleep(ms) {
    return new Promise(resolve => window.setTimeout(resolve, ms));
  }

  async function waitFor(predicate, options = {}) {
    const timeout = options.timeout || 10000;
    const interval = options.interval || 50;
    const start = Date.now();
    while (Date.now() - start < timeout) {
      try {
        const value = await predicate();
        if (value) return value;
      } catch (_) {
      }
      await sleep(interval);
    }
    throw new Error(options.errorMessage || 'Timed out while waiting for docs capture state');
  }

  function dispatchInputEvents(element) {
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function getFixedHeaderOffset() {
    const header = document.querySelector('.fixed-header');
    if (header) {
      const style = window.getComputedStyle(header);
      if (style.display !== 'none' && style.visibility !== 'hidden') {
        const rect = header.getBoundingClientRect();
        if (rect.height > 0) return Math.ceil(rect.height);
      }
    }
    const cssOffset = window.getComputedStyle(document.documentElement).getPropertyValue('--fixed-header-offset').trim();
    const parsed = Number.parseInt(cssOffset, 10);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function restoreCaptureUi() {
    cleanupCaptureStage();
    if (state.captureUi.headerHidden) {
      const header = document.querySelector('.fixed-header');
      const previewBackButton = document.querySelector('.preview-back-floating');
      if (header) {
        header.style.display = state.captureUi.headerDisplay || '';
      }
      if (previewBackButton) {
        previewBackButton.style.display = state.captureUi.previewBackDisplay || '';
      }
      if (state.captureUi.fixedHeaderOffset) {
        document.documentElement.style.setProperty('--fixed-header-offset', state.captureUi.fixedHeaderOffset);
      } else {
        document.documentElement.style.removeProperty('--fixed-header-offset');
      }
    }
    if (state.captureUi.documentOverflow) {
      document.documentElement.style.overflow = state.captureUi.documentOverflow;
    } else {
      document.documentElement.style.removeProperty('overflow');
    }
    if (state.captureUi.bodyOverflow) {
      document.body.style.overflow = state.captureUi.bodyOverflow;
    } else {
      document.body.style.removeProperty('overflow');
    }
    state.captureUi.headerHidden = false;
    state.captureUi.fixedHeaderOffset = '';
    state.captureUi.headerDisplay = '';
    state.captureUi.previewBackDisplay = '';
    state.captureUi.documentOverflow = '';
    state.captureUi.bodyOverflow = '';
  }

  function cleanupCaptureStage() {
    if (state.captureStage && state.captureStage.overlay && state.captureStage.overlay.parentNode) {
      state.captureStage.overlay.parentNode.removeChild(state.captureStage.overlay);
    }
    state.captureStage = null;
  }

  function applyCaptureUi(capture = {}) {
    restoreCaptureUi();
    if (!capture) return;
    const header = document.querySelector('.fixed-header');
    const previewBackButton = document.querySelector('.preview-back-floating');
    state.captureUi.fixedHeaderOffset = document.documentElement.style.getPropertyValue('--fixed-header-offset');
    state.captureUi.documentOverflow = document.documentElement.style.overflow;
    state.captureUi.bodyOverflow = document.body.style.overflow;
    if (capture.hideFixedHeader) {
      state.captureUi.headerHidden = true;
      state.captureUi.headerDisplay = header ? header.style.display : '';
      state.captureUi.previewBackDisplay = previewBackButton ? previewBackButton.style.display : '';
      if (header) {
        header.style.display = 'none';
      }
      if (previewBackButton) {
        previewBackButton.style.display = 'none';
      }
      document.documentElement.style.setProperty('--fixed-header-offset', '0px');
    }
    if (capture.mode === 'crop') {
      document.documentElement.style.overflow = 'hidden';
      document.body.style.overflow = 'hidden';
    }
  }

  function clearCaptureMarkers() {
    document.querySelectorAll('[' + TARGET_ATTR + '], [' + INCLUDE_ATTR + ']').forEach(element => {
      element.removeAttribute(TARGET_ATTR);
      element.removeAttribute(INCLUDE_ATTR);
    });
  }

  function getQueryShotId() {
    const url = new URL(window.location.href);
    return (url.searchParams.get('docsShot') || '').trim();
  }

  function buildShotUrl(shot) {
    const page = shot.page || {};
    const pagePath = page.path || 'boothcsv.html';
    const url = new URL(pagePath, window.location.href);
    const query = page.query || {};
    Object.entries(query).forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== '') {
        url.searchParams.set(key, String(value));
      }
    });
    url.searchParams.set('docsShot', shot.id);
    return url.toString();
  }

  async function loadDefinitions() {
    if (state.definitions) return state.definitions;
    const response = await fetch(SHOTS_PATH, { cache: 'no-store' });
    if (!response.ok) {
      throw new Error('Failed to load docs capture definitions: ' + response.status);
    }
    const definitions = await response.json();
    if (!Array.isArray(definitions)) {
      throw new Error('Docs capture definitions must be an array');
    }
    state.definitions = definitions;
    return definitions;
  }

  async function listShots() {
    const definitions = await loadDefinitions();
    return definitions.map(definition => ({
      id: definition.id,
      output: definition.output,
      viewport: definition.viewport || null,
      page: definition.page || null,
      capture: definition.capture || null,
      url: buildShotUrl(definition)
    }));
  }

  async function getShotDefinition(shotId) {
    const definitions = await loadDefinitions();
    const shot = definitions.find(definition => definition.id === shotId);
    if (!shot) {
      throw new Error('Unknown docs capture shot: ' + shotId);
    }
    return shot;
  }

  async function ensureAppReady() {
    await waitFor(() => typeof window.importCSVText === 'function', { errorMessage: 'importCSVText is not available' });
    await waitFor(() => window.CustomLabels && typeof window.CustomLabels.initialize === 'function', { errorMessage: 'CustomLabels is not ready' });
    await waitFor(() => document.getElementById('appSidebar'), { errorMessage: 'Sidebar is not ready' });
    await waitFor(() => window.orderImageDropZone && typeof window.orderImageDropZone.setImage === 'function', { errorMessage: 'Global image drop zone is not ready' });
    state.initialized = true;
  }

  async function setCheckbox(id, value) {
    const element = document.getElementById(id);
    if (!element) return;
    if (element.checked !== value) {
      element.checked = value;
      element.dispatchEvent(new Event('change', { bubbles: true }));
      await sleep(80);
    }
    if (window.settingsCache) {
      if (id === 'labelyn') window.settingsCache.labelyn = !!value;
      if (id === 'customLabelEnable') window.settingsCache.customLabelEnable = !!value;
      if (id === 'orderImageEnable') window.settingsCache.orderImageEnable = !!value;
      if (id === 'sortByPaymentDate') window.settingsCache.sortByPaymentDate = !!value;
    }
    if (id === 'customLabelEnable' && typeof window.toggleCustomLabelRow === 'function') {
      window.toggleCustomLabelRow(!!value);
    }
    if (id === 'orderImageEnable' && typeof window.toggleOrderImageRow === 'function') {
      window.toggleOrderImageRow(!!value);
    }
    if (id === 'labelyn' && typeof window.updateExtensionUIVisibility === 'function') {
      window.updateExtensionUIVisibility();
    }
  }

  async function setInputValue(id, value) {
    const element = document.getElementById(id);
    if (!element) return;
    const nextValue = String(value);
    if (element.value !== nextValue) {
      element.value = nextValue;
      dispatchInputEvents(element);
      await sleep(80);
    }
    if (window.settingsCache && id === 'labelskipnum') {
      window.settingsCache.labelskip = Number(value) || 0;
    }
  }

  async function ensureSidebar(mode) {
    const sidebar = document.getElementById('appSidebar');
    const toggle = document.getElementById('sidebarToggle');
    const pin = document.getElementById('sidebarPin');
    const close = document.getElementById('sidebarClose');
    if (!sidebar || !toggle || !pin || !close) return;

    const isOpen = () => sidebar.classList.contains('open');
    const isDocked = () => document.body.classList.contains('sidebar-docked');

    if (mode === 'docked') {
      if (!isOpen()) {
        toggle.click();
        await sleep(120);
      }
      if (!isDocked()) {
        pin.click();
        await sleep(160);
      }
      return;
    }

    if (mode === 'open') {
      if (!isOpen()) {
        toggle.click();
        await sleep(120);
      }
      if (isDocked()) {
        pin.click();
        await sleep(160);
      }
      return;
    }

    if (mode === 'closed') {
      if (isDocked()) {
        pin.click();
        await sleep(160);
      }
      if (isOpen()) {
        close.click();
        await sleep(120);
      }
    }
  }

  function clearFileInput() {
    const fileInput = document.getElementById('file');
    if (!fileInput) return;
    fileInput.value = '';
    fileInput.dispatchEvent(new Event('change', { bubbles: true }));
  }

  async function clearGlobalImageFixture() {
    if (window.StorageManager && typeof window.StorageManager.clearGlobalOrderImageBinary === 'function') {
      await window.StorageManager.clearGlobalOrderImageBinary();
    }
    const imageDropZone = window.orderImageDropZone?.element || document.querySelector('#imageDropZone .order-image-drop') || document.getElementById('imageDropZone');
    if (imageDropZone) {
      imageDropZone.innerHTML = '';
      const placeholder = document.createElement('p');
      placeholder.className = 'order-image-default-msg';
      placeholder.textContent = '注文明細の余白に表示したい画像をここにドロップ or クリックでファイルを選択';
      imageDropZone.appendChild(placeholder);
    }
  }

  async function resetTransientUi() {
    restoreCaptureUi();
    clearCaptureMarkers();
    document.querySelectorAll('.custom-label-context-menu').forEach(node => node.remove());
    if (window.getSelection) {
      const selection = window.getSelection();
      if (selection) selection.removeAllRanges();
    }
    document.activeElement?.blur?.();
    if (typeof window.clearPreviousResults === 'function') {
      window.clearPreviousResults();
    }
    window.currentDisplayedOrderNumbers = [];
    window.lastOrderSelection = [];
    window.lastPreviewConfig = null;
    if (typeof window.setPreviewSource === 'function') {
      window.setPreviewSource('');
    }
    const processedOrdersPanel = document.getElementById('processedOrdersPanel');
    if (processedOrdersPanel) {
      processedOrdersPanel.hidden = true;
    }
    const productOrderImagesBlock = document.getElementById('productOrderImagesBlock');
    const productOrderImagesContainer = document.getElementById('productOrderImagesContainer');
    const productOrderImagesSummary = document.getElementById('productOrderImagesSummary');
    if (productOrderImagesBlock) {
      productOrderImagesBlock.hidden = true;
    }
    if (productOrderImagesContainer) {
      productOrderImagesContainer.innerHTML = '';
    }
    if (productOrderImagesSummary) {
      productOrderImagesSummary.textContent = '商品ID画像: 0件';
    }
    if (typeof window.showQuickGuide === 'function') {
      window.showQuickGuide();
    }
    await clearGlobalImageFixture();
    await sleep(40);
  }

  async function applyUi(ui = {}) {
    if (ui.clearResults) {
      await resetTransientUi();
    }
    if (ui.clearFile) {
      clearFileInput();
      await sleep(120);
    }
    if (ui.sidebar) {
      await ensureSidebar(ui.sidebar);
    }
    if (ui.labelyn !== undefined) {
      await setCheckbox('labelyn', !!ui.labelyn);
    }
    if (ui.customLabel !== undefined) {
      await setCheckbox('customLabelEnable', !!ui.customLabel);
      if (!!ui.customLabel) {
        await waitFor(() => document.getElementById('customLabelsBlock') && !document.getElementById('customLabelsBlock').hidden, {
          timeout: 3000,
          errorMessage: 'Custom labels block did not become visible'
        });
      }
    }
    if (ui.orderImage !== undefined) {
      await setCheckbox('orderImageEnable', !!ui.orderImage);
      if (!!ui.orderImage) {
        await waitFor(() => document.getElementById('imageDropZone'), {
          timeout: 3000,
          errorMessage: 'Image drop zone is not available'
        });
      }
    }
    if (ui.sortByPaymentDate !== undefined) {
      await setCheckbox('sortByPaymentDate', !!ui.sortByPaymentDate);
    }
    const labelSkip = ui.labelSkip !== undefined ? ui.labelSkip : ui.labelskip;
    if (labelSkip !== undefined) {
      await setInputValue('labelskipnum', labelSkip);
    }
    await sleep(160);
  }

  async function fetchText(path) {
    const response = await fetch(path, { cache: 'no-store' });
    if (!response.ok) {
      throw new Error('Failed to fetch text fixture: ' + path);
    }
    return await response.text();
  }

  async function fetchBinaryFixture(path) {
    const response = await fetch(path, { cache: 'no-store' });
    if (!response.ok) {
      throw new Error('Failed to fetch binary fixture: ' + path);
    }
    const blob = await response.blob();
    return {
      blob,
      buffer: await blob.arrayBuffer(),
      objectUrl: URL.createObjectURL(blob)
    };
  }

  async function setCustomLabelsFixture(fixture) {
    const labels = Array.isArray(fixture.labels) ? fixture.labels : [];
    window.CustomLabels.initialize(labels);
    window.CustomLabels.save();
    await window.CustomLabels.updateSummary();
    if (typeof window.updateCustomLabelsPreview === 'function') {
      await window.updateCustomLabelsPreview();
    }
    await sleep(160);
  }

  async function resetStoredDataFixture(fixture = {}) {
    const clearOrders = fixture.orders !== false;
    const clearProductImages = fixture.productImages !== false;
    const clearGlobalImage = fixture.globalImage !== false;
    const clearCustomLabels = fixture.customLabels === true;

    if (window.StorageManager && typeof window.StorageManager.ensureDatabase === 'function') {
      const db = await window.StorageManager.ensureDatabase();
      if (clearOrders) {
        if (window.orderRepository?.cache instanceof Map) {
          window.orderRepository.cache.clear();
          if (typeof window.orderRepository.emit === 'function') {
            window.orderRepository.emit();
          }
        }
        if (db && typeof db.clearAllOrders === 'function') {
          await db.clearAllOrders();
        } else if (typeof window.StorageManager.clearAllOrders === 'function') {
          await window.StorageManager.clearAllOrders();
        }
      }
      if (clearProductImages) {
        if (db && typeof db.clearAllProductImages === 'function') {
          await db.clearAllProductImages();
        }
      }
      if (clearGlobalImage && typeof window.StorageManager.clearGlobalOrderImageBinary === 'function') {
        await window.StorageManager.clearGlobalOrderImageBinary();
      }
      if (clearCustomLabels && typeof window.StorageManager.clearAllCustomLabels === 'function') {
        await window.StorageManager.clearAllCustomLabels();
      }
    }

    await resetTransientUi();
    await sleep(fixture.delay || 120);
  }

  async function importCsvFixture(fixture) {
    const csvText = await fetchText(fixture.path);
    await window.importCSVText(csvText, { sourceName: fixture.sourceName || fixture.path.split('/').pop() || 'sample.csv' });
    await sleep(fixture.delay || 240);
  }

  async function setGlobalImageFixture(fixture) {
    const image = await fetchBinaryFixture(fixture.path);
    window.orderImageDropZone.setImage(image.objectUrl, image.buffer);
    await waitFor(() => document.querySelector('#imageDropZone img'), {
      timeout: 3000,
      errorMessage: 'Global image preview did not render'
    });
    await sleep(fixture.delay || 180);
  }

  function resolveFixtureOrderNumber(fixture = {}) {
    if (fixture.orderNumber) {
      return String(fixture.orderNumber);
    }

    const orderIndex = Number.isInteger(fixture.orderIndex) ? fixture.orderIndex : 0;
    if (Array.isArray(window.currentDisplayedOrderNumbers) && window.currentDisplayedOrderNumbers[orderIndex]) {
      return String(window.currentDisplayedOrderNumbers[orderIndex]);
    }
    if (Array.isArray(window.lastOrderSelection) && window.lastOrderSelection[orderIndex]) {
      return String(window.lastOrderSelection[orderIndex]);
    }

    const repoOrders = typeof window.orderRepository?.getAll === 'function'
      ? window.orderRepository.getAll().map(record => record && record.orderNumber).filter(Boolean)
      : [];
    if (repoOrders[orderIndex]) {
      return String(repoOrders[orderIndex]);
    }

    return null;
  }

  async function setOrderImageFixture(fixture) {
    const orderNumber = resolveFixtureOrderNumber(fixture);
    if (!orderNumber) {
      throw new Error('Target order number for order image fixture was not found');
    }
    const image = await fetchBinaryFixture(fixture.path);
    if (!window.orderRepository || typeof window.orderRepository.setOrderImage !== 'function') {
      throw new Error('orderRepository.setOrderImage is not available');
    }
    await window.orderRepository.setOrderImage(orderNumber, {
      data: image.buffer,
      mimeType: fixture.mimeType || image.blob.type || 'image/png'
    });
    await sleep(fixture.delay || 180);
  }

  async function injectQrToFirstLabelFixture(fixture) {
    const image = await fetchBinaryFixture(fixture.path);
    const targetOrderNumber = await waitFor(() => {
      if (Array.isArray(window.currentDisplayedOrderNumbers) && window.currentDisplayedOrderNumbers.length > 0) {
        return window.currentDisplayedOrderNumbers[0];
      }
      if (Array.isArray(window.lastOrderSelection) && window.lastOrderSelection.length > 0) {
        return window.lastOrderSelection[0];
      }
      return null;
    }, {
      timeout: 5000,
      errorMessage: 'Target order number for QR fixture was not found'
    });

    const img = document.createElement('img');
    img.src = image.objectUrl;
    img.style.maxWidth = '100%';
    img.style.height = 'auto';
    await new Promise(resolve => {
      img.onload = resolve;
      img.onerror = resolve;
    });

    const originalConfirm = window.confirm;
    window.confirm = function(message) {
      if (typeof message === 'string' && message.includes('このQRコードは既に以下の注文で使用されています')) {
        return true;
      }
      return typeof originalConfirm === 'function' ? originalConfirm.call(window, message) : true;
    };

    try {
      if (window.BoothCSVExtensionBridge && typeof window.BoothCSVExtensionBridge.importOrderQRCodeImage === 'function') {
        await window.BoothCSVExtensionBridge.importOrderQRCodeImage(String(targetOrderNumber), img, { allowDuplicate: true });
      } else if (typeof window.saveQRCodeForOrder === 'function') {
        await window.saveQRCodeForOrder(String(targetOrderNumber), img, { allowDuplicate: true });
      } else if (typeof window.readQR === 'function') {
        const dropZone = await waitFor(() => document.querySelector('.dropzone'), {
          timeout: 5000,
          errorMessage: 'QR drop zone was not found'
        });
        dropZone.innerHTML = '';
        dropZone.appendChild(img);
        await window.readQR(img);
      }
    } finally {
      window.confirm = originalConfirm;
    }

    await sleep(fixture.delay || 300);
  }

  async function applyFixture(fixture) {
    switch (fixture.type) {
      case 'reset-storage':
        await resetStoredDataFixture(fixture);
        return;
      case 'csv':
        await importCsvFixture(fixture);
        return;
      case 'custom-labels':
        await setCustomLabelsFixture(fixture);
        return;
      case 'global-image':
        await setGlobalImageFixture(fixture);
        return;
      case 'order-image':
        await setOrderImageFixture(fixture);
        return;
      case 'qr-to-first-label':
        await injectQrToFirstLabelFixture(fixture);
        return;
      default:
        throw new Error('Unsupported docs capture fixture: ' + fixture.type);
    }
  }

  function getCustomLabelEditor(index) {
    const items = Array.from(document.querySelectorAll('#customLabelsContainer .custom-label-item'));
    const item = items[index || 0];
    if (!item) return null;
    return item.querySelector('.rich-text-editor');
  }

  async function showCustomLabelContextMenu(action) {
    const editor = getCustomLabelEditor(action.itemIndex || 0);
    if (!editor) {
      throw new Error('Custom label editor was not found');
    }
    const target = action.selector ? editor.querySelector(action.selector) : editor;
    if (!target) {
      throw new Error('Custom label context menu target was not found');
    }
    const range = document.createRange();
    range.selectNodeContents(target);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    const rect = target.getBoundingClientRect();
    const event = new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      view: window,
      button: 2,
      clientX: rect.left + 8,
      clientY: rect.top + 8
    });
    target.dispatchEvent(event);
    await sleep(action.delay || 260);
  }

  async function closeCustomLabelContextMenu() {
    document.querySelectorAll('.custom-label-context-menu').forEach(node => node.remove());
    const selection = window.getSelection();
    if (selection) selection.removeAllRanges();
    await sleep(80);
  }

  async function expandFontSection() {
    const content = document.getElementById('fontSectionContent');
    const arrow = document.getElementById('fontSectionArrow');
    if (!content) return;
    content.hidden = false;
    content.style.overflow = 'visible';
    content.style.maxHeight = 'none';
    if (arrow) {
      arrow.style.transform = 'rotate(0deg)';
    }
    await sleep(120);
  }

  async function collapseFontSection() {
    const content = document.getElementById('fontSectionContent');
    const arrow = document.getElementById('fontSectionArrow');
    if (!content) return;
    content.hidden = true;
    content.style.overflow = '';
    content.style.maxHeight = '';
    if (arrow) {
      arrow.style.transform = 'rotate(-90deg)';
    }
    await sleep(120);
  }

  async function showProcessedOrdersPanelAction(action = {}) {
    clearFileInput();
    if (typeof window.hideQuickGuide === 'function') {
      window.hideQuickGuide();
    }
    if (typeof window.clearPreviousResults === 'function') {
      window.clearPreviousResults();
    }
    if (window.ProcessedOrdersPanel && typeof window.ProcessedOrdersPanel.clearSelection === 'function') {
      window.ProcessedOrdersPanel.clearSelection();
    }
    if (typeof window.refreshProcessedOrdersPanel === 'function') {
      await window.refreshProcessedOrdersPanel();
    }
    if (typeof window.updateProcessedOrdersVisibility === 'function') {
      window.updateProcessedOrdersVisibility();
    }
    if (typeof window.hideQuickGuide === 'function') {
      window.hideQuickGuide();
    }
    const panel = await waitFor(() => {
      const element = document.getElementById('processedOrdersPanel');
      if (!element || element.hidden) return null;
      const rowCount = element.querySelectorAll('#processedOrdersBody tr').length;
      return rowCount > 0 ? element : null;
    }, {
      timeout: 5000,
      errorMessage: 'Processed orders panel did not become visible'
    });
    if (action.selectAll) {
      const selectAll = document.getElementById('processedOrdersSelectAll');
      if (selectAll && !selectAll.checked) {
        selectAll.click();
        await sleep(120);
      }
    }
    panel.hidden = false;
    await sleep(action.delay || 180);
  }

  async function runAction(action) {
    switch (action.type) {
      case 'set-checkbox':
        await setCheckbox(action.id, !!action.value);
        return;
      case 'show-custom-label-context-menu':
        await showCustomLabelContextMenu(action);
        return;
      case 'close-custom-label-context-menu':
        await closeCustomLabelContextMenu();
        return;
      case 'expand-font-section':
        await expandFontSection();
        return;
      case 'collapse-font-section':
        await collapseFontSection();
        return;
      case 'show-processed-orders-panel':
        await showProcessedOrdersPanelAction(action);
        return;
      case 'sleep':
        await sleep(action.ms || 100);
        return;
      default:
        throw new Error('Unsupported docs capture action: ' + action.type);
    }
  }

  function resolveSnapTarget(target) {
    const anchors = Array.from(document.querySelectorAll('.snap-anchor[data-snap="' + target.snap + '"]'));
    if (!anchors.length) return null;
    const visibleAnchor = anchors.find(anchor => {
      const rect = anchor.getBoundingClientRect();
      return rect.top >= 0 && rect.top < window.innerHeight;
    });
    const anchor = visibleAnchor || anchors[0];
    if (!anchor) return null;
    if (target.selector) {
      const closest = anchor.closest(target.selector);
      if (closest) return closest;
    }
    return anchor.closest('section, .sidebar-section, .custom-font-section') || anchor.parentElement;
  }

  function resolveTargetElement(target = {}) {
    if (target.snap) {
      const bySnap = resolveSnapTarget(target);
      if (bySnap) return bySnap;
    }
    if (target.selector) {
      const bySelector = document.querySelector(target.selector);
      if (bySelector) return bySelector;
    }
    return null;
  }

  function markTargetElement(targetElement, shotId) {
    clearCaptureMarkers();
    if (!targetElement) return null;
    targetElement.setAttribute(TARGET_ATTR, shotId);
    return {
      selector: '[' + TARGET_ATTR + '="' + shotId + '"]',
      tagName: targetElement.tagName,
      id: targetElement.id || ''
    };
  }

  function sanitizeClone(root) {
    if (!root) return root;
    root.querySelectorAll('script').forEach(node => node.remove());
    root.querySelectorAll('[' + TARGET_ATTR + ']').forEach(node => node.removeAttribute(TARGET_ATTR));
    root.querySelectorAll('[' + INCLUDE_ATTR + ']').forEach(node => node.removeAttribute(INCLUDE_ATTR));
    return root;
  }

  function syncFormState(sourceRoot, cloneRoot) {
    if (!sourceRoot || !cloneRoot) return;
    const sourceFields = Array.from(sourceRoot.querySelectorAll('input, textarea, select'));
    const cloneFields = Array.from(cloneRoot.querySelectorAll('input, textarea, select'));
    sourceFields.forEach((sourceField, index) => {
      const cloneField = cloneFields[index];
      if (!cloneField) return;
      if (sourceField instanceof HTMLInputElement) {
        cloneField.checked = sourceField.checked;
        cloneField.value = sourceField.value;
      } else if (sourceField instanceof HTMLTextAreaElement || sourceField instanceof HTMLSelectElement) {
        cloneField.value = sourceField.value;
      }
    });
  }

  function getIncludeElements(capture = {}) {
    const includeSelectors = Array.isArray(capture.includeSelectors) ? capture.includeSelectors : [];
    const elements = [];
    includeSelectors.forEach(selector => {
      if (!selector) return;
      document.querySelectorAll(selector).forEach(node => {
        const rect = node.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          elements.push({ node, rect });
        }
      });
    });
    return elements;
  }

  async function waitForImagesInElement(root) {
    if (!root) return;
    const images = Array.from(root.querySelectorAll('img'));
    if (!images.length) return;
    await Promise.all(images.map(image => {
      if (image.complete && image.naturalWidth > 0) {
        return Promise.resolve();
      }
      return new Promise(resolve => {
        const finish = () => {
          image.removeEventListener('load', finish);
          image.removeEventListener('error', finish);
          resolve();
        };
        image.addEventListener('load', finish, { once: true });
        image.addEventListener('error', finish, { once: true });
      });
    }));
  }

  async function createCaptureStage(targetElement, shot) {
    cleanupCaptureStage();
    const capture = shot.capture || {};
    const padding = capture.padding !== undefined ? Number(capture.padding) || 0 : 10;
    const targetRect = targetElement.getBoundingClientRect();
    const includeElements = getIncludeElements(capture);
    let minLeft = targetRect.left;
    let minTop = targetRect.top;
    let maxRight = targetRect.right;
    let maxBottom = targetRect.bottom;

    includeElements.forEach(({ rect }) => {
      minLeft = Math.min(minLeft, rect.left);
      minTop = Math.min(minTop, rect.top);
      maxRight = Math.max(maxRight, rect.right);
      maxBottom = Math.max(maxBottom, rect.bottom);
    });

    const contentWidth = Math.ceil(maxRight - minLeft);
    const contentHeight = Math.ceil(maxBottom - minTop);

    const overlay = document.createElement('div');
    overlay.id = 'docsCaptureStageOverlay';
    overlay.setAttribute('aria-hidden', 'true');
    overlay.style.position = 'fixed';
    overlay.style.left = '0';
    overlay.style.top = '0';
    overlay.style.zIndex = '2147483647';
    overlay.style.pointerEvents = 'none';
    overlay.style.background = '#ffffff';
    overlay.style.width = Math.ceil(contentWidth + (padding * 2)) + 'px';
    overlay.style.height = Math.ceil(contentHeight + (padding * 2)) + 'px';
    overlay.style.overflow = 'hidden';

    const frame = document.createElement('div');
    frame.id = 'docsCaptureStageFrame';
    frame.style.position = 'absolute';
    frame.style.left = '0';
    frame.style.top = '0';
    frame.style.background = '#ffffff';
    frame.style.boxSizing = 'border-box';
    frame.style.padding = padding + 'px';
    frame.style.display = 'block';
    frame.style.overflow = 'hidden';
    frame.style.width = Math.ceil(contentWidth + (padding * 2)) + 'px';
    frame.style.minHeight = Math.ceil(contentHeight + (padding * 2)) + 'px';

    let contentRoot = frame;
    const sidebarHost = targetElement.closest('#appSidebar');
    if (sidebarHost) {
      const sidebarContext = document.createElement('div');
      sidebarContext.id = 'appSidebar';
      sidebarContext.className = sidebarHost.className;
      sidebarContext.style.position = 'relative';
      sidebarContext.style.left = '0';
      sidebarContext.style.top = '0';
      sidebarContext.style.width = Math.ceil(contentWidth) + 'px';
      sidebarContext.style.minHeight = '0';
      sidebarContext.style.height = 'auto';
      sidebarContext.style.maxHeight = 'none';
      sidebarContext.style.display = 'block';
      sidebarContext.style.overflow = 'visible';
      sidebarContext.style.transform = 'none';
      sidebarContext.style.right = 'auto';
      sidebarContext.style.bottom = 'auto';
      frame.appendChild(sidebarContext);
      contentRoot = sidebarContext;
    }

    const targetClone = sanitizeClone(targetElement.cloneNode(true));
    syncFormState(targetElement, targetClone);
    targetClone.style.position = 'absolute';
    targetClone.style.left = Math.ceil(targetRect.left - minLeft) + 'px';
    targetClone.style.top = Math.ceil(targetRect.top - minTop) + 'px';
    targetClone.style.boxSizing = 'border-box';
    targetClone.style.width = Math.ceil(targetRect.width) + 'px';
    targetClone.style.maxWidth = 'none';
    contentRoot.appendChild(targetClone);

    includeElements.forEach(({ node, rect }) => {
      const clone = sanitizeClone(node.cloneNode(true));
      syncFormState(node, clone);
      clone.style.position = 'absolute';
      clone.style.left = Math.ceil(rect.left - minLeft) + 'px';
      clone.style.top = Math.ceil(rect.top - minTop) + 'px';
      clone.style.width = Math.ceil(rect.width) + 'px';
      clone.style.maxWidth = 'none';
      contentRoot.appendChild(clone);
    });

    overlay.appendChild(frame);
    document.body.appendChild(overlay);
    await waitForImagesInElement(frame);
    state.captureStage = { overlay, frame, padding, shotId: shot.id };
    return frame;
  }

  async function scrollTargetIntoView(targetElement, target = {}, capture = {}) {
    if (!targetElement) return;
    const scrollMode = target.scroll || 'center';
    const padding = target.padding !== undefined ? Number(target.padding) || 0 : 8;

    if (scrollMode === 'start') {
      const rect = targetElement.getBoundingClientRect();
      const headerOffset = capture.hideFixedHeader ? 0 : getFixedHeaderOffset();
      const top = Math.max(0, Math.round(window.scrollY + rect.top - headerOffset - padding));
      window.scrollTo({ top, behavior: 'auto' });
      await sleep(180);
      return;
    }

    targetElement.scrollIntoView({ behavior: 'auto', block: scrollMode, inline: 'nearest' });
    await sleep(180);

    if (scrollMode === 'center' && !capture.hideFixedHeader) {
      const rect = targetElement.getBoundingClientRect();
      const headerOffset = getFixedHeaderOffset();
      if (rect.top < headerOffset) {
        const top = Math.max(0, Math.round(window.scrollY + rect.top - headerOffset - padding));
        window.scrollTo({ top, behavior: 'auto' });
        await sleep(180);
      }
    }
  }

  async function finalizeTarget(shot) {
    const targetElement = resolveTargetElement(shot.target || {});
    if (!targetElement) {
      return null;
    }

    await scrollTargetIntoView(targetElement, shot.target || {}, shot.capture || {});

    let effectiveTarget = targetElement;
    if ((shot.capture || {}).mode === 'crop') {
      effectiveTarget = await createCaptureStage(targetElement, shot);
    }

    const marker = markTargetElement(effectiveTarget, shot.id);
    return {
      marker,
      target: shot.target || null,
      rect: effectiveTarget.getBoundingClientRect().toJSON ? effectiveTarget.getBoundingClientRect().toJSON() : null
    };
  }

  async function runShot(shotId) {
    await ensureAppReady();
    const shot = await getShotDefinition(shotId);
    state.currentShot = {
      id: shot.id,
      status: 'preparing',
      shot,
      startedAt: new Date().toISOString()
    };
    document.body.dataset.docsShot = shot.id;

    await resetTransientUi();
    await applyUi(shot.ui || {});

    for (const fixture of shot.fixtures || []) {
      await applyFixture(fixture);
    }

    for (const action of shot.actions || []) {
      await runAction(action);
    }

    applyCaptureUi(shot.capture || {});
    const finalizedTarget = await finalizeTarget(shot);
    await sleep(120);

    state.currentShot = {
      id: shot.id,
      status: 'ready',
      shot,
      viewport: shot.viewport || null,
      output: shot.output || '',
      capture: shot.capture || null,
      target: finalizedTarget,
      url: buildShotUrl(shot),
      completedAt: new Date().toISOString()
    };

    window.dispatchEvent(new CustomEvent('boothcsv:docs-shot-ready', { detail: state.currentShot }));
    return state.currentShot;
  }

  function getCurrentShot() {
    return state.currentShot;
  }

  async function initializeFromQuery() {
    const shotId = getQueryShotId();
    if (!shotId) return null;
    try {
      log('Preparing docs shot from query:', shotId);
      return await runShot(shotId);
    } catch (error) {
      state.currentShot = {
        id: shotId,
        status: 'error',
        error: error.message,
        failedAt: new Date().toISOString()
      };
      console.error('[docsCapture] failed to prepare shot', error);
      window.dispatchEvent(new CustomEvent('boothcsv:docs-shot-error', { detail: state.currentShot }));
      throw error;
    }
  }

  window.BoothCSVDocsCapture = {
    loadDefinitions,
    listShots,
    getShotDefinition,
    buildShotUrl,
    runShot,
    getCurrentShot,
    initializeFromQuery
  };

  async function bootstrap() {
    if (!getQueryShotId()) return;
    await ensureAppReady();
    await loadDefinitions();
    await initializeFromQuery();
  }

  if (document.readyState === 'complete') {
    bootstrap().catch(error => console.error('[docsCapture] bootstrap error', error));
  } else {
    window.addEventListener('load', function() {
      bootstrap().catch(error => console.error('[docsCapture] bootstrap error', error));
    }, { once: true });
  }

})();
