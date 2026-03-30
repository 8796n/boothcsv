// アプリケーション定数
const CONSTANTS = {
  LABEL: {
    TOTAL_LABELS_PER_SHEET: 44,
    LABELS_PER_ROW: 4
  },
  IMAGE: {
    SUPPORTED_FORMATS: /^image\/(jpeg|png|svg\+xml)$/,
    ACCEPTED_TYPES: 'image/jpeg, image/png, image/svg+xml'
  },
  FONT: {
    SUPPORTED_FORMATS: /\.(ttf|otf|woff|woff2)$/i,
    ACCEPTED_TYPES: '.ttf,.otf,.woff,.woff2'
  },
  QR: {
    EXPECTED_PARTS: 3
  },
  CSV: {
    PRODUCT_COLUMN: "商品ID / 数量 / 商品名",
    ORDER_NUMBER_COLUMN: "注文番号",
    PAYMENT_DATE_COLUMN: "支払い日時"
  },  
};
// 他モジュールからも参照できるように露出
window.CONSTANTS = CONSTANTS;

// デバッグフラグの取得
const DEBUG_MODE = (() => {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get('debug') === '1';
})();

function getForcedAppMode() {
  const urlParams = new URLSearchParams(window.location.search);
  const appMode = (urlParams.get('appMode') || '').trim().toLowerCase();
  return ['extension', 'web', 'file'].includes(appMode) ? appMode : '';
}

function isRealChromeExtensionApp() {
  return window.location.protocol === 'chrome-extension:' && !!(window.chrome && chrome.runtime && chrome.runtime.id);
}

function detectActualAppMode() {
  if (isRealChromeExtensionApp()) return 'extension';
  if (window.location.protocol === 'file:') return 'file';
  return 'web';
}

function resolveAppMode() {
  return getForcedAppMode() || detectActualAppMode();
}

function getAppMode() {
  return resolveAppMode();
}

function canUseExtensionBridge() {
  return isRealChromeExtensionApp();
}

// プロトコル/能力判定ユーティリティ
const PROTOCOL_UTILS = {
  isFileProtocol: () => detectActualAppMode() === 'file',
  isHttpProtocol: () => detectActualAppMode() === 'web',
  canDragFromExternalSites: () => detectActualAppMode() !== 'file'
};

window.BoothCSVAppRuntime = Object.assign(window.BoothCSVAppRuntime || {}, {
  getForcedAppMode,
  detectActualAppMode,
  getAppMode: resolveAppMode,
  canUseExtensionBridge,
  isRealChromeExtensionApp
});

// 初回クイックガイド表示制御（グローバルからどこでも呼べるユーティリティ）
function hideQuickGuide(){ const el = document.getElementById('initialQuickGuide'); if(el) el.hidden = true; }
function showQuickGuide(){ const el = document.getElementById('initialQuickGuide'); if(el) el.hidden = false; }

// リアルタイム更新制御フラグ
// isEditingCustomLabel は custom-labels.js で window プロパティとして定義される
// 直近描画に使用した設定・注文番号を保持し、再プレビューへ再利用
window.lastPreviewConfig = null; // { labelyn, labelskip, sortByPaymentDate, customLabelEnable, customLabels }
window.lastOrderSelection = []; // [orderNumber, ...]
window.currentPreviewSource = ''; // '', 'processed-orders'
let pendingUpdateTimer = null;

// デバッグログ用関数
const DEBUG_FLAGS = {
  csv: true,          // CSV 読み込み関連
  repo: true,         // OrderRepository 連携
  label: false,       // ラベル生成詳細（大量になるので初期OFF）
  font: false,        // フォント読み込み
  customLabel: false, // カスタムラベルUI
  image: false,       // 画像/ドロップゾーン
  // v6 debug: 一時的に true で画像保存経路を追跡 (問題解決後 false に戻しても良い)
  image: true,
  general: true       // 一般的な進行ログ
};

// デバッグログ出力（カテゴリ別フィルタリング対応）
function debugLog(catOrMsg, ...rest){
  if(!DEBUG_MODE) return;
  let cat = 'general';
  let msgArgs;
  if(typeof catOrMsg === 'string' && catOrMsg.startsWith('[')){
    // 形式: [cat] メッセージ
    const m = catOrMsg.match(/^\[([^\]]+)\]\s?(.*)$/);
    if(m){
      cat = m[1];
      const tail = m[2];
      msgArgs = tail ? [tail, ...rest] : rest;
    } else {
      msgArgs = [catOrMsg, ...rest];
    }
  } else if(typeof catOrMsg === 'string' && DEBUG_FLAGS[catOrMsg] !== undefined){
    cat = catOrMsg; msgArgs = rest;
  } else {
    msgArgs = [catOrMsg, ...rest];
  }
  if(!DEBUG_FLAGS[cat]) return;
  console.log(`[${cat}]`, ...msgArgs);
}

// HTMLエスケープ（フォントメタ表示用）
function escapeHTML(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// CSVの行データから注文番号を取得
function getOrderNumberFromCSVRow(row){
  if (!row || !row[CONSTANTS.CSV.ORDER_NUMBER_COLUMN]) return '';
  return String(row[CONSTANTS.CSV.ORDER_NUMBER_COLUMN]).trim();
}
// 表示用の『注文番号 : 』プレフィクスは CSS の .注文番号::before で付与するため
// ここでのフォーマット関数は不要になった。
// 注文番号の有効性をチェック
function isValidOrderNumber(orderNumber){
  return !!(orderNumber && String(orderNumber).trim());
}

function updateSelectedSourceInfo(label, hasData = false) {
  const fileSelectedInfoCompact = document.getElementById('fileSelectedInfoCompact');
  if (!fileSelectedInfoCompact) return;
  fileSelectedInfoCompact.textContent = label || '未選択';
  fileSelectedInfoCompact.classList.toggle('has-file', !!hasData);
}

function getRenderedSheetState() {
  const allSheets = document.querySelectorAll('section.sheet');
  const labelSheets = document.querySelectorAll('section.sheet.label-sheet');
  const hasAnySheets = allSheets.length > 0;
  const hasLabelSheets = labelSheets.length > 0;
  const hasOrderSheets = allSheets.length > labelSheets.length;
  return {
    hasAnySheets,
    hasLabelSheets,
    hasOrderSheets
  };
}

function hasEnabledCustomLabelContent(config = null) {
  const source = config && typeof config === 'object' ? config : buildCurrentPreviewConfig(null);
  if (!source.labelyn || !source.customLabelEnable) return false;
  return Array.isArray(source.customLabels)
    && source.customLabels.some(label => label && label.enabled !== false && (parseInt(label.count, 10) || 0) > 0 && String(label.text || '').trim() !== '');
}

function updatePreviewReturnButtonVisibility() {
  const button = document.getElementById('previewBackButton');
  if (!button) return;
  const hasCurrentPreview = Array.isArray(window.currentDisplayedOrderNumbers)
    && window.currentDisplayedOrderNumbers.length > 0;
  const { hasOrderSheets } = getRenderedSheetState();
  button.hidden = !(hasCurrentPreview || hasOrderSheets);
}

function setPreviewSource(source) {
  window.currentPreviewSource = source || '';
  updatePreviewReturnButtonVisibility();
}

function updateExtensionUIVisibility() {
  const appMode = getAppMode();
  const isExtensionMode = appMode === 'extension';
  document.body.dataset.appMode = appMode;
  document.body.classList.toggle('is-extension-app', isExtensionMode);
  document.body.classList.toggle('is-web-app', appMode === 'web');
  document.body.classList.toggle('is-file-app', appMode === 'file');
  const labelEnabled = !!document.getElementById('labelyn')?.checked;
  const csvDownloadWrapper = document.getElementById('boothOrdersOpenWrapper');
  const csvFileInputWrapper = document.getElementById('csvFileInputWrapper');
  const extensionHeaderCsvControls = document.getElementById('extensionHeaderCsvControls');
  const extensionLabelPrintTools = document.getElementById('extensionLabelPrintTools');
  const yamatoSection = document.getElementById('yamatoIssueSettingsSection');

  if (csvDownloadWrapper) csvDownloadWrapper.hidden = isExtensionMode;
  if (csvFileInputWrapper) csvFileInputWrapper.hidden = isExtensionMode;
  if (extensionHeaderCsvControls) extensionHeaderCsvControls.hidden = !isExtensionMode;
  if (extensionLabelPrintTools) extensionLabelPrintTools.hidden = !(isExtensionMode && labelEnabled);
  if (yamatoSection) yamatoSection.hidden = !(isExtensionMode && labelEnabled);
}

function initializeFixedHeaderOffset() {
  const header = document.querySelector('.fixed-header');
  if (!header) return;

  const updateOffset = function() {
    const rect = header.getBoundingClientRect();
    const nextOffset = Math.max(0, Math.ceil(rect.height));
    document.documentElement.style.setProperty('--fixed-header-offset', `${nextOffset}px`);
  };

  updateOffset();
  window.addEventListener('resize', updateOffset, { passive: true });

  if (typeof ResizeObserver === 'function') {
    const observer = new ResizeObserver(updateOffset);
    observer.observe(header);
  }
}

function buildCurrentPreviewConfig(file = null) {
  return {
    file,
    labelyn: settingsCache.labelyn,
    labelskip: settingsCache.labelskip,
    sortByPaymentDate: settingsCache.sortByPaymentDate,
    customLabelEnable: settingsCache.customLabelEnable,
    customLabels: settingsCache.customLabelEnable ? CustomLabels.getFromUI().filter(label => label.enabled) : []
  };
}

// CSV解析ユーティリティ
class CSVAnalyzer {
  // CSVファイルの行数を取得（非同期）
  static async getRowCount(file) {
    return new Promise((resolve, reject) => {
      if (!file) {
        resolve(0);
        return;
      }
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: function (results) {
          resolve(results.data.length);
        },
        error: function (error) {
          console.error('CSV解析エラー:', error);
          reject(error);
        }
      });
    });
  }
  
  // CSVファイルの基本情報を取得（非同期）
  static async getFileInfo(file) {
    if (!file) {
      return { rowCount: 0, fileName: '', fileSize: 0 };
    }
    try {
      const rowCount = await this.getRowCount(file);
      return { rowCount, fileName: file.name, fileSize: file.size };
    } catch (error) {
      console.error('CSVファイル情報取得エラー:', error);
      return { rowCount: 0, fileName: file.name, fileSize: file.size };
    }
  }
}

// （重複回避）バイナリ変換ユーティリティは storage.js に集約しました

// IndexedDBを使用した統合ストレージ管理クラス（破壊的移行版）
// （重複回避）UnifiedDatabase は storage.js のものを使用します

// （重複回避）unifiedDB 初期化とクリーンアップは storage.js に集約しました

// フォント管理は custom-labels-font.js (window.CustomLabelFont) に移動
// 後方互換のため window.fontManager 参照が残っている場合は初期化後に同期される

// 統合ストレージ管理クラス（UnifiedDatabaseのラッパー）
// （重複回避）StorageManager は storage.js のものを使用します

// 初期化処理（破壊的移行対応）
// グローバル設定キャッシュ（UIとIndexedDBの間の単純なメモリ反映）
window.settingsCache = {
  labelyn: true,
  labelskip: 0,
  sortByPaymentDate: false,
  customLabelEnable: false,
  orderImageEnable: false,
  shippingMessageTemplate: '',
  yamatoPackageSizeId: '16',
  yamatoCodeType: '0',
  yamatoDescription: '',
  yamatoDescriptionHistory: [],
  yamatoIncludeOrderNumber: true,
  yamatoHandlingCodes: []
};

const YAMATO_DESCRIPTION_HISTORY_LIMIT = 10;

const YAMATO_HANDLING_CHECKBOXES = [
  'yamatoHandlingPrecisionEquipment',
  'yamatoHandlingFragile',
  'yamatoHandlingDoNotStack',
  'yamatoHandlingDoNotTurnOver',
  'yamatoHandlingRawFood'
];

window.addEventListener("load", async function(){
  // プロトコル判定とファイルプロトコル警告の表示
  if (PROTOCOL_UTILS.isFileProtocol()) {
    const warningElement = document.getElementById('fileProtocolWarning');
    if (warningElement) {
      warningElement.hidden = false;
      debugLog('ファイルプロトコル警告を表示');
    }
  }

  // IndexedDB 初期化（失敗時は利用不可とし終了）
  await StorageManager.ensureDatabase();
  if (!window.unifiedDB) {
    console.error('IndexedDB 未利用のためアプリを継続できません');
    alert('この環境では IndexedDB が利用できないためツールを使用できません。ブラウザ設定を確認してください。');
    return; // 以降の機能を停止
  }

  const settings = await StorageManager.getSettingsAsync();

  const elLabel = document.getElementById("labelyn"); if (elLabel) elLabel.checked = settings.labelyn;
  const elSkip = document.getElementById("labelskipnum"); if (elSkip) elSkip.value = settings.labelskip;
  const elSort = document.getElementById("sortByPaymentDate"); if (elSort) elSort.checked = settings.sortByPaymentDate;
  const elCustom = document.getElementById("customLabelEnable"); if (elCustom) elCustom.checked = settings.customLabelEnable;
  const elImg = document.getElementById("orderImageEnable"); if (elImg) elImg.checked = settings.orderImageEnable;
  const elShippingTemplate = document.getElementById("shippingMessageTemplate");
  if (elShippingTemplate) elShippingTemplate.value = settings.shippingMessageTemplate || '';
  const elYamatoSize = document.getElementById('yamatoPackageSizeId'); if (elYamatoSize) elYamatoSize.value = settings.yamatoPackageSizeId || '16';
  const elYamatoCodeType = document.getElementById('yamatoCodeType'); if (elYamatoCodeType) elYamatoCodeType.value = settings.yamatoCodeType || '0';
  const elYamatoDescription = document.getElementById('yamatoDescription'); if (elYamatoDescription) elYamatoDescription.value = settings.yamatoDescription || '';
  renderYamatoDescriptionHistory(settings.yamatoDescriptionHistory || []);
  const elYamatoIncludeOrderNumber = document.getElementById('yamatoIncludeOrderNumber'); if (elYamatoIncludeOrderNumber) elYamatoIncludeOrderNumber.checked = settings.yamatoIncludeOrderNumber !== false;
  YAMATO_HANDLING_CHECKBOXES.forEach(function(id) {
    const input = document.getElementById(id);
    if (!input) return;
    const code = input.dataset.handlingCode;
    input.checked = Array.isArray(settings.yamatoHandlingCodes) && settings.yamatoHandlingCodes.includes(code);
  });

  Object.assign(window.settingsCache, {
    labelyn: settings.labelyn,
    labelskip: settings.labelskip,
    sortByPaymentDate: settings.sortByPaymentDate,
    customLabelEnable: settings.customLabelEnable,
    orderImageEnable: settings.orderImageEnable,
    shippingMessageTemplate: settings.shippingMessageTemplate || '',
    yamatoPackageSizeId: settings.yamatoPackageSizeId || '16',
    yamatoCodeType: settings.yamatoCodeType || '0',
    yamatoDescription: settings.yamatoDescription || '',
    yamatoDescriptionHistory: normalizeYamatoDescriptionHistory(settings.yamatoDescriptionHistory),
    yamatoIncludeOrderNumber: settings.yamatoIncludeOrderNumber !== false,
    yamatoHandlingCodes: Array.isArray(settings.yamatoHandlingCodes) ? settings.yamatoHandlingCodes.slice() : []
  });

  initializeShippingMessageTemplateUI(settings.shippingMessageTemplate || '');
  initializeYamatoIssueSettingsUI(settings);

  await setupProcessedOrdersPanel();
  updateProcessedOrdersVisibility();
  updateExtensionUIVisibility();

  toggleCustomLabelRow(settings.customLabelEnable);
  toggleOrderImageRow(settings.orderImageEnable);
  console.log('🎉 アプリケーション初期化完了 (fallback 無し)');

  // 初回クイックガイド制御: 以下のいずれかなら非表示
  // 1) 既に CSV データがある 2) 起動時点でラベルシート(section.sheet)が存在 3) カスタムラベル設定が残っていて生成済み
  // まだ DOM 生成前の場合を考慮し、初回判定と遅延再判定を実施
  const hasExistingSelection = Array.isArray(window.lastOrderSelection) && window.lastOrderSelection.length > 0;
  const hasSheetsNow = !!document.querySelector('section.sheet');
  // カスタムラベル設定有無だけでは非表示にしない（ユーザ要望）
  if (hasExistingSelection || hasSheetsNow) {
    hideQuickGuide();
  } else {
    showQuickGuide();
  }

  initializeFixedHeaderOffset();
  // カスタムラベル初期化 / シート生成後の再チェック（0ms + 300ms 両方）
  setTimeout(() => {
    if (document.querySelector('section.sheet')) hideQuickGuide();
  }, 0);
  setTimeout(() => {
    if (document.querySelector('section.sheet')) hideQuickGuide();
  }, 300);

  // 動的に label シートが生成されたタイミングでクイックガイドを自動非表示にする監視
  // （ユーザが「ラベルシールも印刷する」を後からチェックした場合など）
  (function setupSheetAppearObserver(){
    // 既に非表示なら不要
    const guide = document.getElementById('initialQuickGuide');
    if (!guide || guide.hidden) return;
    // すでに存在すれば即非表示
    if (document.querySelector('section.sheet')) { hideQuickGuide(); return; }
    const observer = new MutationObserver((mutations)=>{
      if (guide.hidden) { observer.disconnect(); return; }
      for (const m of mutations) {
        if (m.type === 'childList') {
          // 追加ノードおよびその子孫に sheet が無いか確認
          for (const n of m.addedNodes) {
            if (!(n instanceof HTMLElement)) continue;
            if (n.matches && n.matches('section.sheet')) { hideQuickGuide(); observer.disconnect(); return; }
            if (n.querySelector && n.querySelector('section.sheet')) { hideQuickGuide(); observer.disconnect(); return; }
          }
        }
      }
    });
    observer.observe(document.body, { childList: true, subtree: true });
    // window に保持し、後で明示解除も可能に
    window.__quickGuideSheetObserver = observer;
  })();

  // 複数のカスタムラベルを初期化（これによりシートが生成される場合、上の遅延チェックでガイドが非表示化される）
  CustomLabels.initialize(settings.customLabels);

  // 固定ヘッダーを初期表示（0枚でも表示）
  updatePrintCountDisplay(0, 0, 0);

   // 画像ドロップゾーンの初期化
  async function initGlobalDropZone(){
    const host = document.getElementById('imageDropZone');
  if (!host) return false;
    if (host.children.length>0) return true;
    const dz = await createOrderImageDropZone();
    if (dz && dz.element) { host.appendChild(dz.element); window.orderImageDropZone = dz; return true; }
    return false;
  }
  if (!await initGlobalDropZone()) {
    setTimeout(initGlobalDropZone, 100);
    setTimeout(initGlobalDropZone, 300);
  }
  await renderProductOrderImagesManager();
  // 画像表示トグルとドロップゾーンの表示を同期
  try {
    const cb = document.getElementById('orderImageEnable');
    if (cb) {
      toggleOrderImageRow(cb.checked);
      cb.addEventListener('change', () => toggleOrderImageRow(cb.checked));
    }
  } catch {}

  // フォント関連初期化（移行後API）
  if (window.CustomLabelFont) {
    CustomLabelFont.initializeFontSection?.();
    CustomLabelFont.initializeFontManager?.().then(() => {
      CustomLabelFont.loadCustomFontsCSS?.();
      CustomLabelFont.updateFontList?.(); // 初期表示
    });
    CustomLabelFont.initializeFontDropZone?.();
  }

  // 全ての注文データをクリアするボタン (QR統合後: orders + 関連QR情報)
  const clearAllButton = document.getElementById('clearAllButton');
  if (clearAllButton) {
    clearAllButton.onclick = async () => {
      if (!confirm('本当に全ての注文データ (QR含む) をクリアしますか？この操作は取り消せません。')) return;
      try {
        if (window.orderRepository && window.orderRepository.db && window.orderRepository.db.clearAllOrders) {
          await window.orderRepository.db.clearAllOrders();
          // repository のキャッシュもリセット
          window.orderRepository.cache.clear();
          window.orderRepository.emit();
        } else if (StorageManager && StorageManager.clearAllOrders) {
          // フォールバック: まだユーティリティがある場合
          await StorageManager.clearAllOrders();
        } else {
          alert('注文データクリア機能が利用できません');
          return;
        }
        alert('全ての注文データを削除しました');
        location.reload();
      } catch (e) {
        alert('注文データクリア中にエラーが発生しました: ' + e.message);
      }
    };
  }

  // バックアップ & リストア
  const backupBtn = document.getElementById('backupDBButton');
  const restoreBtn = document.getElementById('restoreDBButton');
  const restoreFile = document.getElementById('restoreDBFile');
  if (backupBtn) {
    backupBtn.onclick = async () => {
      try {
        const data = await StorageManager.exportAllData();
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const ts = new Date().toISOString().replace(/[:.]/g,'-');
        a.download = `boothcsv-backup-${ts}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      } catch (e) {
        alert('バックアップに失敗しました: ' + e.message);
      }
    };
  }
  if (restoreBtn && restoreFile) {
    restoreBtn.onclick = () => restoreFile.click();
    restoreFile.onchange = async () => {
      if (!restoreFile.files || restoreFile.files.length === 0) return;
      const file = restoreFile.files[0];
      if (!confirm('バックアップをリストアすると現在の全データは上書きされます。続行しますか？')) { restoreFile.value=''; return; }
      try {
        const text = await file.text();
        const json = JSON.parse(text);
        await StorageManager.importAllData(json, { clearExisting: true });
        alert('リストアが完了しました。ページを再読み込みします。');
        location.reload();
      } catch (e) {
        alert('リストアに失敗しました: ' + e.message);
      } finally { restoreFile.value=''; }
    };
  }

  // 全てのカスタムフォントをクリアするボタンのイベントリスナーを追加
  const clearAllFontsButton = document.getElementById('clearAllFontsButton');
  if (clearAllFontsButton) {
    clearAllFontsButton.onclick = async () => {
      if (confirm('本当に全てのカスタムフォントをクリアしますか？')) {
        try {
          if (!fontManager) {
            await initializeFontManager();
          }
          await fontManager.clearAllFonts();
          await loadCustomFontsCSS();
          updateFontList();
          alert('全てのカスタムフォントをクリアしました');
        } catch (error) {
          console.error('フォントクリアエラー:', error);
          alert('フォントのクリア中にエラーが発生しました');
        }
      }
    };
  }

  // 設定UIイベントをマッピングで一括登録（重複リスナー整理）
  registerSettingChangeHandlers();

  // カスタムラベル機能のイベントリスナー（遅延実行）
  setTimeout(function() {
  CustomLabels.setupEvents();
  }, 100);

  // ボタンの初期状態を設定
  CustomLabels.updateButtonStates();

  // labelskipnum の input イベントも統合ハンドラ内で処理（updateButtonStates sideEffect）

  // 初期カスタムラベルプレビューの更新（遅延実行）
  setTimeout(async function() {
    await updateCustomLabelsPreview();
  }, 200);

}, false);

// 自動処理関数（ファイル選択時や設定変更時に呼ばれる）
async function autoProcessCSV() {
  return new Promise(async (resolve, reject) => {
    try {
      const fileInput = document.getElementById("file");
      if (!fileInput.files || fileInput.files.length === 0) {
        console.log('ファイルが選択されていません。自動処理をスキップします。');
        updateProcessedOrdersVisibility();
        try {
          await renderProductOrderImagesManager();
          await updateCustomLabelsPreview();
          resolve();
        } catch (e) {
          reject(e);
        }
        return;
      }

      // カスタムラベルのバリデーション（エラー表示なし）
      // バリデーションエラーがあってもCSV処理は継続する
  const hasValidCustomLabels = CustomLabels.validateQuiet();
      if (!hasValidCustomLabels) {
        console.log('カスタムラベルにエラーがありますが、CSV処理は継続します。');
      }
      
      console.log('自動CSV処理を開始します...');
      setPreviewSource('');
      clearPreviousResults();
      updateProcessedOrdersVisibility();
    const config = buildCurrentPreviewConfig(document.getElementById('file').files[0]);
      
      Papa.parse(config.file, {
        header: true,
        skipEmptyLines: true,
        complete: async function(results) {
          try {
            await processCSVResults(results, config);
            console.log('自動CSV処理が完了しました。');
            // ペイント後に解決してスクロール処理が安定するようにする
            requestAnimationFrame(() => requestAnimationFrame(resolve));
          } catch (e) {
            reject(e);
          }
        }
      });
    } catch (error) {
      console.error('自動処理中にエラーが発生しました:', error);
      reject(error);
    }
  });
}

async function importCSVText(csvText, options = {}) {
  const sourceName = options.sourceName || 'BOOTH CSV';
  if (typeof csvText !== 'string' || !csvText.trim()) {
    throw new Error('CSVテキストが空です');
  }

  CustomLabels.updateButtonStates();
  await CustomLabels.updateSummary();
  updateSelectedSourceInfo(sourceName, true);
  updateProcessedOrdersVisibility();
  hideQuickGuide();
  setPreviewSource('');
  clearPreviousResults();

  const config = buildCurrentPreviewConfig(null);
  return await new Promise((resolve, reject) => {
    Papa.parse(csvText, {
      header: true,
      skipEmptyLines: true,
      complete: async function(results) {
        try {
          await processCSVResults(results, config);
          resolve({
            rowCount: Array.isArray(results.data) ? results.data.length : 0,
            sourceName
          });
        } catch (error) {
          reject(error);
        }
      },
      error: reject
    });
  });
}

function syncPreviewSurfaceState() {
  const { hasAnySheets, hasOrderSheets } = getRenderedSheetState();
  const hasCurrentPreview = Array.isArray(window.currentDisplayedOrderNumbers)
    && window.currentDisplayedOrderNumbers.length > 0;
  const hasSelection = Array.isArray(window.lastOrderSelection)
    && window.lastOrderSelection.length > 0;

  if (!hasOrderSheets && !hasCurrentPreview) {
    window.currentDisplayedOrderNumbers = [];
    if (!hasSelection) {
      setPreviewSource('');
    }
  }

  if (!hasAnySheets && !hasCurrentPreview && !hasSelection) {
    setPreviewSource('');
  }

  updatePreviewReturnButtonVisibility();
  updateProcessedOrdersVisibility();
}

// カスタムラベルのリアルタイムプレビュー更新
async function updateCustomLabelsPreview() {
  // 編集中は更新をスキップ
  if (isEditingCustomLabel) {
    debugLog('カスタムラベル編集中のため、プレビュー更新をスキップ');
    return;
  }

  try {
      const config = {
        file: document.getElementById('file').files[0],
        labelyn: settingsCache.labelyn,
        labelskip: settingsCache.labelskip,
        sortByPaymentDate: settingsCache.sortByPaymentDate,
        customLabelEnable: settingsCache.customLabelEnable,
  customLabels: settingsCache.customLabelEnable ? CustomLabels.getFromUI().filter(l=>l.enabled) : []
      };
    window.lastPreviewConfig = sanitizePreviewConfig(config);
    const hasSelection = Array.isArray(window.lastOrderSelection) && window.lastOrderSelection.length > 0;

    // 選択済み注文プレビューがある場合: ラベルON/OFFに関わらず注文明細を再生成する
    if (hasSelection) {
      // 再生成（renderPreviewFromRepository が labelyn を見てラベル出力の有無を切り替える）
      clearPreviousResults();
      await renderPreviewFromRepository(window.lastOrderSelection, config);
      return;
    }

    // CSV が無い場合 (カスタムラベル単独プレビュー)
    if (!config.labelyn) {
      clearPreviousResults();
      updatePrintCountDisplay(0, 0, 0);
      return;
    }
    if (!config.customLabelEnable) {
      // カスタムラベル無効でCSV無し → 何も描画しない
      clearPreviousResults();
      updatePrintCountDisplay(0, 0, 0);
      return;
    }
    const enabledLabels = (config.customLabels || []).filter(label => label.enabled && label.text.trim() !== '');
    if (enabledLabels.length > 0) {
      clearPreviousResults();
      await processCustomLabelsOnly(config, true);
    } else {
      clearPreviousResults();
      updatePrintCountDisplay(0, 0, 0);
    }
  } catch (error) {
    console.error('カスタムラベルプレビュー更新エラー:', error);
  } finally {
    syncPreviewSurfaceState();
  }
}

// 前回の処理結果をクリア
function clearPreviousResults() {
  // 結果セクションだけを削除（サイドバー等の一般sectionは残す）
  document.querySelectorAll('section.sheet').forEach(sec => sec.remove());
  
  // 印刷枚数表示もクリア
  clearPrintCountDisplay();
  updatePreviewReturnButtonVisibility();
}

async function returnToProcessedOrdersPanel() {
  clearPreviousResults();

  const fileInput = document.getElementById('file');
  if (fileInput) {
    fileInput.value = '';
  }

  updateSelectedSourceInfo('未選択', false);
  window.currentDisplayedOrderNumbers = [];
  window.lastOrderSelection = [];
  const customLabelOnlyConfig = buildCurrentPreviewConfig(null);
  window.lastPreviewConfig = hasEnabledCustomLabelContent(customLabelOnlyConfig)
    ? sanitizePreviewConfig(customLabelOnlyConfig)
    : null;
  setPreviewSource('');
  await renderProductOrderImagesManager();

  if (hasEnabledCustomLabelContent(customLabelOnlyConfig)) {
    await processCustomLabelsOnly(customLabelOnlyConfig, true);
  }

  await refreshProcessedOrdersPanel();
  syncPreviewSurfaceState();
}

async function ensureOrderRepository() {
  if (window.orderRepository) {
    if (typeof window.orderRepository.init === 'function' && !window.orderRepository.initialized) {
      await window.orderRepository.init();
    }
    return window.orderRepository;
  }
  if (typeof OrderRepository === 'undefined') {
    console.error('OrderRepository 未読込');
    return null;
  }
  const db = await StorageManager.ensureDatabase();
  if (!db) return null;
  const repo = new OrderRepository(db);
  await repo.init();
  window.orderRepository = repo;
  return repo;
}

async function setupProcessedOrdersPanel() {
  if (!window.ProcessedOrdersPanel || typeof window.ProcessedOrdersPanel.init !== 'function') return;
  await window.ProcessedOrdersPanel.init({
    ensureOrderRepository,
    clearPreviousResults,
    hideQuickGuide,
    renderPreview: renderPreviewFromRepository,
    setPreviewSource,
    notifyShipment: notifySelectedOrdersShipment
  });
}

async function refreshProcessedOrdersPanel() {
  if (!window.ProcessedOrdersPanel || typeof window.ProcessedOrdersPanel.refresh !== 'function') return;
  await window.ProcessedOrdersPanel.refresh();
}

function updateProcessedOrdersVisibility() {
  if (!window.ProcessedOrdersPanel || typeof window.ProcessedOrdersPanel.updateVisibility !== 'function') return;
  window.ProcessedOrdersPanel.updateVisibility();
}

function formatShippingCommentPreview(messageTemplate) {
  if (typeof messageTemplate !== 'string') return '未設定';
  const normalized = messageTemplate.replace(/\r\n?/g, '\n').trim();
  return normalized || '未設定';
}

async function collectShipmentConfirmationStatus(orderNumbers, bridge) {
  const responses = new Map();
  const shipped = [];
  const pending = [];
  const failed = [];

  for (const orderNumber of orderNumbers) {
    try {
      const response = await bridge.collectOrderShipmentStatus(orderNumber);
      responses.set(orderNumber, response);
      const shippedAtValue = response && response.ok ? (response.shippedAt || response.shippedAtRaw || '') : '';
      if (response && response.ok && shippedAtValue) {
        shipped.push(orderNumber);
      } else if (response && response.ok) {
        pending.push(orderNumber);
      } else {
        const diagnostic = response && response.diagnosticsSummary ? ` | ${response.diagnosticsSummary}` : '';
        failed.push(`${orderNumber}: ${response && response.error ? response.error : '発送状態の取得に失敗しました'}${diagnostic}`);
      }
    } catch (error) {
      failed.push(`${orderNumber}: ${error && error.message ? error.message : '発送状態の取得に失敗しました'}`);
    }
  }

  return { responses, shipped, pending, failed };
}

function buildShipmentConfirmationMessage(orderNumbers, messageTemplate, useExtensionBridge, statusSummary) {
  const count = Array.isArray(orderNumbers) ? orderNumbers.length : 0;
  const commentPreview = formatShippingCommentPreview(messageTemplate);

  if (useExtensionBridge) {
    const shippedCount = statusSummary && Array.isArray(statusSummary.shipped) ? statusSummary.shipped.length : 0;
    const pendingCount = statusSummary && Array.isArray(statusSummary.pending) ? statusSummary.pending.length : 0;
    const failedCount = statusSummary && Array.isArray(statusSummary.failed) ? statusSummary.failed.length : 0;

    return [
      '選択した注文を確認し、未通知の注文のみ BOOTH で発送通知します。',
      '',
      `対象件数: ${count}件`,
      `未通知で BOOTH へ発送通知する注文: ${pendingCount}件`,
      `通知済みで発送日時を再取得する注文: ${shippedCount}件`,
      failedCount > 0 ? `発送状態の取得に失敗した注文: ${failedCount}件` : '',
      '',
      '通知コメント:',
      commentPreview,
      '',
      '通知済みの注文は BOOTH 側で再通知できないため、発送日時の再取得のみ行います。',
      failedCount > 0 ? '発送状態を取得できなかった注文は、続行後に再度確認したうえで処理します。' : '',
      'この操作を実行すると、未通知の注文は BOOTH 側で発送完了通知まで実施します。',
      '続行しますか？'
    ].filter(Boolean).join('\n');
  }

  return [
    '選択した注文の発送日時を記録します。',
    '',
    `対象件数: ${count}件`,
    '',
    '通知コメント:',
    commentPreview,
    '',
    'この環境では発送日時のみ反映します。',
    'BOOTH の注文詳細で、上記コメントを使って手動で発送完了通知を行ってください。',
    '続行しますか？'
  ].join('\n');
}

async function notifySelectedOrdersShipment(orderNumbers) {
  const repo = await ensureOrderRepository();
  if (!repo) throw new Error('注文データを読み込めませんでした');

  const selectedOrderNumbers = normalizeOrderSelection(orderNumbers);
  if (selectedOrderNumbers.length === 0) return;

  const useExtensionBridge = canUseExtensionBridge();
  const shippingMessageTemplate = (settingsCache && typeof settingsCache.shippingMessageTemplate === 'string')
    ? settingsCache.shippingMessageTemplate
    : '';

  if (!useExtensionBridge) {
    if (!confirm(buildShipmentConfirmationMessage(selectedOrderNumbers, shippingMessageTemplate, false))) {
      return;
    }

    const shippedAt = new Date().toISOString();
    for (const orderNumber of selectedOrderNumbers) {
      await repo.markShipped(orderNumber, shippedAt);
    }
    await refreshProcessedOrdersPanel();
    alert('発送日時を記録しました。\nBOOTHの注文詳細で発送完了を通知してください。');
    return;
  }

  const bridge = window.BoothCSVExtensionBridge;
  if (!bridge || typeof bridge.collectOrderShipmentStatus !== 'function' || typeof bridge.notifyOrderShipment !== 'function') {
    throw new Error('Chrome拡張との連携が初期化されていません');
  }

  const confirmationStatus = await collectShipmentConfirmationStatus(selectedOrderNumbers, bridge);
  if (!confirm(buildShipmentConfirmationMessage(selectedOrderNumbers, shippingMessageTemplate, true, confirmationStatus))) {
    return;
  }

  const updated = [];
  const pending = [];
  const failed = [];
  const submitted = [];

  for (const orderNumber of selectedOrderNumbers) {
    try {
      const response = confirmationStatus.responses.has(orderNumber)
        ? confirmationStatus.responses.get(orderNumber)
        : await bridge.collectOrderShipmentStatus(orderNumber);
      const shippedAtValue = response && response.ok ? (response.shippedAt || response.shippedAtRaw || '') : '';
      if (response && response.ok && shippedAtValue) {
        await repo.markShipped(orderNumber, shippedAtValue);
        updated.push({ orderNumber, shippedAt: response.shippedAtRaw || response.shippedAt || shippedAtValue });
      } else if (response && response.ok) {
        const notifyResponse = await bridge.notifyOrderShipment(orderNumber, shippingMessageTemplate);
        const notifiedShippedAt = notifyResponse ? (notifyResponse.shippedAt || notifyResponse.shippedAtRaw || '') : '';
        if (notifyResponse && notifyResponse.ok && notifiedShippedAt) {
          await repo.markShipped(orderNumber, notifiedShippedAt);
          updated.push({ orderNumber, shippedAt: notifyResponse.shippedAtRaw || notifyResponse.shippedAt || notifiedShippedAt });
          if (notifyResponse.submitted) {
            submitted.push(orderNumber);
          }
        } else {
          const diagnostic = notifyResponse && notifyResponse.diagnosticsSummary ? ` (${notifyResponse.diagnosticsSummary})` : '';
          pending.push(`${orderNumber}${diagnostic}`);
        }
      } else {
        const diagnostic = response && response.diagnosticsSummary ? ` | ${response.diagnosticsSummary}` : '';
        failed.push(`${orderNumber}: ${response && response.error ? response.error : '発送状態の取得に失敗しました'}${diagnostic}`);
      }
    } catch (error) {
      failed.push(`${orderNumber}: ${error && error.message ? error.message : '発送状態の取得に失敗しました'}`);
    }
  }

  await refreshProcessedOrdersPanel();

  const messages = [];
  if (updated.length > 0) {
    messages.push(`発送日時を反映: ${updated.length}件`);
  }
  if (submitted.length > 0) {
    messages.push(`発送通知を送信: ${submitted.length}件`);
  }
  if (pending.length > 0) {
    messages.push(`未発送または発送日時未取得: ${pending.length}件`);
    messages.push('発送通知後の状態確認が完了しなかった注文があります。必要に応じて BOOTH の注文詳細を確認してください。');
    messages.push(pending.slice(0, 3).join('\n'));
  }
  if (failed.length > 0) {
    messages.push(`取得失敗: ${failed.length}件`);
    messages.push(failed.slice(0, 3).join('\n'));
  }
  if (messages.length === 0) {
    messages.push('発送日時を反映できる注文はありませんでした。');
  }
  alert(messages.join('\n'));
}

async function persistCsvToRepository(results) {
  const repo = await ensureOrderRepository();
  if (!repo) return { orderNumbers: [], totalRows: 0 };

  const rows = results && Array.isArray(results.data) ? results.data : [];

  if (DEBUG_MODE && rows.length > 0) {
    const first = rows[0];
    if (first) {
      debugLog('[csv] 先頭行キー一覧', Object.keys(first));
    } else {
      debugLog('[csv] CSVにデータ行がありません');
    }
  }

  await repo.bulkUpsert(rows);
  await refreshProcessedOrdersPanel();

  const orderNumbers = [];
  let debugValidOrderCount = 0;
  for (const row of rows) {
    const raw = getOrderNumberFromCSVRow(row);
    if (!raw) continue;
    const normalized = OrderRepository.normalize(raw);
    orderNumbers.push(normalized);
    debugValidOrderCount++;
    if (DEBUG_MODE) {
      const rec = repo.get(normalized);
      debugLog('[repo] 読み込み注文', { raw, normalized, exists: !!rec, printedAt: rec ? rec.printedAt : null });
    }
  }

  if (DEBUG_MODE) {
    debugLog('[csv] 行数サマリ', {
      totalRows: rows.length,
      withOrderNumber: debugValidOrderCount,
      repositoryStored: repo.getAll().length
    });
  }

  return { orderNumbers, totalRows: rows.length };
}

function sanitizePreviewConfig(config) {
  const fallbackSkip = (config && typeof config.labelskip !== 'undefined') ? config.labelskip : settingsCache.labelskip;
  return {
    labelyn: !!config.labelyn,
    labelskip: fallbackSkip,
    sortByPaymentDate: !!config.sortByPaymentDate,
    customLabelEnable: !!config.customLabelEnable,
    customLabels: Array.isArray(config.customLabels) ? config.customLabels.map(label => ({
      text: label.text,
      html: label.html,
      count: label.count,
      fontSize: label.fontSize,
      enabled: label.enabled !== false
    })) : []
  };
}

function normalizeOrderSelection(orderNumbers) {
  if (!Array.isArray(orderNumbers)) return [];
  const seen = new Set();
  const normalized = [];
  for (const num of orderNumbers) {
    const key = OrderRepository.normalize(num);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    normalized.push(key);
  }
  return normalized;
}

async function renderPreviewFromRepository(orderNumbers, config) {
  const repo = await ensureOrderRepository();
  if (!repo) return;

  const baseSelection = Array.isArray(orderNumbers) ? orderNumbers : (window.lastOrderSelection || []);
  const selection = normalizeOrderSelection(baseSelection);
  const sourceConfig = config && typeof config === 'object' ? config : (window.lastPreviewConfig || {});
  const previewConfig = sanitizePreviewConfig(sourceConfig);
  window.lastPreviewConfig = previewConfig;

  const records = selection.map(num => repo.get(num)).filter(rec => !!rec && !!rec.row);
  window.lastOrderSelection = records.map(rec => rec.orderNumber);

  const detailRecords = [...records];
  if (previewConfig.sortByPaymentDate) {
    detailRecords.sort((a, b) => {
      const timeA = (a.row && a.row[CONSTANTS.CSV.PAYMENT_DATE_COLUMN]) || '';
      const timeB = (b.row && b.row[CONSTANTS.CSV.PAYMENT_DATE_COLUMN]) || '';
      const cmp = timeA.localeCompare(timeB);
      if (cmp !== 0) return cmp;
      return a.orderNumber.localeCompare(b.orderNumber, 'ja', { numeric: true, sensitivity: 'base' });
    });
  }

  const detailRows = detailRecords.map(rec => rec.row).filter(Boolean);
  const unprintedRecords = detailRecords.filter(rec => !rec.printedAt);
  const labelRows = unprintedRecords.map(rec => rec.row).filter(Boolean);

  window.currentDisplayedOrderNumbers = detailRecords.map(rec => rec.orderNumber);
  await renderProductOrderImagesManager();

  const csvRowCountForLabels = labelRows.length;
  const effectiveCustomLabels = previewConfig.customLabelEnable ? (previewConfig.customLabels || []) : [];
  const enabledCustomLabels = effectiveCustomLabels.filter(label => label.enabled !== false);
  const totalCustomLabelCount = enabledCustomLabels.reduce((sum, label) => sum + (parseInt(label.count, 10) || 0), 0);
  const skipCount = parseInt(previewConfig.labelskip, 10) || 0;
  const totalLabelsNeeded = skipCount + csvRowCountForLabels + totalCustomLabelCount;
  const requiredSheets = csvRowCountForLabels === 0 && totalCustomLabelCount === 0
    ? 0
    : Math.ceil(totalLabelsNeeded / CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET);

  const labelSet = new Set(unprintedRecords.map(rec => rec.orderNumber));

  await generateOrderDetails(detailRows, Array(skipCount).fill(''), labelSet);

  let totalLabelArray;
  if (previewConfig.labelyn) {
    totalLabelArray = new Array(skipCount).fill('');
    const numbersInOrder = unprintedRecords.map(rec => rec.orderNumber);
    totalLabelArray.push(...numbersInOrder);

    if (previewConfig.customLabelEnable && enabledCustomLabels.length > 0) {
      for (const customLabel of enabledCustomLabels) {
        const count = parseInt(customLabel.count, 10) || 0;
        for (let i = 0; i < count; i++) {
          totalLabelArray.push({
            type: 'custom',
            content: customLabel.html || customLabel.text,
            fontSize: customLabel.fontSize || '10pt'
          });
        }
      }
    }

    if (totalLabelArray.length > 0) {
      await generateLabels(totalLabelArray, { skipOnFirstSheet: skipCount });
    }
  }

  if (DEBUG_MODE) {
    debugLog('[label] ラベル生成サマリ', {
      skipCount,
      unprintedCount: unprintedRecords.length,
      detailCount: detailRows.length,
      labelRowCount: labelRows.length,
      customLabelCount: totalCustomLabelCount,
      totalLabelArrayLength: Array.isArray(totalLabelArray) ? totalLabelArray.length : 0,
      requiredSheets
    });
  }

  const labelSheetsForDisplay = previewConfig.labelyn ? requiredSheets : 0;
  const customFacesForDisplay = (previewConfig.labelyn && previewConfig.customLabelEnable)
    ? totalCustomLabelCount
    : 0;

  updatePrintCountDisplay(unprintedRecords.length, labelSheetsForDisplay, customFacesForDisplay);
  await CustomLabels.updateSummary();
  CustomLabels.updateButtonStates();
  updatePreviewReturnButtonVisibility();
  updateProcessedOrdersVisibility();

  return {
    detailCount: detailRows.length,
    unprintedCount: unprintedRecords.length
  };
}

// CSVデータを処理して注文詳細とラベルを生成
async function processCSVResults(results, config) {
  const { orderNumbers } = await persistCsvToRepository(results);
  return await renderPreviewFromRepository(orderNumbers, config);
}

// カスタムラベルのみを処理（CSV無し）
async function processCustomLabelsOnly(config, isPreviewMode = false) {
  // 複数カスタムラベルの総面数を計算
  const totalCustomLabelCount = config.customLabels.reduce((sum, label) => sum + label.count, 0);
  const labelskipNum = parseInt(settingsCache.labelskip, 10) || 0; // settingsCache 参照
  
  // 有効なカスタムラベルがあるかチェック
  const validLabels = config.customLabels.filter(label => label.text.trim() !== '');
  if (validLabels.length === 0) {
    if (!isPreviewMode) {
      alert('印刷する文字列を入力してください。');
    }
    return;
  }
  
  if (totalCustomLabelCount === 0) {
    if (!isPreviewMode) {
      alert('印刷する面数を1以上に設定してください。');
    }
    return;
  }
  
  // 複数シートの分散計算
  const sheetsInfo = CustomLabelCalculator.calculateMultiSheetDistribution(totalCustomLabelCount, labelskipNum);
  const totalSheets = sheetsInfo.length;
  
  // 各シート用のラベル配列を作成
  let remainingLabels = [...config.customLabels]; // コピーを作成
  let currentSkip = labelskipNum;
  
  for (let sheetIndex = 0; sheetIndex < totalSheets; sheetIndex++) {
    const sheetInfo = sheetsInfo[sheetIndex];
    const labelarr = [];
    
    // スキップラベルを追加（最初のシートのみ）
    if (sheetIndex === 0 && currentSkip > 0) {
      for (let i = 0; i < currentSkip; i++) {
        labelarr.push("");
      }
    }
    
    // このシートに配置するラベル数
    let labelsToPlaceInSheet = sheetInfo.labelCount;
    
    // カスタムラベルを配置
    for (let labelIndex = 0; labelIndex < remainingLabels.length && labelsToPlaceInSheet > 0; labelIndex++) {
      const customLabel = remainingLabels[labelIndex];
      
      if (customLabel.count > 0) {
        const placedCount = Math.min(customLabel.count, labelsToPlaceInSheet);
        
        for (let i = 0; i < placedCount; i++) {
          labelarr.push({ 
            type: 'custom', 
            content: customLabel.html || customLabel.text,
            fontSize: customLabel.fontSize || '10pt'
          });
        }
        
        customLabel.count -= placedCount;
        labelsToPlaceInSheet -= placedCount;
      }
    }
    
    // このシートのラベルを生成
    if (labelarr.length > 0) {
      await generateLabels(labelarr, { skipOnFirstSheet: labelskipNum });
    }
    
    // 使い切ったラベルを削除
    remainingLabels = remainingLabels.filter(label => label.count > 0);
    currentSkip = 0; // 2シート目以降はスキップなし
  }
  
  // ヘッダーの印刷枚数表示を更新（カスタムラベルのみ）
  if (!isPreviewMode) {
    updatePrintCountDisplay(0, sheetsInfo.length, totalCustomLabelCount);
  } else {
    // プレビューモードでも表示を更新
    updatePrintCountDisplay(0, sheetsInfo.length, totalCustomLabelCount);
  }
  
  // ボタンの状態を更新
  CustomLabels.updateButtonStates();
}

// ヘッダーの印刷枚数表示を更新する関数
function updatePrintCountDisplay(orderSheetCount = 0, labelSheetCount = 0, customLabelCount = 0) {
  const displayElement = document.getElementById('printCountDisplay');
  const orderCountElement = document.getElementById('orderSheetCount');
  const labelCountElement = document.getElementById('labelSheetCount');
  const customLabelCountElement = document.getElementById('customLabelCount');
  const customLabelItem = document.getElementById('customLabelCountItem');
  
  debugLog('general', `updatePrintCountDisplay呼び出し: ラベル:${labelSheetCount}枚, 普通紙:${orderSheetCount}枚, カスタム:${customLabelCount}面`);
  
  if (!displayElement) {
    console.error('printCountDisplay要素が見つかりません');
    return;
  }
  
  // 値を更新
  if (orderCountElement) {
    orderCountElement.textContent = orderSheetCount;
  } else {
    console.error('orderSheetCount要素が見つかりません');
  }
  
  if (labelCountElement) {
    labelCountElement.textContent = labelSheetCount;
  } else {
    console.error('labelSheetCount要素が見つかりません');
  }
  
  if (customLabelCountElement) {
    customLabelCountElement.textContent = customLabelCount;
  } else {
    console.error('customLabelCount要素が見つかりません');
  }
  
  // カスタムラベルの表示/非表示を制御
  if (customLabelItem) {
    customLabelItem.style.display = customLabelCount > 0 ? 'flex' : 'none';
  }
  
  // 全体を常に表示（0枚でも表示）
  displayElement.style.display = 'flex';
  
  debugLog('general', `印刷枚数更新完了: ラベル:${labelSheetCount}枚, 普通紙:${orderSheetCount}枚, カスタム:${customLabelCount}面`);
}

// 印刷枚数をクリアする関数
function clearPrintCountDisplay() {
  updatePrintCountDisplay(0, 0, 0);
}

// 注文データから注文明細のHTML要素を生成
async function generateOrderDetails(data, labelarr, labelSet = null, printedAtMap = null) {
  const tOrder = document.querySelector('#注文明細');
  
  for (let row of data) {
    const cOrder = document.importNode(tOrder.content, true);
    let orderNumber = '';
    // --- setOrderInfo inline 化 ---
    for (let c of Object.keys(row).filter(key => key != CONSTANTS.CSV.PRODUCT_COLUMN)) {
      const divc = cOrder.querySelector('.' + c);
      if (!divc) continue;
      if (c === CONSTANTS.CSV.ORDER_NUMBER_COLUMN) {
        orderNumber = getOrderNumberFromCSVRow(row);
        divc.textContent = orderNumber; // 生の番号のみ（装飾はCSS）
      } else if (row[c]) {
        divc.textContent = row[c];
      }
    }
    // section にアンカーID付与
    if (orderNumber) {
      const sectionEl = cOrder.querySelector('section.sheet');
      if (sectionEl) {
        const normalized = String(orderNumber).trim();
        sectionEl.id = `order-${normalized}`;
      }
    }

    // 注文明細ごとの非印刷パネル（印刷日時）のセットアップ
    try {
      await setupOrderPrintedAtPanel(cOrder, orderNumber);
    } catch (e) {
      console.warn('印刷日時パネル設定エラー:', e);
    }
    
    
    // 個別画像ドロップゾーンの作成
    await createIndividualImageDropZone(cOrder, orderNumber);
    
    // 商品項目の処理
    processProductItems(cOrder, row);
    
    // 画像表示の処理
    await displayOrderImage(cOrder, orderNumber, row);
    
    // 追加前にルートsectionを特定
    const rootSection = cOrder.querySelector('section.sheet');
    // まずDOMに追加
    document.body.appendChild(cOrder);
    // 印刷状態でクラスを付与
    try {
  const normalized = (orderNumber == null) ? '' : String(orderNumber).trim();
      if (rootSection && normalized) {
        if (window.orderRepository) {
          const rec = window.orderRepository.get(normalized);
          if (rec?.printedAt) rootSection.classList.add('is-printed');
          else rootSection.classList.remove('is-printed');
        }
      }
    } catch (e) { console.warn('printed state apply error', e); }
  }
}

// 各注文明細の非印刷パネルをセットアップ（印刷日時表示とクリア機能）
async function setupOrderPrintedAtPanel(cOrder, orderNumber) {
  const panel = cOrder.querySelector('.order-print-info');
  if (!panel) return;
  const dateEl = panel.querySelector('.printed-at');
  const markPrintedBtn = panel.querySelector('.mark-printed');
  const clearBtn = panel.querySelector('.clear-printed-at');
  const normalized = (orderNumber == null) ? '' : String(orderNumber).trim();
  if (!normalized) {
    if (dateEl) dateEl.textContent = '未印刷';
    if (markPrintedBtn) { markPrintedBtn.style.display = ''; markPrintedBtn.disabled = false; }
    if (clearBtn) { clearBtn.style.display = 'none'; clearBtn.disabled = true; }
    return;
  }
  const order = (window.orderRepository) ? window.orderRepository.get(normalized) : null;
  const printedAt = order?.printedAt || null;
  if (dateEl) {
    dateEl.textContent = printedAt ? new Date(printedAt).toLocaleString() : '未印刷';
  }
  // ボタン表示の切り替え
  if (printedAt) {
    if (markPrintedBtn) markPrintedBtn.style.display = 'none';
    if (clearBtn) clearBtn.style.display = '';
  } else {
    if (markPrintedBtn) markPrintedBtn.style.display = '';
    if (clearBtn) clearBtn.style.display = 'none';
  }

  // 「印刷済みにする」
  if (markPrintedBtn) {
    markPrintedBtn.disabled = !!printedAt;
    markPrintedBtn.onclick = async () => {
      try {
        const now = new Date().toISOString();
        const anchorOrder = normalized;
        const doc = document.scrollingElement || document.documentElement;
        debugLog('🟢 [mark] click', { order: anchorOrder, beforeScrollY: window.scrollY, beforeScrollH: doc.scrollHeight, sections: document.querySelectorAll('section.sheet').length });
        if (window.orderRepository) await window.orderRepository.markPrinted(normalized, now);

  // 部分更新：UI更新＋該当セクションをグレーアウト、ラベル/枚数再計算
  if (dateEl) dateEl.textContent = new Date(now).toLocaleString();
  if (markPrintedBtn) { markPrintedBtn.style.display = 'none'; }
  if (clearBtn) { clearBtn.style.display = ''; clearBtn.disabled = false; }
  const sectionEl = panel.closest('section.sheet');
  if (sectionEl) sectionEl.classList.add('is-printed');
  await regenerateLabelsFromDB();
  recalcAndUpdateCounts();
      } catch (e) {
        alert('印刷済みへの更新中にエラーが発生しました');
        console.error(e);
      }
    };
  }

  // 「印刷日時をクリア」
  if (clearBtn) {
    clearBtn.disabled = !printedAt;
    clearBtn.onclick = async () => {
      const ok = confirm(`注文 ${normalized} の印刷日時をクリアしますか？`);
      if (!ok) return;
      try {
        const anchorOrder = normalized;
        const doc = document.scrollingElement || document.documentElement;
        debugLog('🟠 [clear] click', { order: anchorOrder, beforeScrollY: window.scrollY, beforeScrollH: doc.scrollHeight, sections: document.querySelectorAll('section.sheet').length });
        if (window.orderRepository) await window.orderRepository.clearPrinted(normalized);

  // 部分更新：UI更新＋グレーアウト解除、ラベル/枚数再計算
  if (dateEl) dateEl.textContent = '未印刷';
  if (markPrintedBtn) { markPrintedBtn.style.display = ''; markPrintedBtn.disabled = false; }
  if (clearBtn) { clearBtn.style.display = 'none'; }
  const sectionEl = panel.closest('section.sheet');
  if (sectionEl) sectionEl.classList.remove('is-printed');
  await regenerateLabelsFromDB();
  recalcAndUpdateCounts();
      } catch (e) {
        alert('印刷日時のクリア中にエラーが発生しました');
        console.error(e);
      }
    };
  }
}

// 設定変更イベントを一括登録し重複ロジックを削減
function registerSettingChangeHandlers() {
  const defs = [
    { id: 'labelyn', key: 'LABEL_SETTING', type: 'checkbox' },
    { id: 'labelskipnum', key: 'LABEL_SKIP', type: 'number' },
    { id: 'sortByPaymentDate', key: 'SORT_BY_PAYMENT', type: 'checkbox' },
    { id: 'orderImageEnable', key: 'ORDER_IMAGE_ENABLE', type: 'checkbox', sideEffects: [
        (val) => toggleOrderImageRow(val),
        async (val) => { if (val) await renderProductOrderImagesManager(); },
        (val) => { 
          // CSSクラスをトグルするだけで再描画はしない
          document.body.classList.toggle('order-image-hidden', !val);
        }
      ] },
    // customLabelEnable はカスタムラベルUI群と関係が深く遅延初期化 setupCustomLabelEvents() 内に既存処理があるためここでは扱わない
  ];

  for (const def of defs) {
    const el = document.getElementById(def.id);
    if (!el) continue;
    const keyConst = StorageManager.KEYS[def.key];
    el.addEventListener('change', async function() {
      debugLog('general', `⚙️ 設定変更ハンドラがトリガーされました: ${def.id}`);
      let value;
      if (def.type === 'checkbox') value = this.checked;
      else if (def.type === 'number') value = parseInt(this.value, 10) || 0;
      else value = this.value;
      await StorageManager.set(keyConst, value);
      switch(def.id){
        case 'labelyn': settingsCache.labelyn = value; break;
        case 'labelskipnum': settingsCache.labelskip = value; break;
        case 'sortByPaymentDate': settingsCache.sortByPaymentDate = value; break;
        case 'orderImageEnable': settingsCache.orderImageEnable = value; break;
      }
      if (def.id === 'labelskipnum') { CustomLabels.updateButtonStates(); }
      if (def.id === 'labelyn') { updateExtensionUIVisibility(); }
      if (Array.isArray(def.sideEffects)) {
        for (const fx of def.sideEffects) {
          try { await fx(value); } catch(e) { console.error('sideEffect error', def.id, e); }
        }
      }
      // orderImageEnable の変更では autoProcessCSV を呼ばない
      if (def.id !== 'orderImageEnable') {
        await autoProcessCSV();
      }
    });

    if (def.id === 'labelskipnum') {
      el.addEventListener('input', function() { CustomLabels.updateButtonStates(); });
    }
  }
}

// 現在の「読み込んだファイル全て表示」のON/OFFを返す

function initializeShippingMessageTemplateUI(initialValue) {
  const textarea = document.getElementById('shippingMessageTemplate');
  const saveBtn = document.getElementById('saveShippingMessageTemplate');
  const copyBtn = document.getElementById('copyShippingMessageTemplate');

  if (textarea) {
    textarea.value = initialValue || '';
  }

  if (initializeShippingMessageTemplateUI._initialized) {
    return;
  }
  initializeShippingMessageTemplateUI._initialized = true;

  if (saveBtn) {
    saveBtn.addEventListener('click', async () => {
      const value = textarea ? textarea.value : '';
      try {
        await StorageManager.set(StorageManager.KEYS.SHIPPING_MESSAGE_TEMPLATE, value);
        settingsCache.shippingMessageTemplate = value;
        alert('出荷連絡の定型文を保存しました');
      } catch (e) {
        console.error('shipping template save error', e);
        alert('定型文の保存に失敗しました: ' + (e && e.message ? e.message : e));
      }
    });
  }

  if (copyBtn) {
    copyBtn.addEventListener('click', async () => {
      const value = textarea ? textarea.value : '';
      if (!value) {
        alert('コピーする定型文がありません。');
        return;
      }
      try {
        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
          await navigator.clipboard.writeText(value);
        } else if (textarea) {
          textarea.focus();
          textarea.select();
          const copied = document.execCommand ? document.execCommand('copy') : false;
          if (!copied) {
            throw new Error('copyCommandFailed');
          }
          if (window.getSelection) {
            const selection = window.getSelection();
            if (selection && selection.removeAllRanges) {
              selection.removeAllRanges();
            }
          }
        } else {
          throw new Error('clipboardUnavailable');
        }
        alert('定型文をコピーしました');
      } catch (e) {
        console.error('shipping template copy error', e);
        alert('コピーに失敗しました。ブラウザの許可設定を確認してください。');
      }
    });
  }
}

function getYamatoHandlingCodesFromUI() {
  return YAMATO_HANDLING_CHECKBOXES.map(function(id) {
    const input = document.getElementById(id);
    if (!input || !input.checked) return null;
    return input.dataset.handlingCode || null;
  }).filter(Boolean);
}

function normalizeYamatoDescriptionHistory(history) {
  const normalized = Array.isArray(history)
    ? history.map(function(value) { return String(value || '').trim(); }).filter(Boolean)
    : [];
  return Array.from(new Set(normalized)).slice(0, YAMATO_DESCRIPTION_HISTORY_LIMIT);
}

function mergeYamatoDescriptionHistory(value, history) {
  const normalizedValue = String(value || '').trim();
  const normalizedHistory = normalizeYamatoDescriptionHistory(history).filter(function(entry) {
    return entry !== normalizedValue;
  });
  return normalizedValue
    ? [normalizedValue].concat(normalizedHistory).slice(0, YAMATO_DESCRIPTION_HISTORY_LIMIT)
    : normalizedHistory;
}

function renderYamatoDescriptionHistory(history) {
  const datalist = document.getElementById('yamatoDescriptionHistory');
  if (!datalist) return;
  datalist.innerHTML = '';
  normalizeYamatoDescriptionHistory(history).forEach(function(entry) {
    const option = document.createElement('option');
    option.value = entry;
    datalist.appendChild(option);
  });
}

function getCurrentYamatoIssueSettingsFromUI() {
  return {
    packageSizeId: document.getElementById('yamatoPackageSizeId')?.value || settingsCache.yamatoPackageSizeId || '',
    codeType: document.getElementById('yamatoCodeType')?.value || settingsCache.yamatoCodeType || '',
    description: (document.getElementById('yamatoDescription')?.value || settingsCache.yamatoDescription || '').trim(),
    includeOrderNumber: !!document.getElementById('yamatoIncludeOrderNumber')?.checked,
    handlingCodes: getYamatoHandlingCodesFromUI()
  };
}

async function persistYamatoIssueSettings(settings, options = {}) {
  const normalized = {
    packageSizeId: String(settings.packageSizeId || '').trim(),
    codeType: String(settings.codeType || '').trim(),
    description: String(settings.description || '').trim(),
    includeOrderNumber: settings.includeOrderNumber !== false,
    handlingCodes: Array.isArray(settings.handlingCodes) ? Array.from(new Set(settings.handlingCodes.filter(Boolean))) : []
  };

  await StorageManager.set(StorageManager.KEYS.YAMATO_PACKAGE_SIZE_ID, normalized.packageSizeId);
  await StorageManager.set(StorageManager.KEYS.YAMATO_CODE_TYPE, normalized.codeType);
  await StorageManager.set(StorageManager.KEYS.YAMATO_DESCRIPTION, normalized.description);
  await StorageManager.set(StorageManager.KEYS.YAMATO_INCLUDE_ORDER_NUMBER, normalized.includeOrderNumber);
  await StorageManager.set(StorageManager.KEYS.YAMATO_HANDLING_CODES, normalized.handlingCodes);

  settingsCache.yamatoPackageSizeId = normalized.packageSizeId;
  settingsCache.yamatoCodeType = normalized.codeType;
  settingsCache.yamatoDescription = normalized.description;
  settingsCache.yamatoIncludeOrderNumber = normalized.includeOrderNumber;
  settingsCache.yamatoHandlingCodes = normalized.handlingCodes.slice();

  if (options.updateHistory && normalized.description) {
    const nextHistory = mergeYamatoDescriptionHistory(normalized.description, settingsCache.yamatoDescriptionHistory);
    await StorageManager.set(StorageManager.KEYS.YAMATO_DESCRIPTION_HISTORY, nextHistory);
    settingsCache.yamatoDescriptionHistory = nextHistory;
    renderYamatoDescriptionHistory(nextHistory);
  }

  return normalized;
}

function focusElementById(elementId) {
  const element = document.getElementById(elementId);
  if (!element) return;
  element.focus();
  if (typeof element.scrollIntoView === 'function') {
    element.scrollIntoView({ block: 'center', behavior: 'smooth' });
  }
}

function validateYamatoIssueSettings(settings) {
  const packageSizeId = typeof settings.packageSizeId === 'string' ? settings.packageSizeId.trim() : '';
  const codeType = typeof settings.codeType === 'string' ? settings.codeType.trim() : '';
  const description = typeof settings.description === 'string' ? settings.description.trim() : '';
  const handlingCodes = Array.isArray(settings.handlingCodes) ? Array.from(new Set(settings.handlingCodes.filter(Boolean))) : [];

  const packageSizeOptions = Array.from(document.getElementById('yamatoPackageSizeId')?.options || []).map(function(option) { return option.value; });
  if (!packageSizeId || !packageSizeOptions.includes(packageSizeId)) {
    focusElementById('yamatoPackageSizeId');
    throw new Error('発送コード発行設定の「サイズ」を選択してください');
  }

  const codeTypeOptions = Array.from(document.getElementById('yamatoCodeType')?.options || []).map(function(option) { return option.value; });
  if (!codeType || !codeTypeOptions.includes(codeType)) {
    focusElementById('yamatoCodeType');
    throw new Error('発送コード発行設定の「発送場所」を選択してください');
  }

  if (!description) {
    focusElementById('yamatoDescription');
    throw new Error('発送コード発行設定の「品名」を入力してください');
  }

  if (description.length > 8) {
    focusElementById('yamatoDescription');
    throw new Error('発送コード発行設定の「品名」は8文字以内で入力してください');
  }

  if (handlingCodes.length > 2) {
    focusElementById('yamatoHandlingPrecisionEquipment');
    throw new Error('発送コード発行設定の「荷扱い」は2つまで選択できます');
  }

  return {
    packageSizeId,
    codeType,
    description,
    includeOrderNumber: settings.includeOrderNumber !== false,
    handlingCodes
  };
}

function initializeYamatoIssueSettingsUI(initialSettings) {
  if (initializeYamatoIssueSettingsUI._initialized) {
    return;
  }
  initializeYamatoIssueSettingsUI._initialized = true;

  const definitions = [
    { id: 'yamatoPackageSizeId', key: StorageManager.KEYS.YAMATO_PACKAGE_SIZE_ID, getValue: function(el) { return el.value || '16'; }, cacheKey: 'yamatoPackageSizeId' },
    { id: 'yamatoCodeType', key: StorageManager.KEYS.YAMATO_CODE_TYPE, getValue: function(el) { return el.value || '0'; }, cacheKey: 'yamatoCodeType' },
    { id: 'yamatoDescription', key: StorageManager.KEYS.YAMATO_DESCRIPTION, getValue: function(el) { return (el.value || '').trim(); }, cacheKey: 'yamatoDescription' },
    { id: 'yamatoIncludeOrderNumber', key: StorageManager.KEYS.YAMATO_INCLUDE_ORDER_NUMBER, getValue: function(el) { return !!el.checked; }, cacheKey: 'yamatoIncludeOrderNumber' }
  ];

  definitions.forEach(function(def) {
    const element = document.getElementById(def.id);
    if (!element) return;
    if (def.id === 'yamatoDescription') {
      element.addEventListener('input', async function() {
        const value = def.getValue(element);
        settingsCache[def.cacheKey] = value;
        await StorageManager.set(def.key, value);
      });
      element.addEventListener('change', async function() {
        const value = def.getValue(element);
        settingsCache[def.cacheKey] = value;
        await StorageManager.set(def.key, value);
        const nextHistory = mergeYamatoDescriptionHistory(value, settingsCache.yamatoDescriptionHistory);
        await StorageManager.set(StorageManager.KEYS.YAMATO_DESCRIPTION_HISTORY, nextHistory);
        settingsCache.yamatoDescriptionHistory = nextHistory;
        renderYamatoDescriptionHistory(nextHistory);
      });
      return;
    }
    element.addEventListener('change', async function() {
      const value = def.getValue(element);
      await StorageManager.set(def.key, value);
      settingsCache[def.cacheKey] = value;
    });
  });

  YAMATO_HANDLING_CHECKBOXES.forEach(function(id) {
    const input = document.getElementById(id);
    if (!input) return;
    input.addEventListener('change', async function() {
      let handlingCodes = getYamatoHandlingCodesFromUI();
      if (input.checked && handlingCodes.length > 2) {
        input.checked = false;
        handlingCodes = getYamatoHandlingCodesFromUI();
        alert('荷扱いは2つまで選択できます');
      }
      await StorageManager.set(StorageManager.KEYS.YAMATO_HANDLING_CODES, handlingCodes);
      settingsCache.yamatoHandlingCodes = handlingCodes;
    });
  });

  if (initialSettings && Array.isArray(initialSettings.yamatoHandlingCodes)) {
    settingsCache.yamatoHandlingCodes = initialSettings.yamatoHandlingCodes.slice();
  }
}

// 既存のDOMからラベル部分だけ再生成（CSVデータはDBから復元）
async function regenerateLabelsFromDB() {
  // --- Stage 3: repository ベース未印刷抽出 (第一弾) ---
  try {
    document.querySelectorAll('section.sheet.label-sheet').forEach(sec => sec.remove());
  } catch {}
  // ラベルシート枚数カウンタをリセット（C: 内部カウンタ化）
  window.currentLabelSheetCount = 0;

  const settings = await StorageManager.getSettingsAsync();
  if (!settings.labelyn) return; // ラベル印刷OFFなら終了

  const repo = window.orderRepository || null;
  const displayed = Array.isArray(window.currentDisplayedOrderNumbers) ? window.currentDisplayedOrderNumbers : [];
  let sourceNumbers = displayed;
  if (displayed.length === 0 && repo) {
    // フォールバック: 初期化前などは repository 全件（理論上少数）
    sourceNumbers = repo.getAll().map(r => r.orderNumber);
  }
  const unprintedOrderNumbers = repo
    ? sourceNumbers.filter(n => { const rec = repo.get(n); return rec && !rec.printedAt; })
    : []; // repository 前提。無い場合は空。

  const skip = parseInt(settings.labelskip || '0', 10) || 0;
  // 未印刷 0 件ならラベルシートは表示しない（プレビュー用にも残さない方針）
  if (unprintedOrderNumbers.length === 0) {
    recalcAndUpdateCounts();
    return;
  }
  const labelarr = new Array(skip).fill("");
  for (const num of unprintedOrderNumbers) labelarr.push(num);

  if (settings.labelyn && settings.customLabelEnable && settings.customLabels?.length) {
    for (const cl of settings.customLabels.filter(l => l.enabled)) {
      for (let i = 0; i < cl.count; i++) {
        labelarr.push({ type: 'custom', content: cl.html || cl.text, fontSize: cl.fontSize || '10pt' });
      }
    }
  }

  if (labelarr.length > 0) await generateLabels(labelarr, { skipOnFirstSheet: skip });
}

// 画面上の枚数表示（固定ヘッダー）を再計算して更新
function recalcAndUpdateCounts() {
  const repo = window.orderRepository || null;
  // C: ラベルシート枚数は generateLabels が管理する内部カウンタを利用（UI 依存排除）
  const labelSheetCount = (typeof window.currentLabelSheetCount === 'number')
    ? window.currentLabelSheetCount
    : document.querySelectorAll('section.sheet.label-sheet').length; // 互換フォールバック
  StorageManager.getSettingsAsync().then(settings => {
    const labelSheetsForDisplay = settings.labelyn ? labelSheetCount : 0;
    const customCountForDisplay = (settings.labelyn && settings.customLabelEnable && Array.isArray(settings.customLabels))
      ? settings.customLabels.filter(l => l.enabled).reduce((s, l) => s + (parseInt(l.count, 10) || 0), 0)
      : 0;
    // 表示対象注文番号リスト (processCSVResults で保持) に基づき repository から未印刷数を算出
    let orderSheetCount = 0;
    if (repo && Array.isArray(window.currentDisplayedOrderNumbers)) {
      for (const num of window.currentDisplayedOrderNumbers) {
        const rec = repo.get(num);
        if (rec && !rec.printedAt) orderSheetCount++;
      }
    }
    updatePrintCountDisplay(orderSheetCount, labelSheetsForDisplay, customCountForDisplay);
  });
}

// getOrderSection は scrollToOrderSection 内にインライン化済み（id=order-<番号>）

// 注文明細内に個別注文画像ドロップゾーンを作成
async function createIndividualImageDropZone(cOrder, orderNumber) {
  debugLog(`個別画像ドロップゾーン作成開始 - 注文番号: "${orderNumber}"`);
  
  const individualDropZoneContainer = cOrder.querySelector('.individual-image-dropzone');
  const individualZone = cOrder.querySelector('.individual-order-image-zone');
  
  debugLog(`ドロップゾーンコンテナ発見: ${!!individualDropZoneContainer}`);
  debugLog(`個別ゾーン発見: ${!!individualZone}`);
  
  // 表示/非表示はCSSクラス `order-image-hidden` で body タグレベルで制御する。
  // ここでは常にドロップゾーンを作成する。

  if (individualDropZoneContainer && isValidOrderNumber(orderNumber)) {
    // 注文番号を正規化
  const normalizedOrderNumber = (orderNumber == null) ? '' : String(orderNumber).trim();
    
    try {
      const individualImageDropZone = await createIndividualOrderImageDropZone(normalizedOrderNumber);
      if (individualImageDropZone && individualImageDropZone.element) {
        individualDropZoneContainer.appendChild(individualImageDropZone.element);
        debugLog(`個別画像ドロップゾーン作成成功: ${normalizedOrderNumber}`);
      } else {
        debugLog(`個別画像ドロップゾーン作成失敗: ${normalizedOrderNumber}`);
      }
    } catch (error) {
      debugLog(`個別画像ドロップゾーン作成エラー: ${error.message}`);
      console.error('個別画像ドロップゾーン作成エラー:', error);
    }
  } else {
    debugLog(`個別画像ドロップゾーン作成スキップ - コンテナ: ${!!individualDropZoneContainer}, 注文番号: "${orderNumber}"`);
  }
}

function createObjectUrlFromBinary(data, mimeType = 'image/png') {
  if (!(data instanceof ArrayBuffer) || data.byteLength === 0) return null;
  try {
    return URL.createObjectURL(new Blob([data], { type: mimeType || 'image/png' }));
  } catch (error) {
    console.error('画像Blob URL生成失敗', error);
    return null;
  }
}

function extractProductItemsFromCSVRow(row) {
  if (!row || !row[CONSTANTS.CSV.PRODUCT_COLUMN]) return [];
  const items = [];
  for (const itemrow of String(row[CONSTANTS.CSV.PRODUCT_COLUMN]).split('\n')) {
    const productInfo = parseProductItemData(itemrow);
    if (productInfo.itemId && productInfo.quantity) {
      items.push(productInfo);
    }
  }
  return items;
}

function getOrderImageSourceLabel(resolved) {
  if (!resolved) return '適用画像: なし';
  switch (resolved.sourceType) {
    case 'individual':
      return '適用画像: 個別';
    case 'product':
      return `適用画像: 商品ID ${resolved.productId}`;
    case 'global':
      return '適用画像: 全体';
    default:
      return '適用画像: なし';
  }
}

function updateOrderImageSourceLabel(cOrder, resolved) {
  const labelEl = cOrder.querySelector('.individual-image-source-label');
  if (labelEl) {
    labelEl.textContent = getOrderImageSourceLabel(resolved);
  }
}

function renderOrderImageIntoContainer(container, imageUrl) {
  if (!container) return;
  container.innerHTML = '';
  if (!imageUrl) return;
  const imageDiv = document.createElement('div');
  imageDiv.classList.add('order-image');
  const img = document.createElement('img');
  img.src = imageUrl;
  imageDiv.appendChild(img);
  container.appendChild(imageDiv);
}

async function resolveOrderImageSource(orderNumber, row = null) {
  const normalizedOrderNumber = orderNumber == null ? '' : String(orderNumber).trim();
  const repositoryRecord = normalizedOrderNumber && window.orderRepository ? window.orderRepository.get(normalizedOrderNumber) : null;
  const sourceRow = row || (repositoryRecord ? repositoryRecord.row : null);

  if (repositoryRecord && repositoryRecord.image && repositoryRecord.image.data instanceof ArrayBuffer) {
    return {
      sourceType: 'individual',
      imageUrl: createObjectUrlFromBinary(repositoryRecord.image.data, repositoryRecord.image.mimeType),
      orderNumber: normalizedOrderNumber
    };
  }

  const products = extractProductItemsFromCSVRow(sourceRow);
  for (const product of products) {
    try {
      const productImage = await StorageManager.getProductOrderImage(product.itemId);
      if (productImage) {
        return {
          sourceType: 'product',
          imageUrl: createObjectUrlFromBinary(productImage.data, productImage.mimeType),
          productId: product.itemId,
          productName: product.productName || productImage.productName || ''
        };
      }
    } catch (error) {
      console.error('商品ID画像取得失敗', product.itemId, error);
    }
  }

  const globalImage = await getGlobalOrderImage();
  if (globalImage) {
    return {
      sourceType: 'global',
      imageUrl: globalImage
    };
  }

  return {
    sourceType: 'none',
    imageUrl: null
  };
}

function buildProductImageCandidateMap(orderNumbers) {
  const map = new Map();
  const repo = window.orderRepository;
  if (!repo) return map;
  let sequence = 0;

  for (const orderNumber of orderNumbers) {
    const rec = repo.get(orderNumber);
    if (!rec || !rec.row) continue;
    const seenInOrder = new Set();
    for (const product of extractProductItemsFromCSVRow(rec.row)) {
      const productId = product.itemId;
      if (!productId) continue;
      let entry = map.get(productId);
      if (!entry) {
        entry = {
          productId,
          productName: product.productName || '',
          count: 0,
          sequence: sequence++
        };
        map.set(productId, entry);
      }
      if (!entry.productName && product.productName) {
        entry.productName = product.productName;
      }
      if (!seenInOrder.has(productId)) {
        entry.count += 1;
        seenInOrder.add(productId);
      }
    }
  }
  return map;
}

async function collectProductImageCandidates() {
  const repo = await ensureOrderRepository();
  const sourceNumbers = Array.isArray(window.currentDisplayedOrderNumbers) && window.currentDisplayedOrderNumbers.length > 0
    ? window.currentDisplayedOrderNumbers
    : (repo ? repo.getAll().map(rec => rec.orderNumber) : []);
  const candidateMap = buildProductImageCandidateMap(sourceNumbers);
  const savedImages = await StorageManager.getAllProductOrderImages();

  for (const savedImage of savedImages) {
    if (!savedImage || !savedImage.productId) continue;
    if (!candidateMap.has(savedImage.productId)) {
      candidateMap.set(savedImage.productId, {
        productId: savedImage.productId,
        productName: savedImage.productName || '',
        count: 0,
        sequence: Number.MAX_SAFE_INTEGER
      });
    } else if (!candidateMap.get(savedImage.productId).productName && savedImage.productName) {
      candidateMap.get(savedImage.productId).productName = savedImage.productName;
    }
  }

  return Array.from(candidateMap.values()).sort((a, b) => {
    if (a.sequence !== b.sequence) return a.sequence - b.sequence;
    return a.productId.localeCompare(b.productId, 'ja', { numeric: true, sensitivity: 'base' });
  });
}

async function renderProductOrderImagesManager() {
  const block = document.getElementById('productOrderImagesBlock');
  const container = document.getElementById('productOrderImagesContainer');
  const summary = document.getElementById('productOrderImagesSummary');
  if (!block || !container || !summary) return;

  const enabled = !!settingsCache.orderImageEnable;
  block.hidden = !enabled;
  if (!enabled) return;

  const [candidates, savedImages] = await Promise.all([
    collectProductImageCandidates(),
    StorageManager.getAllProductOrderImages()
  ]);
  const savedMap = new Map(savedImages.map(item => [item.productId, item]));
  summary.textContent = `商品ID画像: ${savedImages.length}件`;

  container.innerHTML = '';
  if (candidates.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'product-order-image-empty';
    empty.textContent = 'CSVを読み込むと商品IDごとの画像設定欄がここに表示されます。';
    container.appendChild(empty);
    return;
  }

  for (const candidate of candidates) {
    const item = document.createElement('div');
    item.className = 'product-order-image-item';

    const header = document.createElement('div');
    header.className = 'product-order-image-item-header';

    const title = document.createElement('div');
    title.className = 'product-order-image-item-title';

    const idLine = document.createElement('div');
    idLine.className = 'product-order-image-item-id';
    idLine.textContent = `商品ID: ${candidate.productId}`;
    title.appendChild(idLine);

    if (candidate.productName) {
      const nameLine = document.createElement('div');
      nameLine.className = 'product-order-image-item-name';
      nameLine.textContent = candidate.productName;
      title.appendChild(nameLine);
    }

    const meta = document.createElement('div');
    meta.className = 'product-order-image-item-meta';
    meta.textContent = candidate.count > 0 ? `現在の注文で ${candidate.count} 件に出現` : '保存済み設定';
    title.appendChild(meta);
    header.appendChild(title);
    item.appendChild(header);

    const saved = savedMap.get(candidate.productId);
    const dropzoneHost = document.createElement('div');
    dropzoneHost.className = 'product-order-image-dropzone';
    const dropzone = await createProductOrderImageDropZone(candidate.productId, candidate.productName || (saved && saved.productName) || '');
    if (dropzone && dropzone.element) {
      dropzoneHost.appendChild(dropzone.element);
    }

    item.appendChild(dropzoneHost);
    container.appendChild(item);
  }
}

// 注文の商品アイテムリストを処理してHTMLに追加
function processProductItems(cOrder, row) {
  const tItems = cOrder.querySelector('#商品');
  const trSpace = cOrder.querySelector('.spacerow');
  
  for (const productInfo of extractProductItemsFromCSVRow(row)) {
    const cItem = document.importNode(tItems.content, true);
    if (productInfo.itemId && productInfo.quantity) {
      setProductItemElements(cItem, productInfo);
      trSpace.parentNode.parentNode.insertBefore(cItem, trSpace.parentNode);
    }
  }
}

// 商品行データを解析して商品情報オブジェクトを返す
function parseProductItemData(itemrow) {
  const firstSplit = itemrow.split(' / ');
  const itemIdSplit = firstSplit[0].split(':');
  const itemId = itemIdSplit.length > 1 ? itemIdSplit[1].trim() : '';
  const quantitySplit = firstSplit[1] ? firstSplit[1].split(':') : [];
  const quantity = quantitySplit.length > 1 ? quantitySplit[1].trim() : '';
  const productName = firstSplit.slice(2).join(' / ');
  
  return { itemId, quantity, productName };
}

// 商品情報をHTML要素に設定
function setProductItemElements(cItem, productInfo) {
  const tdId = cItem.querySelector(".商品ID");
  if (tdId) {
    tdId.textContent = productInfo.itemId;
  }
  const tdQuantity = cItem.querySelector(".数量");
  if (tdQuantity) {
    tdQuantity.textContent = productInfo.quantity;
  }
  const tdName = cItem.querySelector(".商品名");
  if (tdName) {
    tdName.textContent = productInfo.productName;
  }
}

// 注文明細に注文画像を表示
async function displayOrderImage(cOrder, orderNumber, row = null) {
  const resolved = await resolveOrderImageSource(orderNumber, row);
  updateOrderImageSourceLabel(cOrder, resolved);
  renderOrderImageIntoContainer(cOrder.querySelector('.order-image-container'), resolved.imageUrl);
}

// ラベルシートを生成してDOMに追加
async function generateLabels(labelarr, options = {}) {
  const opts = {
    skipOnFirstSheet: 0,
    ...options
  };
  if (!Array.isArray(labelarr) || labelarr.length === 0) return; // 何も生成しない
  // 全てが空スキップ要素("" など falsy)のみなら生成しない（全注文印刷済みで skip 指定だけのケースで空シートが出るのを防止）
  const hasMeaningful = labelarr.some(l => {
    if (!l) return false; // 空文字や null
    if (typeof l === 'string') return l.trim() !== '';
    // オブジェクト（カスタムラベルなど）は有効とみなす
    return true;
  });
  if (!hasMeaningful) return;
  // シートをちょうど埋めるために不足分だけ空ラベルを追加
  if (labelarr.length % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET) {
    const remainder = labelarr.length % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET;
    const toFill = CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET - remainder;
    for (let i = 0; i < toFill; i++) {
      labelarr.push("");
    }
  }
  
  const tL44 = document.querySelector('#L44');
  let cL44 = document.importNode(tL44.content, true);
  // 生成するラベルシートに識別クラスを付与
  cL44.querySelector('section.sheet')?.classList.add('label-sheet');
  let tableL44 = cL44.querySelector("table");
  let tr = document.createElement("tr");
  let i = 0; // 全体インデックス
  let sheetIndex = 0;
  let posInSheet = 0; // 0..43
  // C: 生成開始時にカウンタ初期化（既存を上書き）
  if (typeof window.currentLabelSheetCount !== 'number') window.currentLabelSheetCount = 0;
  
  for (let label of labelarr) {
    if (i > 0 && i % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET == 0) {
      tableL44.appendChild(tr);
      tr = document.createElement("tr");
  document.body.insertBefore(cL44, tL44);
  window.currentLabelSheetCount++;
      cL44 = document.importNode(tL44.content, true);
      cL44.querySelector('section.sheet')?.classList.add('label-sheet');
      tableL44 = cL44.querySelector("table");
      tr = document.createElement("tr");
      sheetIndex++;
      posInSheet = 0;
    } else if (i > 0 && i % CONSTANTS.LABEL.LABELS_PER_ROW == 0) {
      tableL44.appendChild(tr);
      tr = document.createElement("tr");
    }
    const td = await createLabel(label);
    // スキップ面の視覚表示（初回シートの先頭skip数のみ）
    if (sheetIndex === 0 && posInSheet < (opts.skipOnFirstSheet || 0)) {
      td.classList.add('skip-face');
      td.setAttribute('data-label-index', String(posInSheet + 1));
      const indicator = document.createElement('div');
      indicator.className = 'skip-indicator';
      indicator.textContent = String(posInSheet + 1);
      td.appendChild(indicator);
    }
    tr.appendChild(td);
    posInSheet++;
    i++;
  }
  tableL44.appendChild(tr);
  document.body.insertBefore(cL44, tL44);
  window.currentLabelSheetCount++;
}

// div要素にpタグを追加するヘルパー関数
function addP(div, text){
  const p = document.createElement("p");
  p.innerText = text;
  div.appendChild(p);
}

// div要素を作成するヘルパー関数
function createDiv(classname="", text=""){
  const div = document.createElement('div');
  if(classname){
    div.classList.add(classname);
  }
  if(text){
    addP(div,text);
  }
  return div;
}
// 汎用 template クローンヘルパー
function cloneTemplate(id) {
  const tpl = document.getElementById(id);
  if (!tpl || !tpl.content || !tpl.content.firstElementChild) {
    throw new Error(id + ' template が見つかりません');
  }
  return tpl.content.firstElementChild.cloneNode(true);
}
// QRペーストプレースホルダを template から生成（プロトコルに応じて適切なテンプレートを選択）
function buildQRPastePlaceholder() {
  if (getAppMode() === 'extension') {
    return cloneTemplate('qrFetchPlaceholderExtension');
  }
  const templateId = PROTOCOL_UTILS.canDragFromExternalSites() ? 'qrDropPlaceholderHttp' : 'qrDropPlaceholder';
  return cloneTemplate(templateId);
}

function setExtensionQrFetchState(dropzone, isBusy) {
  if (!dropzone) return;
  const button = dropzone.querySelector('.qr-fetch-placeholder-button');
  if (!button) return;
  button.textContent = isBusy ? '取得中...' : 'QR取得';
  button.disabled = !!isBusy;
  dropzone.classList.toggle('is-busy', !!isBusy);
}

function setupExtensionQrFetchEvents(dropzone) {
  dropzone.addEventListener('click', async function(event) {
    if (event.target.tagName === 'IMG') return;

    const td = dropzone.closest('td');
    const orderNumber = getOrderNumberFromLabelCell(td);
    if (!orderNumber) {
      alert('注文番号を特定できないため、QR取得を実行できません');
      return;
    }

    const bridge = window.BoothCSVExtensionBridge;
    if (!bridge || typeof bridge.collectOrderQRCode !== 'function') {
      alert('Chrome拡張のQR取得機能を利用できません');
      return;
    }

    setExtensionQrFetchState(dropzone, true);
    try {
      await bridge.collectOrderQRCode(orderNumber);
    } catch (error) {
      alert(error && error.message ? error.message : 'QR取得に失敗しました');
      setExtensionQrFetchState(dropzone, false);
    }
  });
}

// ペーストゾーンにクリップボード画像ペーストイベントを設定
function setupPasteZoneEvents(dropzone) {
  // クリップボードからの画像ペーストを受け付ける
  dropzone.addEventListener("paste", function (event) {
    try {
      event.preventDefault();
      const cd = event.clipboardData;
      if (!cd) return;

      // ファイルが含まれる場合はファイル優先で処理
      if (cd.files && cd.files.length > 0) {
        for (const file of cd.files) {
          if (!file.type || !file.type.match(/^image\/((jpeg|png)|svg\+xml)$/)) continue;
          const fr = new FileReader();
          fr.onload = function (e) {
            const elImage = document.createElement('img');
            elImage.src = e.target.result;
            elImage.onload = function () {
              dropzone.parentNode.appendChild(elImage);
              readQR(elImage);
              dropzone.classList.remove('dropzone');
              dropzone.style.zIndex = -1;
              elImage.style.zIndex = 9;
              addEventQrReset(elImage);
            };
          };
          fr.readAsDataURL(file);
        }
      } else if (cd.types && (cd.types.includes('text/plain') || cd.types.includes('text/uri-list'))) {
        // クリップボードに URL がある場合（外部画像のURLを貼り付けたケースなど）
        const text = cd.getData('text/uri-list') || cd.getData('text/plain');
        if (text) handleExternalImageUrl(text.trim());
      }
    } catch (e) {
      console.error('paste handler error', e);
    } finally {
      // プレースホルダを復元
      dropzone.innerHTML = '';
      dropzone.appendChild(buildQRPastePlaceholder());
    }
  });

  // ドラッグ操作の見た目制御
  dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
  dropzone.addEventListener('dragleave', () => { dropzone.classList.remove('dragover'); });

  // ドロップサポート: ファイル or URL
  dropzone.addEventListener('drop', function (e) {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    const dt = e.dataTransfer;
    if (!dt) return;

    // ファイルがあれば優先して処理
    if (dt.files && dt.files.length > 0) {
      const file = dt.files[0];
      if (file && file.type && file.type.match(/^image\/((jpeg|png)|svg\+xml)$/)) {
        const reader = new FileReader();
        reader.onload = async (ev) => {
          const dataUrl = ev.target.result;
          const elImage = document.createElement('img');
          elImage.src = dataUrl;
          elImage.onload = function () {
            dropzone.parentNode.appendChild(elImage);
            readQR(elImage);
            dropzone.classList.remove('dropzone');
            dropzone.style.zIndex = -1;
            elImage.style.zIndex = 9;
            addEventQrReset(elImage);
          };
        };
        reader.readAsDataURL(file);
        return;
      }
    }

    // URL ドロップの処理 (ブラウザ間ドラッグで URL が来る場合)
    const url = (dt.getData && (dt.getData('text/uri-list') || dt.getData('text/plain'))) || null;
    if (url) {
      handleExternalImageUrl(url.trim());
    }
  });

  function handleExternalImageUrl(rawUrl) {
    if (!rawUrl) return;
    // data: スキームは直接使用
    if (rawUrl.startsWith('data:')) {
      const elImage = document.createElement('img');
      elImage.src = rawUrl;
      elImage.onload = function () {
        dropzone.parentNode.appendChild(elImage);
        readQR(elImage);
        dropzone.classList.remove('dropzone');
        dropzone.style.zIndex = -1;
        elImage.style.zIndex = 9;
        addEventQrReset(elImage);
      };
      return;
    }

    // HTTP/HTTPS はサーバープロキシ経由で取得（CORS回避）
    if (/^https?:\/\//i.test(rawUrl)) {
      try {
        const proxied = '/proxy?url=' + encodeURIComponent(rawUrl);
        const elImage = document.createElement('img');
        elImage.crossOrigin = 'anonymous';
        elImage.src = proxied;
        elImage.onload = function () {
          dropzone.parentNode.appendChild(elImage);
          readQR(elImage);
          dropzone.classList.remove('dropzone');
          dropzone.style.zIndex = -1;
          elImage.style.zIndex = 9;
          addEventQrReset(elImage);
        };
        elImage.onerror = function () { console.error('外部画像の読み込みに失敗しました', rawUrl); };
      } catch (e) {
        console.error('外部画像処理エラー', e);
      }
    }
  }
}

function createDropzone(div){ // 互換のため名称維持（内部はペースト専用）
  const divDrop = createDiv('dropzone');
  divDrop.appendChild(buildQRPastePlaceholder());
  if (getAppMode() === 'extension') {
    divDrop.classList.add('extension-qr-fetch-mode');
    setupExtensionQrFetchEvents(divDrop);
  } else {
    divDrop.setAttribute("contenteditable", "true");
    // ペースト専用イベントを設定
    setupPasteZoneEvents(divDrop);
  }
  
  div.appendChild(divDrop);
}

// ラベル要素を作成（注文番号またはカスタムラベル）
async function createLabel(labelData=""){
  // カスタムラベル判定を先に
  if (typeof labelData === 'object' && labelData?.type === 'custom') {
    const base = cloneTemplate('customLabelCell');
    // base は <td class="qrlabel custom-mode"> ...
    const td = base.matches('td') ? base : base.querySelector('td.qrlabel');
    const contentWrap = td.querySelector('.custom-content');
    contentWrap.innerHTML = labelData.content || '';
    if (labelData.fontSize) contentWrap.style.fontSize = labelData.fontSize;
    return td;
  }

  // 通常ラベル（注文番号 or 空）
  const base = cloneTemplate('labelCell');
  const tdLabel = base.matches('td') ? base : base.querySelector('td.qrlabel');
  const divQr = tdLabel.querySelector('.qr');
  const divOrdernum = tdLabel.querySelector('.ordernum');
  const divYamato = tdLabel.querySelector('.yamato');

  if (typeof labelData === 'string' && labelData) {
    // 注文番号をリンク化
    const orderLink = document.createElement('a');
    orderLink.href = `https://manage.booth.pm/orders/${labelData}`;
    orderLink.textContent = labelData;
    orderLink.target = '_blank';
    orderLink.rel = 'noopener';
    divOrdernum.appendChild(orderLink);

    const repo = window.orderRepository || null;
    const rec = repo ? repo.get(labelData) : null;
    const qr = rec ? rec.qr : null;
    if (qr && qr['qrimage'] != null) {
      const elImage = document.createElement('img');
      // 永続化後は ArrayBuffer が入る想定。isBinary フラグに依存せず型で判定する
      try {
        if (qr.qrimage instanceof ArrayBuffer) {
          const blob = new Blob([qr.qrimage], { type: qr.qrimageType || 'image/png' });
          elImage.src = URL.createObjectURL(blob);
          elImage.addEventListener('error', () => console.error('QR画像Blob URL読み込み失敗'));
        } else if (typeof qr.qrimage === 'string') {
          // 互換: dataURL などの文字列
          elImage.src = qr.qrimage;
        } else {
          console.warn('未知のQR画像型、ドロップゾーンにフォールバック');
          createDropzone(divQr);
          return tdLabel;
        }
      } catch (e) {
        console.error('QR画像生成失敗', e);
        createDropzone(divQr);
        return tdLabel;
      }
      divQr.appendChild(elImage);
      addP(divYamato, qr['receiptnum']);
      addP(divYamato, qr['receiptpassword']);
      addEventQrReset(elImage);
    } else {
      createDropzone(divQr);
    }
  }
  // 空文字 / falsy の場合はパディングセル: 何も入れない
  return tdLabel;
}

function computeQRHash(content) {
  let hash = 0;
  if (!content) return hash.toString();
  for (let i = 0; i < content.length; i++) {
    const ch = content.charCodeAt(i);
    hash = ((hash << 5) - hash) + ch;
    hash &= hash;
  }
  return hash.toString();
}

async function rasterizeQRCodeImageSource(imageSource) {
  return await new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = async function() {
      try {
        const canv = document.createElement('canvas');
        const context = canv.getContext('2d');
        canv.width = img.width;
        canv.height = img.height;
        context.drawImage(img, 0, 0, canv.width, canv.height);

        const blob = await new Promise(resolveBlob => canv.toBlob(resolveBlob, 'image/png'));
        let arrayBuffer = null;
        if (blob) {
          arrayBuffer = await blob.arrayBuffer();
        } else {
          const tmpB64 = canv.toDataURL('image/png');
          const bin = atob(tmpB64.split(',')[1]);
          const len = bin.length;
          const u8 = new Uint8Array(len);
          for (let i = 0; i < len; i++) {
            u8[i] = bin.charCodeAt(i);
          }
          arrayBuffer = u8.buffer;
        }

        resolve({
          arrayBuffer,
          imageData: context.getImageData(0, 0, canv.width, canv.height),
          width: canv.width,
          height: canv.height,
          qrimageType: 'image/png'
        });
      } catch (error) {
        reject(error);
      }
    };

    img.onerror = function() {
      reject(new Error('画像の読み込みに失敗しました'));
    };

    img.src = imageSource instanceof HTMLImageElement ? imageSource.src : imageSource;
  });
}

async function decodeQRCodeImageSource(imageSource) {
  const rasterized = await rasterizeQRCodeImageSource(imageSource);
  const barcode = jsQR(rasterized.imageData.data, rasterized.imageData.width, rasterized.imageData.height);
  if (!barcode) {
    throw new Error('QRコードを読み取れませんでした');
  }

  const parts = String(barcode.data).replace(/^\s+|\s+$/g, '').replace(/ +/g, ' ').split(' ');
  if (parts.length !== CONSTANTS.QR.EXPECTED_PARTS) {
    throw new Error('QRコードの形式が正しくありません');
  }

  return {
    barcodeData: barcode.data,
    receiptnum: parts[1],
    receiptpassword: parts[2],
    qrimage: rasterized.arrayBuffer,
    qrimageType: rasterized.qrimageType,
    qrhash: computeQRHash(barcode.data)
  };
}

async function buildQRCodeDataFromMetadata(imageSource, options = {}) {
  const rasterized = await rasterizeQRCodeImageSource(imageSource);
  const receiptnum = options.receiptnum ? String(options.receiptnum).trim() : '';
  const receiptpassword = options.receiptpassword ? String(options.receiptpassword).trim() : '';
  if (!receiptnum || !receiptpassword) {
    throw new Error('受付番号またはパスワードが不足しています');
  }

  const barcodeData = options.barcodeData
    ? String(options.barcodeData)
    : `issued ${receiptnum} ${receiptpassword}`;

  return {
    barcodeData,
    receiptnum,
    receiptpassword,
    qrimage: rasterized.arrayBuffer,
    qrimageType: rasterized.qrimageType,
    qrhash: computeQRHash(barcodeData)
  };
}

function updateLabelCellWithQR(orderNumber, qrData) {
  const normalized = OrderRepository.normalize(orderNumber);
  const orderLink = Array.from(document.querySelectorAll('.ordernum a')).find(link => OrderRepository.normalize(link.textContent) === normalized);
  if (!orderLink) return false;

  const td = orderLink.closest('td');
  if (!td) return false;

  const qrDiv = td.querySelector('.qr');
  const yamatoDiv = td.querySelector('.yamato');
  if (!qrDiv || !yamatoDiv) return false;

  qrDiv.innerHTML = '';
  yamatoDiv.innerHTML = '';

  const image = document.createElement('img');
  const blob = new Blob([qrData.qrimage], { type: qrData.qrimageType || 'image/png' });
  image.src = URL.createObjectURL(blob);
  qrDiv.appendChild(image);
  addP(yamatoDiv, qrData.receiptnum);
  addP(yamatoDiv, qrData.receiptpassword);
  addEventQrReset(image);
  return true;
}

async function saveQRCodeForOrder(orderNumber, imageSource, options = {}) {
  const normalized = OrderRepository.normalize(orderNumber);
  if (!normalized) throw new Error('注文番号が不正です');

  let qrData;
  try {
    qrData = await decodeQRCodeImageSource(imageSource);
  } catch (error) {
    if (!options.receiptnum || !options.receiptpassword) {
      throw error;
    }
    qrData = await buildQRCodeDataFromMetadata(imageSource, options);
  }
  const duplicates = window.orderRepository
    ? await window.orderRepository.checkQRDuplicate(qrData.barcodeData, normalized)
    : [];

  if (duplicates.length > 0 && !options.allowDuplicate) {
    throw new Error(`既に他の注文で使われているQRコードです: ${duplicates.join(', ')}`);
  }

  if (window.orderRepository) {
    await window.orderRepository.setOrderQRData(normalized, qrData);
  }
  updateLabelCellWithQR(normalized, qrData);
  return {
    orderNumber: normalized,
    receiptnum: qrData.receiptnum,
    receiptpassword: qrData.receiptpassword,
    duplicates
  };
}

function getPendingQrOrderNumbers() {
  const sourceNumbers = Array.isArray(window.currentDisplayedOrderNumbers) && window.currentDisplayedOrderNumbers.length > 0
    ? window.currentDisplayedOrderNumbers
    : (window.lastOrderSelection || []);
  const normalized = normalizeOrderSelection(sourceNumbers);
  if (!window.orderRepository) return normalized;
  return normalized.filter(orderNumber => {
    const record = window.orderRepository.get(orderNumber);
    return !!record && !record.qr;
  });
}

function getOrderNumberFromLabelCell(td) {
  if (!td) return null;
  const orderNumberContainer = td.querySelector('.ordernum');
  if (!orderNumberContainer) return null;

  const orderLink = orderNumberContainer.querySelector('a');
  if (orderLink && orderLink.textContent) {
    return String(orderLink.textContent).trim();
  }

  const orderParagraph = orderNumberContainer.querySelector('p');
  if (orderParagraph && orderParagraph.textContent) {
    return String(orderParagraph.textContent).trim();
  }

  const text = orderNumberContainer.textContent;
  return text ? String(text).trim() : null;
}

// QRコード画像にクリックでリセット機能を追加
function addEventQrReset(elImage){
    elImage.addEventListener('click', async function(event) {
      event.preventDefault();
      
      // 親要素のQRセクションを取得
      const qrDiv = elImage.parentNode;
      const td = qrDiv.closest('td');
      
      if (td) {
        // 注文番号を取得
  const orderNumber = getOrderNumberFromLabelCell(td);
        
        // 保存されたQRデータを削除
        if (orderNumber) {
          try {
            if (window.orderRepository) {
              await window.orderRepository.clearOrderQRData(orderNumber);
            }
            console.log(`QRデータを削除しました: ${orderNumber}`);
            updateProcessedOrdersVisibility();
          } catch (error) {
            console.error('QRデータ削除エラー:', error);
          }
        }
        
        // ヤマト運輸情報をクリア
        const yamatoDiv = td.querySelector('.yamato');
        if (yamatoDiv) {
          yamatoDiv.innerHTML = '';
        }
        
        // QR画像を削除
        elImage.remove();
        
        // ドロップゾーンを復元
        qrDiv.innerHTML = '';
        createDropzone(qrDiv);
      }
    });
}

// 画像からQRコードを読み取り、伝票データを抽出
async function readQR(elImage){
  try {
    const orderNumElem = elImage.closest("td").querySelector(".ordernum a");
    const rawOrderNum = orderNumElem ? orderNumElem.textContent : null;
    const ordernum = (rawOrderNum == null) ? '' : String(rawOrderNum).trim();
    const decoded = await decodeQRCodeImageSource(elImage);

    const duplicates = window.orderRepository ? await window.orderRepository.checkQRDuplicate(decoded.barcodeData, ordernum) : [];
    if (duplicates.length > 0) {
      const duplicateList = duplicates.join(', ');
      const confirmMessage = `警告: このQRコードは既に以下の注文で使用されています:\n${duplicateList}\n\n同じQRコードを使用すると配送ミスの原因となる可能性があります。\n続行しますか？`;

      if (!confirm(confirmMessage)) {
        console.log('QRコード登録がキャンセルされました');
        const parentQr = elImage.parentNode;
        const dropzone = parentQr.querySelector('.dropzone, div[class="dropzone"]');

        if (dropzone) {
          dropzone.classList.add('dropzone');
          dropzone.style.zIndex = '99';
          dropzone.innerHTML = '';
          dropzone.appendChild(buildQRPastePlaceholder());
          dropzone.style.display = 'block';
        } else {
          const newDropzone = document.createElement('div');
          newDropzone.className = 'dropzone';
          newDropzone.contentEditable = 'true';
          newDropzone.setAttribute('effectallowed', 'move');
          newDropzone.style.zIndex = '99';
          newDropzone.innerHTML = '';
          newDropzone.appendChild(buildQRPastePlaceholder());
          parentQr.appendChild(newDropzone);
          setupPasteZoneEvents(newDropzone);
        }

        elImage.parentNode.removeChild(elImage);
        return;
      }
    }

    await saveQRCodeForOrder(ordernum, elImage, { allowDuplicate: true });
  } catch (error) {
    console.error('QR読み取り関数エラー:', error);
    alert(error && error.message ? error.message : 'QRコードの処理に失敗しました');
  }
}

window.BoothCSVExtensionBridge = Object.assign(window.BoothCSVExtensionBridge || {}, {
  importCSVText,
  importOrderQRCodeImage: saveQRCodeForOrder,
  getPendingQrOrderNumbers,
  getCurrentOrderNumbers: () => normalizeOrderSelection(window.currentDisplayedOrderNumbers || window.lastOrderSelection || []),
  getValidatedYamatoIssueSettings: () => validateYamatoIssueSettings(getCurrentYamatoIssueSettingsFromUI()),
  prepareYamatoIssueSettings: async () => {
    const validated = validateYamatoIssueSettings(getCurrentYamatoIssueSettingsFromUI());
    return await persistYamatoIssueSettings(validated, { updateHistory: true });
  },
  getYamatoIssueSettings: () => ({
    packageSizeId: getCurrentYamatoIssueSettingsFromUI().packageSizeId || '16',
    codeType: getCurrentYamatoIssueSettingsFromUI().codeType || '0',
    description: getCurrentYamatoIssueSettingsFromUI().description || '',
    includeOrderNumber: getCurrentYamatoIssueSettingsFromUI().includeOrderNumber !== false,
    handlingCodes: getCurrentYamatoIssueSettingsFromUI().handlingCodes
  })
});

// QRコード用ペーストゾーンのみドラッグ&ドロップを抑止し、他領域（注文画像/フォント等）は従来どおり許可
// ...existing code...

// 設定管理
const CONFIG = {
  SUPPORTED_IMAGE_TYPES: ['image/jpeg', 'image/png', 'image/svg+xml'],
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
  
  isImageFile(file) {
    return file && this.SUPPORTED_IMAGE_TYPES.includes(file.type);
  },
  
  isValidFileSize(file) {
    return file && file.size <= this.MAX_FILE_SIZE;
  },
  
  validateImageFile(file) {
    if (!file) {
      throw new Error('ファイルが選択されていません');
    }
    
    if (!this.isImageFile(file)) {
      throw new Error('サポートされていないファイル形式です（JPEG、PNG、SVGのみ）');
    }
    
    if (!this.isValidFileSize(file)) {
      throw new Error('ファイルサイズが大きすぎます（10MB以下）');
    }
    
    return true;
  }
};

// v6: グローバル注文画像を settings (IndexedDB) にバイナリ保存するユーティリティ
async function setGlobalOrderImage(arrayBuffer, mimeType='image/png') {
  try {
    if(!(arrayBuffer instanceof ArrayBuffer)) throw new Error('ArrayBuffer 以外');
    await StorageManager.setGlobalOrderImageBinary(arrayBuffer, mimeType);
  debugLog('[image] グローバル画像保存完了 size=' + arrayBuffer.byteLength + ' mime=' + mimeType);
  } catch(e){ console.error('グローバル画像保存失敗', e); }
}

// グローバル注文画像をIndexedDBから取得してURL生成
async function getGlobalOrderImage(){
  try {
  const v = await StorageManager.getGlobalOrderImageBinary();
  if(!v || !(v.data instanceof ArrayBuffer) || v.data.byteLength===0) return null;
    const blob = new Blob([v.data], { type: v.mimeType || 'image/png' });
    return URL.createObjectURL(blob);
  } catch(e){ console.error('グローバル画像取得失敗', e); return null; }
}

// 共通のドラッグ&ドロップ機能を提供するベース関数
async function createBaseImageDropZone(options = {}) {
  const {
    isIndividual = false,
    orderNumber = null,
    productId = null,
    productName = '',
    containerClass = 'order-image-drop',
    defaultMessage = '画像をドロップ or クリックで選択'
  } = options;

  debugLog(`画像ドロップゾーン作成: 個別=${isIndividual} 注文番号=${orderNumber||'-'} 商品ID=${productId||'-'}`);

  const dropZone = document.createElement('div');
  dropZone.classList.add(containerClass);
  
  if (isIndividual) {
    dropZone.style.cssText = 'min-height: 80px; border: 1px dashed #999; padding: 5px; background: #f9f9f9; cursor: pointer;';
  }

  let droppedImage = null;           // 表示用URL (Blob URL)
  let droppedImageBuffer = null;     // 保存用 ArrayBuffer

  // 保存された画像
  let savedImage = null;
  if (orderNumber && window.orderRepository) {
    const rec = window.orderRepository.get(orderNumber);
    if (rec && rec.image && rec.image.data instanceof ArrayBuffer) {
      try { const blob = new Blob([rec.image.data], { type: rec.image.mimeType || 'image/png' }); savedImage = URL.createObjectURL(blob); } catch(e){ console.error('保存画像Blob生成失敗', e); }
    }
  }
  if (productId && !savedImage) {
    try {
      const rec = await StorageManager.getProductOrderImage(productId);
      if (rec && rec.data instanceof ArrayBuffer) {
        savedImage = createObjectUrlFromBinary(rec.data, rec.mimeType);
      }
    } catch (e) {
      console.error('商品ID画像復元失敗', e);
    }
  }
  // グローバル（非個別）時は settings から復元
  if (!orderNumber && !productId && !isIndividual && !savedImage) {
    try { savedImage = await getGlobalOrderImage(); if (savedImage) debugLog('[image] 初期グローバル画像復元'); } catch(e){ console.error('初期グローバル画像取得失敗', e); }
  }
  if (savedImage) {
  debugLog('保存された画像を復元');
    let restoredUrl = savedImage;
    if (savedImage instanceof ArrayBuffer) {
      try {
        const blob = new Blob([savedImage], { type: 'image/png' }); // 型情報は未保存のためデフォルトPNG扱い
        restoredUrl = URL.createObjectURL(blob);
      } catch(e){ console.error('保存画像復元失敗', e); }
    }
    await updatePreview(restoredUrl, null); // ArrayBuffer を URL 化済み
  } else {
    if (isIndividual) {
  dropZone.innerHTML = '';
  const templateId = PROTOCOL_UTILS.canDragFromExternalSites() ? 'orderImageDropDefault' : 'orderImageDropDefaultFile';
  const node = cloneTemplate(templateId);
      node.textContent = defaultMessage; // 個別はメッセージ差替
      dropZone.appendChild(node);
    } else {
  dropZone.innerHTML = '';
      const templateId = PROTOCOL_UTILS.canDragFromExternalSites() ? 'orderImageDropDefault' : 'orderImageDropDefaultFile';
  dropZone.appendChild(cloneTemplate(templateId));
    }
    debugLog(`初期メッセージを設定: ${isIndividual ? defaultMessage : 'デフォルトコンテンツ'} プロトコル: ${window.location.protocol}`);
  }

  // 全ての注文明細の画像を更新する関数
  async function updateAllOrderImages() {
    // 注文画像表示機能が無効の場合は何もしない
    const settings = await StorageManager.getSettingsAsync();
       if (!settings.orderImageEnable) {
      return;
    }

    const allOrderSections = document.querySelectorAll('section.sheet');
    for (const orderSection of allOrderSections) {
      const imageContainer = orderSection.querySelector('.order-image-container');
      if (!imageContainer) continue;

      // 統一化された方法で注文番号を取得
      const orderNumber = (orderSection.id && orderSection.id.startsWith('order-')) ? orderSection.id.substring(6) : '';

      // 個別画像があるかチェック（個別画像を最優先）
      let imageToShow = null;
      if (orderNumber) {
        const r = window.orderRepository ? window.orderRepository.get(orderNumber) : null;
        const resolved = await resolveOrderImageSource(orderNumber, r ? r.row : null);
        imageToShow = resolved.imageUrl;
        updateOrderImageSourceLabel(orderSection, resolved);
      } else {
        // 注文番号がない場合はグローバル画像を使用
  const globalImage = await getGlobalOrderImage();
        debugLog('注文番号なし、グローバル画像を使用', globalImage ? 'あり' : 'なし');
        imageToShow = globalImage;
      }

      // 画像コンテナを更新
      renderOrderImageIntoContainer(imageContainer, imageToShow);
    }
  }

  async function updatePreview(imageUrl, arrayBuffer, mimeType = null) {
    // 既存の Blob URL を解放
    if (droppedImage && typeof droppedImage === 'string' && droppedImage.startsWith('blob:') && droppedImage !== imageUrl) {
      try { URL.revokeObjectURL(droppedImage); } catch {}
    }
    droppedImage = imageUrl;
    if (arrayBuffer instanceof ArrayBuffer) {
      droppedImageBuffer = arrayBuffer;
      if (isIndividual && orderNumber && window.orderRepository) {
        // 個別注文画像保存
        try { await window.orderRepository.setOrderImage(orderNumber, { data: arrayBuffer, mimeType }); }
        catch(e){ console.error('画像保存失敗', e); }
      } else if (productId) {
        try { await StorageManager.setProductOrderImage(productId, { data: arrayBuffer, mimeType, productName }); }
        catch(e){ console.error('商品ID画像保存失敗', e); }
      } else if (!isIndividual) {
        // グローバル画像保存
  debugLog('[image] グローバル画像保存開始 isIndividual=' + isIndividual + ' orderNumber=' + orderNumber + ' size=' + arrayBuffer.byteLength + ' mime=' + mimeType);
        try { await setGlobalOrderImage(arrayBuffer, mimeType||'image/png'); }
        catch(e){ console.error('グローバル画像保存失敗', e); }
      }
    } else if (arrayBuffer === null) {
      // URLのみ（復元時）: 保存操作は不要
    }
    dropZone.innerHTML = '';
    const preview = document.createElement('img');
    preview.src = imageUrl;
    
    if (isIndividual) {
      preview.style.cssText = 'max-width: 100%; max-height: 60px; cursor: pointer;';
    } else {
      preview.classList.add('preview-image');
    }
    
    preview.title = 'クリックでリセット';
    dropZone.appendChild(preview);
  // setOrderImage は上で ArrayBuffer の場合のみ呼んでいる

    // 個別画像の場合は即座に表示を更新
    if (isIndividual && orderNumber) {
      await updateOrderImageDisplay();
    } else if (productId && arrayBuffer instanceof ArrayBuffer) {
      await updateAllOrderImages();
      await renderProductOrderImagesManager();
    } else if (!isIndividual) {
      // グローバル画像の場合は全ての注文明細の画像を更新
      await updateAllOrderImages();
    }

    // 画像クリックでリセット
    preview.addEventListener('click', async (e) => {
      e.stopPropagation();
      if (isIndividual && orderNumber && window.orderRepository) {
        try { await window.orderRepository.clearOrderImage(orderNumber); } catch(e){ console.error('画像削除失敗', e); }
      } else if (productId) {
        try { await StorageManager.clearProductOrderImage(productId); } catch(e){ console.error('商品ID画像削除失敗', e); }
      } else if (!isIndividual) {
        // グローバル画像クリア（0バイトは保存せず null を保存）
        try { await StorageManager.clearGlobalOrderImageBinary(); } catch(e){ console.error('グローバル画像削除失敗', e); }
      }
      droppedImage = null;
      // Blob URL であれば解放
      if (preview.src && preview.src.startsWith('blob:')) {
        try { URL.revokeObjectURL(preview.src); } catch {}
      }
      
      if (isIndividual) {
  dropZone.innerHTML = '';
  const node = cloneTemplate('orderImageDropDefault');
        node.textContent = defaultMessage;
        dropZone.appendChild(node);
      await updateOrderImageDisplay();
        } else if (productId) {
      dropZone.innerHTML = '';
      const node = cloneTemplate('orderImageDropDefault');
        node.textContent = defaultMessage;
        dropZone.appendChild(node);
        await updateAllOrderImages();
        await renderProductOrderImagesManager();
      } else {
  dropZone.innerHTML = '';
  dropZone.appendChild(cloneTemplate('orderImageDropDefault'));
  await updateAllOrderImages();
      }
    });
  }

  // 個別画像用の表示更新関数
  async function updateOrderImageDisplay() {
    // 注文画像表示機能が無効の場合は何もしない
    const settings = await StorageManager.getSettingsAsync();
    if (!settings.orderImageEnable) {
      return;
    }

    const orderSection = dropZone.closest('section');
    if (!orderSection) return;

    const imageContainer = orderSection.querySelector('.order-image-container');
    if (!imageContainer) return;
    const sectionOrderNumber = (orderSection.id && orderSection.id.startsWith('order-')) ? orderSection.id.substring(6) : orderNumber;
    const rec = window.orderRepository && sectionOrderNumber ? window.orderRepository.get(sectionOrderNumber) : null;
    const resolved = await resolveOrderImageSource(sectionOrderNumber, rec ? rec.row : null);
    updateOrderImageSourceLabel(orderSection, resolved);
    renderOrderImageIntoContainer(imageContainer, resolved.imageUrl);
  }

  // 共通のイベントリスナー設定
  setupDragAndDropEvents(dropZone, updatePreview, isIndividual);
  setupClickEvent(dropZone, updatePreview, () => droppedImage);

  // メソッドを持つオブジェクトを返す
  return {
    element: dropZone,
    getImage: () => droppedImage,
    setImage: (imageData, buffer) => { droppedImage = imageData; if (buffer instanceof ArrayBuffer) droppedImageBuffer = buffer; updatePreview(imageData, buffer || null); }
  };
}

// ドラッグ&ドロップイベントの共通設定
function setupDragAndDropEvents(dropZone, updatePreview, isIndividual) {
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    if (isIndividual) {
      dropZone.style.backgroundColor = '#e6f3ff';
    } else {
      dropZone.classList.add('dragover');
    }
  });

  dropZone.addEventListener('dragleave', () => {
    if (isIndividual) {
      dropZone.style.backgroundColor = '#f9f9f9';
    } else {
      dropZone.classList.remove('dragover');
    }
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    if (isIndividual) {
      dropZone.style.backgroundColor = '#f9f9f9';
    } else {
      dropZone.classList.remove('dragover');
    }

    const file = e.dataTransfer.files[0];
    if (file && file.type.match(/^image\/(jpeg|png|svg\+xml)$/)) {
      const reader = new FileReader();
      reader.onload = async (ev) => {
        const buf = ev.target.result;
        try {
              const blob = new Blob([buf], { type: file.type });
              const url = URL.createObjectURL(blob);
              await updatePreview(url, buf, file.type);
        } catch(err){ console.error('ドロップ画像処理失敗', err); }
      };
      reader.readAsArrayBuffer(file);
    }
  });
}

// クリックイベントの共通設定
function setupClickEvent(dropZone, updatePreview, getDroppedImage) {
  dropZone.addEventListener('click', (e) => {
    // 画像が表示されている場合はクリックイベントを無視（画像のクリックリセット用）
    if (e.target.tagName === 'IMG') {
      return;
    }
    
    const currentImage = getDroppedImage();
    debugLog(`ドロップゾーンクリック - 現在の画像: ${currentImage ? 'あり' : 'なし'}`);
    
    if (!currentImage) {
      const fileInput = document.createElement('input');
      fileInput.type = 'file';
      fileInput.accept = 'image/jpeg, image/png, image/svg+xml';
      fileInput.style.display = 'none';
      document.body.appendChild(fileInput);
      
      debugLog('ファイル選択ダイアログを表示');
      fileInput.click();

      fileInput.addEventListener('change', () => {
        const file = fileInput.files[0];
        debugLog(`ファイル選択: ${file ? file.name : 'なし'}`);
        
        if (file && file.type.match(/^image\/(jpeg|png|svg\+xml)$/)) {
          const reader = new FileReader();
          reader.onload = async (ev) => {
            const buf = ev.target.result;
            debugLog(`画像読み込み完了 - サイズ: ${buf.byteLength} bytes`);
            try {
              const blob = new Blob([buf], { type: file.type });
              const url = URL.createObjectURL(blob);
              await updatePreview(url, buf, file.type);
            } catch(err){ console.error('画像読込処理失敗', err); }
          };
          reader.readAsArrayBuffer(file);
        } else if (file) {
          alert('JPEG、PNG、SVGファイルのみサポートしています。');
        }
        document.body.removeChild(fileInput);
      });
    }
  });
}

// グローバル注文画像ドロップゾーンを作成
async function createOrderImageDropZone() {
  return await createBaseImageDropZone({
    storageKey: 'orderImage',
    isIndividual: false,
    containerClass: 'order-image-drop'
  });
}

async function createProductOrderImageDropZone(productId, productName = '') {
  const defaultMessage = PROTOCOL_UTILS.canDragFromExternalSites()
    ? 'この商品ID用の画像を設定'
    : 'この商品ID用の画像を選択';

  return await createBaseImageDropZone({
    productId,
    productName,
    containerClass: 'order-image-drop',
    defaultMessage
  });
}

// 個別注文用の画像ドロップゾーンを作成する関数（リファクタリング済み）
async function createIndividualOrderImageDropZone(orderNumber) {
  const defaultMessage = PROTOCOL_UTILS.canDragFromExternalSites() 
    ? '画像をドロップ or クリックで選択'
    : '画像をクリックで選択 (ドラッグは簡易サーバー経由で可能)';
    
  return await createBaseImageDropZone({
    storageKey: `orderImage_${orderNumber}`,
    isIndividual: true,
    orderNumber: orderNumber,
    containerClass: 'individual-order-image-drop',
    defaultMessage: defaultMessage
  });
}

document.getElementById("file").addEventListener("change", async function() {
  CustomLabels.updateButtonStates();
  await CustomLabels.updateSummary();
  
  // 固定ヘッダーのファイル選択状態を更新
  const fileInput = this;
  if (fileInput.files && fileInput.files.length > 0) {
      const fileName = fileInput.files[0].name;
      // コンパクト表示用に短縮
      const shortName = fileName.length > 15 ? fileName.substring(0, 12) + '...' : fileName;
      updateSelectedSourceInfo(shortName, true);
      updateProcessedOrdersVisibility();
  setPreviewSource('');
  hideQuickGuide();
      
      // CSVファイルが選択されたら自動的に処理を実行
      console.log('CSVファイルが選択されました。自動処理を開始します:', fileName);
      await autoProcessCSV();
  } else {
      updateSelectedSourceInfo('未選択', false);
      updateProcessedOrdersVisibility();
  setPreviewSource('');
  showQuickGuide();
      
      // ファイルがクリアされた場合は結果もクリア
      clearPreviousResults();
      window.lastOrderSelection = [];
      window.lastPreviewConfig = null;
      window.currentDisplayedOrderNumbers = [];
  }
});

// ページロード時にボタン状態を設定
window.addEventListener("load", function() {
  // 初期状態でボタンを設定
  CustomLabels.updateButtonStates();
});

// グローバルエラーハンドリング
window.addEventListener('error', function(event) {
  console.error('予期しないエラーが発生しました:', event.error);
});

window.addEventListener('unhandledrejection', function(event) {
  console.error('未処理のPromise拒否:', event.reason);
});

// パフォーマンス監視（開発時のみ）
if (window.performance && window.performance.mark) {
  window.performance.mark('app-start');
}

// カスタムラベル設定行の表示/非表示を切り替え
function toggleCustomLabelRow(enabled) {
  // 新: 統合されたカスタムラベルブロック
  const block = document.getElementById('customLabelsBlock');
  if (block) block.hidden = !enabled;
  // 旧IDへの互換（残っていても非表示に）
  const legacy = document.getElementById('sidebarCustomLabelSection');
  if (legacy) legacy.hidden = true;
}

// 注文画像表示機能の関数群
function toggleOrderImageRow(enabled) {
  const orderImageRow = document.getElementById('orderImageRow');
  if (orderImageRow) orderImageRow.style.display = enabled ? 'table-row' : 'none';
  const imageDropZone = document.getElementById('imageDropZone');
  const imageDropGroup = imageDropZone ? imageDropZone.closest('.sidebar-group') || imageDropZone : null;
  if (imageDropGroup) imageDropGroup.style.display = enabled ? '' : 'none';
  const productBlock = document.getElementById('productOrderImagesBlock');
  if (productBlock) productBlock.hidden = !enabled;
}

// 全ての注文明細の画像表示可視性を更新
async function updateAllOrderImagesVisibility(enabled) {
  // この関数は不要になったため、中身を空にするか、呼び出し元を削除します。
  // 今回は呼び出し元を修正したため、この関数はもう呼ばれません。
  // 念のため残しますが、処理は行いません。
  debugLog('updateAllOrderImagesVisibility はアーキテクチャ変更により不要になりました。');
}

// ---- カスタムラベル関連ロジック整理済み ----
//  編集 / 書式 / フォント / 検証 / 保存 / 集計 / イベント初期化 / 遅延プレビュー
//  -> custom-labels.js / custom-labels-font.js に集約
//  boothcsv.js 側は高レベルの CSV 処理と最終描画のみ担当
// ------------------------------------------

// 必須テンプレートの存在を起動時に検証
function verifyRequiredTemplates() {
  const required = [
    'qrDropPlaceholder',
    'orderImageDropDefault',
    'customLabelsInstructionTemplate',
  'fontListItemTemplate',
  'labelCell',
  'customLabelCell'
  ];
  // cloneTemplate 内部で存在検証。失敗時は即 throw。
  required.forEach(id => cloneTemplate(id));
}

// 印刷ボタンのイベントリスナーを変更
document.addEventListener('DOMContentLoaded', function() {
  // 必須 template の存在確認（不足時は起動中断）
  verifyRequiredTemplates();
  const fileInput = document.getElementById('file');
  const boothOrdersOpenButton = document.getElementById('boothOrdersOpenButton');
  const selectCsvFileButton = document.getElementById('selectCsvFileButton');
  const quickGuideDownloadCsvButton = document.getElementById('quickGuideDownloadCsvButton');
  const quickGuideSelectCsvButton = document.getElementById('quickGuideSelectCsvButton');
  const fontSectionHeader = document.getElementById('fontSectionHeader');
  const printButton = document.getElementById('printButton');
  const printButtonCompact = document.getElementById('printButtonCompact');
  const printBtn = document.getElementById('print-btn'); // 新しい印刷ボタン
  const previewBackButton = document.getElementById('previewBackButton');

  function openBoothOrdersPage() {
    window.open('https://manage.booth.pm/orders?state=paid', '_blank', 'noopener');
  }

  function openCsvFilePicker() {
    if (fileInput) {
      fileInput.click();
    }
  }

  [boothOrdersOpenButton, quickGuideDownloadCsvButton].forEach(function(button) {
    if (!button) return;
    button.addEventListener('click', openBoothOrdersPage);
  });

  [selectCsvFileButton, quickGuideSelectCsvButton].forEach(function(button) {
    if (!button) return;
    button.addEventListener('click', openCsvFilePicker);
  });

  if (fontSectionHeader) {
    fontSectionHeader.addEventListener('click', function() {
      toggleFontSection();
    });
    fontSectionHeader.addEventListener('keydown', function(event) {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      toggleFontSection();
    });
  }

  if (previewBackButton) {
    previewBackButton.addEventListener('click', function() {
      returnToProcessedOrdersPanel().catch(error => {
        console.error('一覧への戻り処理エラー:', error);
        alert(error && error.message ? error.message : '処理済み注文一覧への戻りに失敗しました');
      });
    });
  }

  updatePreviewReturnButtonVisibility();
  
  // 既存の印刷ボタンがある場合
  if (printButton) {
    // 既存のonClickを保存
    const originalOnClick = printButton.onclick;
    
    // 新しい処理を設定
    printButton.onclick = function() {
      // 印刷前の確認
      if (!confirmPrint()) {
        return; // キャンセルされた場合は印刷しない
      }
      
      // 印刷を実行
      window.print();
      
      // 印刷ダイアログが閉じた後に実行される
      setTimeout(() => {
        // 印刷後の処理があれば実行
        if (originalOnClick) {
          originalOnClick.call(this);
        }
      }, 1000);
    };
  }
  
  // 固定ヘッダーの印刷ボタンがある場合
  if (printButtonCompact) {
    printButtonCompact.onclick = function() {
      // 印刷前の確認
      if (!confirmPrint()) {
        return; // キャンセルされた場合は印刷しない
      }
      
      // 印刷を実行
      window.print();
      
      // 印刷ダイアログが閉じた後に実行される
      setTimeout(() => {
        // 印刷完了の確認
        if (confirm('印刷が完了しましたか？完了した場合、次回のスキップ枚数を更新します。')) {
          updateSkipCount();
        }
      }, 1000);
    };
  }
  
  // 新しい印刷ボタン（print-btn）の処理
  if (printBtn) {
    printBtn.onclick = function() {
      // 印刷前の確認
      if (!confirmPrint()) {
        return; // キャンセルされた場合は印刷しない
      }
      
      // 印刷を実行
      window.print();
      
      // 印刷ダイアログが閉じた後に実行される
      setTimeout(() => {
        // 印刷完了の確認
        if (confirm('印刷が完了しましたか？完了した場合、次回のスキップ枚数を更新します。')) {
          updateSkipCount();
        }
      }, 1000);
    };
  }
});

// 現在の印刷枚数を取得する関数
function getCurrentPrintCounts() {
  const orderCountElement = document.getElementById('orderSheetCount');
  const labelCountElement = document.getElementById('labelSheetCount');
  const customLabelCountElement = document.getElementById('customLabelCount');
  
  const orderSheets = orderCountElement ? parseInt(orderCountElement.textContent, 10) || 0 : 0;
  const labelSheets = labelCountElement ? parseInt(labelCountElement.textContent, 10) || 0 : 0;
  const customLabels = customLabelCountElement ? parseInt(customLabelCountElement.textContent, 10) || 0 : 0;
  
  return {
    orderSheets,
    labelSheets,
    customLabels
  };
}

// 印刷前の確認を行う関数
function confirmPrint() {
  const counts = getCurrentPrintCounts();
  
  let message = '印刷を開始します。プリンターに以下の用紙をセットしてください：\n\n';
  
  // 印刷順序に合わせて表示: ラベルシート→普通紙
  if (counts.labelSheets > 0) {
    message += `🏷️ A4ラベルシート(44面): ${counts.labelSheets}枚\n`;
    if (counts.customLabels > 0) {
      message += `   (うちカスタムラベル: ${counts.customLabels}面)\n`;
    }
  }
  
  if (counts.orderSheets > 0) {
    message += `📄 A4普通紙: ${counts.orderSheets}枚\n`;
  }
  
  if (counts.orderSheets === 0 && counts.labelSheets === 0) {
    message += '印刷するものがありません。\n';
    alert(message);
    return false;
  }
  
  message += '\n用紙の準備ができましたら「OK」を押してください。';
  
  return confirm(message);
}

// スキップ枚数を更新する関数
async function updateSkipCount() {
  try {
    // 現在のスキップ枚数を取得
    const currentSkip = parseInt(document.getElementById("labelskipnum").value, 10) || 0;
    
    // 固定ヘッダーから実際に印刷された枚数を取得
    const counts = getCurrentPrintCounts();
    
    // ラベルシートが印刷されていない場合は何もしない
    if (counts.labelSheets === 0) {
      alert('ラベルシートが印刷されていないため、スキップ枚数の更新はありません。');
      return;
    }
    
    // 実際に使用したラベル面数を計算
    let totalUsedLabels = 0;
    
    // 未印刷の注文明細数を取得（印刷済みはスキップ面数に含めない）
    let unprintedOrderCount = 0;
    const repo = window.orderRepository || null;
    if (repo && Array.isArray(window.currentDisplayedOrderNumbers)) {
      for (const num of window.currentDisplayedOrderNumbers) {
        const rec = repo.get(num);
        if (rec && !rec.printedAt) unprintedOrderCount++;
      }
    } else {
      const sections = document.querySelectorAll('section.sheet');
      for (const section of sections) {
        if (section.classList.contains('is-printed')) continue;
        const id = section.id || '';
        if (id.startsWith('order-')) unprintedOrderCount++;
      }
    }
    totalUsedLabels += unprintedOrderCount;
    
    // 有効なカスタムラベル面数を取得
    if (document.getElementById("customLabelEnable").checked) {
  const customLabels = CustomLabels.getFromUI();
      const enabledCustomLabels = customLabels.filter(label => label.enabled);
      const totalCustomCount = enabledCustomLabels.reduce((sum, label) => sum + (parseInt(label.count, 10) || 0), 0);
      totalUsedLabels += totalCustomCount;
    }
    
    // 全体の使用面数を計算（現在のスキップ + 新たに使用した面数）
    const totalUsedWithSkip = currentSkip + totalUsedLabels;
    
    // 44面シートでの余り面数を計算
    const newSkipValue = totalUsedWithSkip % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET;
    
    console.log(`スキップ枚数更新計算:
      現在のスキップ: ${currentSkip}面
      未印刷注文明細: ${unprintedOrderCount}面
      カスタムラベル: ${totalUsedLabels - unprintedOrderCount}面
      合計使用面数: ${totalUsedLabels}面
      総使用面数(スキップ含む): ${totalUsedWithSkip}面
      新しいスキップ値: ${newSkipValue}面`);
    
    // 新しいスキップ枚数を設定
    document.getElementById("labelskipnum").value = newSkipValue;
    await StorageManager.set(StorageManager.KEYS.LABEL_SKIP, newSkipValue);
    
    // settingsCacheも更新（CSV再読み込み時の不整合を防ぐ）
    settingsCache.labelskip = newSkipValue;
    
    // lastPreviewConfigも更新（カスタムラベルプレビューの不整合を防ぐ）
    if (window.lastPreviewConfig) {
      window.lastPreviewConfig.labelskip = newSkipValue;
    }
    
    // カスタムラベルの上限も更新（エラーハンドリング付き）
    try {
  await CustomLabels.updateSummary();
      console.log('✅ カスタムラベルサマリー更新完了');
    } catch (summaryError) {
      console.error('⚠️ カスタムラベルサマリー更新エラー:', summaryError);
      // サマリー更新エラーは致命的ではないので、処理を継続
    }

    // 印刷済み注文番号の印刷日時を repository 経由で記録
    try {
      const repo = window.orderRepository;
      if (repo) {
        const now = new Date().toISOString();
        // 注文明細セクションは section.sheet にID "order-<番号>" が付与される
        const sections = document.querySelectorAll('section.sheet');
        for (const section of sections) {
          // 既に印刷済みのものは今回対象外（@media print でも除外されている）
          if (section.classList.contains('is-printed')) continue;
          const id = section.id || '';
          const orderNumber = id.startsWith('order-') ? id.substring(6) : '';
          if (orderNumber) await repo.markPrinted(orderNumber, now);
        }
        console.log('✅ 印刷済み注文番号の印刷日時を保存しました (repository)');
      }
    } catch (e) {
      console.error('❌ 印刷済み注文番号保存エラー(repository):', e);
    }
    
    // 印刷枚数表示を再更新
    updatePrintCountDisplay();
    console.log('✅ スキップ枚数更新後の印刷枚数表示を更新しました');
    
    // カスタムラベルプレビューも再更新
    try {
      await updateCustomLabelsPreview();
      console.log('✅ スキップ枚数更新後のカスタムラベルプレビューを更新しました');
    } catch (previewError) {
      console.error('⚠️ カスタムラベルプレビュー更新エラー:', previewError);
      // プレビュー更新エラーは致命的ではないので、処理を継続
    }
    
    // 更新完了メッセージ
    alert(`次回のスキップ枚数を ${newSkipValue} 面に更新しました。\n\n詳細:\n・印刷前スキップ: ${currentSkip}面\n・今回使用: ${totalUsedLabels}面\n・合計: ${totalUsedWithSkip}面\n・次回スキップ: ${newSkipValue}面`);

    await returnToProcessedOrdersPanel();
    
  } catch (error) {
    console.error('スキップ枚数更新エラー:', error);
    alert(`スキップ枚数の更新中にエラーが発生しました。\n\nエラー詳細: ${error.message || error}`);
  }
}

// カスタムフォント管理機能
// initializeFontDropZone / handleFontFile は custom-labels-font.js に移動


// (getFontMimeType / addFontToCSS / getFontFormat) は未使用のため削除済み

let _fontFaceLoadToken = 0;

// --- 汎用入力 Enter 抑止 (フォーム内で Enter がフォント削除ボタン等にフォーカス飛ぶ副作用対策) ---
function setupPreventEnterOnSimpleInputs() {
  const selector = 'input[type="number"], input[type="text"], input[type="search"], input[type="email"], input[type="url"], input[type="tel"], input[type="password"]';
  document.querySelectorAll(selector).forEach(el => {
    if (el.dataset.preventEnterBound) return;
    el.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        e.preventDefault();
        e.stopPropagation();
        // 明示的に確定処理が必要ならここで呼び出せる（現状なし）
      }
    });
    el.dataset.preventEnterBound = '1';
  });
}

// 初期化後に呼び出し
setTimeout(setupPreventEnterOnSimpleInputs, 0);

// 動的追加入力にも対応 (MutationObserver)
const preventEnterObserver = new MutationObserver(() => {
  setupPreventEnterOnSimpleInputs();
});
preventEnterObserver.observe(document.documentElement, { childList: true, subtree: true });

// ボタンで Enter を無効化（フォント削除等）。スペース / クリック操作のみ許可。
document.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    const target = e.target;
    if (target instanceof HTMLButtonElement) {
      // type="button" のアクションボタンで Enter を無効化
      if ((target.getAttribute('type') || 'button') === 'button') {
        e.preventDefault();
        e.stopPropagation();
      }
    }
  }
}, true);

// applyStyleToSelection などのスタイル編集関数は custom-labels.js (CustomLabelStyle) に移動

// updateSpanStyle は移動

// 部分選択時のspan分割処理
// handlePartialSpanSelection は移動

// 複数span要素選択時の統合処理（改行保持版）
// handleMultiSpanSelection は移動

// 新しいspan要素を作成するヘルパー関数
// createNewSpanForSelection は移動

// BRやゼロ幅スペースを保持しつつ、選択範囲内のテキストノードにだけスタイルを適用
// applyStylePreservingBreaks は移動

// フォントファミリーを選択範囲に適用（統合された関数を使用）
// applyFontFamilyToSelection は移動

// 選択範囲解析（簡略再定義）
// analyzeSelectionRange は移動

// デフォルトフォントに戻す専用関数
// applyDefaultFontToSelection は移動

// 空のspan要素やネストしたspan要素を掃除（改良版）
// cleanupEmptySpans は移動（簡略版 custom-labels.js 内）

// CSSスタイル文字列をMapに変換するヘルパー関数
// parseStyles は移動

// フォントサイズを選択範囲に適用（統合された関数を使用）
// applyFontSizeToSelection は移動

// span要素のスタイルを統合するヘルパー関数
// mergeSpanStyles は移動

// スタイル文字列を正規化するヘルパー関数
// normalizeStyle は移動

// スタイル文字列をMapに変換するヘルパー関数
// parseStyleString は移動

// 指定した要素配下の全てのspanから、指定スタイルプロパティを取り除く
// removeStyleFromDescendants は移動

// フォントサイズを選択範囲に適用（統合された関数を使用）
// applyFontSizeToSelection は移動（重複）

// ===========================================
// IndexedDBフォント機能の初期化とヘルパー関数
// ===========================================

// フォントセクションの折りたたみ機能
async function toggleFontSection() {
  const content = document.getElementById('fontSectionContent');
  const arrow = document.getElementById('fontSectionArrow');
  
  debugLog('toggleFontSection called');
  debugLog('Current maxHeight:', content.style.maxHeight);
  debugLog('ScrollHeight:', content.scrollHeight);
  debugLog('Content element:', content);
  debugLog('Arrow element:', arrow);
  
  if (content.style.maxHeight && content.style.maxHeight !== '0px') {
    // 折りたたむ
    debugLog('Collapsing font section');
    content.style.maxHeight = '0px';
    arrow.style.transform = 'rotate(-90deg)';
    await StorageManager.setFontSectionCollapsed(true);
    debugLog('Font section collapsed, state saved as true');
  } else {
    // 展開する
    debugLog('Expanding font section');
    // 一時的にトランジションを無効化
    content.style.transition = 'none';
    content.style.maxHeight = content.scrollHeight + 'px';
    // トランジションを再有効化
    setTimeout(() => {
      content.style.transition = 'max-height 0.3s ease-out';
    }, 10);
    arrow.style.transform = 'rotate(0deg)';
    await StorageManager.setFontSectionCollapsed(false);
    debugLog('Font section expanded to', content.scrollHeight + 'px', 'state saved as false');
    
    // 確認のため再度ログ出力
    setTimeout(() => {
      debugLog('After expansion - maxHeight:', content.style.maxHeight, 'computedHeight:', getComputedStyle(content).height);
    }, 50);
  }
}

// フォントセクションの初期状態を設定
async function initializeFontSection() {
  const content = document.getElementById('fontSectionContent');
  const arrow = document.getElementById('fontSectionArrow');
  const isCollapsed = await StorageManager.getFontSectionCollapsed();
  
  debugLog('initializeFontSection called');
  debugLog('Stored collapsed state:', isCollapsed);
  debugLog('Content element:', content);
  debugLog('Arrow element:', arrow);
  
  // 初期状態は折りたたみ（明示的に展開が設定されている場合のみ展開）
  if (!isCollapsed) {
    // 展開状態
    debugLog('Setting initial state to expanded');
    setTimeout(() => {
      content.style.maxHeight = content.scrollHeight + 'px';
      debugLog('Font section initialized to expanded height:', content.scrollHeight + 'px');
    }, 100);
    arrow.style.transform = 'rotate(0deg)';
  } else {
    // 折りたたみ状態（デフォルト）
    debugLog('Setting initial state to collapsed');
    content.style.maxHeight = '0px';
    arrow.style.transform = 'rotate(-90deg)';
    debugLog('Font section initialized to collapsed');
  }
}

// フォントリスト更新時にセクション高さを調整
function adjustFontSectionHeight() {
  const content = document.getElementById('fontSectionContent');
  debugLog('adjustFontSectionHeight called');
  debugLog('Content element:', content);
  debugLog('Current maxHeight:', content?.style.maxHeight);
  debugLog('ScrollHeight:', content?.scrollHeight);
  
  if (content && content.style.maxHeight !== '0px') {
    content.style.maxHeight = content.scrollHeight + 'px';
    debugLog('Font section height adjusted to:', content.scrollHeight + 'px');
  } else {
    debugLog('Font section height not adjusted (collapsed or element missing)');
  }
}

// ================================
// サイドバー開閉ロジック（印刷非影響）
// ================================
(function setupSidebar() {
  function initSidebarOnce(){
    if (initSidebarOnce._ran) return; initSidebarOnce._ran = true;

  const toggleBtn = document.getElementById('sidebarToggle');
    const sidebar = document.getElementById('appSidebar');
  const overlay = document.getElementById('sidebarOverlay');
  const closeBtn = document.getElementById('sidebarClose');
  const pinBtn = document.getElementById('sidebarPin');
    if (!toggleBtn || !sidebar || !overlay || !closeBtn) return;

    const focusableSelectors = 'a[href], button:not([disabled]), input, select, textarea, [tabindex]:not([tabindex="-1"])';
    let lastFocused = null;

    function openSidebar() {
      lastFocused = document.activeElement;
      sidebar.classList.add('open');
      sidebar.removeAttribute('aria-hidden');
      toggleBtn.setAttribute('aria-expanded', 'true');
      if (!document.body.classList.contains('sidebar-docked')) {
        overlay.hidden = false;
        overlay.classList.add('show');
        overlay.style.display = 'block';
        document.body.dataset.scrollLock = '1';
        document.body.style.overflow = 'hidden';
      }
      // フォーカス移動
      const first = sidebar.querySelector(focusableSelectors);
      (first || sidebar).focus();
    }

    function closeSidebar() {
      sidebar.classList.remove('open');
      sidebar.setAttribute('aria-hidden', 'true');
      toggleBtn.setAttribute('aria-expanded', 'false');
      if (!document.body.classList.contains('sidebar-docked')) {
        overlay.classList.remove('show');
        overlay.hidden = true;
        overlay.style.display = 'none';
        if (document.body.dataset.scrollLock === '1') {
          delete document.body.dataset.scrollLock;
          document.body.style.overflow = '';
        }
      }
      if (lastFocused && typeof lastFocused.focus === 'function') {
        lastFocused.focus();
      } else {
        toggleBtn.focus();
      }
    }

    toggleBtn.addEventListener('click', () => {
      if (sidebar.classList.contains('open')) closeSidebar(); else openSidebar();
    });
    overlay.addEventListener('click', closeSidebar);
    closeBtn.addEventListener('click', closeSidebar);

    // ピン留め（ドック）切り替え
    function applyDockedState(docked){
      if (docked) {
        document.body.classList.add('sidebar-docked');
        pinBtn?.setAttribute('aria-pressed','true');
  // ドック時はサイドバーは常時表示・オーバーレイ無効
        sidebar.classList.add('open');
        sidebar.removeAttribute('aria-hidden');
        overlay.classList.remove('show');
        overlay.hidden = true;
        overlay.style.display = 'none';
        // スクロールロック解除
        if (document.body.dataset.scrollLock === '1') {
          delete document.body.dataset.scrollLock;
          document.body.style.overflow = '';
        }
      } else {
        document.body.classList.remove('sidebar-docked');
        pinBtn?.setAttribute('aria-pressed','false');
  // 閉状態に戻す（必要なら）
        sidebar.classList.remove('open');
        sidebar.setAttribute('aria-hidden','true');
      }
    }

    // 保存された状態を復元（IndexedDB優先、未設定なら初期値: ドックON）
    (async () => {
      let persisted = null;
      try { if (window.StorageManager && typeof StorageManager.getSidebarDocked==='function') { persisted = await StorageManager.getSidebarDocked(); } } catch {}
      if (persisted === null) {
        // 初期既定: ドックON
        applyDockedState(true);
        try {
          if (window.StorageManager && typeof StorageManager.setSidebarDocked==='function') await StorageManager.setSidebarDocked(true);
        } catch {}
      } else {
        applyDockedState(!!persisted);
      }
    })();

  pinBtn?.addEventListener('click', async () => {
      const docked = !document.body.classList.contains('sidebar-docked');
      applyDockedState(docked);
      try {
        if (window.StorageManager && typeof StorageManager.setSidebarDocked==='function') await StorageManager.setSidebarDocked(!!docked);
      } catch {}
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && sidebar.classList.contains('open')) {
        e.preventDefault();
        closeSidebar();
      } else if (e.key === 'Tab' && sidebar.classList.contains('open')) {
        // フォーカストラップ
        const nodes = Array.from(sidebar.querySelectorAll(focusableSelectors)).filter(el => el.offsetParent !== null);
        if (nodes.length === 0) return;
        const first = nodes[0];
        const last = nodes[nodes.length - 1];
        if (e.shiftKey && document.activeElement === first) {
          e.preventDefault(); last.focus();
        } else if (!e.shiftKey && document.activeElement === last) {
          e.preventDefault(); first.focus();
        }
      }
    });
  }
  // DOM が既に準備済みなら即時実行、そうでなければ DOMContentLoaded で実行
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSidebarOnce, { once: true });
  } else {
    initSidebarOnce();
  }
})();
