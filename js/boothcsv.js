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
  }
};

// デバッグフラグの取得
const DEBUG_MODE = (() => {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get('debug') === '1';
})();

// プロトコル判定ユーティリティ
const PROTOCOL_UTILS = {
  isFileProtocol: () => window.location.protocol === 'file:',
  isHttpProtocol: () => window.location.protocol === 'http:' || window.location.protocol === 'https:',
  canDragFromExternalSites: () => !PROTOCOL_UTILS.isFileProtocol()
};

// 初回クイックガイド表示制御（グローバルからどこでも呼べるユーティリティ）
function hideQuickGuide(){ const el = document.getElementById('initialQuickGuide'); if(el) el.hidden = true; }
function showQuickGuide(){ const el = document.getElementById('initialQuickGuide'); if(el) el.hidden = false; }

// リアルタイム更新制御フラグ
// isEditingCustomLabel は custom-labels.js で window プロパティとして定義される
// 直近描画に使用した設定・注文番号を保持し、再プレビューへ再利用
window.lastPreviewConfig = null; // { labelyn, labelskip, sortByPaymentDate, customLabelEnable, customLabels }
window.lastOrderSelection = []; // [orderNumber, ...]
let pendingUpdateTimer = null;

const processedOrdersState = {
  sortKey: 'paymentDate',
  sortDirection: 'desc',
  currentPage: 1,
  pageSize: 20,
  unprintedOnly: false,
  totalItems: 0,
  totalPages: 1,
  pageItems: []
};
const processedOrdersSelection = new Set();
const processedOrdersUI = {
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
let processedOrdersUnsubscribe = null;

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
  shippingMessageTemplate: ''
};

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

  Object.assign(window.settingsCache, {
    labelyn: settings.labelyn,
    labelskip: settings.labelskip,
    sortByPaymentDate: settings.sortByPaymentDate,
    customLabelEnable: settings.customLabelEnable,
    orderImageEnable: settings.orderImageEnable,
    shippingMessageTemplate: settings.shippingMessageTemplate || ''
  });

  initializeShippingMessageTemplateUI(settings.shippingMessageTemplate || '');

  await setupProcessedOrdersPanel();
  updateProcessedOrdersVisibility();

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
  // 画像表示トグルとドロップゾーンの表示を同期
  try {
    const dzHost = document.getElementById('imageDropZone');
    const group = dzHost ? dzHost.closest('.sidebar-group') || dzHost : null;
    const cb = document.getElementById('orderImageEnable');
    const apply = (on) => { if (group) group.style.display = on ? '' : 'none'; };
    if (cb) {
      apply(cb.checked);
      cb.addEventListener('change', () => apply(cb.checked));
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
      clearPreviousResults();
      updateProcessedOrdersVisibility();
    const config = {
      file: document.getElementById('file').files[0],
      labelyn: settingsCache.labelyn,
      labelskip: settingsCache.labelskip,
      sortByPaymentDate: settingsCache.sortByPaymentDate,
      customLabelEnable: settingsCache.customLabelEnable,
  customLabels: settingsCache.customLabelEnable ? CustomLabels.getFromUI().filter(l=>l.enabled) : []
    };
      
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

    // CSV が読み込まれている場合: CSVラベル + (必要なら) カスタムラベル を再生成
    if (hasSelection) {
      if (!config.labelyn) {
        // ラベル印刷OFFなら表示だけクリア
        clearPreviousResults();
        updatePrintCountDisplay(0, 0, 0);
        return;
      }
      // 再生成（processCSVResults 内で clearPreviousResults していないので先に消す）
      clearPreviousResults();
      await renderPreviewFromRepository(window.lastOrderSelection, config);
      return; // ここで終了（CSV再表示 + カスタムラベル反映済）
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
  }
}

// 前回の処理結果をクリア
function clearPreviousResults() {
  // 結果セクションだけを削除（サイドバー等の一般sectionは残す）
  document.querySelectorAll('section.sheet').forEach(sec => sec.remove());
  
  // 印刷枚数表示もクリア
  clearPrintCountDisplay();
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

function parseDateToTimestamp(value) {
  if (!value) return NaN;
  const direct = Date.parse(value);
  if (!Number.isNaN(direct)) return direct;
  const normalized = value.replace(/T/, ' ').replace(/-/g, '/');
  const fallback = Date.parse(normalized);
  return Number.isNaN(fallback) ? NaN : fallback;
}

function formatDateTime(value) {
  if (!value) return '';
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) return date.toLocaleString();
  const normalized = value.replace(/T/, ' ').replace(/-/g, '/');
  const fallback = new Date(normalized);
  if (!Number.isNaN(fallback.getTime())) return fallback.toLocaleString();
  return value;
}

async function setupProcessedOrdersPanel() {
  const template = document.getElementById('processedOrdersPanelTemplate');
  if (!template) return;
  const repo = await ensureOrderRepository();
  if (!repo) return;

  if (!processedOrdersUI.panel) {
    const fragment = template.content.cloneNode(true);
    const guide = document.getElementById('initialQuickGuide');
    if (guide && guide.parentNode) {
      guide.parentNode.insertBefore(fragment, guide.nextSibling);
    } else {
      const main = document.getElementById('mainContent');
      if (main) main.appendChild(fragment);
    }

    processedOrdersUI.panel = document.getElementById('processedOrdersPanel');
    processedOrdersUI.body = document.getElementById('processedOrdersBody');
    processedOrdersUI.selectAll = document.getElementById('processedOrdersSelectAll');
    processedOrdersUI.deleteButton = document.getElementById('processedOrdersDeleteButton');
    processedOrdersUI.previewButton = document.getElementById('processedOrdersPreviewButton');
    processedOrdersUI.prev = document.getElementById('processedOrdersPrev');
    processedOrdersUI.next = document.getElementById('processedOrdersNext');
    processedOrdersUI.pageInfo = document.getElementById('processedOrdersPageInfo');
    processedOrdersUI.empty = document.getElementById('processedOrdersEmpty');
    processedOrdersUI.filter = document.getElementById('processedOrdersUnprintedOnly');
    processedOrdersUI.headerCells = processedOrdersUI.panel ? Array.from(processedOrdersUI.panel.querySelectorAll('th.sortable')) : [];

    if (processedOrdersUI.selectAll) {
      processedOrdersUI.selectAll.addEventListener('change', handleProcessedOrdersSelectAll);
    }
    if (processedOrdersUI.filter) {
      processedOrdersUI.filter.addEventListener('change', handleProcessedOrdersFilterChange);
    }
    if (processedOrdersUI.prev) {
      processedOrdersUI.prev.addEventListener('click', handleProcessedOrdersPrevPage);
    }
    if (processedOrdersUI.next) {
      processedOrdersUI.next.addEventListener('click', handleProcessedOrdersNextPage);
    }
    if (processedOrdersUI.deleteButton) {
      processedOrdersUI.deleteButton.addEventListener('click', handleProcessedOrdersDelete);
    }
    if (processedOrdersUI.previewButton) {
      processedOrdersUI.previewButton.addEventListener('click', handleProcessedOrdersPreview);
    }
    processedOrdersUI.headerCells.forEach(cell => {
      const key = cell.dataset.sort;
      if (!key) return;
      cell.addEventListener('click', () => handleProcessedOrdersSort(key));
    });
  }

  if (processedOrdersUnsubscribe) {
    processedOrdersUnsubscribe();
    processedOrdersUnsubscribe = null;
  }
  processedOrdersUnsubscribe = repo.onUpdate(() => {
    refreshProcessedOrdersPanel();
  });

  await refreshProcessedOrdersPanel();
}

function compareProcessedOrders(a, b) {
  const dir = processedOrdersState.sortDirection === 'asc' ? 1 : -1;
  let result = 0;
  if (processedOrdersState.sortKey === 'paymentDate') {
    const av = Number.isNaN(a.paymentDateValue) ? -Infinity : a.paymentDateValue;
    const bv = Number.isNaN(b.paymentDateValue) ? -Infinity : b.paymentDateValue;
    if (av < bv) result = -1;
    else if (av > bv) result = 1;
  } else if (processedOrdersState.sortKey === 'printedAt') {
    const av = Number.isNaN(a.printedAtValue) ? -Infinity : a.printedAtValue;
    const bv = Number.isNaN(b.printedAtValue) ? -Infinity : b.printedAtValue;
    if (av < bv) result = -1;
    else if (av > bv) result = 1;
  } else {
    result = a.orderNumber.localeCompare(b.orderNumber, 'ja', { numeric: true, sensitivity: 'base' });
  }
  if (result === 0) {
    result = a.orderNumber.localeCompare(b.orderNumber, 'ja', { numeric: true, sensitivity: 'base' });
  }
  return result * dir;
}

async function refreshProcessedOrdersPanel() {
  const repo = await ensureOrderRepository();
  if (!repo || !processedOrdersUI.panel) return;

  const records = repo.getAll();
  const data = records.map(rec => {
    const paymentRaw = rec && rec.row ? rec.row[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || '' : '';
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

  const sorted = [...data].sort(compareProcessedOrders);
  const filtered = processedOrdersState.unprintedOnly ? sorted.filter(item => !item.printed) : sorted;

  processedOrdersState.totalItems = filtered.length;
  processedOrdersState.totalPages = filtered.length === 0 ? 1 : Math.ceil(filtered.length / processedOrdersState.pageSize);
  if (processedOrdersState.currentPage > processedOrdersState.totalPages) {
    processedOrdersState.currentPage = processedOrdersState.totalPages;
  }
  if (filtered.length === 0) {
    processedOrdersState.currentPage = 1;
  }
  const startIndex = (processedOrdersState.currentPage - 1) * processedOrdersState.pageSize;
  const pageItems = filtered.slice(startIndex, startIndex + processedOrdersState.pageSize);
  processedOrdersState.pageItems = pageItems;

  const validKeys = new Set(data.map(item => item.orderNumber));
  processedOrdersSelection.forEach(key => {
    if (!validKeys.has(key)) processedOrdersSelection.delete(key);
  });

  if (processedOrdersUI.body) {
    processedOrdersUI.body.textContent = '';
    pageItems.forEach(item => {
      const tr = document.createElement('tr');

      const selectTd = document.createElement('td');
      selectTd.className = 'col-select';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.dataset.orderNumber = item.orderNumber;
      checkbox.checked = processedOrdersSelection.has(item.orderNumber);
      checkbox.setAttribute('aria-label', `注文 ${item.orderNumber} を選択`);
      checkbox.addEventListener('change', (ev) => {
        handleProcessedOrdersRowSelection(item.orderNumber, ev.target.checked);
      });
      selectTd.appendChild(checkbox);
      tr.appendChild(selectTd);

      const orderTd = document.createElement('td');
      const link = document.createElement('a');
      link.href = `https://manage.booth.pm/orders/${encodeURIComponent(item.orderNumber)}`;
      link.target = '_blank';
      link.rel = 'noopener';
      link.className = 'order-link';
      link.textContent = item.orderNumber;
      orderTd.appendChild(link);
      tr.appendChild(orderTd);

      const paymentTd = document.createElement('td');
      paymentTd.textContent = item.paymentDateRaw || '';
      tr.appendChild(paymentTd);

      const printedTd = document.createElement('td');
      printedTd.textContent = item.printedAtRaw ? formatDateTime(item.printedAtRaw) : '';
      tr.appendChild(printedTd);

      processedOrdersUI.body.appendChild(tr);
    });
  }

  if (processedOrdersUI.empty) {
    if (filtered.length === 0) {
      processedOrdersUI.empty.hidden = false;
      processedOrdersUI.empty.textContent = processedOrdersState.unprintedOnly ? '未印刷の注文はありません。' : '保存されている注文はありません。';
    } else {
      processedOrdersUI.empty.hidden = true;
    }
  }

  syncProcessedOrdersSelectAll(pageItems);
  syncProcessedOrdersActionButtons();
  updateProcessedOrdersPagination(filtered.length);
  updateProcessedOrdersSortIndicators();
  updateProcessedOrdersVisibility();
}

function updateProcessedOrdersSortIndicators() {
  if (!Array.isArray(processedOrdersUI.headerCells)) return;
  processedOrdersUI.headerCells.forEach(cell => {
    cell.classList.remove('sort-asc', 'sort-desc');
    if (cell.dataset.sort === processedOrdersState.sortKey) {
      cell.classList.add(processedOrdersState.sortDirection === 'asc' ? 'sort-asc' : 'sort-desc');
    }
  });
}

function syncProcessedOrdersSelectAll(pageItems) {
  if (!processedOrdersUI.selectAll) return;
  if (pageItems.length === 0) {
    processedOrdersUI.selectAll.checked = false;
    processedOrdersUI.selectAll.indeterminate = false;
    processedOrdersUI.selectAll.disabled = processedOrdersState.totalItems === 0;
    return;
  }
  const selectedCount = pageItems.filter(item => processedOrdersSelection.has(item.orderNumber)).length;
  processedOrdersUI.selectAll.disabled = false;
  processedOrdersUI.selectAll.checked = selectedCount === pageItems.length;
  processedOrdersUI.selectAll.indeterminate = selectedCount > 0 && selectedCount < pageItems.length;
}

function syncProcessedOrdersActionButtons() {
  const count = processedOrdersSelection.size;
  if (processedOrdersUI.deleteButton) {
    processedOrdersUI.deleteButton.disabled = count === 0;
    processedOrdersUI.deleteButton.textContent = count > 0 ? `選択した注文を削除 (${count})` : '選択した注文を削除';
  }
  if (processedOrdersUI.previewButton) {
    processedOrdersUI.previewButton.disabled = count === 0;
    processedOrdersUI.previewButton.textContent = count > 0 ? `選択した注文を印刷プレビュー (${count})` : '選択した注文を印刷プレビュー';
  }
}

function updateProcessedOrdersPagination(totalItems) {
  if (processedOrdersUI.prev) {
    processedOrdersUI.prev.disabled = processedOrdersState.currentPage <= 1 || totalItems === 0;
  }
  if (processedOrdersUI.next) {
    processedOrdersUI.next.disabled = processedOrdersState.currentPage >= processedOrdersState.totalPages || totalItems === 0;
  }
  if (processedOrdersUI.pageInfo) {
    if (totalItems === 0) {
      processedOrdersUI.pageInfo.textContent = '0件';
    } else {
      processedOrdersUI.pageInfo.textContent = `${processedOrdersState.currentPage} / ${processedOrdersState.totalPages} ページ（全${totalItems}件）`;
    }
  }
}

function handleProcessedOrdersSort(key) {
  if (!key) return;
  if (processedOrdersState.sortKey === key) {
    processedOrdersState.sortDirection = processedOrdersState.sortDirection === 'asc' ? 'desc' : 'asc';
  } else {
    processedOrdersState.sortKey = key;
    processedOrdersState.sortDirection = key === 'orderNumber' ? 'asc' : 'desc';
  }
  processedOrdersState.currentPage = 1;
  refreshProcessedOrdersPanel();
}

function handleProcessedOrdersSelectAll(event) {
  const checked = event.target.checked;
  processedOrdersState.pageItems.forEach(item => {
    if (checked) processedOrdersSelection.add(item.orderNumber);
    else processedOrdersSelection.delete(item.orderNumber);
  });
  if (processedOrdersUI.body) {
    processedOrdersUI.body.querySelectorAll('input[type="checkbox"][data-order-number]').forEach(box => {
      box.checked = processedOrdersSelection.has(box.dataset.orderNumber);
    });
  }
  syncProcessedOrdersSelectAll(processedOrdersState.pageItems);
    syncProcessedOrdersActionButtons();
}

function handleProcessedOrdersRowSelection(orderNumber, checked) {
  if (!orderNumber) return;
  if (checked) processedOrdersSelection.add(orderNumber);
  else processedOrdersSelection.delete(orderNumber);
  syncProcessedOrdersSelectAll(processedOrdersState.pageItems);
    syncProcessedOrdersActionButtons();
}

function handleProcessedOrdersFilterChange(event) {
  processedOrdersState.unprintedOnly = !!event.target.checked;
  processedOrdersState.currentPage = 1;
  refreshProcessedOrdersPanel();
}

function handleProcessedOrdersPrevPage() {
  if (processedOrdersState.currentPage <= 1) return;
  processedOrdersState.currentPage -= 1;
  refreshProcessedOrdersPanel();
}

function handleProcessedOrdersNextPage() {
  if (processedOrdersState.currentPage >= processedOrdersState.totalPages) return;
  processedOrdersState.currentPage += 1;
  refreshProcessedOrdersPanel();
}

async function handleProcessedOrdersPreview() {
  if (processedOrdersSelection.size === 0) return;
  try {
    const orderNumbers = Array.from(processedOrdersSelection);
    const config = {
      labelyn: settingsCache.labelyn,
      labelskip: settingsCache.labelskip,
      sortByPaymentDate: settingsCache.sortByPaymentDate,
      customLabelEnable: settingsCache.customLabelEnable,
      customLabels: settingsCache.customLabelEnable ? CustomLabels.getFromUI().filter(label => label.enabled) : []
    };
    clearPreviousResults();
    hideQuickGuide();
    await renderPreviewFromRepository(orderNumbers, config);
  } catch (e) {
    console.error('order preview error', e);
    alert('プレビューの生成に失敗しました: ' + (e && e.message ? e.message : e));
  }
}

async function handleProcessedOrdersDelete() {
  if (processedOrdersSelection.size === 0) return;
  const orderNumbers = Array.from(processedOrdersSelection);
  if (!confirm(`選択した${orderNumbers.length}件の注文を削除しますか？`)) return;
  try {
    const repo = await ensureOrderRepository();
    if (!repo) return;
    await repo.deleteMany(orderNumbers);
    processedOrdersSelection.clear();
    alert('選択した注文を削除しました');
  } catch (e) {
    console.error('order delete error', e);
    alert('注文の削除に失敗しました: ' + (e && e.message ? e.message : e));
  }
  refreshProcessedOrdersPanel();
}

function updateProcessedOrdersVisibility() {
  if (!processedOrdersUI.panel) return;
  const fileInput = document.getElementById('file');
  const hasFile = !!(fileInput && fileInput.files && fileInput.files.length > 0);
  processedOrdersUI.panel.hidden = hasFile;
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
    await displayOrderImage(cOrder, orderNumber);
    
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

// 注文の商品アイテムリストを処理してHTMLに追加
function processProductItems(cOrder, row) {
  const tItems = cOrder.querySelector('#商品');
  const trSpace = cOrder.querySelector('.spacerow');
  
  if (row[CONSTANTS.CSV.PRODUCT_COLUMN]) {
    for (let itemrow of row[CONSTANTS.CSV.PRODUCT_COLUMN].split('\n')) {
      const cItem = document.importNode(tItems.content, true);
      const productInfo = parseProductItemData(itemrow);
      
      if (productInfo.itemId && productInfo.quantity) {
        setProductItemElements(cItem, productInfo);
        trSpace.parentNode.parentNode.insertBefore(cItem, trSpace.parentNode);
      }
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
async function displayOrderImage(cOrder, orderNumber) {
  // 表示/非表示はCSSクラスで制御するため、ここでは常に画像データを取得・表示する。

  let imageToShow = null;
  if (isValidOrderNumber(orderNumber)) {
    // 注文番号を正規化
  const normalizedOrderNumber = (orderNumber == null) ? '' : String(orderNumber).trim();
    
    // 個別画像: repository から
    let individualImage = null;
    if (window.orderRepository) {
      const rec = window.orderRepository.get(normalizedOrderNumber);
      if (rec && rec.image && rec.image.data instanceof ArrayBuffer) {
        try { const blob = new Blob([rec.image.data], { type: rec.image.mimeType || 'image/png' }); individualImage = URL.createObjectURL(blob); } catch {}
      }
    }
    if (individualImage) imageToShow = individualImage; else imageToShow = await getGlobalOrderImage();
  }

  if (imageToShow) {
    const imageDiv = document.createElement('div');
    imageDiv.classList.add('order-image');

    const img = document.createElement('img');
    img.src = imageToShow;

    imageDiv.appendChild(img);
    const container = cOrder.querySelector('.order-image-container');
    if (container) {
      container.appendChild(imageDiv);
    }
  }
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
  const templateId = PROTOCOL_UTILS.canDragFromExternalSites() ? 'qrDropPlaceholderHttp' : 'qrDropPlaceholder';
  return cloneTemplate(templateId);
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
  divDrop.setAttribute("contenteditable", "true");
  // ペースト専用イベントを設定
  setupPasteZoneEvents(divDrop);
  
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

// QRコード画像にクリックでリセット機能を追加
function addEventQrReset(elImage){
    elImage.addEventListener('click', async function(event) {
      event.preventDefault();
      
      // 親要素のQRセクションを取得
      const qrDiv = elImage.parentNode;
      const td = qrDiv.closest('td');
      
      if (td) {
        // 注文番号を取得
        const ordernumDiv = td.querySelector('.ordernum p');
  const orderNumber = ordernumDiv ? String(ordernumDiv.textContent || '').trim() : null;
        
        // 保存されたQRデータを削除
        if (orderNumber) {
          try {
            if (window.orderRepository) {
              await window.orderRepository.clearOrderQRData(orderNumber);
            }
            console.log(`QRデータを削除しました: ${orderNumber}`);
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
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.src = elImage.src;
    
    img.onload = async function() {
      try {
        const canv = document.createElement("canvas");
        const context = canv.getContext("2d");
        canv.width = img.width;
        canv.height = img.height;
        context.drawImage(img, 0, 0, canv.width, canv.height);
        // Canvas から直接 ArrayBuffer を取得（データURLは非採用）
        const blobPromise = new Promise(resolve => canv.toBlob(resolve, 'image/png'));
        const blob = await blobPromise;
        let arrayBuffer = null;
        if (blob) {
          arrayBuffer = await blob.arrayBuffer();
        } else {
          console.warn('QR: canvas toBlob が取得できませんでした。フォールバックとして dataURL を ArrayBuffer 化');
          const tmpB64 = canv.toDataURL('image/png');
          const bin = atob(tmpB64.split(',')[1]);
            const len = bin.length; const u8 = new Uint8Array(len); for (let i=0;i<len;i++){u8[i]=bin.charCodeAt(i);} arrayBuffer = u8.buffer;
        }
        
        const imageData = context.getImageData(0, 0, canv.width, canv.height);
        const barcode = jsQR(imageData.data, imageData.width, imageData.height);
        
        if(barcode){
          const b = String(barcode.data).replace(/^\s+|\s+$/g,'').replace(/ +/g,' ').split(" ");
          
          if(b.length === CONSTANTS.QR.EXPECTED_PARTS){
            // 注文番号リンク化対応: .ordernum a から取得
            const orderNumElem = elImage.closest("td").querySelector(".ordernum a");
            const rawOrderNum = orderNumElem ? orderNumElem.textContent : null;
            const ordernum = (rawOrderNum == null) ? '' : String(rawOrderNum).trim();
            
            // 重複チェック
            const duplicates = window.orderRepository ? await window.orderRepository.checkQRDuplicate(barcode.data, ordernum) : [];
            if (duplicates.length > 0) {
              const duplicateList = duplicates.join(', ');
              const confirmMessage = `警告: このQRコードは既に以下の注文で使用されています:\n${duplicateList}\n\n同じQRコードを使用すると配送ミスの原因となる可能性があります。\n続行しますか？`;
              
              if (!confirm(confirmMessage)) {
                console.log('QRコード登録がキャンセルされました');
                // 画像を削除して元の状態に戻す
                const parentQr = elImage.parentNode;
                const dropzone = parentQr.querySelector('.dropzone, div[class="dropzone"]');
                
                if (dropzone) {
                  // 既存のドロップゾーンを復元
                  dropzone.classList.add('dropzone');
                  dropzone.style.zIndex = '99';
                  dropzone.innerHTML = '';
                  dropzone.appendChild(buildQRPastePlaceholder());
                  dropzone.style.display = 'block';
                } else {
                  // ドロップゾーンが見つからない場合は新しく作成
                  const newDropzone = document.createElement('div');
                  newDropzone.className = 'dropzone';
                  newDropzone.contentEditable = 'true';
                  newDropzone.setAttribute('effectallowed', 'move');
                  newDropzone.style.zIndex = '99';
                  newDropzone.innerHTML = '';
                  newDropzone.appendChild(buildQRPastePlaceholder());
                  parentQr.appendChild(newDropzone);
                  
                  // ドロップゾーンのイベントリスナーを再設定
                  setupPasteZoneEvents(newDropzone);
                }
                
                // 画像を削除
                elImage.parentNode.removeChild(elImage);
                return;
              }
            }
            
            const d = elImage.closest("td").querySelector(".yamato");
            d.innerHTML = "";
            
            for(let i = 1; i < CONSTANTS.QR.EXPECTED_PARTS; i++){
              const p = document.createElement("p");
              p.innerText = b[i];
              d.appendChild(p);
            }
            
            const qrData = {
              receiptnum: b[1],
              receiptpassword: b[2],
              qrimage: arrayBuffer,
              qrimageType: 'image/png',
              qrhash: (function(content){ let hash=0; if(!content) return hash.toString(); for(let i=0;i<content.length;i++){ const ch=content.charCodeAt(i); hash=((hash<<5)-hash)+ch; hash&=hash; } return hash.toString(); })(barcode.data)
            };
            if (window.orderRepository) {
              await window.orderRepository.setOrderQRData(ordernum, qrData);
            }
          } else {
            console.warn('QRコードの形式が正しくありません');
          }
        }
      } catch (error) {
        console.error('QRコード処理エラー:', error);
      }
    };
    
    img.onerror = function() {
      console.error('画像の読み込みに失敗しました');
      alert('画像の取得に失敗しました。外部画像の場合はCORS制限の可能性があります。Ctrl+Vで貼り付けるか、ローカル画像を使用してください。');
    };
  } catch (error) {
    console.error('QR読み取り関数エラー:', error);
  }
}

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
    containerClass = 'order-image-drop',
    defaultMessage = '画像をドロップ or クリックで選択'
  } = options;

  debugLog(`画像ドロップゾーン作成: 個別=${isIndividual} 注文番号=${orderNumber||'-'}`);

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
  // グローバル（非個別）時は settings から復元
  if (!orderNumber && !isIndividual && !savedImage) {
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
        let individualImage = null;
        if (window.orderRepository) {
          const r = window.orderRepository.get(orderNumber);
          if (r && r.image && r.image.data instanceof ArrayBuffer) {
            try { const blob = new Blob([r.image.data], { type: r.image.mimeType || 'image/png' }); individualImage = URL.createObjectURL(blob); } catch(e){ console.error('個別画像Blob生成失敗', e); }
          }
  }
  const globalImage = await getGlobalOrderImage();
        
        debugLog(`注文番号: ${orderNumber}`);
        debugLog(`個別画像: ${individualImage ? 'あり' : 'なし'}`);
        debugLog(`グローバル画像: ${globalImage ? 'あり' : 'なし'}`);
        
        if (individualImage) {
          debugLog(`個別画像を優先使用: ${orderNumber}`);
          imageToShow = individualImage;
        } else {
          // 個別画像がない場合のみグローバル画像を使用
          debugLog(`個別画像なし、グローバル画像を使用: ${orderNumber}`);
          imageToShow = globalImage;
        }
      } else {
        // 注文番号がない場合はグローバル画像を使用
  const globalImage = await getGlobalOrderImage();
        debugLog('注文番号なし、グローバル画像を使用', globalImage ? 'あり' : 'なし');
        imageToShow = globalImage;
      }

      // 画像コンテナを更新
      imageContainer.innerHTML = '';
      if (imageToShow) {
        const imageDiv = document.createElement('div');
        imageDiv.classList.add('order-image');
        const img = document.createElement('img');
        img.src = imageToShow;
        imageDiv.appendChild(img);
        imageContainer.appendChild(imageDiv);
      }
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
      await updateOrderImageDisplay(imageUrl);
    } else if (!isIndividual) {
      // グローバル画像の場合は全ての注文明細の画像を更新
      await updateAllOrderImages();
    }

    // 画像クリックでリセット
    preview.addEventListener('click', async (e) => {
      e.stopPropagation();
      if (isIndividual && orderNumber && window.orderRepository) {
        try { await window.orderRepository.clearOrderImage(orderNumber); } catch(e){ console.error('画像削除失敗', e); }
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
  await updateOrderImageDisplay(null);
      } else {
  dropZone.innerHTML = '';
  dropZone.appendChild(cloneTemplate('orderImageDropDefault'));
  await updateAllOrderImages();
      }
    });
  }

  // 個別画像用の表示更新関数
  async function updateOrderImageDisplay(imageUrl) {
    // 注文画像表示機能が無効の場合は何もしない
    const settings = await StorageManager.getSettingsAsync();
    if (!settings.orderImageEnable) {
      return;
    }

    const orderSection = dropZone.closest('section');
    if (!orderSection) return;

    const imageContainer = orderSection.querySelector('.order-image-container');
    if (!imageContainer) return;

    imageContainer.innerHTML = '';

    if (imageUrl) {
      const imageDiv = document.createElement('div');
      imageDiv.classList.add('order-image');
      const img = document.createElement('img');
      img.src = imageUrl;
      imageDiv.appendChild(img);
      imageContainer.appendChild(imageDiv);
    } else {
      // 個別画像がない場合はグローバル画像を表示
      const globalImage = await getGlobalOrderImage();
      if (globalImage) {
        const imageDiv = document.createElement('div');
        imageDiv.classList.add('order-image');
        const img = document.createElement('img');
        img.src = globalImage;
        imageDiv.appendChild(img);
        imageContainer.appendChild(imageDiv);
      }
    }
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
  const fileSelectedInfoCompact = document.getElementById('fileSelectedInfoCompact');
  if (fileSelectedInfoCompact) {
    if (fileInput.files && fileInput.files.length > 0) {
      const fileName = fileInput.files[0].name;
      // コンパクト表示用に短縮
      const shortName = fileName.length > 15 ? fileName.substring(0, 12) + '...' : fileName;
      fileSelectedInfoCompact.textContent = shortName;
      fileSelectedInfoCompact.classList.add('has-file');
      updateProcessedOrdersVisibility();
  hideQuickGuide();
      
      // CSVファイルが選択されたら自動的に処理を実行
      console.log('CSVファイルが選択されました。自動処理を開始します:', fileName);
      await autoProcessCSV();
    } else {
      fileSelectedInfoCompact.textContent = '未選択';
      fileSelectedInfoCompact.classList.remove('has-file');
      updateProcessedOrdersVisibility();
  showQuickGuide();
      
      // ファイルがクリアされた場合は結果もクリア
      clearPreviousResults();
      window.lastOrderSelection = [];
      window.lastPreviewConfig = null;
      window.currentDisplayedOrderNumbers = [];
    }
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
  const printButton = document.getElementById('printButton');
  const printButtonCompact = document.getElementById('printButtonCompact');
  const printBtn = document.getElementById('print-btn'); // 新しい印刷ボタン
  
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
    
    // CSV行数を取得（注文明細の数）
    const orderPages = document.querySelectorAll(".page");
    const csvRowCount = orderPages.length;
    totalUsedLabels += csvRowCount;
    
    // 有効なカスタムラベル面数を取得
    if (document.getElementById("customLabelEnable").checked) {
  const customLabels = CustomLabels.getFromUI();
      const enabledCustomLabels = customLabels.filter(label => label.enabled);
      const totalCustomCount = enabledCustomLabels.reduce((sum, label) => sum + label.count, 0);
      totalUsedLabels += totalCustomCount;
    }
    
    // 全体の使用面数を計算（現在のスキップ + 新たに使用した面数）
    const totalUsedWithSkip = currentSkip + totalUsedLabels;
    
    // 44面シートでの余り面数を計算
    const newSkipValue = totalUsedWithSkip % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET;
    
    console.log(`スキップ枚数更新計算:
      現在のスキップ: ${currentSkip}面
      CSV行数: ${csvRowCount}面
      カスタムラベル: ${totalUsedLabels - csvRowCount}面
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
    
    // CSVファイルが読み込まれている場合は、CSV印刷プレビューも再生成
    const fileInput = document.getElementById("file");
    if (fileInput && fileInput.files && fileInput.files.length > 0) {
      try {
        console.log('📄 CSVファイルが読み込まれているため、印刷プレビューを再生成します...');
        await autoProcessCSV();
        console.log('✅ スキップ枚数更新後のCSV印刷プレビューを更新しました');
      } catch (csvError) {
        console.error('⚠️ CSV印刷プレビュー更新エラー:', csvError);
        // CSV更新エラーは致命的ではないので、処理を継続
      }
    }
    
    // 更新完了メッセージ
    alert(`次回のスキップ枚数を ${newSkipValue} 面に更新しました。\n\n詳細:\n・印刷前スキップ: ${currentSkip}面\n・今回使用: ${totalUsedLabels}面\n・合計: ${totalUsedWithSkip}面\n・次回スキップ: ${newSkipValue}面`);
    
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
