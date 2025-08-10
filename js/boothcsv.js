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

// リアルタイム更新制御フラグ
// isEditingCustomLabel は custom-labels.js で window プロパティとして定義される
// 直近読み込んだCSV結果を保持してカスタムラベル編集時に再レンダリングへ再利用
window.lastCSVResults = null; // { data: [...] }
window.lastCSVBaseConfig = null; // { labelyn, labelskip, sortByPaymentDate }
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

// 注文番号ユーティリティ（軽量化のためオブジェクト廃止）
function getOrderNumberFromCSVRow(row){
  if (!row || !row[CONSTANTS.CSV.ORDER_NUMBER_COLUMN]) return '';
  return String(row[CONSTANTS.CSV.ORDER_NUMBER_COLUMN]).trim();
}
// 表示用の『注文番号 : 』プレフィクスは CSS の .注文番号::before で付与するため
// ここでのフォーマット関数は不要になった。
function isValidOrderNumber(orderNumber){
  return !!(orderNumber && String(orderNumber).trim());
}

// カスタムラベルの複数シート計算ユーティリティ
class CustomLabelCalculator {
  // 複数シートにわたるカスタムラベルの配置計算
  static calculateMultiSheetDistribution(totalLabels, skipCount) {
    const sheetsInfo = [];
    let remainingLabels = totalLabels;
    let currentSkip = skipCount;
    let sheetNumber = 1;
    
    while (remainingLabels > 0) {
      const availableInSheet = CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET - currentSkip;
      const labelsInThisSheet = Math.min(remainingLabels, availableInSheet);
      const remainingInSheet = availableInSheet - labelsInThisSheet;
      
      sheetsInfo.push({
        sheetNumber,
        skipCount: currentSkip,
        labelCount: labelsInThisSheet,
        remainingCount: remainingInSheet,
        totalInSheet: currentSkip + labelsInThisSheet
      });
      
      remainingLabels -= labelsInThisSheet;
      currentSkip = 0; // 2シート目以降はスキップなし
      sheetNumber++;
    }
    
    return sheetsInfo;
  }
  
  // 最終シートの情報を取得
  static getLastSheetInfo(totalLabels, skipCount) {
    const sheetsInfo = this.calculateMultiSheetDistribution(totalLabels, skipCount);
    return sheetsInfo[sheetsInfo.length - 1] || null;
  }
  
  // 総シート数を計算
  static calculateTotalSheets(totalLabels, skipCount) {
    const sheetsInfo = this.calculateMultiSheetDistribution(totalLabels, skipCount);
    return sheetsInfo.length;
  }
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
  orderImageEnable: false
};

window.addEventListener("load", async function(){
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

  Object.assign(window.settingsCache, {
    labelyn: settings.labelyn,
    labelskip: settings.labelskip,
    sortByPaymentDate: settings.sortByPaymentDate,
    customLabelEnable: settings.customLabelEnable,
    orderImageEnable: settings.orderImageEnable
  });

  toggleCustomLabelRow(settings.customLabelEnable);
  toggleOrderImageRow(settings.orderImageEnable);
  console.log('🎉 アプリケーション初期化完了 (fallback 無し)');

  // 複数のカスタムラベルを初期化
  CustomLabels.initialize(settings.customLabels);

  // 固定ヘッダーを初期表示（0枚でも表示）
  updatePrintCountDisplay(0, 0, 0);

   // 画像ドロップゾーンの初期化
  const imageDropZoneElement = document.getElementById('imageDropZone');
  const imageDropZone = await createOrderImageDropZone();
  imageDropZoneElement.appendChild(imageDropZone.element);
  window.orderImageDropZone = imageDropZone;

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
        // OrderRepository に統合されたため、orders ストアを全削除
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

  // 全ての注文画像をクリアするボタンのイベントリスナーを追加
  // 個別クリアボタン廃止: 全注文データクリアに統合済み

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

  // showAllOrders 廃止

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
    const hasCSVLoaded = !!(window.lastCSVResults && window.lastCSVResults.data && window.lastCSVResults.data.length > 0);

    // CSV が読み込まれている場合: CSVラベル + (必要なら) カスタムラベル を再生成
    if (hasCSVLoaded) {
      if (!config.labelyn) {
        // ラベル印刷OFFなら表示だけクリア
        clearPreviousResults();
        updatePrintCountDisplay(0, 0, 0);
        return;
      }
      // 再生成（processCSVResults 内で clearPreviousResults していないので先に消す）
      clearPreviousResults();
      // ベース構成を更新して再利用
      window.lastCSVBaseConfig = { labelyn: config.labelyn, labelskip: config.labelskip, sortByPaymentDate: config.sortByPaymentDate };
      await processCSVResults(window.lastCSVResults, config);
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

// scheduleDelayedPreviewUpdate は CustomLabels.schedulePreview に移動

// 静かなバリデーション関数（アラート表示なし）
// validateCustomLabelsQuiet は custom-labels.js へ移動（後方互換のためグローバル関数は存続）

function clearPreviousResults() {
  for (let sheet of document.querySelectorAll('section')) {
    sheet.parentNode.removeChild(sheet);
  }
  
  // 印刷枚数表示もクリア
  clearPrintCountDisplay();
}

// collectConfig 廃止：settingsCache を利用

async function processCSVResults(results, config) {
  // --- Stage B: OrderRepository 利用 ---
  const db = await StorageManager.ensureDatabase();
  if (!window.orderRepository) {
    if (typeof OrderRepository === 'undefined') {
      console.error('OrderRepository 未読込');
    } else {
      window.orderRepository = new OrderRepository(db);
      await window.orderRepository.init();
    }
  }
  if (!window.orderRepository) return; // フェイルセーフ

  // 直近CSV保持（プレビュー再描画用）
  try {
    window.lastCSVResults = { data: Array.isArray(results.data) ? results.data : [] };
    window.lastCSVBaseConfig = { labelyn: config.labelyn, labelskip: config.labelskip, sortByPaymentDate: config.sortByPaymentDate };
  } catch(e) { /* ignore */ }

  // デバッグ: 先頭行の列キー確認（BOM混入/名称ズレ検出用）
  if (DEBUG_MODE && results && Array.isArray(results.data)) {
    const first = results.data[0];
    if (first) {
  debugLog('[csv] 先頭行キー一覧', Object.keys(first));
    } else {
  debugLog('[csv] CSVにデータ行がありません');
    }
  }

  // CSV の順序（= BOOTH ダウンロードの注文番号降順）を保持したまま repository に反映
  await window.orderRepository.bulkUpsert(results.data);
  const csvOrderKeys = [];
  let debugValidOrderCount = 0;
  for (const row of results.data) {
  const num = getOrderNumberFromCSVRow(row);
    if (!num) continue;
    csvOrderKeys.push(OrderRepository.normalize(num));
    debugValidOrderCount++;
    if (DEBUG_MODE) {
      const normalized = OrderRepository.normalize(num);
      const rec = window.orderRepository.get(normalized);
  debugLog('[repo] 読み込み注文', { raw:num, normalized, exists: !!rec, printedAt: rec ? rec.printedAt : null });
    }
  }
  if (DEBUG_MODE) {
  debugLog('[csv] 行数サマリ', {
      totalRows: results.data.length,
      withOrderNumber: debugValidOrderCount,
      repositoryStored: window.orderRepository.getAll().length
    });
  }
  // 表示対象は今回の CSV に含まれる注文のみ（従来挙動に合わせ、過去 CSV の注文は表示しない）
  const orderObjs = csvOrderKeys.map(k => window.orderRepository.get(k)).filter(o => !!o);
  const unprinted = orderObjs.filter(o => !o.printedAt);
  let detailRows = orderObjs.map(o => o.row); // 表示用（CSV 順維持）
  let labelRows = unprinted.map(o => o.row); // 未印刷のみ（CSV 順維持）
  const csvRowCountForLabels = labelRows.length;
  // 複数カスタムラベルの総面数を計算
  const totalCustomLabelCount = config.customLabels.reduce((sum, label) => sum + label.count, 0);
  // 複数シート対応：1シートの制限を撤廃
  // CSVデータとカスタムラベルの合計で必要なシート数を計算
  const skipCount = parseInt(settingsCache.labelskip, 10) || 0;
  const totalLabelsNeeded = skipCount + csvRowCountForLabels + totalCustomLabelCount;
  const requiredSheets = Math.ceil(totalLabelsNeeded / CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET);

  // データの並び替え
  if (config.sortByPaymentDate) {
    // 支払い日時ソート有効時のみ上書き
    labelRows.sort((a, b) => {
      const timeA = a[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      const timeB = b[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      return timeA.localeCompare(timeB);
    });
    detailRows.sort((a, b) => {
      const timeA = a[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      const timeB = b[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      return timeA.localeCompare(timeB);
    });
  }

  // 注文明細の生成
  // ラベル対象の注文番号セットを作成
  // repository から未印刷を判定済みなので、labelRows に含まれる行の orderNumber を repository キャッシュから逆引き
  const labelSet = new Set();
  if (window.orderRepository) {
    const all = window.orderRepository.getAll();
    // row オブジェクト参照比較で対応（CSV parse 再利用時に新インスタンスなら fallback）
    const rowToOrder = new Map();
    for (const rec of all) {
      if (rec.row) rowToOrder.set(rec.row, rec.orderNumber);
    }
    for (const r of labelRows) {
  const num = rowToOrder.get(r) || getOrderNumberFromCSVRow(r);
      if (num) labelSet.add(String(num).trim());
    }
  } else {
    for (const r of labelRows) {
  const num = getOrderNumberFromCSVRow(r);
      if (num) labelSet.add(String(num).trim());
    }
  }
  const baseLabelArr = Array(skipCount).fill("");
  await generateOrderDetails(detailRows, baseLabelArr, labelSet);

  // 各注文明細パネルはgenerateOrderDetails内で個別に更新済み

  // ラベル生成（注文分＋カスタムラベル）- 複数シート対応
  if (config.labelyn) {
  let totalLabelArray = [...baseLabelArr];

    // 明細表示順 (detailRows の順) に合わせて repository から未印刷のみ追加
    const repo = window.orderRepository;
    const numbersInOrder = detailRows
  .map(r => getOrderNumberFromCSVRow(r))
      .map(n => OrderRepository.normalize(n))
      .filter(n => {
        const rec = repo ? repo.get(n) : null; return rec && !rec.printedAt;
      });
    totalLabelArray.push(...numbersInOrder);
    // 表示対象注文番号リストをグローバルに保持（再計算用）
    window.currentDisplayedOrderNumbers = detailRows
  .map(r => getOrderNumberFromCSVRow(r))
      .map(n => OrderRepository.normalize(n));

    // カスタムラベルが有効な場合は追加
    if (config.customLabelEnable && config.customLabels.length > 0) {
      for (const customLabel of config.customLabels) {
        for (let i = 0; i < customLabel.count; i++) {
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
      unprintedCount: unprinted.length,
      detailCount: detailRows.length,
      labelRowCount: labelRows.length,
      customLabelCount: totalCustomLabelCount,
      totalLabelArrayLength: (typeof totalLabelArray !== 'undefined') ? totalLabelArray.length : 0,
      requiredSheets
    });
  }

  // 印刷枚数の表示（複数シート対応）
  // showCSVWithCustomLabelPrintSummary(csvRowCount, totalCustomLabelCount, skipCount, requiredSheets);

  // ヘッダーの印刷枚数表示を更新
  // ラベル印刷がオフの場合はラベル枚数・カスタム面数ともに0を表示する
  const labelSheetsForDisplay = config.labelyn ? requiredSheets : 0;
  const customFacesForDisplay = (config.labelyn && config.customLabelEnable) ? totalCustomLabelCount : 0;
  // 普通紙（注文明細）は未印刷のみ
  updatePrintCountDisplay(unprinted.length, labelSheetsForDisplay, customFacesForDisplay);

  // CSV処理完了後のカスタムラベルサマリー更新（複数シート対応）
  await CustomLabels.updateSummary();

  // ボタンの状態を更新
  CustomLabels.updateButtonStates();
}

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
  
  // 印刷枚数の表示（複数シート対応）
  // showMultiSheetCustomLabelPrintSummary(totalCustomLabelCount, labelskipNum, sheetsInfo);
  
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
  
  console.log(`updatePrintCountDisplay呼び出し: ラベル:${labelSheetCount}枚, 普通紙:${orderSheetCount}枚, カスタム:${customLabelCount}面`);
  
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
  
  console.log(`印刷枚数更新完了: ラベル:${labelSheetCount}枚, 普通紙:${orderSheetCount}枚, カスタム:${customLabelCount}面`);
}

// 印刷枚数をクリアする関数
function clearPrintCountDisplay() {
  updatePrintCountDisplay(0, 0, 0);
}

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
        async (val) => { await updateAllOrderImagesVisibility(val); }
      ] },
    // customLabelEnable はカスタムラベルUI群と関係が深く遅延初期化 setupCustomLabelEvents() 内に既存処理があるためここでは扱わない
  ];

  for (const def of defs) {
    const el = document.getElementById(def.id);
    if (!el) continue;
    const keyConst = StorageManager.KEYS[def.key];
    el.addEventListener('change', async function() {
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
      await autoProcessCSV();
    });

    if (def.id === 'labelskipnum') {
  el.addEventListener('input', function() { CustomLabels.updateButtonStates(); });
    }
  }
}

// 以前は設定変更時にスクロール位置復元をしていたが、設定UIはページ上部のみで再描画影響が小さいためロジック削除

// 現在の「読み込んだファイル全て表示」のON/OFFを返す
// showAllOrders 廃止

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
  // A: DOM 走査を廃止し、表示中注文番号リスト + repository のみで未印刷抽出
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
    // 既存 label-sheet は既に削除済みなので枚数再計算のみ
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

async function createIndividualImageDropZone(cOrder, orderNumber) {
  debugLog(`個別画像ドロップゾーン作成開始 - 注文番号: "${orderNumber}"`);
  
  const individualDropZoneContainer = cOrder.querySelector('.individual-image-dropzone');
  const individualZone = cOrder.querySelector('.individual-order-image-zone');
  
  debugLog(`ドロップゾーンコンテナ発見: ${!!individualDropZoneContainer}`);
  debugLog(`個別ゾーン発見: ${!!individualZone}`);
  
  // 注文画像表示機能が無効の場合は個別画像ゾーン全体を非表示
  const settings = await StorageManager.getSettingsAsync();
  debugLog(`注文画像表示設定: ${settings.orderImageEnable}`);
  
  if (!settings.orderImageEnable) {
    if (individualZone) {
      individualZone.style.display = 'none';
      debugLog('注文画像表示が無効のため個別ゾーンを非表示');
    }
    return;
  }

  // 有効な場合は表示
  if (individualZone) {
    individualZone.style.display = 'block';
    debugLog('注文画像表示が有効のため個別ゾーンを表示');
  }

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

function parseProductItemData(itemrow) {
  const firstSplit = itemrow.split(' / ');
  const itemIdSplit = firstSplit[0].split(':');
  const itemId = itemIdSplit.length > 1 ? itemIdSplit[1].trim() : '';
  const quantitySplit = firstSplit[1] ? firstSplit[1].split(':') : [];
  const quantity = quantitySplit.length > 1 ? quantitySplit[1].trim() : '';
  const productName = firstSplit.slice(2).join(' / ');
  
  return { itemId, quantity, productName };
}

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

async function displayOrderImage(cOrder, orderNumber) {
  // 注文画像表示機能が無効の場合は何もしない
  const settings = await StorageManager.getSettingsAsync();
  if (!settings.orderImageEnable) {
    return;
  }

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

// 旧: グローバル印刷日時パネルは廃止（各注文明細内に移行）

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

// 以下の関数は廃止されました（印刷枚数は固定ヘッダーにリアルタイム表示）
// function showPrintSummary() { ... }
// function showCustomLabelPrintSummary() { ... }
// function showMultiSheetCustomLabelPrintSummary() { ... }
// function showCSVWithCustomLabelPrintSummary() { ... }

function addP(div, text){
  const p = document.createElement("p");
  p.innerText = text;
  div.appendChild(p);
}
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
// QRペーストプレースホルダを template から生成
function buildQRPastePlaceholder() {
  return cloneTemplate('qrDropPlaceholder');
}
function setupPasteZoneEvents(dropzone) {
  // クリップボードからの画像ペーストのみを受け付ける
  dropzone.addEventListener("paste", function(event){
    event.preventDefault();
    if (!event.clipboardData 
            || !event.clipboardData.types
            || (!event.clipboardData.types.includes("Files"))) {
            return true;
    }
    
    for(let item of event.clipboardData.items){
      if(item["kind"] == "file"){
        const imageFile = item.getAsFile();
        const fr = new FileReader();
        fr.onload = function(e) {
          const elImage = document.createElement('img');
          elImage.src = e.target.result;
          elImage.onload = function() {
            dropzone.parentNode.appendChild(elImage);
            readQR(elImage);
            dropzone.classList.remove('dropzone');
            dropzone.style.zIndex = -1;
            elImage.style.zIndex = 9;
            addEventQrReset(elImage);
          };
        };
        fr.readAsDataURL(imageFile);
      }
    }
    
    // 画像以外がペーストされたときのために、元に戻しておく
    this.textContent = '';
  this.innerHTML = '';
  this.appendChild(buildQRPastePlaceholder());
  });
}

function createDropzone(div){ // 互換のため名称維持（内部はペースト専用）
  const divDrop = createDiv('dropzone');
  divDrop.appendChild(buildQRPastePlaceholder());
  divDrop.setAttribute("contenteditable", "true");
  // ペースト専用イベントを設定
  setupPasteZoneEvents(divDrop);
  
  div.appendChild(divDrop);
}

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
    addP(divOrdernum, labelData);
    const repo = window.orderRepository || null;
    const rec = repo ? repo.get(labelData) : null;
    const qr = rec ? rec.qr : null;
    if (qr && qr['qrimage']) {
      const elImage = document.createElement('img');
      let srcValue = qr['qrimage'];
      if (qr.isBinary && qr.qrimage instanceof ArrayBuffer) {
        try {
          const blob = new Blob([qr.qrimage], { type: qr.qrimageType || 'image/png' });
          srcValue = URL.createObjectURL(blob);
          elImage.addEventListener('error', () => console.error('QR画像Blob URL読み込み失敗'));
        } catch (e) {
          console.error('QR画像Blob生成失敗', e);
        }
      }
      elImage.src = srcValue;
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

// ドラッグ＆ドロップ廃止に伴い showDropping / hideDropping は削除

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
            const rawOrderNum = elImage.closest("td").querySelector(".ordernum p").innerHTML;
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
    };
  } catch (error) {
    console.error('QR読み取り関数エラー:', error);
  }
}

// drag&drop 廃止に伴い attachImage は不要となったため削除
// QRコード用ペーストゾーンのみドラッグ&ドロップを抑止し、他領域（注文画像/フォント等）は従来どおり許可
document.addEventListener('dragover', e => {
  const target = e.target instanceof HTMLElement ? e.target.closest('.dropzone') : null;
  if (target) {
    e.preventDefault();
  }
});
document.addEventListener('drop', e => {
  const target = e.target instanceof HTMLElement ? e.target.closest('.dropzone') : null;
  if (target) {
    e.preventDefault();
    console.warn('QRコード領域ではドラッグ＆ドロップ不可。Ctrl+V で貼り付けてください。');
  }
});

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
// 以前の base64 JSON 方式は廃止。Blob URL を都度生成（簡易実装）。必要ならキャッシュ最適化可。
async function setGlobalOrderImage(arrayBuffer, mimeType='image/png') {
  try {
    if(!(arrayBuffer instanceof ArrayBuffer)) throw new Error('ArrayBuffer 以外');
    await StorageManager.setGlobalOrderImageBinary(arrayBuffer, mimeType);
  debugLog('[image] グローバル画像保存完了 size=' + arrayBuffer.byteLength + ' mime=' + mimeType);
  } catch(e){ console.error('グローバル画像保存失敗', e); }
}
async function getGlobalOrderImage(){
  try {
    const v = await StorageManager.getGlobalOrderImageBinary();
    if(!v || !(v.data instanceof ArrayBuffer)) return null;
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
  // フォールバック廃止 (images ストア削除)
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
  const node = cloneTemplate('orderImageDropDefault');
      node.textContent = defaultMessage; // 個別はメッセージ差替
      dropZone.appendChild(node);
    } else {
  dropZone.innerHTML = '';
  dropZone.appendChild(cloneTemplate('orderImageDropDefault'));
    }
    debugLog(`初期メッセージを設定: ${isIndividual ? defaultMessage : 'デフォルトコンテンツ'}`);
  }

  // 全ての注文明細の画像を更新する関数
  async function updateAllOrderImages() {
    // 注文画像表示機能が無効の場合は何もしない
    const settings = await StorageManager.getSettingsAsync();
    if (!settings.orderImageEnable) {
      return;
    }

    const allOrderSections = document.querySelectorAll('section');
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
        // グローバル画像クリア (空バッファ)
        try { await StorageManager.setGlobalOrderImageBinary(new ArrayBuffer(0), 'application/octet-stream'); } catch(e){ console.error('グローバル画像削除失敗', e); }
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

async function createOrderImageDropZone() {
  return await createBaseImageDropZone({
    storageKey: 'orderImage',
    isIndividual: false,
    containerClass: 'order-image-drop'
  });
}

// 個別注文用の画像ドロップゾーンを作成する関数（リファクタリング済み）
async function createIndividualOrderImageDropZone(orderNumber) {
  return await createBaseImageDropZone({
    storageKey: `orderImage_${orderNumber}`,
    isIndividual: true,
    orderNumber: orderNumber,
    containerClass: 'individual-order-image-drop',
    defaultMessage: '画像をドロップ or クリックで選択'
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
      
      // CSVファイルが選択されたら自動的に処理を実行
      console.log('CSVファイルが選択されました。自動処理を開始します:', fileName);
      await autoProcessCSV();
    } else {
      fileSelectedInfoCompact.textContent = '未選択';
      fileSelectedInfoCompact.classList.remove('has-file');
      
      // ファイルがクリアされた場合は結果もクリア
      clearPreviousResults();
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

// カスタムラベル機能の関数群
function toggleCustomLabelRow(enabled) {
  const customLabelRow = document.getElementById('customLabelRow');
  customLabelRow.style.display = enabled ? 'table-row' : 'none';
}

// 注文画像表示機能の関数群
function toggleOrderImageRow(enabled) {
  const orderImageRow = document.getElementById('orderImageRow');
  orderImageRow.style.display = enabled ? 'table-row' : 'none';
}

// 全ての注文明細の画像表示可視性を更新
async function updateAllOrderImagesVisibility(enabled) {
  debugLog(`画像表示機能が${enabled ? '有効' : '無効'}に変更されました`);
  
  const allOrderSections = document.querySelectorAll('section');
  for (const orderSection of allOrderSections) {
    const imageContainer = orderSection.querySelector('.order-image-container');
    const individualZone = orderSection.querySelector('.individual-order-image-zone');
    const individualDropZoneContainer = orderSection.querySelector('.individual-image-dropzone');
    
    if (enabled) {
      // 有効な場合：画像を表示し、個別画像ゾーンも表示
      if (individualZone) {
        individualZone.style.display = 'block';
      }
      
      // 統一化された方法で注文番号を取得
  const orderNumber = (orderSection.id && orderSection.id.startsWith('order-')) ? orderSection.id.substring(6) : '';
      
      // 個別画像ドロップゾーンが存在するが中身が空の場合、ドロップゾーンを作成
      if (individualDropZoneContainer && orderNumber && individualDropZoneContainer.children.length === 0) {
        debugLog(`個別画像ドロップゾーンを後から作成: ${orderNumber}`);
        try {
          const individualImageDropZone = await createIndividualOrderImageDropZone(orderNumber);
          if (individualImageDropZone && individualImageDropZone.element) {
            individualDropZoneContainer.appendChild(individualImageDropZone.element);
            debugLog(`個別画像ドロップゾーン作成成功: ${orderNumber}`);
          }
        } catch (error) {
          debugLog(`個別画像ドロップゾーン作成エラー: ${error.message}`);
          console.error('個別画像ドロップゾーン作成エラー:', error);
        }
      }
      
      if (imageContainer) {
        // 画像を表示
        let imageToShow = null;
        if (orderNumber) {
          // 個別画像
          if (window.orderRepository) {
            const rec = window.orderRepository.get(orderNumber);
            if (rec && rec.image && rec.image.data instanceof ArrayBuffer) {
              try { const blob = new Blob([rec.image.data], { type: rec.image.mimeType || 'image/png' }); imageToShow = URL.createObjectURL(blob); debugLog(`個別画像を表示: ${orderNumber}`); } catch {}
            }
          }
          if (!imageToShow) {
            const globalImage = await getGlobalOrderImage();
            if (globalImage) { imageToShow = globalImage; debugLog(`グローバル画像を表示: ${orderNumber}`); }
          }
        }
        
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
    } else {
      // 無効な場合：画像を非表示にし、個別画像ゾーンも非表示
      if (imageContainer) {
        imageContainer.innerHTML = '';
      }
      if (individualZone) {
        individualZone.style.display = 'none';
      }
    }
  }
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
        const orderPages = document.querySelectorAll('.page');
        for (const page of orderPages) {
          const orderNumber = (page.id && page.id.startsWith('order-')) ? page.id.substring(6) : '';
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

// handleFontFile は移動済み

// localStorage ベースの容量チェックは廃止 (IndexedDB 専用化に伴い削除済み)


// (getFontMimeType / addFontToCSS / getFontFormat) は未使用のため削除済み

let _fontFaceLoadToken = 0;
// loadCustomFontsCSS は移動済み

// updateFontList は移動済み

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

// removeFontFromList は移動済み

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

// cleanupEmptySpans は custom-labels.js (CustomLabelStyle) に統合済み

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

// 成功メッセージ表示ヘルパー
// showSuccessMessage は移動済み

// ローディング表示ヘルパー
// showFontUploadProgress は移動済み

// CSS用のpulseアニメーションを追加
// フォントアニメーション定義は移動済み
