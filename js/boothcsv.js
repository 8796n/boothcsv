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
let isEditingCustomLabel = false;
let pendingUpdateTimer = null;

// デバッグログ用関数
function debugLog(...args) {
  if (DEBUG_MODE) {
    console.log(...args);
  }
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

// 注文番号処理を統一管理するクラス
class OrderNumberManager {
  // 注文番号の正規化（"注文番号 : 66463556" → "66463556"）
  static normalize(orderNumber) {
    if (!orderNumber || typeof orderNumber !== 'string') {
      return '';
    }
    
    const normalized = orderNumber.replace(/^.*?:\s*/, '').trim();
    return normalized;
  }
  
  // DOM要素から注文番号を取得（注文明細用）
  static getFromOrderSection(orderSection) {
    if (!orderSection) {
      return null;
    }
    
    // 方法1: .注文番号クラスから取得
    const orderNumberElement = orderSection.querySelector('.注文番号');
    if (orderNumberElement) {
      const rawOrderNumber = orderNumberElement.textContent.trim();
      const normalized = this.normalize(rawOrderNumber);
      return normalized;
    }
    
    // 方法2: .ordernum pから取得（ラベル用）
    const ordernumElement = orderSection.querySelector('.ordernum p');
    if (ordernumElement) {
      const rawOrderNumber = ordernumElement.textContent.trim();
      const normalized = this.normalize(rawOrderNumber);
      return normalized;
    }
    
    return null;
  }
  
  // CSV行データから注文番号を取得（表示用フォーマット付き）
  static getFromCSVRow(row) {
    if (!row || !row[CONSTANTS.CSV.ORDER_NUMBER_COLUMN]) {
      return '';
    }
    
    const orderNumber = row[CONSTANTS.CSV.ORDER_NUMBER_COLUMN];
    return orderNumber;
  }
  
  // 表示用フォーマットを生成（"注文番号 : 66463556"）
  static createDisplayFormat(orderNumber) {
    if (!orderNumber) {
      return '';
    }
    
    // 既に表示用フォーマットの場合はそのまま返す
    if (orderNumber.includes('注文番号')) {
      return orderNumber;
    }
    
    const formatted = `注文番号 : ${orderNumber}`;
    return formatted;
  }
  
  // 注文番号の妥当性チェック
  static isValid(orderNumber) {
    const normalized = this.normalize(orderNumber);
    const isValid = normalized && normalized.length > 0;
    return isValid;
  }
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
        complete: function(results) {
          const rowCount = results.data.length;
          resolve(rowCount);
        },
        error: function(error) {
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
      return {
        rowCount,
        fileName: file.name,
        fileSize: file.size
      };
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

// グローバルなフォントマネージャー（UnifiedDatabaseを使用）
let fontManager = null;

// 統合フォント管理クラス
class FontManager {
  constructor(unifiedDB) {
    this.unifiedDB = unifiedDB;
  }

  // フォントを保存
  async saveFont(fontName, fontData, metadata = {}) {
    const fontObject = {
      name: fontName,
      data: fontData, // ArrayBufferを直接保存
      type: metadata.type || 'font/ttf',
      originalName: metadata.originalName || fontName,
      size: fontData.byteLength || fontData.length,
      createdAt: Date.now()
    };
    
    await this.unifiedDB.setFont(fontName, fontObject);
    console.log(`フォント保存完了: ${fontName}`);
    return fontObject;
  }

  // フォントを取得
  async getFont(fontName) {
    return await this.unifiedDB.getFont(fontName);
  }

  // すべてのフォントを取得（オブジェクト形式で返す）
  async getAllFonts() {
    const fonts = await this.unifiedDB.getAllFonts();
    const fontMap = {};
    
    // 配列をオブジェクトマップに変換
    fonts.forEach(font => {
      if (font && font.name) {
        fontMap[font.name] = {
          data: font.data,
          metadata: {
            type: font.type,
            originalName: font.originalName,
            size: font.size,
            createdAt: font.createdAt
          }
        };
      }
    });
    
    return fontMap;
  }

  // フォントを削除
  async deleteFont(fontName) {
    await this.unifiedDB.deleteFont(fontName);
    console.log(`フォント削除完了: ${fontName}`);
  }

  // すべてのフォントを削除
  async clearAllFonts() {
    await this.unifiedDB.clearAllFonts();
    console.log('全フォント削除完了');
  }

  // ストレージ使用量を取得
  async getStorageInfo() {
    const fonts = await this.unifiedDB.getAllFonts();
    const totalSize = fonts.reduce((sum, font) => sum + (font.size || 0), 0);
    const fontCount = fonts.length;
    
    return {
      fontCount,
      totalSize,
      totalSizeMB: Math.round(totalSize / (1024 * 1024) * 100) / 100,
      fonts: fonts.map(f => ({
        name: f.name,
        size: f.size,
        sizeMB: Math.round(f.size / (1024 * 1024) * 100) / 100,
        createdAt: f.createdAt
      }))
    };
  }
}

// フォントマネージャーを初期化
async function initializeFontManager() {
  try {
    if (!window.unifiedDB) {
      await initializeUnifiedDatabase();
    }
    fontManager = new FontManager(window.unifiedDB);
    console.log('フォントマネージャー初期化完了');
    return fontManager;
  } catch (error) {
    console.error('フォントマネージャー初期化失敗:', error);
    alert('フォント管理システムの初期化に失敗しました。');
    return null;
  }
}

// 統合ストレージ管理クラス（UnifiedDatabaseのラッパー）
// （重複回避）StorageManager は storage.js のものを使用します

// 初期化処理（破壊的移行対応）
window.addEventListener("load", async function(){
  let settings;
  
  try {
    // StorageManagerを通じて統合データベースを初期化（重複回避）
    await StorageManager.ensureDatabase();
    
    // 設定の取得（非同期）
    settings = await StorageManager.getSettingsAsync();
    
    document.getElementById("labelyn").checked = settings.labelyn;
    document.getElementById("labelskipnum").value = settings.labelskip;
  // showAllOrders 廃止
    document.getElementById("sortByPaymentDate").checked = settings.sortByPaymentDate;
    document.getElementById("customLabelEnable").checked = settings.customLabelEnable;
    document.getElementById("orderImageEnable").checked = settings.orderImageEnable;

    // カスタムラベル行の表示/非表示
    toggleCustomLabelRow(settings.customLabelEnable);

    // 注文画像行の表示/非表示
    toggleOrderImageRow(settings.orderImageEnable);

    console.log('🎉 アプリケーション初期化完了');
    
  } catch (error) {
    console.error('初期化エラー:', error);
    
    // フォールバック: 従来の方式
    console.log('フォールバック初期化を実行');
    settings = StorageManager.getDefaultSettings();
    
    document.getElementById("labelyn").checked = settings.labelyn;
    document.getElementById("labelskipnum").value = settings.labelskip;
  // showAllOrders 廃止
    document.getElementById("sortByPaymentDate").checked = settings.sortByPaymentDate;
    document.getElementById("customLabelEnable").checked = settings.customLabelEnable;
    document.getElementById("orderImageEnable").checked = settings.orderImageEnable;

    toggleCustomLabelRow(settings.customLabelEnable);
    toggleOrderImageRow(settings.orderImageEnable);
  }

  // 複数のカスタムラベルを初期化
  initializeCustomLabels(settings.customLabels);

  // 固定ヘッダーを初期表示（0枚でも表示）
  updatePrintCountDisplay(0, 0, 0);

   // 画像ドロップゾーンの初期化
  const imageDropZoneElement = document.getElementById('imageDropZone');
  const imageDropZone = await createOrderImageDropZone();
  imageDropZoneElement.appendChild(imageDropZone.element);
  window.orderImageDropZone = imageDropZone;

  // フォントドロップゾーンの初期化
  initializeFontDropZone();
  
  // フォントセクションの初期状態設定
  await initializeFontSection();
  
  // フォントマネージャーの初期化と カスタムフォントのCSS読み込み（非同期で実行）
  setTimeout(async () => {
    try {
      await initializeFontManager();
      await loadCustomFontsCSS();
    } catch (error) {
      console.warn('フォント初期化エラー:', error);
    }
  }, 100); // 少し遅らせて確実にunifiedDBが初期化されるのを待つ

  // 全ての画像をクリアするボタンのイベントリスナーを追加
  const clearAllButton = document.getElementById('clearAllButton');
  if (clearAllButton) {
    clearAllButton.onclick = async () => {
      if (confirm('本当に全てのQR画像をクリアしますか？')) {
  try { const count = await StorageManager.clearQRImages(); alert(`QR画像を削除しました: ${count}件`); } catch (e) { alert('QR画像クリア中にエラーが発生しました: ' + e.message); }
  location.reload();
      }
    };
  }

  // 全ての注文画像をクリアするボタンのイベントリスナーを追加
  const clearAllOrderImagesButton = document.getElementById('clearAllOrderImagesButton');
  if (clearAllOrderImagesButton) {
    clearAllOrderImagesButton.onclick = async () => {
      if (confirm('本当に全ての注文画像（グローバル画像と個別画像）をクリアしますか？')) {
  try { const count = await StorageManager.clearOrderImages(); alert(`注文画像を削除しました: ${count}件`); } catch (e) { alert('注文画像クリア中にエラーが発生しました: ' + e.message); }
  location.reload();
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

   // チェックボックスの状態が変更されたときにStorageManagerに保存 + 自動再処理
  document.getElementById("labelyn").addEventListener("change", async function() {
    const restoreScroll = captureAndRestoreScrollPosition();
    await StorageManager.set(StorageManager.KEYS.LABEL_SETTING, this.checked);
    await autoProcessCSV(); // 設定変更時に自動再処理
    restoreScroll();
  });

  document.getElementById("labelskipnum").addEventListener("change", async function() {
    const restoreScroll = captureAndRestoreScrollPosition();
    await StorageManager.set(StorageManager.KEYS.LABEL_SKIP, parseInt(this.value, 10) || 0);
    await autoProcessCSV(); // 設定変更時に自動再処理
    restoreScroll();
  });

  document.getElementById("sortByPaymentDate").addEventListener("change", async function() {
    const restoreScroll = captureAndRestoreScrollPosition();
    await StorageManager.set(StorageManager.KEYS.SORT_BY_PAYMENT, this.checked);
    await autoProcessCSV(); // 設定変更時に自動再処理
    restoreScroll();
  });

  // showAllOrders 廃止

   // 注文画像表示機能のイベントリスナー
   document.getElementById("orderImageEnable").addEventListener("change", async function() {
     await StorageManager.set(StorageManager.KEYS.ORDER_IMAGE_ENABLE, this.checked);
     toggleOrderImageRow(this.checked);
     
     // 画像表示をリアルタイムで更新
     await updateAllOrderImagesVisibility(this.checked);
     
     // 設定変更時に自動再処理
     await autoProcessCSV();
   });

  // カスタムラベル機能のイベントリスナー（遅延実行）
  setTimeout(function() {
    setupCustomLabelEvents();
  }, 100);

  // ボタンの初期状態を設定
  updateButtonStates();

  // スキップ数変更時の処理を追加
  document.getElementById("labelskipnum").addEventListener("input", function() {
    updateButtonStates();
  });

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
      const hasValidCustomLabels = validateCustomLabelsQuiet();
      if (!hasValidCustomLabels) {
        console.log('カスタムラベルにエラーがありますが、CSV処理は継続します。');
      }
      
      console.log('自動CSV処理を開始します...');
      clearPreviousResults();
      const config = await getConfigFromUI();
      
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
    const config = await getConfigFromUI();
    
    // ラベル印刷が無効またはカスタムラベルが無効な場合
    if (!config.labelyn || !config.customLabelEnable) {
      clearPreviousResults();
      return;
    }

    // 有効なカスタムラベルがある場合のみプレビューを生成
    if (config.customLabels && config.customLabels.length > 0) {
      const enabledLabels = config.customLabels.filter(label => label.enabled && label.text.trim() !== '');
      
      if (enabledLabels.length > 0) {
        // 既存の結果をクリアしてからプレビューを生成
        clearPreviousResults();
        // カスタムラベルのみの処理を実行（プレビュー用）
        await processCustomLabelsOnly(config, true); // 第2引数でプレビューモードを指定
      } else {
        // 有効なカスタムラベルがない場合は結果をクリアし、固定ヘッダーも更新
        clearPreviousResults();
        updatePrintCountDisplay(0, 0, 0); // カスタム面数を0にリセット
      }
    } else {
      // カスタムラベルが存在しない場合も結果をクリアし、固定ヘッダーを更新
      clearPreviousResults();
      updatePrintCountDisplay(0, 0, 0); // カスタム面数を0にリセット
    }
  } catch (error) {
    console.error('カスタムラベルプレビュー更新エラー:', error);
  }
}

// 遅延更新を実行する関数
function scheduleDelayedPreviewUpdate(delay = 500) {
  // 既存のタイマーをクリア
  if (pendingUpdateTimer) {
    clearTimeout(pendingUpdateTimer);
  }
  
  // 新しいタイマーを設定
  pendingUpdateTimer = setTimeout(async () => {
    debugLog('遅延プレビュー更新を実行');
    await updateCustomLabelsPreview();
    pendingUpdateTimer = null;
  }, delay);
}

// 静かなバリデーション関数（アラート表示なし）
function validateCustomLabelsQuiet() {
  const customLabelEnable = document.getElementById('customLabelEnable').checked;
  
  if (!customLabelEnable) {
    return true; // カスタムラベルが無効なら常にOK
  }

  const customLabels = getCustomLabelsFromUI();
  const enabledLabels = customLabels.filter(label => label.enabled);
  
  // 有効なラベルがあるかチェック
  if (enabledLabels.length === 0) {
    return false;
  }

  // 各ラベルの内容をチェック
  for (const label of enabledLabels) {
    if (!label.text || label.text.trim() === '') {
      return false;
    }
    if (!label.count || label.count <= 0) {
      return false;
    }
  }

  return true;
}

function clearPreviousResults() {
  for (let sheet of document.querySelectorAll('section')) {
    sheet.parentNode.removeChild(sheet);
  }
  
  // 印刷枚数表示もクリア
  clearPrintCountDisplay();
}

async function getConfigFromUI() {
  const file = document.getElementById("file").files[0];
  const labelyn = document.getElementById("labelyn").checked;
  const labelskip = document.getElementById("labelskipnum").value;
  const sortByPaymentDate = document.getElementById("sortByPaymentDate").checked;
  const customLabelEnable = document.getElementById("customLabelEnable").checked;
  
  // 複数のカスタムラベルを取得（有効なもののみ）
  const allCustomLabels = getCustomLabelsFromUI();
  const customLabels = customLabelEnable ? allCustomLabels.filter(label => label.enabled) : [];
  
  await StorageManager.set(StorageManager.KEYS.LABEL_SETTING, labelyn);
  await StorageManager.set(StorageManager.KEYS.LABEL_SKIP, labelskip);
  // showAllOrders 廃止
  await StorageManager.set(StorageManager.KEYS.CUSTOM_LABEL_ENABLE, customLabelEnable);
  await StorageManager.setCustomLabels(allCustomLabels); // 全てのラベルを保存（有効/無効問わず）
  
  const labelarr = [];
  const labelskipNum = parseInt(labelskip, 10) || 0;
  if (labelskipNum > 0) {
    for (let i = 0; i < labelskipNum; i++) {
      labelarr.push("");
    }
  }
  
  return { 
    file, 
    labelyn, 
    labelskip, 
    sortByPaymentDate, 
    labelarr, 
    customLabelEnable, 
    customLabels
  };
}

async function processCSVResults(results, config) {

  // IndexedDB注文データ保存＆印刷済み注文除外（既存注文は保持・新規のみ追加）
  const db = await StorageManager.ensureDatabase();
  // 既存注文をMapで取得
  const existingOrdersArr = await db.getAllOrders();
  const existingOrders = new Map();
  for (const o of existingOrdersArr) {
    if (o && o.orderNumber) existingOrders.set(String(o.orderNumber), o);
  }
  const filteredData = [];
  const allData = [];
  for (const row of results.data) {
    const orderNumber = OrderNumberManager.getFromCSVRow(row);
    if (!orderNumber) continue;
    const key = String(orderNumber);
    let printedAt = null;
    let createdAt = new Date().toISOString();
    // 既存注文があればprintedAt等を引き継ぐ
    if (existingOrders.has(key)) {
      const old = existingOrders.get(key);
      printedAt = old.printedAt || null;
      createdAt = old.createdAt || createdAt;
    }
    await db.saveOrder({
      orderNumber: key,
      row,
      printedAt,
      createdAt
    });
  // 未印刷は印刷対象、全件は画面表示用
  if (!printedAt) filteredData.push(row);
  allData.push(row);
  }
  // 画面は常に全件表示、印刷（ラベル含む）は未印刷のみ
  const detailRows = allData;
  const labelRows = filteredData;
  const csvRowCountForLabels = labelRows.length;
  // 複数カスタムラベルの総面数を計算
  const totalCustomLabelCount = config.customLabels.reduce((sum, label) => sum + label.count, 0);
  // 複数シート対応：1シートの制限を撤廃
  // CSVデータとカスタムラベルの合計で必要なシート数を計算
  const skipCount = parseInt(config.labelskip, 10) || 0;
  const totalLabelsNeeded = skipCount + csvRowCountForLabels + totalCustomLabelCount;
  const requiredSheets = Math.ceil(totalLabelsNeeded / CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET);

  // データの並び替え
  if (config.sortByPaymentDate) {
    filteredData.sort((a, b) => {
      const timeA = a[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      const timeB = b[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      return timeA.localeCompare(timeB);
    });
    allData.sort((a, b) => {
      const timeA = a[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      const timeB = b[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      return timeA.localeCompare(timeB);
    });
  }

  // 注文明細の生成
  // ラベル対象の注文番号セットを作成
  const labelSet = new Set(labelRows.map(r => String(OrderNumberManager.getFromCSVRow(r)).trim()));
  await generateOrderDetails(detailRows, config.labelarr, labelSet);

  // 各注文明細パネルはgenerateOrderDetails内で個別に更新済み

  // ラベル生成（注文分＋カスタムラベル）- 複数シート対応
  if (config.labelyn) {
    let totalLabelArray = [...config.labelarr];

    // 明細の並び順に合わせて未印刷のみの注文番号を追加
    const visibleUnprintedSections = Array.from(document.querySelectorAll('template#注文明細 ~ section.sheet:not(.is-printed)'));
    const numbersInOrder = visibleUnprintedSections.map(sec => sec.dataset.orderNumber).filter(Boolean);
    totalLabelArray.push(...numbersInOrder);

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

  // 印刷枚数の表示（複数シート対応）
  // showCSVWithCustomLabelPrintSummary(csvRowCount, totalCustomLabelCount, skipCount, requiredSheets);

  // ヘッダーの印刷枚数表示を更新
  // ラベル印刷がオフの場合はラベル枚数・カスタム面数ともに0を表示する
  const labelSheetsForDisplay = config.labelyn ? requiredSheets : 0;
  const customFacesForDisplay = (config.labelyn && config.customLabelEnable) ? totalCustomLabelCount : 0;
  // 普通紙（注文明細）は未印刷のみ
  updatePrintCountDisplay(filteredData.length, labelSheetsForDisplay, customFacesForDisplay);

  // CSV処理完了後のカスタムラベルサマリー更新（複数シート対応）
  await updateCustomLabelsSummary();

  // ボタンの状態を更新
  updateButtonStates();
}

async function processCustomLabelsOnly(config, isPreviewMode = false) {
  // 複数カスタムラベルの総面数を計算
  const totalCustomLabelCount = config.customLabels.reduce((sum, label) => sum + label.count, 0);
  const labelskipNum = parseInt(config.labelskip, 10) || 0;
  
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
  updateButtonStates();
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
    
    // 注文情報の設定
    orderNumber = setOrderInfo(cOrder, row, labelarr, labelSet);

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
      const normalized = OrderNumberManager.normalize(orderNumber);
      if (rootSection && normalized) {
  if (!window.unifiedDB) await StorageManager.ensureDatabase();
  const o = await window.unifiedDB.getOrder(normalized);
        if (o?.printedAt) rootSection.classList.add('is-printed');
        else rootSection.classList.remove('is-printed');
      }
    } catch {}
  }
}

// 各注文明細の非印刷パネルをセットアップ（印刷日時表示とクリア機能）
async function setupOrderPrintedAtPanel(cOrder, orderNumber) {
  const panel = cOrder.querySelector('.order-print-info');
  if (!panel) return;
  const dateEl = panel.querySelector('.printed-at');
  const markPrintedBtn = panel.querySelector('.mark-printed');
  const clearBtn = panel.querySelector('.clear-printed-at');
  const normalized = OrderNumberManager.normalize(orderNumber);
  if (!normalized) {
    if (dateEl) dateEl.textContent = '未印刷';
    if (markPrintedBtn) { markPrintedBtn.style.display = ''; markPrintedBtn.disabled = false; }
    if (clearBtn) { clearBtn.style.display = 'none'; clearBtn.disabled = true; }
    return;
  }

  if (!window.unifiedDB) await StorageManager.ensureDatabase();
  const order = await window.unifiedDB.getOrder(normalized);
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
  await window.unifiedDB.setPrintedAt(normalized, now);

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
  await window.unifiedDB.setPrintedAt(normalized, null);

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

// スクロール位置を保持・復元する（再描画でDOMが差し替わってもUXを維持）
function captureAndRestoreScrollPosition() {
  const doc = document.scrollingElement || document.documentElement;
  const x = window.scrollX || doc.scrollLeft || 0;
  const y = window.scrollY || doc.scrollTop || 0;
  const prevScrollHeight = doc.scrollHeight || document.body.scrollHeight || 0;
  const viewportH = window.innerHeight || doc.clientHeight || 0;
  const prevScrollable = (prevScrollHeight - viewportH) > 2; // 実質スクロール可能か
  debugLog('📌 captureScroll', { x, y, prevScrollHeight, viewportH, prevScrollable });

  return function restore() {
    const docNow = document.scrollingElement || document.documentElement;
    const newScrollHeight = docNow.scrollHeight || document.body.scrollHeight || 0;
    const newViewportH = window.innerHeight || docNow.clientHeight || 0;
    const newScrollable = (newScrollHeight - newViewportH) > 2;

    // どちらも実質スクロール不可なら復元不要
    if (!prevScrollable && !newScrollable) {
      debugLog('↩️ restoreScroll skip: not scrollable');
      return;
    }

    const maxPrev = Math.max(prevScrollHeight - viewportH, 1);
    const ratio = Math.min(Math.max(y / maxPrev, 0), 1);
    const maxNew = Math.max(newScrollHeight - newViewportH, 0);
    const targetY = Math.min(Math.max(Math.round(ratio * maxNew), 0), maxNew);

    const doScroll = () => {
      const currentY = window.scrollY || docNow.scrollTop || 0;
      const currentX = window.scrollX || docNow.scrollLeft || 0;
      // ほぼ同位置ならスキップ
      if (Math.abs(currentY - targetY) < 2 && Math.abs(currentX - x) < 2) {
        debugLog('↩️ restoreScroll skip: no-op', { targetY, currentY });
        return;
      }
      debugLog('↩️ restoreScroll', { ratio, targetY, maxNew, newScrollHeight });
      try {
        window.scrollTo({ left: x, top: targetY, behavior: 'auto' });
      } catch {
        window.scrollTo(x, targetY);
      }
    };

    // レイアウト確定後に1回だけ復元（二重RAF）
    requestAnimationFrame(() => requestAnimationFrame(doScroll));
  };
}

// 注文番号のセクションにスクロール（固定ヘッダー分オフセット考慮）
function scrollToOrderSection(normalizedOrder) {
  if (!normalizedOrder) return;
  debugLog('🎯 scrollToOrderSection request', { normalizedOrder });
  const target = document.querySelector(`section.sheet[data-order-number="${CSS.escape(normalizedOrder)}"]`);
  if (!target) {
    debugLog('🎯 target not found', { normalizedOrder });
    return false;
  }
  const header = document.querySelector('.fixed-header');
  const headerHeight = header && getComputedStyle(header).display !== 'none' ? header.offsetHeight : 0;
  const rect = target.getBoundingClientRect();
  const y = window.scrollY + rect.top - Math.max(headerHeight + 8, 0);
  debugLog('🎯 scrolling', { headerHeight, rectTop: rect.top, to: y });
  try {
    window.scrollTo({ top: Math.max(y, 0), behavior: 'auto' });
  } catch {
    window.scrollTo(0, Math.max(y, 0));
  }
  return true;
}

// 現在の「読み込んだファイル全て表示」のON/OFFを返す
// showAllOrders 廃止

// 既存のDOMからラベル部分だけ再生成（CSVデータはDBから復元）
async function regenerateLabelsFromDB() {
  try {
    // 既存のラベルセクションを削除（テンプレート位置に依存せず、専用クラスで判別）
    document.querySelectorAll('section.sheet.label-sheet').forEach(sec => sec.remove());
  } catch {}

  // 現在の画面上の未印刷注文明細の並び順をそのままラベルに反映
  const orderSections = Array.from(document.querySelectorAll('template#注文明細 ~ section.sheet:not(.is-printed)'));
  const orderNumbers = orderSections
    .map(sec => sec.dataset.orderNumber)
    .filter(Boolean);

  // 設定取得
  const settings = await StorageManager.getSettingsAsync();
  // ラベル印刷が無効ならここで終了（既存は削除済み）
  if (!settings.labelyn) {
    return;
  }

  // ラベル配列の再構築（スキップ数 + 未印刷の注文明細順。カスタムラベルは設定から）
  const skip = parseInt(settings.labelskip || '0', 10) || 0;
  const labelarr = new Array(skip).fill("");
  for (const num of orderNumbers) {
    labelarr.push(num);
  }
  // カスタムラベル（ON のとき）
  if (settings.labelyn && settings.customLabelEnable && settings.customLabels?.length) {
    for (const cl of settings.customLabels.filter(l => l.enabled)) {
      for (let i = 0; i < cl.count; i++) {
        labelarr.push({ type: 'custom', content: cl.html || cl.text, fontSize: cl.fontSize || '10pt' });
      }
    }
  }

  if (labelarr.length > 0) {
    await generateLabels(labelarr, { skipOnFirstSheet: skip });
  }
}

// 画面上の枚数表示（固定ヘッダー）を再計算して更新
function recalcAndUpdateCounts() {
  const orderSheetCount = document.querySelectorAll('template#注文明細 ~ section.sheet:not(.is-printed)').length;
  const labelSheetCount = document.querySelectorAll('section.sheet.label-sheet').length;
  // カスタム面数を設定から再計算
  StorageManager.getSettingsAsync().then(settings => {
    // ラベル印刷がOFFの場合はラベル/カスタムとも0表示にする
    const labelSheetsForDisplay = settings.labelyn ? labelSheetCount : 0;
    const customCountForDisplay = (settings.labelyn && settings.customLabelEnable && Array.isArray(settings.customLabels))
      ? settings.customLabels.filter(l => l.enabled).reduce((s, l) => s + (parseInt(l.count, 10) || 0), 0)
      : 0;
    updatePrintCountDisplay(orderSheetCount, labelSheetsForDisplay, customCountForDisplay);
  });
}

function setOrderInfo(cOrder, row, labelarr, labelSet = null) {
  let orderNumber = '';
  
  for (let c of Object.keys(row).filter(key => key != CONSTANTS.CSV.PRODUCT_COLUMN)) {
    const divc = cOrder.querySelector("." + c);
    if (divc) {
      if (c == CONSTANTS.CSV.ORDER_NUMBER_COLUMN) {
        orderNumber = OrderNumberManager.getFromCSVRow(row);
        const displayFormat = OrderNumberManager.createDisplayFormat(orderNumber);
        divc.textContent = displayFormat;
  // 以前はここで labelarr に未印刷の注文番号を追加していたが、
  // 現在は DOM 上の未印刷セクションの並びから再収集して重複を避けるため追加しない
      } else if (row[c]) {
        divc.textContent = row[c];
      }
    }
  }
  // セクションに注文アンカーを付与
  try {
    const sectionEl = cOrder.querySelector('section.sheet');
    if (sectionEl && orderNumber) {
      const normalized = OrderNumberManager.normalize(String(orderNumber));
      sectionEl.dataset.orderNumber = normalized;
      // sectionEl.id = `order-${normalized}`; // 必要ならidも付与
    }
  } catch {}
  
  return orderNumber;
}

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

  if (individualDropZoneContainer && OrderNumberManager.isValid(orderNumber)) {
    // 注文番号を正規化
    const normalizedOrderNumber = OrderNumberManager.normalize(orderNumber);
    
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
  if (OrderNumberManager.isValid(orderNumber)) {
    // 注文番号を正規化
    const normalizedOrderNumber = OrderNumberManager.normalize(orderNumber);
    
    // 個別画像があるかチェック
    const individualImage = await StorageManager.getOrderImage(normalizedOrderNumber);
    if (individualImage) {
      imageToShow = individualImage;
    } else {
      // 個別画像がない場合はグローバル画像を使用
      const globalImage = window.orderImageDropZone?.getImage();
      if (globalImage) {
        imageToShow = globalImage;
      }
    }
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
  
  for (let label of labelarr) {
    if (i > 0 && i % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET == 0) {
      tableL44.appendChild(tr);
      tr = document.createElement("tr");
      document.body.insertBefore(cL44, tL44);
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
function setupDropzoneEvents(dropzone) {
  dropzone.addEventListener('dragover', function(event) {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
    showDropping(this);
  });

  dropzone.addEventListener('dragleave', function(event) {
    hideDropping(this);
  });

  dropzone.addEventListener('drop', function (event) {
    event.preventDefault();
    hideDropping(this);
    const elImage = document.createElement('img');
    
    if(event.dataTransfer.types.includes("text/uri-list")){
      const url = event.dataTransfer.getData('text/uri-list');
      elImage.src = url;
      this.parentNode.appendChild(elImage);
      readQR(elImage);
    } else {
      const file = event.dataTransfer.files[0];
      if(file && file.type.indexOf('image/') === 0){
        this.parentNode.appendChild(elImage);
        attachImage(file, elImage);
      }
    }
    
    this.classList.remove('dropzone');
    this.style.zIndex = -1;
    elImage.style.zIndex = 9;
    addEventQrReset(elImage);
  });

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
    const tpl = document.getElementById('qrDropPlaceholder');
    if (tpl && tpl.content.firstElementChild) {
      this.innerHTML = '';
      this.appendChild(tpl.content.firstElementChild.cloneNode(true));
    } else {
      this.innerHTML = '<p>Paste QR image here!</p>';
    }
  });
}

function createDropzone(div){
  const divDrop = createDiv('dropzone');
  const tpl = document.getElementById('qrDropPlaceholder');
  if (tpl && tpl.content.firstElementChild) {
    divDrop.appendChild(tpl.content.firstElementChild.cloneNode(true));
  } else {
    divDrop.textContent = 'Paste QR image here!';
  }
  divDrop.setAttribute("contenteditable", "true");
  divDrop.setAttribute("effectAllowed", "move");
  
  // 共通のイベントリスナーを設定
  setupDropzoneEvents(divDrop);
  
  div.appendChild(divDrop);
}

async function createLabel(labelData=""){
  const divQr = createDiv('qr');
  const divOrdernum = createDiv('ordernum');
  const divYamato = createDiv('yamato');

  // ラベルデータが文字列の場合（既存の注文番号）
  if (typeof labelData === 'string') {
    if (labelData) {
      addP(divOrdernum, labelData);
      const qr = await StorageManager.getQRData(labelData);
      if(qr && qr['qrimage']){
        // 保存されたQR画像がある場合は画像を表示
        const elImage = document.createElement('img');
        elImage.src = qr['qrimage'];
        divQr.appendChild(elImage);
        addP(divYamato, qr['receiptnum']);
        addP(divYamato, qr['receiptpassword']);
        addEventQrReset(elImage);
      } else {
        // QR画像がない場合のみドロップゾーンを作成
        createDropzone(divQr);
      }
    }
  } 
  // ラベルデータがオブジェクトの場合（カスタムラベル）
  else if (typeof labelData === 'object' && labelData.type === 'custom') {
    divOrdernum.classList.add('custom-label');
    const contentDiv = document.createElement('div');
    contentDiv.classList.add('custom-label-text');
    contentDiv.innerHTML = labelData.content;
    
    // 文字サイズを適用
    if (labelData.fontSize) {
      contentDiv.style.fontSize = labelData.fontSize;
    }
    
    divOrdernum.appendChild(contentDiv);
    
    // カスタムラベルの場合はQRコードエリアとヤマトエリアを非表示
    divQr.style.display = 'none';
    divYamato.style.display = 'none';
    
    // カスタムラベルが全体を覆うようにスタイル調整
    divOrdernum.style.position = 'absolute';
    divOrdernum.style.width = '100%';
    divOrdernum.style.height = '100%';
    divOrdernum.style.top = '0';
    divOrdernum.style.left = '0';
    divOrdernum.style.display = 'flex';
    divOrdernum.style.alignItems = 'center';
    divOrdernum.style.justifyContent = 'center';
    divOrdernum.style.padding = '2px';
    divOrdernum.style.boxSizing = 'border-box';
  }

  const tdLabel = document.createElement('td');
  tdLabel.classList.add('qrlabel');
  tdLabel.appendChild(divQr);
  tdLabel.appendChild(divOrdernum);
  tdLabel.appendChild(divYamato);

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
        const orderNumber = ordernumDiv ? OrderNumberManager.normalize(ordernumDiv.textContent) : null;
        
        // 保存されたQRデータを削除
        if (orderNumber) {
          try {
            await StorageManager.setQRData(orderNumber, null);
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

function showDropping(elDrop) {
        elDrop.classList.add('dropover');
}

function hideDropping(elDrop) {
        elDrop.classList.remove('dropover');
}

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
        const b64data = canv.toDataURL("image/png");
        
        const imageData = context.getImageData(0, 0, canv.width, canv.height);
        const barcode = jsQR(imageData.data, imageData.width, imageData.height);
        
        if(barcode){
          const b = String(barcode.data).replace(/^\s+|\s+$/g,'').replace(/ +/g,' ').split(" ");
          
          if(b.length === CONSTANTS.QR.EXPECTED_PARTS){
            const rawOrderNum = elImage.closest("td").querySelector(".ordernum p").innerHTML;
            const ordernum = OrderNumberManager.normalize(rawOrderNum);
            
            // 重複チェック
            const duplicates = await StorageManager.checkQRDuplicate(barcode.data, ordernum);
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
                  dropzone.innerHTML = '<p>Paste QR image here!</p>';
                  dropzone.style.display = 'block';
                } else {
                  // ドロップゾーンが見つからない場合は新しく作成
                  const newDropzone = document.createElement('div');
                  newDropzone.className = 'dropzone';
                  newDropzone.contentEditable = 'true';
                  newDropzone.setAttribute('effectallowed', 'move');
                  newDropzone.style.zIndex = '99';
                  newDropzone.innerHTML = '<p>Paste QR image here!</p>';
                  parentQr.appendChild(newDropzone);
                  
                  // ドロップゾーンのイベントリスナーを再設定
                  setupDropzoneEvents(newDropzone);
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
              "receiptnum": b[1], 
              "receiptpassword": b[2], 
              "qrimage": b64data,
              "qrhash": StorageManager.generateQRHash(barcode.data)
            };
            
            await StorageManager.setQRData(ordernum, qrData);
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

function attachImage(file, elImage) {
  if (!file || !elImage) {
    console.error('ファイルまたは画像要素が無効です');
    return;
  }

  const reader = new FileReader();
  reader.onload = function(event) {
    try {
      const src = event.target.result;
      elImage.src = src;
      elImage.setAttribute('title', file.name);
      elImage.onload = function() {
        readQR(elImage);
      };
    } catch (error) {
      console.error('画像の読み込みエラー:', error);
    }
  };
  
  reader.onerror = function() {
    console.error('ファイル読み込みエラー');
  };
  
  reader.readAsDataURL(file);
}

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

// 共通のドラッグ&ドロップ機能を提供するベース関数
async function createBaseImageDropZone(options = {}) {
  const {
    storageKey = 'orderImage',
    isIndividual = false,
    orderNumber = null,
    containerClass = 'order-image-drop',
    defaultMessage = '画像をドロップ or クリックで選択'
  } = options;

  debugLog(`ベース画像ドロップゾーン作成: ${storageKey}, 個別: ${isIndividual}, 注文番号: ${orderNumber}`);

  const dropZone = document.createElement('div');
  dropZone.classList.add(containerClass);
  
  if (isIndividual) {
    dropZone.style.cssText = 'min-height: 80px; border: 1px dashed #999; padding: 5px; background: #f9f9f9; cursor: pointer;';
  }

  let droppedImage = null;

  // StorageManagerから保存された画像を読み込む
  const savedImage = await StorageManager.getOrderImage(orderNumber);
  if (savedImage) {
    debugLog(`保存された画像を復元: ${storageKey}`);
    await updatePreview(savedImage);
  } else {
    if (isIndividual) {
      const imgTpl = document.getElementById('orderImageDropDefault');
      if (imgTpl && imgTpl.content.firstElementChild) {
        dropZone.innerHTML = '';
        const node = imgTpl.content.firstElementChild.cloneNode(true);
        node.textContent = defaultMessage; // 個別はメッセージ差替
        dropZone.appendChild(node);
      } else {
        dropZone.innerHTML = `<p class="order-image-default-msg">${defaultMessage}</p>`;
      }
    } else {
      const defaultContentElement = document.getElementById('dropZoneDefaultContent');
      if (defaultContentElement) {
        dropZone.innerHTML = '';
        // 既存の defaultContent を template へ移行していればそちら優先
        const tpl = document.getElementById('orderImageDropDefault');
        if (tpl && tpl.content.firstElementChild) {
          dropZone.appendChild(tpl.content.firstElementChild.cloneNode(true));
        } else {
          dropZone.innerHTML = defaultContentElement.innerHTML;
        }
      } else {
        dropZone.innerHTML = `<p class="order-image-default-msg">${defaultMessage}</p>`;
      }
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
      const orderNumber = OrderNumberManager.getFromOrderSection(orderSection);

      // 個別画像があるかチェック（個別画像を最優先）
      let imageToShow = null;
      if (orderNumber) {
        const individualImage = await StorageManager.getOrderImage(orderNumber);
        const globalImage = await StorageManager.getOrderImage(); // グローバル画像を取得
        
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
        const globalImage = await StorageManager.getOrderImage(); // グローバル画像を取得
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

  async function updatePreview(imageUrl) {
    droppedImage = imageUrl;
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
    await StorageManager.setOrderImage(imageUrl, orderNumber);

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
      await StorageManager.removeOrderImage(orderNumber);
      droppedImage = null;
      
      if (isIndividual) {
        const imgTpl = document.getElementById('orderImageDropDefault');
        dropZone.innerHTML = '';
        if (imgTpl && imgTpl.content.firstElementChild) {
          const node = imgTpl.content.firstElementChild.cloneNode(true);
          node.textContent = defaultMessage;
          dropZone.appendChild(node);
        } else {
          dropZone.innerHTML = `<p class="order-image-default-msg">${defaultMessage}</p>`;
        }
        await updateOrderImageDisplay(null);
      } else {
        const tplGlobal = document.getElementById('orderImageDropDefault');
        dropZone.innerHTML = '';
        if (tplGlobal && tplGlobal.content.firstElementChild) {
          dropZone.appendChild(tplGlobal.content.firstElementChild.cloneNode(true));
        } else {
          dropZone.innerHTML = `<p class="order-image-default-msg">${defaultMessage}</p>`;
        }
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
      const globalImage = window.orderImageDropZone?.getImage();
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
    setImage: (imageData) => {
      droppedImage = imageData;
      updatePreview(imageData);
    }
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
      reader.onload = async (e) => {
        await updatePreview(e.target.result);
      };
      reader.readAsDataURL(file);
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
          reader.onload = async (e) => {
            debugLog(`画像読み込み完了 - サイズ: ${e.target.result.length} bytes`);
            await updatePreview(e.target.result);
          };
          reader.readAsDataURL(file);
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
  updateButtonStates();
  await updateCustomLabelsSummary();
  
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
  updateButtonStates();
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
      const orderNumber = OrderNumberManager.getFromOrderSection(orderSection);
      
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
          const individualImage = await StorageManager.getOrderImage(orderNumber);
          if (individualImage) {
            imageToShow = individualImage;
            debugLog(`個別画像を表示: ${orderNumber}`);
          } else {
            const globalImage = await StorageManager.getOrderImage();
            if (globalImage) {
              imageToShow = globalImage;
              debugLog(`グローバル画像を表示: ${orderNumber}`);
            }
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

// 複数カスタムラベルの初期化
function initializeCustomLabels(customLabels) {
  const container = document.getElementById('customLabelsContainer');
  container.innerHTML = '';
  
  // 説明文を一番上に追加
  const instTpl = document.getElementById('customLabelsInstructionTemplate');
  if (instTpl && instTpl.content.firstElementChild) {
    container.appendChild(instTpl.content.firstElementChild.cloneNode(true));
  }
  
  if (customLabels && customLabels.length > 0) {
    customLabels.forEach((label, index) => {
      addCustomLabelItem(label.html || label.text, label.count, index, label.enabled !== false);
      
      // 文字サイズを復元
      if (label.fontSize) {
        const item = container.children[index];
        const editor = item.querySelector('.rich-text-editor');
        
        if (editor) {
          // フォントサイズが既にpt単位なら そのまま使用、数値のみなら pt を追加
          const fontSize = label.fontSize.toString().includes('pt') ? label.fontSize : label.fontSize + 'pt';
          editor.style.fontSize = fontSize;
        }
      }
    });
  } else {
    // デフォルトで1つ追加
    addCustomLabelItem('', 1, 0, true);
  }
  
  // 非同期でサマリーを更新
  updateCustomLabelsSummary().catch(console.error);
}

// カスタムラベル項目を追加
function addCustomLabelItem(text = '', count = 1, index = null, enabled = true) {
  debugLog('addCustomLabelItem関数が呼び出されました'); // デバッグ用
  debugLog('引数:', { text, count, index }); // デバッグ用
  
  const container = document.getElementById('customLabelsContainer');
  debugLog('container要素:', container); // デバッグ用
  if (!container) {
    console.error('customLabelsContainer要素が見つかりません');
    return;
  }
  
  // テンプレートを取得
  const template = document.getElementById('customLabelItem');
  if (!template) {
    console.error('customLabelItemテンプレートが見つかりません');
    return;
  }
  
  // テンプレートをクローン
  const item = template.content.cloneNode(true);
  const itemDiv = item.querySelector('.custom-label-item');
  
  const itemIndex = index !== null ? index : container.children.length;
  debugLog('itemIndex:', itemIndex); // デバッグ用
  
  // データ属性を設定
  itemDiv.dataset.index = itemIndex;
  
  // チェックボックスの設定
  const checkbox = item.querySelector('.custom-label-enabled');
  checkbox.id = `customLabel_${itemIndex}_enabled`;
  checkbox.dataset.index = itemIndex;
  checkbox.checked = enabled;
  
  // ラベルのfor属性を設定
  const label = item.querySelector('.custom-label-item-title');
  label.setAttribute('for', `customLabel_${itemIndex}_enabled`);
  
  // エディタの設定
  const editor = item.querySelector('.rich-text-editor');
  editor.dataset.index = itemIndex;
  
  // 枚数入力の設定
  const countInput = item.querySelector('input[type="number"]');
  countInput.value = count;
  countInput.dataset.index = itemIndex;
  
  // 削除ボタンの設定
  const removeBtn = item.querySelector('.btn-remove');
  removeBtn.onclick = () => removeCustomLabelItem(itemIndex);
  
  container.appendChild(item);
  debugLog('item要素がコンテナに追加されました'); // デバッグ用
  
  // リッチテキストエディタのイベントリスナーを設定
  const editorElement = container.querySelector(`[data-index="${itemIndex}"].rich-text-editor`);
  debugLog('editorElement:', editorElement); // デバッグ用
  if (editorElement) {
    // テキスト内容を設定（HTMLとして）
    if (text && text.trim() !== '') {
      editorElement.innerHTML = text;
    }
    
    setupRichTextFormatting(editorElement);
    setupTextOnlyEditor(editorElement);
    
    // 編集開始時のイベント
    editorElement.addEventListener('focus', function() {
      debugLog('カスタムラベル編集開始');
      isEditingCustomLabel = true;
    });
    
    // 編集終了時のイベント
    editorElement.addEventListener('blur', async function() {
      debugLog('カスタムラベル編集終了');
      isEditingCustomLabel = false;
      
      // 編集終了後に遅延更新を実行
      scheduleDelayedPreviewUpdate(300);
    });
    
    editorElement.addEventListener('input', async function() {
      // 文字列が入力されたら強調表示をクリア
      const item = editorElement.closest('.custom-label-item');
      if (item && editorElement.textContent.trim() !== '') {
        item.classList.remove('error');
      }
      
      saveCustomLabels();
      updateButtonStates();
      await updateCustomLabelsSummary();
      
      // 編集中は即座の自動再処理をスキップし、遅延更新をスケジュール
      if (!isEditingCustomLabel) {
        await autoProcessCSV();
      } else {
        scheduleDelayedPreviewUpdate(1000); // 編集中は1秒の遅延
      }
    });
  } else {
    console.error('editorElement要素が見つかりません');
  }
  
  // チェックボックスの状態を設定とイベントリスナー追加
  const enabledCheckbox = container.querySelector(`[data-index="${itemIndex}"].custom-label-enabled`);
  if (enabledCheckbox) {
    enabledCheckbox.checked = enabled;
    enabledCheckbox.addEventListener('change', async function() {
      saveCustomLabels();
      await updateCustomLabelsSummary();
      
      // チェックボックス変更時に自動再処理
      await autoProcessCSV();
    });
  }
  
  // 枚数入力のイベントリスナーを設定
  const countInputElement = container.querySelector(`[data-index="${itemIndex}"][type="number"]`);
  debugLog('countInputElement要素:', countInputElement); // デバッグ用
  if (countInputElement) {
    countInputElement.addEventListener('input', async function() {
      saveCustomLabels();
      updateButtonStates();
      await updateCustomLabelsSummary();
      
      // 枚数変更時に自動再処理
      await autoProcessCSV();
    });
  } else {
    console.error('countInput要素が見つかりません');
  }
  
  updateCustomLabelsSummary().catch(console.error);
}

// カスタムラベル項目を削除
function removeCustomLabelItem(index) {
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  
  if (items.length <= 1) {
    alert('最低1つのカスタムラベルは必要です。');
    return;
  }
  
  // 指定されたインデックスの項目を削除
  const itemToRemove = container.querySelector(`[data-index="${index}"]`);
  if (itemToRemove) {
    itemToRemove.remove();
  }
  
  // インデックスを再設定
  reindexCustomLabelItems();
  saveCustomLabels();
  updateCustomLabelsSummary().catch(console.error);
  updateButtonStates();
  
  // 項目削除時に自動再処理
  autoProcessCSV().catch(console.error);
}

// UIからカスタムラベルデータを取得
function getCustomLabelsFromUI() {
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  const labels = [];
  
  items.forEach(item => {
    const editor = item.querySelector('.rich-text-editor');
    const countInput = item.querySelector('input[type="number"]');
    const enabledCheckbox = item.querySelector('.custom-label-enabled');
    const text = editor.innerHTML.trim();
    const count = parseInt(countInput.value, 10) || 1;
    const enabled = enabledCheckbox ? enabledCheckbox.checked : true;
    
    // フォントサイズをエディタのスタイルから取得
    const computedStyle = window.getComputedStyle(editor);
    const fontSize = computedStyle.fontSize || '12pt';
    
    // テキストがある場合、または設定を保持するため常に保存
    labels.push({ 
      text, 
      count, 
      fontSize,
      html: text, // HTMLフォーマットも保存
      enabled // 印刷有効フラグを追加
    });
  });
  
  return labels;
}

// カスタムラベルデータを保存
function saveCustomLabels() {
  const labels = getCustomLabelsFromUI();
  StorageManager.setCustomLabels(labels);
}

// カスタムラベルの総計を更新（非同期対応）
async function updateCustomLabelsSummary() {
  const labels = getCustomLabelsFromUI();
  const enabledLabels = labels.filter(label => label.enabled);
  const totalCount = enabledLabels.reduce((sum, label) => sum + label.count, 0);
  const skipCount = parseInt(document.getElementById("labelskipnum").value, 10) || 0;
  const fileInput = document.getElementById("file");
  
  const summary = document.getElementById('customLabelsSummary');
  
  if (totalCount === 0) {
    // CSVファイルが選択されている場合はCSV行数も表示
    if (fileInput.files.length > 0) {
      try {
        const csvInfo = await CSVAnalyzer.getFileInfo(fileInput.files[0]);
        const remainingLabels = CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET - skipCount - csvInfo.rowCount;
        summary.innerHTML = `カスタムラベルなし。<br>44面シート中 スキップ${skipCount} + CSV${csvInfo.rowCount} = ${skipCount + csvInfo.rowCount}面使用済み。<br>残り${Math.max(0, remainingLabels)}面設定可能。`;
      } catch (error) {
        summary.innerHTML = `カスタムラベルなし。<br>CSVファイル選択済み（行数解析中...）。`;
      }
    } else {
      summary.innerHTML = `カスタムラベルなし。<br>44面シート中 スキップ${skipCount}面使用済み。`;
    }
    summary.style.color = '#666';
    summary.style.fontWeight = 'normal';
    return;
  }
  
  // CSVファイルが選択されている場合も複数シート対応
  if (fileInput.files.length > 0) {
    try {
      // CSV行数を非同期で取得
      const csvInfo = await CSVAnalyzer.getFileInfo(fileInput.files[0]);
      const csvRowCount = csvInfo.rowCount;
      
      // CSV+カスタムラベルの複数シート分散計算
      const totalLabels = csvRowCount + totalCount;
      const sheetsInfo = CustomLabelCalculator.calculateMultiSheetDistribution(totalLabels, skipCount);
      const totalSheets = sheetsInfo.length;
      const lastSheet = sheetsInfo[sheetsInfo.length - 1];
      
      if (totalSheets === 1) {
        summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>1シート使用: スキップ${skipCount} + CSV${csvRowCount} + カスタム${totalCount} = ${skipCount + csvRowCount + totalCount}面<br>最終シート残り${lastSheet.remainingCount}面。`;
      } else {
        summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>${totalSheets}シート使用: CSV${csvRowCount} + カスタム${totalCount} = ${csvRowCount + totalCount}面<br>最終シート残り${lastSheet.remainingCount}面。`;
      }
      
      summary.style.color = '#666';
      summary.style.fontWeight = 'normal';
    } catch (error) {
      console.error('CSV解析エラー:', error);
      summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>CSVファイル選択済み（行数解析エラー）<br>CSV処理実行後に最終配置が決定されます。`;
      summary.style.color = '#ffc107';
      summary.style.fontWeight = 'normal';
    }
  } else {
    // カスタムラベルのみの場合は複数シート対応
    const sheetsInfo = CustomLabelCalculator.calculateMultiSheetDistribution(totalCount, skipCount);
    const totalSheets = sheetsInfo.length;
    const lastSheet = sheetsInfo[sheetsInfo.length - 1];
    
    if (totalSheets === 1) {
      // 1シートのみの場合
      summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>1シート使用: スキップ${skipCount} + カスタム${totalCount} = ${skipCount + totalCount}面<br>最終シート残り${lastSheet.remainingCount}面。`;
    } else {
      // 複数シートの場合
      summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>${totalSheets}シート使用: カスタム${totalCount}面<br>最終シート残り${lastSheet.remainingCount}面。`;
    }
    
    summary.style.color = '#666';
    summary.style.fontWeight = 'normal';
  }
}

// 総枚数用の調整関数
async function adjustCustomLabelsForTotal(customLabels, maxTotalLabels) {
  let remaining = maxTotalLabels;
  
  for (let i = 0; i < customLabels.length; i++) {
    if (remaining <= 0) {
      customLabels[i].count = 0;
    } else if (customLabels[i].count > remaining) {
      customLabels[i].count = remaining;
    }
    remaining -= customLabels[i].count;
  }
  
  // UIを更新
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  items.forEach((item, index) => {
    const countInput = item.querySelector('input[type="number"]');
    if (customLabels[index]) {
      countInput.value = customLabels[index].count;
    }
  });
  
  saveCustomLabels();
  await updateCustomLabelsSummary();
}

function setupCustomLabelEvents() {
  // 初期カスタムラベルエディターのイベントリスナー設定
  const initialEditor = document.getElementById('initialCustomLabelEditor');
  if (initialEditor) {
    // 編集開始時のイベント
    initialEditor.addEventListener('focus', function() {
      debugLog('初期カスタムラベル編集開始');
      isEditingCustomLabel = true;
    });
    
    // 編集終了時のイベント
    initialEditor.addEventListener('blur', async function() {
      debugLog('初期カスタムラベル編集終了');
      isEditingCustomLabel = false;
      
      // 編集終了後に遅延更新を実行
      scheduleDelayedPreviewUpdate(300);
    });
  }

  // カスタムラベル有効化チェックボックス
  const customLabelEnable = document.getElementById('customLabelEnable');
  if (customLabelEnable) {
    customLabelEnable.addEventListener('change', async function() {
      toggleCustomLabelRow(this.checked);
      await StorageManager.set(StorageManager.KEYS.CUSTOM_LABEL_ENABLE, this.checked);
      updateButtonStates();
      
      // 設定変更時に自動再処理
      await autoProcessCSV();
    });
  }

  // カスタムラベル追加ボタン
  const addButton = document.getElementById('addCustomLabelBtn');
  if (addButton) {
    addButton.addEventListener('click', async function() {
      debugLog('ラベル追加ボタンがクリックされました'); // デバッグ用
      addCustomLabelItem('', 1, null, true);
      saveCustomLabels();
      updateButtonStates();
      
      // ラベル追加時に自動再処理
      await autoProcessCSV();
    });
  } else {
    console.error('addCustomLabelBtn要素が見つかりません');
  }

  // カスタムラベル全削除ボタン
  const clearButton = document.getElementById('clearCustomLabelsBtn');
  if (clearButton) {
    clearButton.addEventListener('click', async function() {
      debugLog('ラベル全削除ボタンがクリックされました'); // デバッグ用
      if (confirm('本当に全てのカスタムラベルを削除しますか？')) {
        await clearAllCustomLabels();
        
        // ラベル削除時に自動再処理
        await autoProcessCSV();
      }
    });
  } else {
    console.error('clearCustomLabelsBtn要素が見つかりません');
  }
}

// 全てのカスタムラベルを削除
async function clearAllCustomLabels() {
  const container = document.getElementById('customLabelsContainer');
  
  // 説明文以外の全ての項目を削除
  const items = container.querySelectorAll('.custom-label-item');
  items.forEach(item => item.remove());
  
  // デフォルトで1つ追加
  addCustomLabelItem('', 1, 0, true);
  
  // 保存とUI更新
  saveCustomLabels();
  await updateCustomLabelsSummary();
  updateButtonStates();
}

async function updateButtonStates() {
  const fileInput = document.getElementById("file");
  const printButton = document.getElementById("printButton");
  
  // 固定ヘッダーのボタン要素も取得
  const printButtonCompact = document.getElementById("printButtonCompact");

  // 印刷ボタンの状態（何らかのコンテンツが生成されている場合に有効）
  const hasSheets = document.querySelectorAll('.sheet').length > 0;
  const hasLabels = document.querySelectorAll('.label44').length > 0;
  const hasContent = hasSheets || hasLabels;
  if (printButton) printButton.disabled = !hasContent;
  if (printButtonCompact) printButtonCompact.disabled = !hasContent;

  // カスタムラベル枚数の上限を更新
  await updateCustomLabelsSummary();
}

// カスタムラベルに内容があるかチェック
function hasCustomLabelsWithContent() {
  const labels = getCustomLabelsFromUI();
  return labels.length > 0 && labels.some(label => label.text.trim() !== '');
}

// カスタムラベルが有効だが内容が未設定の項目があるかチェック
function hasEmptyEnabledCustomLabels() {
  const customLabelEnable = document.getElementById("customLabelEnable");
  if (!customLabelEnable.checked) {
    return false; // カスタムラベル機能が無効の場合はチェック不要
  }
  
  const labels = getCustomLabelsFromUI();
  const enabledLabels = labels.filter(label => label.enabled);
  
  // 有効なラベルで文字列が空のものがあるかチェック
  return enabledLabels.some(label => label.text.trim() === '');
}

// 未設定のカスタムラベル項目を削除
function removeEmptyCustomLabels() {
  const labels = getCustomLabelsFromUI();
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  
  let removedCount = 0;
  
  // 後ろから削除して、インデックスのずれを防ぐ
  for (let i = labels.length - 1; i >= 0; i--) {
    const label = labels[i];
    if (label.enabled && label.text.trim() === '') {
      // 未設定の有効ラベルを削除
      const item = items[i];
      if (item) {
        item.remove();
        removedCount++;
      }
    }
  }
  
  // インデックスを再設定
  reindexCustomLabelItems();
  
  // 保存とUI更新
  saveCustomLabels();
  updateCustomLabelsSummary().catch(console.error);
  
  // 削除後にラベルが全くなくなった場合は、デフォルトで1つ追加
  const remainingItems = container.querySelectorAll('.custom-label-item');
  if (remainingItems.length === 0) {
    addCustomLabelItem('', 1, 0, true);
  }
  
  // ユーザーに削除結果を通知
  if (removedCount > 0) {
    alert(`未設定のカスタムラベル ${removedCount} 項目を削除しました。設定済みのラベルのみで処理を続行します。`);
  }
}

// 未設定のカスタムラベル項目を強調表示
function highlightEmptyCustomLabels() {
  const labels = getCustomLabelsFromUI();
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  
  items.forEach((item, index) => {
    if (labels[index] && labels[index].enabled && labels[index].text.trim() === '') {
      item.classList.add('error');
    } else {
      item.classList.remove('error');
    }
  });
}

// カスタムラベル項目の強調表示をクリア
function clearCustomLabelHighlights() {
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  
  items.forEach(item => {
    item.classList.remove('error');
  });
}

// カスタムラベル項目のインデックスを再設定
function reindexCustomLabelItems() {
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  
  items.forEach((item, index) => {
    item.dataset.index = index;
    
    // 削除ボタンのonclick属性も更新
    const deleteButton = item.querySelector('.btn-danger');
    if (deleteButton) {
      deleteButton.setAttribute('onclick', `removeCustomLabelItem(${index})`);
    }
  });
}

function setupRichTextFormatting(editor) {
  debugLog('setupRichTextFormatting関数が呼び出されました, editor:', editor); // デバッグ用
  if (!editor) {
    console.error('setupRichTextFormatting: editor要素がnullです');
    return;
  }

  // 日本語入力（IME）中のEnterは改行処理を抑止するためのフラグ
  let isComposing = false;
  editor.addEventListener('compositionstart', () => { 
    isComposing = true; 
    debugLog('[editor] compositionstart');
  });
  editor.addEventListener('compositionend', () => { 
    isComposing = false; 
    debugLog('[editor] compositionend');
  });

  // 共通の改行挿入ヘルパー
  const insertLineBreak = (source = 'unknown') => {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const range = selection.getRangeAt(0);
    const beforeHTML = editor.innerHTML;

    // 末尾かどうかを判定（テキストベース）
    let atEnd = false;
    try {
      const tail = document.createRange();
      tail.selectNodeContents(editor);
      tail.setStart(range.endContainer, range.endOffset);
      const remainingText = tail.toString();
      atEnd = remainingText.length === 0;
    } catch {}

    // 選択範囲がある場合は削除
    range.deleteContents();
    // 改行を挿入
    const br = document.createElement('br');
    range.insertNode(br);

    if (atEnd) {
      // 末尾では <br> の直後にゼロ幅スペースを1つだけ入れて視覚的な改行を保証
      const zwsp = document.createTextNode('\u200B');
      if (br.parentNode) {
        if (br.nextSibling && br.nextSibling.nodeType === Node.TEXT_NODE && br.nextSibling.nodeValue.startsWith('\u200B')) {
          // 既にゼロ幅スペースがある場合は重複挿入しない
        } else {
          br.parentNode.insertBefore(zwsp, br.nextSibling);
          // キャレットをゼロ幅スペースの後ろに
          range.setStartAfter(zwsp);
        }
      }
    } else {
      // キャレットを改行の直後へ
      range.setStartAfter(br);
    }

    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);

    const afterHTML = editor.innerHTML;
    debugLog(`[editor] insertLineBreak via ${source}`, { beforeLen: beforeHTML.length, afterLen: afterHTML.length, atEnd });
  };

  // 近代ブラウザ: beforeinput で Enter(=insertParagraph) を横取りし <br> を1回だけ入れる
  const supportsBeforeInput = 'onbeforeinput' in editor;
  if (supportsBeforeInput) {
    debugLog('[editor] supports beforeinput = true');
    editor.addEventListener('beforeinput', function(e) {
      if (e.inputType === 'insertParagraph') {
        debugLog('[editor] beforeinput insertParagraph', { isComposing, evIsComposing: e.isComposing, cancelable: e.cancelable });
        if (isComposing || e.isComposing) {
          // IME確定Enterは改行処理しない
          return;
        }
        e.preventDefault();
        insertLineBreak('beforeinput');
      }
    });
  }
  else {
    debugLog('[editor] supports beforeinput = false');
  }
  
  // Enterキーでの改行処理を改善
  editor.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      debugLog('[editor] keydown Enter', { supportsBeforeInput, isComposing, evIsComposing: e.isComposing, repeat: e.repeat, shift: e.shiftKey, alt: e.altKey, ctrl: e.ctrlKey, meta: e.metaKey });
      // beforeinputが使える環境ではそちらで一元処理する
      if (supportsBeforeInput) return;
      // IME入力中はブラウザに任せる（確定用Enterを奪わない）
      if (isComposing || e.isComposing) return;
      // フォールバック: ここで1回だけ<br>を入れる
      e.preventDefault();
      insertLineBreak('keydown');
    }
  });

  // コンテキストメニューの追加
  editor.addEventListener('contextmenu', async function(e) {
    e.preventDefault();
    
    // 選択範囲の有無を確認
    const selection = window.getSelection();
    const hasSelection = selection.toString().length > 0;
    
    const menu = await createFontSizeMenu(e.clientX, e.clientY, editor, hasSelection);
    document.body.appendChild(menu);
    
    // クリック外でメニューを閉じる
    setTimeout(() => {
      document.addEventListener('click', function closeMenu() {
        closeContextMenu(menu);
        document.removeEventListener('click', closeMenu);
      });
    }, 100);
  });
}

// メニューを閉じるヘルパー関数（フェードアウト効果付き）
function closeContextMenu(menu) {
  if (menu && menu.parentNode) {
    menu.style.opacity = '0';
    setTimeout(() => {
      if (menu.parentNode) {
        menu.parentNode.removeChild(menu);
      }
    }, 200);
  }
}

// フォントサイズ選択メニューを作成（非同期）
async function createFontSizeMenu(x, y, editor, hasSelection = true) {
  const menu = document.createElement('div');
  
  // 初期スタイル設定（位置は後で調整）
  menu.style.cssText = `
    position: fixed;
    background: white;
    border: 1px solid #ccc;
    border-radius: 6px;
    padding: 8px 0;
    box-shadow: 0 4px 20px rgba(0,0,0,0.15);
    z-index: 10000;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    min-width: 160px;
    max-width: 250px;
    max-height: 400px;
    overflow-y: auto;
    visibility: hidden;
    opacity: 0;
    transition: opacity 0.2s ease;
  `;
  
  // 一時的に追加してサイズを測定
  document.body.appendChild(menu);
  
  // フォーマットオプション
  const formatOptions = [
    { label: '太字', command: 'bold', style: 'font-weight: bold;' },
    { label: '斜体', command: 'italic', style: 'font-style: italic;' },
    { label: '下線', command: 'underline', style: 'text-decoration: underline;' },
    { label: 'すべてクリア', command: 'clear', style: 'color: #dc3545; font-weight: bold;' }
  ];
  
  // 選択範囲がない場合はクリア機能のみ表示
  const availableOptions = hasSelection ? formatOptions : [formatOptions[3]]; // クリア機能のみ
  
  // フォーマットボタンを追加
  availableOptions.forEach(option => {
    const item = document.createElement('div');
    item.textContent = option.label;
    item.style.cssText = `
      padding: 8px 15px;
      cursor: pointer;
      font-size: 12px;
      transition: background-color 0.2s;
      ${option.style}
    `;
    
    item.addEventListener('mouseenter', function() {
      this.style.backgroundColor = '#f0f0f0';
    });
    
    item.addEventListener('mouseleave', function() {
      this.style.backgroundColor = 'transparent';
    });
    
    item.addEventListener('mousedown', function(e) {
      e.preventDefault(); // 選択範囲がクリアされるのを防ぐ
      e.stopPropagation();
    });
    
    item.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      
      // 選択範囲を保持してフォーマットを適用
      setTimeout(() => {
        applyFormatToSelection(option.command, editor);
        
        // メニューを閉じる
        closeContextMenu(menu);
        
        saveCustomLabels();
      }, 10);
    });
    
    menu.appendChild(item);
  });
  
  // 選択範囲がある場合のみフォントサイズオプションを追加
  if (hasSelection) {
    // 区切り線
    const separator = document.createElement('div');
    separator.style.cssText = `
      height: 1px;
      background-color: #ddd;
      margin: 5px 0;
    `;
    menu.appendChild(separator);
    
    // フォントサイズオプション
    const fontSizeLabel = document.createElement('div');
    fontSizeLabel.textContent = 'フォントサイズ';
    fontSizeLabel.style.cssText = `
      padding: 5px 15px;
      font-size: 11px;
      color: #666;
      font-weight: bold;
    `;
    menu.appendChild(fontSizeLabel);
    
    const fontSizes = [6, 8, 10, 12, 14, 16, 18, 20, 24, 28];
    
    fontSizes.forEach(size => {
      const item = document.createElement('div');
      item.textContent = `${size}pt`;
      item.style.cssText = `
        padding: 6px 20px;
        cursor: pointer;
        font-size: 11px;
        transition: background-color 0.2s;
      `;
      
      item.addEventListener('mouseenter', function() {
        this.style.backgroundColor = '#f0f0f0';
      });
      
      item.addEventListener('mouseleave', function() {
        this.style.backgroundColor = 'transparent';
      });
      
      item.addEventListener('mousedown', function(e) {
        e.preventDefault(); // 選択範囲がクリアされるのを防ぐ
        e.stopPropagation();
      });
      
      item.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        // 選択範囲を保持してフォントサイズを適用
        setTimeout(() => {
          applyFontSizeToSelection(size, editor);
          
          // メニューを閉じる
          closeContextMenu(menu);
          
          saveCustomLabels();
        }, 10);
      });
      
      menu.appendChild(item);
    });

    // カスタムフォント選択オプション（統合DBベース）
    try {
      if (fontManager) {
        // フォント用区切り線（常に表示）
        const fontSeparator = document.createElement('div');
        fontSeparator.style.cssText = `
          height: 1px;
          background-color: #ddd;
          margin: 5px 0;
        `;
        menu.appendChild(fontSeparator);

        // フォントファミリーオプション（常に表示）
        const fontFamilyLabel = document.createElement('div');
        fontFamilyLabel.textContent = 'フォント';
        fontFamilyLabel.style.cssText = `
          padding: 5px 15px;
          font-size: 11px;
          color: #666;
          font-weight: bold;
        `;
        menu.appendChild(fontFamilyLabel);

        // デフォルトフォント（常に表示）
        const defaultFontItem = document.createElement('div');
        defaultFontItem.textContent = 'デフォルトフォント（システムフォント）';
        defaultFontItem.style.cssText = `
          padding: 6px 20px;
          cursor: pointer;
          font-size: 11px;
          transition: background-color 0.2s;
          font-family: sans-serif;
          border-bottom: 1px solid #eee;
          font-weight: bold;
          color: #333;
        `;
      
        defaultFontItem.addEventListener('mouseenter', function() {
          this.style.backgroundColor = '#f0f0f0';
        });
        
        defaultFontItem.addEventListener('mouseleave', function() {
          this.style.backgroundColor = 'transparent';
        });
        
        defaultFontItem.addEventListener('mousedown', function(e) {
          e.preventDefault();
          e.stopPropagation();
        });
        
        defaultFontItem.addEventListener('click', function(e) {
          e.preventDefault();
          e.stopPropagation();
          
          setTimeout(() => {
            // より確実にデフォルトフォントに戻す処理
            try {
              const selection = window.getSelection();
              if (selection.rangeCount > 0 && !selection.isCollapsed) {
                const range = selection.getRangeAt(0);
                
                // シンプルにフォント削除（applyFontFamilyToSelectionを使用）
                applyFontFamilyToSelection('', editor);
                
                // 成功メッセージをコンソールに出力
                console.log('選択範囲をデフォルトフォントに戻しました');
              } else {
                // 選択範囲がない場合はエディタ全体をデフォルトに戻す
                const editorStyle = editor.style;
                if (editorStyle.fontFamily) {
                  editorStyle.fontFamily = '';
                  console.log('エディタ全体をデフォルトフォントに戻しました');
                }
              }
            } catch (error) {
              console.error('デフォルトフォント設定エラー:', error);
              // フォールバック: 古い方法で処理
              applyFontFamilyToSelection('', editor);
            }
            
            closeContextMenu(menu);
            
            saveCustomLabels();
          }, 10);
        });
        
        menu.appendChild(defaultFontItem);

        // カスタムフォントを取得
        let customFonts = {};
        if (fontManager) {
          try {
            customFonts = await fontManager.getAllFonts();
          } catch (error) {
            console.error('カスタムフォント取得エラー:', error);
            customFonts = {};
          }
        }

        // システムフォント
        const systemFonts = [
          { name: 'ゴシック（sans-serif）', family: 'sans-serif' },
          { name: '明朝（serif）', family: 'serif' },
          { name: '等幅（monospace）', family: 'monospace' },
          { name: 'Arial', family: 'Arial, sans-serif' },
          { name: 'Times New Roman', family: 'Times New Roman, serif' },
          { name: 'メイリオ', family: 'Meiryo, sans-serif' },
          { name: 'ヒラギノ角ゴ', family: 'Hiragino Kaku Gothic Pro, sans-serif' }
        ];

        if (systemFonts.length > 0) {
          // システムフォントセクションのラベル
          const systemFontLabel = document.createElement('div');
          systemFontLabel.textContent = 'システムフォント';
          systemFontLabel.style.cssText = `
            padding: 5px 15px;
            font-size: 10px;
            color: #666;
            font-weight: bold;
            border-bottom: 1px solid #eee;
          `;
          menu.appendChild(systemFontLabel);

          systemFonts.forEach(font => {
            const fontItem = document.createElement('div');
            fontItem.textContent = font.name;
            fontItem.style.cssText = `
              padding: 6px 20px;
              cursor: pointer;
              font-size: 11px;
              transition: background-color 0.2s;
              font-family: ${font.family};
            `;
            
            fontItem.addEventListener('mouseenter', function() {
              this.style.backgroundColor = '#f0f0f0';
            });
            
            fontItem.addEventListener('mouseleave', function() {
              this.style.backgroundColor = 'transparent';
            });
            
            fontItem.addEventListener('mousedown', function(e) {
              e.preventDefault();
              e.stopPropagation();
            });
            
            fontItem.addEventListener('click', function(e) {
              e.preventDefault();
              e.stopPropagation();
              
              setTimeout(() => {
                applyFontFamilyToSelection(font.family, editor);
                
                closeContextMenu(menu);
                
                saveCustomLabels();
              }, 10);
            });
            
            menu.appendChild(fontItem);
          });
        }

        // カスタムフォント
        if (Object.keys(customFonts).length > 0) {
          // カスタムフォントセクションのラベル
          const customFontLabel = document.createElement('div');
          customFontLabel.textContent = 'カスタムフォント';
          customFontLabel.style.cssText = `
            padding: 5px 15px;
            font-size: 10px;
            color: #666;
            font-weight: bold;
            border-top: 1px solid #eee;
            border-bottom: 1px solid #eee;
          `;
          menu.appendChild(customFontLabel);

          Object.keys(customFonts).forEach(fontName => {
            const fontItem = document.createElement('div');
            fontItem.textContent = fontName;
            fontItem.style.cssText = `
              padding: 6px 20px;
              cursor: pointer;
              font-size: 11px;
              transition: background-color 0.2s;
              font-family: "${fontName}", sans-serif;
            `;
            
            fontItem.addEventListener('mouseenter', function() {
              this.style.backgroundColor = '#f0f0f0';
            });
            
            fontItem.addEventListener('mouseleave', function() {
              this.style.backgroundColor = 'transparent';
            });
            
            fontItem.addEventListener('mousedown', function(e) {
              e.preventDefault();
              e.stopPropagation();
            });
            
            fontItem.addEventListener('click', function(e) {
              e.preventDefault();
              e.stopPropagation();
              
              setTimeout(() => {
                applyFontFamilyToSelection(fontName, editor);
                
                closeContextMenu(menu);
                
                saveCustomLabels();
              }, 10);
            });
            
            menu.appendChild(fontItem);
          });
        }
      }
    } catch (error) {
      console.error('カスタムフォント読み込みエラー:', error);
    }
  }
  
  // メニューのサイズを取得して位置を調整
  const menuRect = menu.getBoundingClientRect();
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const scrollX = window.scrollX || window.pageXOffset;
  const scrollY = window.scrollY || window.pageYOffset;
  
  // 安全なマージンを設定
  const margin = 10;
  
  // 水平位置の調整
  let adjustedX = x;
  if (x + menuRect.width > viewportWidth) {
    // 右端からはみ出る場合は左にずらす
    adjustedX = viewportWidth - menuRect.width - margin;
  }
  // 左端からはみ出る場合は右にずらす
  if (adjustedX < margin) {
    adjustedX = margin;
  }
  
  // 垂直位置の調整
  let adjustedY = y;
  if (y + menuRect.height > viewportHeight) {
    // 下端からはみ出る場合は上にずらす
    adjustedY = y - menuRect.height - margin;
  }
  // 上端からはみ出る場合の処理
  if (adjustedY < scrollY + margin) {
    // 画面上端より上に行く場合は、画面内の適切な位置に配置
    if (y + menuRect.height <= viewportHeight) {
      // 元の位置（下向き）で画面内に収まる場合
      adjustedY = y;
    } else {
      // どちらも画面からはみ出る場合は、画面上端に近い位置に配置
      adjustedY = scrollY + margin;
    }
  }
  
  // 最終的な位置の安全性チェック
  adjustedX = Math.max(margin, Math.min(adjustedX, viewportWidth - menuRect.width - margin));
  adjustedY = Math.max(scrollY + margin, Math.min(adjustedY, scrollY + viewportHeight - menuRect.height - margin));
  
  // 調整後の位置を設定して表示
  menu.style.left = `${adjustedX}px`;
  menu.style.top = `${adjustedY}px`;
  menu.style.visibility = 'visible';
  
  // フェードイン効果
  setTimeout(() => {
    menu.style.opacity = '1';
  }, 10);
  
  // console.log(`メニュー位置調整: 元(${x}, ${y}) → 調整後(${adjustedX}, ${adjustedY}), サイズ: ${menuRect.width}x${menuRect.height}`);
  
  return menu;
}

// 選択範囲にフォーマットを適用（ブラウザ標準のexecCommandを使用）
function applyFormatToSelection(command, editor) {
  if (command === 'clear') {
    // すべてクリア：エディタ全体をクリア
    clearAllContent(editor);
    return;
  }
  
  // ブラウザの標準的なexecCommandを使用（Ctrl+Bと同じ動作）
  try {
    let execCommand;
    switch (command) {
      case 'bold':
        execCommand = 'bold';
        break;
      case 'italic':
        execCommand = 'italic';
        break;
      case 'underline':
        execCommand = 'underline';
        break;
      default:
        return; // 不明なコマンドの場合は何もしない
    }
    
    // execCommandを実行（トグル動作も自動的に行われる）
    document.execCommand(execCommand, false, null);
    
  } catch (error) {
    console.warn('execCommandの実行に失敗しました:', error);
    // フォールバック：従来の方法
    applyFormatToSelectionFallback(command, editor);
  }
  
  // エディタにフォーカスを戻す
  editor.focus();
}

// フォールバック用の従来の実装
function applyFormatToSelectionFallback(command, editor) {
  const selection = window.getSelection();
  
  if (selection.rangeCount === 0 || selection.isCollapsed) {
    return; // 選択範囲がない場合は何もしない
  }
  
  const range = selection.getRangeAt(0);
  
  // 選択範囲が既にフォーマットされているかチェック
  if (isSelectionFormatted(range, command)) {
    // フォーマットが適用されている場合は解除
    removeFormatFromSelection(range, command);
  } else {
    // フォーマットが適用されていない場合は適用
    applyFormatToRange(range, command);
  }
}

// 選択範囲が指定されたフォーマットで装飾されているかチェック
function isSelectionFormatted(range, command) {
  const targetTag = getTargetTagName(command);
  if (!targetTag) return false;
  
  // 選択範囲の開始ノードから親要素を遡って該当するフォーマット要素を探す
  let node = range.startContainer;
  if (node.nodeType === Node.TEXT_NODE) {
    node = node.parentNode;
  }
  
  while (node && node.closest && node.closest('.rich-text-editor')) {
    if (node.tagName === targetTag) {
      return true;
    }
    node = node.parentNode;
  }
  
  return false;
}

// コマンドに対応するHTMLタグ名を取得
function getTargetTagName(command) {
  switch (command) {
    case 'bold':
      return 'STRONG';
    case 'italic':
      return 'EM';
    case 'underline':
      return 'U';
    default:
      return null;
  }
}

// 選択範囲にフォーマットを適用
function applyFormatToRange(range, command) {
  const selectedContent = range.extractContents();
  
  let wrapper;
  switch (command) {
    case 'bold':
      wrapper = document.createElement('strong');
      break;
    case 'italic':
      wrapper = document.createElement('em');
      break;
    case 'underline':
      wrapper = document.createElement('u');
      break;
    default:
      return; // 不明なコマンドの場合は何もしない
  }
  
  wrapper.appendChild(selectedContent);
  range.insertNode(wrapper);
  
  // 新しい選択範囲を設定
  range.selectNodeContents(wrapper);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
}

// 選択範囲からフォーマットを削除
function removeFormatFromSelection(range, command) {
  const targetTag = getTargetTagName(command);
  if (!targetTag) return;
  
  // 選択範囲の開始ノードから親要素を遡って該当するフォーマット要素を探す
  let formatElement = null;
  let node = range.startContainer;
  if (node.nodeType === Node.TEXT_NODE) {
    node = node.parentNode;
  }
  
  while (node && node.closest && node.closest('.rich-text-editor')) {
    if (node.tagName === targetTag) {
      formatElement = node;
      break;
    }
    node = node.parentNode;
  }
  
  if (!formatElement) return;
  
  // フォーマット要素の親と位置を記録
  const parent = formatElement.parentNode;
  const editor = formatElement.closest('.rich-text-editor');
  
  // フォーマット要素の内容を取得
  const content = document.createDocumentFragment();
  const childNodes = Array.from(formatElement.childNodes); // 配列にコピー
  childNodes.forEach(child => content.appendChild(child));
  
  // フォーマット要素を内容で置き換え
  parent.replaceChild(content, formatElement);
  
  // 選択範囲を復元（より安全な方法）
  const selection = window.getSelection();
  selection.removeAllRanges();
  
  if (editor && childNodes.length > 0) {
    try {
      const newRange = document.createRange();
      // 最初の子ノードから最後の子ノードまでを選択
      const firstNode = childNodes[0];
      const lastNode = childNodes[childNodes.length - 1];
      
      if (firstNode.nodeType === Node.TEXT_NODE) {
        newRange.setStart(firstNode, 0);
      } else {
        newRange.setStartBefore(firstNode);
      }
      
      if (lastNode.nodeType === Node.TEXT_NODE) {
        newRange.setEnd(lastNode, lastNode.textContent.length);
      } else {
        newRange.setEndAfter(lastNode);
      }
      
      selection.addRange(newRange);
    } catch (e) {
      // 範囲設定に失敗した場合はエディタの末尾にカーソルを置く
      try {
        const newRange = document.createRange();
        newRange.selectNodeContents(editor);
        newRange.collapse(false); // 末尾に移動
        selection.addRange(newRange);
      } catch (e2) {
        // それでも失敗した場合は何もしない
        console.warn('選択範囲の復元に失敗しました:', e2);
      }
    }
  }
}

// エディタの全内容をクリア（書式も含めて）
function clearAllContent(editor) {
  // 確認ダイアログを表示
  if (confirm('このカスタムラベルの内容と書式をすべてクリアしますか？')) {
    // エディタの内容を完全にクリア
    editor.innerHTML = '';
    
    // デフォルトのスタイルを再設定
    editor.style.fontSize = '12pt';
    editor.style.lineHeight = '1.2';
    editor.style.textAlign = 'center';
    
    // フォーカスを設定
    editor.focus();
    
    // カスタムラベルを保存
    saveCustomLabels();
  }
}

// 選択範囲にフォントサイズを適用
// テキストのみ入力可能にする設定
function setupTextOnlyEditor(editor) {
  debugLog('setupTextOnlyEditor関数が呼び出されました, editor:', editor); // デバッグ用
  if (!editor) {
    console.error('setupTextOnlyEditor: editor要素がnullです');
    return;
  }
  
  // 画像やその他のメディアの貼り付けを防ぐ
  editor.addEventListener('paste', function(e) {
    e.preventDefault();
    
    // プレーンテキストのみを取得
    const text = (e.clipboardData || window.clipboardData).getData('text/plain');
    
    // HTMLタグを除去し、改行を<br>に変換
    const cleanText = text.replace(/<[^>]*>/g, '');
    
    // 現在の選択範囲にテキストを挿入
    const selection = window.getSelection();
    if (selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      range.deleteContents();
      
      // 改行を含むテキストを適切に処理
      const lines = cleanText.split('\n');
      lines.forEach((line, index) => {
        if (index > 0) {
          // 改行を挿入
          const br = document.createElement('br');
          range.insertNode(br);
          range.setStartAfter(br);
        }
        
        if (line.length > 0) {
          // テキストを挿入
          const textNode = document.createTextNode(line);
          range.insertNode(textNode);
          range.setStartAfter(textNode);
        }
      });
      
      // カーソルを最後に移動
      range.collapse(true);
      selection.removeAllRanges();
      selection.addRange(range);
    }
  });
  
  // ドラッグ&ドロップでの画像挿入を防ぐ
  editor.addEventListener('drop', function(e) {
    e.preventDefault();
    
    // ドロップされたファイルがある場合は何もしない
    if (e.dataTransfer.files.length > 0) {
      return false;
    }
    
    // テキストのみ許可
    const text = e.dataTransfer.getData('text/plain');
    if (text) {
      const cleanText = text.replace(/<[^>]*>/g, '');
      
      // 改行を含むテキストを適切に処理
      const selection = window.getSelection();
      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        range.deleteContents();
        
        const lines = cleanText.split('\n');
        lines.forEach((line, index) => {
          if (index > 0) {
            const br = document.createElement('br');
            range.insertNode(br);
            range.setStartAfter(br);
          }
          
          if (line.length > 0) {
            const textNode = document.createTextNode(line);
            range.insertNode(textNode);
            range.setStartAfter(textNode);
          }
        });
        
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
      }
    }
  });
  
  // ドラッグオーバー時の処理
  editor.addEventListener('dragover', function(e) {
    // ファイルのドラッグオーバーの場合は拒否
    if (e.dataTransfer.types.includes('Files')) {
      e.dataTransfer.dropEffect = 'none';
      return false;
    }
    e.preventDefault();
  });
  
  // 直接的なHTML挿入を監視してテキストのみに変換
  const observer = new MutationObserver(function(mutations) {
    mutations.forEach(function(mutation) {
      if (mutation.type === 'childList') {
        mutation.addedNodes.forEach(function(node) {
          if (node.nodeType === Node.ELEMENT_NODE) {
            // 画像やその他の要素が挿入された場合は削除
            if (node.tagName === 'IMG' || node.tagName === 'VIDEO' || node.tagName === 'AUDIO') {
              node.remove();
            }
          }
        });
      }
    });
  });
  
  observer.observe(editor, {
    childList: true,
    subtree: true
  });
}

// 印刷ボタンのイベントリスナーを変更
document.addEventListener('DOMContentLoaded', function() {
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
      const customLabels = getCustomLabelsFromUI();
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
      await updateCustomLabelsSummary();
      console.log('✅ カスタムラベルサマリー更新完了');
    } catch (summaryError) {
      console.error('⚠️ カスタムラベルサマリー更新エラー:', summaryError);
      // サマリー更新エラーは致命的ではないので、処理を継続
    }

    // 印刷済み注文番号の印刷日時をIndexedDBに記録
    try {
      if (!window.unifiedDB) await StorageManager.ensureDatabase();
      const now = new Date().toISOString();
      const orderPages = document.querySelectorAll('.page');
      for (const page of orderPages) {
        const orderNumber = OrderNumberManager.getFromOrderSection(page);
        if (orderNumber) {
          await window.unifiedDB.setPrintedAt(orderNumber, now);
        }
      }
      console.log('✅ 印刷済み注文番号の印刷日時を保存しました');
    } catch (e) {
      console.error('❌ 印刷済み注文番号の保存エラー:', e);
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
function initializeFontDropZone() {
  // FontManagerを初期化（まだ初期化されていない場合）
  if (!fontManager) {
    initializeFontManager();
  }

  const dropZone = document.createElement('div');
  dropZone.style.cssText = `
    border: 2px dashed #ccc;
    border-radius: 8px;
    padding: 20px;
    margin: 10px 0;
    text-align: center;
    cursor: pointer;
    background: #f9f9f9;
    transition: all 0.3s ease;
  `;
  dropZone.innerHTML = `
    <p style="margin: 0 0 10px 0; color: #666;">
      フォントファイルをここにドロップ<br>
      <small>対応形式: TTF, OTF, WOFF, WOFF2</small><br>
      <small>IndexedDBに保存されるため容量制限はありません</small>
    </p>
  `;

  // ドロップイベント
  dropZone.addEventListener('dragover', function(e) {
    e.preventDefault();
    this.style.borderColor = '#007bff';
    this.style.background = '#e7f3ff';
  });

  dropZone.addEventListener('dragleave', function(e) {
    e.preventDefault();
    this.style.borderColor = '#ccc';
    this.style.background = '#f9f9f9';
  });

  dropZone.addEventListener('drop', function(e) {
    e.preventDefault();
    this.style.borderColor = '#ccc';
    this.style.background = '#f9f9f9';
    
    const files = Array.from(e.dataTransfer.files);
    files.forEach(file => {
      if (CONSTANTS.FONT.SUPPORTED_FORMATS.test(file.name)) {
        handleFontFile(file);
      } else {
        alert(`${file.name} は対応していないフォント形式です。`);
      }
    });
  });

  // クリックでファイル選択
  dropZone.addEventListener('click', function(e) {
    // ボタンクリックの場合は無視
    if (e.target.tagName === 'BUTTON') {
      return;
    }
    
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = CONSTANTS.FONT.ACCEPTED_TYPES;
    input.multiple = true;
    input.addEventListener('change', function(e) {
      Array.from(e.target.files).forEach(file => {
        handleFontFile(file);
      });
    });
    input.click();
  });

  // HTMLに追加（既存のfontDropZone要素を使用）
  const fontDropZoneElement = document.getElementById('fontDropZone');
  if (fontDropZoneElement) {
    fontDropZoneElement.appendChild(dropZone);
  }

  // 既存のフォント一覧を表示
  updateFontList();
}

async function handleFontFile(file) {
  try {
    // fontManagerが初期化されていない場合
    if (!fontManager) {
      alert('フォントマネージャーが初期化されていません。\nページを再読み込みしてから再試行してください。');
      return;
    }
    
    // ファイルサイズの警告（100MB以上で警告）
    const warningSize = 100 * 1024 * 1024; // 100MB
    if (file.size > warningSize) {
      const fileSizeMB = (file.size / 1024 / 1024).toFixed(2);
      const proceed = confirm(
        `フォントファイルが非常に大きいです（${fileSizeMB}MB）。\n\n` +
        `処理に時間がかかる可能性がありますが、続行しますか？\n\n` +
        `※ブラウザの動作が重くなる場合があります。`
      );
      if (!proceed) {
        return;
      }
    }

    // ファイル形式チェック
    const supportedFormats = ['.ttf', '.otf', '.woff', '.woff2'];
    const fileName = file.name.toLowerCase();
    const isSupported = supportedFormats.some(format => fileName.endsWith(format));
    
    if (!isSupported) {
      alert(`サポートされていないファイル形式です。\n対応形式: ${supportedFormats.join(', ')}`);
      return;
    }

    // ローディング表示
    showFontUploadProgress(true);

    // ファイルをArrayBufferとして読み込み
    const arrayBuffer = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });

    // MIMEタイプの決定
    let mimeType = file.type;
    if (!mimeType) {
      if (fileName.endsWith('.ttf')) mimeType = 'font/ttf';
      else if (fileName.endsWith('.otf')) mimeType = 'font/otf';
      else if (fileName.endsWith('.woff')) mimeType = 'font/woff';
      else if (fileName.endsWith('.woff2')) mimeType = 'font/woff2';
      else mimeType = 'font/ttf';
    }

    // フォント名の生成（拡張子を除く）
    const fontName = file.name.replace(/\.[^/.]+$/, "");
    
    // FontManagerが初期化されていない場合は初期化
    if (!fontManager) {
      await initializeFontManager();
    }
    
    // 重複チェック
    const existingFont = await fontManager.getFont(fontName);
    if (existingFont) {
      if (!confirm(`フォント "${fontName}" は既に存在します。上書きしますか？`)) {
        showFontUploadProgress(false);
        return;
      }
    }

    // IndexedDBに保存
    await fontManager.saveFont(fontName, arrayBuffer, {
      type: mimeType,
      originalName: file.name
    });

    // CSS更新
    await loadCustomFontsCSS();
    
    // フォントリスト更新
    await updateFontList();

    console.log(`フォント "${fontName}" をIndexedDBに保存しました (${(arrayBuffer.byteLength / 1024 / 1024).toFixed(2)}MB)`);
    
    // 成功通知
    showSuccessMessage(`フォント "${fontName}" をアップロードしました`);

  } catch (error) {
    console.error('フォント処理エラー:', error);
    alert(`フォントファイルの処理中にエラーが発生しました：\n${error.message}`);
  } finally {
    showFontUploadProgress(false);
  }
}

// localStorage ベースの容量チェックは廃止 (IndexedDB 専用化に伴い削除済み)

// ArrayBufferをBase64に変換する関数（スタックオーバーフローを避ける）
function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 8192; // 8KB ずつ処理
  let binary = '';
  
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.slice(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }
  
  return btoa(binary);
}

function getFontMimeType(fileName) {
  const ext = fileName.toLowerCase().split('.').pop();
  const mimeTypes = {
    'ttf': 'font/ttf',
    'otf': 'font/otf',
    'woff': 'font/woff',
    'woff2': 'font/woff2'
  };
  return mimeTypes[ext] || 'font/ttf';
}

function addFontToCSS(fontName, base64Data, mimeType) {
  let styleElement = document.getElementById('custom-fonts-style');
  if (!styleElement) {
    styleElement = document.createElement('style');
    styleElement.id = 'custom-fonts-style';
    document.head.appendChild(styleElement);
  }

  const fontFace = `
    @font-face {
      font-family: "${fontName}";
      src: url("data:${mimeType};base64,${base64Data}") format("${getFontFormat(mimeType)}");
      font-display: swap;
    }
  `;

  styleElement.textContent += fontFace;
}

function getFontFormat(mimeType) {
  const formats = {
    'font/ttf': 'truetype',
    'font/otf': 'opentype',
    'font/woff': 'woff',
    'font/woff2': 'woff2'
  };
  return formats[mimeType] || 'truetype';
}

async function loadCustomFontsCSS() {
  try {
    // fontManagerが初期化されていない場合は何もしない
    if (!fontManager) {
      console.warn('FontManagerが初期化されていません');
      return;
    }
    
    const fonts = await fontManager.getAllFonts();
    
    // 既存のカスタムフォントCSSをクリア
    let styleElement = document.getElementById('custom-fonts-style');
    if (styleElement) {
      styleElement.remove();
    }
    
    if (Object.keys(fonts).length === 0) {
      console.log('カスタムフォントがありません');
      return;
    }
    
    // 新しいスタイル要素を作成
    styleElement = document.createElement('style');
    styleElement.id = 'custom-fonts-style';
    styleElement.type = 'text/css';
    
    let cssContent = '';
    
    for (const [fontName, fontData] of Object.entries(fonts)) {
      try {
        // データ構造の検証
        if (!fontData || !fontData.data || !fontData.metadata) {
          console.warn(`フォント "${fontName}" のデータ構造が不正です:`, fontData);
          continue;
        }
        
        // ArrayBufferからBase64に変換
        const base64Data = arrayBufferToBase64(fontData.data);
        
        const fontFaceRule = `
@font-face {
  font-family: "${fontName}";
  src: url(data:${fontData.metadata.type || 'font/ttf'};base64,${base64Data});
  font-display: swap;
}`;
        
        cssContent += fontFaceRule;
        console.log(`フォント "${fontName}" をCSSに追加しました`);
        
      } catch (error) {
        console.error(`フォント "${fontName}" のCSS追加でエラー:`, error);
      }
    }
    
    styleElement.textContent = cssContent;
    document.head.appendChild(styleElement);
    
    console.log(`${Object.keys(fonts).length}個のカスタムフォントを読み込みました`);
    
  } catch (error) {
    console.error('カスタムフォントCSS読み込みエラー:', error);
  }
}

async function updateFontList() {
  const fontListElement = document.getElementById('fontList');
  if (!fontListElement) return;

  try {
    if (!fontManager) {
      fontListElement.innerHTML = '<div class="font-list-placeholder">フォントマネージャーを初期化中...</div>';
      return;
    }

    const fonts = await fontManager.getAllFonts();
    const entries = Object.entries(fonts);

    if (entries.length === 0) {
      fontListElement.innerHTML = '<div class="font-list-placeholder">フォントファイルをアップロードしてください</div>';
      return;
    }

    // 既存ノードとの差分適用（シンプル版: 全クリア → 再構築）
    fontListElement.textContent = '';
    const tpl = document.getElementById('fontListItemTemplate');
    if (!tpl) {
      console.warn('fontListItemTemplate が見つかりません');
    }
    entries.forEach(([fontName, fontData]) => {
      const metadata = fontData.metadata || {};
      const originalName = metadata.originalName || fontName;
      const createdAt = metadata.createdAt || Date.now();
      const sizeMB = (fontData.data.byteLength / 1024 / 1024).toFixed(2);

      const node = tpl ? tpl.content.firstElementChild.cloneNode(true) : document.createElement('div');
      if (!tpl) node.className = 'font-list-item';
      node.dataset.fontName = fontName;

      const nameEl = node.querySelector('.font-name') || (() => { const d=document.createElement('div'); d.className='font-name'; node.appendChild(d); return d; })();
      nameEl.textContent = fontName;
      nameEl.style.fontFamily = `'${fontName}', sans-serif`;

      const metaEl = node.querySelector('.font-meta') || (() => { const d=document.createElement('div'); d.className='font-meta'; node.appendChild(d); return d; })();
      metaEl.innerHTML = `${escapeHTML(originalName)} <span>• ${new Date(createdAt).toLocaleDateString()} • ${sizeMB}MB</span>`;

      const btn = node.querySelector('.font-remove-btn') || (() => { const b=document.createElement('button'); b.type='button'; b.textContent='削除'; b.className='btn-base btn-danger btn-small font-remove-btn'; node.appendChild(b); return b; })();
      if (!btn._bound) {
        btn.addEventListener('click', () => removeFontFromList(fontName));
        btn._bound = true;
      }

      fontListElement.appendChild(node);
    });
  } catch (error) {
    console.error('フォントリスト更新エラー:', error);
  fontListElement.innerHTML = '<div class="font-list-error">フォントリストの読み込みに失敗しました</div>';
  } finally {
    setTimeout(adjustFontSectionHeight, 100);
  }
}

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

async function removeFontFromList(fontName) {
  try {
    // fontManagerが初期化されていない場合
    if (!fontManager) {
      alert('フォントマネージャーが初期化されていません。');
      return;
    }
    
    const fontData = await fontManager.getFont(fontName);
    
    if (!fontData) {
      alert('指定されたフォントが見つかりません。');
      return;
    }
    
    const originalName = fontData.originalName || fontName;
    const sizeMB = (fontData.size / 1024 / 1024).toFixed(2);
    
    const confirmMessage = `フォント "${fontName}" (${originalName}, ${sizeMB}MB) を削除しますか？\n\n削除すると、このフォントを使用しているカスタムラベルの表示が変わる可能性があります。`;
    
    if (confirm(confirmMessage)) {
      try {
        await fontManager.deleteFont(fontName);
        await updateFontList();
        await loadCustomFontsCSS();
        
        showSuccessMessage(`フォント "${fontName}" を削除しました`);
        
      } catch (error) {
        console.error('フォント削除エラー:', error);
        alert('フォントの削除中にエラーが発生しました。');
      }
    }
  } catch (error) {
    console.error('フォント削除エラー:', error);
    alert('フォントの削除中にエラーが発生しました。');
  }
}

// シンプルで確実なスタイル適用関数（ネスト防止版・改良版）
function applyStyleToSelection(styleProperty, styleValue, editor, isDefault = false) {
  const selection = window.getSelection();
  
  if (selection.rangeCount === 0 || selection.isCollapsed) {
    return; // 選択範囲がない場合は何もしない
  }
  
  try {
    const range = selection.getRangeAt(0);
    debugLog(`スタイル変更: "${styleProperty}: ${styleValue}" を適用中...`);
    
    // 選択範囲のテキスト内容を取得
    const selectedText = range.toString();
    debugLog('選択されたテキスト:', selectedText);
    
    if (!selectedText) {
      debugLog('選択されたテキストがありません');
      return;
    }
    
    // 選択範囲の分析
    const rangeInfo = analyzeSelectionRange(range);
    
    if (rangeInfo.isCompleteSpan) {
      debugLog('既存span要素を更新:', rangeInfo.targetSpan);
      updateSpanStyle(rangeInfo.targetSpan, styleProperty, styleValue, isDefault);
      // 選択全体に統一適用する際、子孫spanに同一プロパティがあれば削除して親の指定を有効化
      if ((styleProperty === 'font-size' || styleProperty === 'font-family')) {
        removeStyleFromDescendants(rangeInfo.targetSpan, styleProperty);
      }
      
    } else if (rangeInfo.isPartialSpan) {
      debugLog('部分選択でspan分割処理');
      handlePartialSpanSelection(range, rangeInfo, styleProperty, styleValue, isDefault);
      
    } else if (rangeInfo.isMultiSpan) {
      debugLog('複数span要素を統合処理:', rangeInfo.multiSpans.length + '個');
      handleMultiSpanSelection(range, rangeInfo, styleProperty, styleValue, isDefault);
      // 各対象spanの子孫からも当該プロパティを除去
      if ((styleProperty === 'font-size' || styleProperty === 'font-family')) {
        rangeInfo.multiSpans.forEach(s => removeStyleFromDescendants(s, styleProperty));
      }
      
    } else {
        debugLog('新しいspan要素を作成（BR保持版）');
        // BRや特殊ノードを壊さない安全な適用
        applyStylePreservingBreaks(range, styleProperty, styleValue, isDefault);
      }
    
    debugLog(`スタイル "${styleProperty}: ${styleValue}" を適用しました`);
    debugLog('処理後のエディタHTML:', editor.innerHTML);
    
    // エディタ全体の空のspan要素を掃除
    cleanupEmptySpans(editor);
    
    debugLog('スタイル変更完了');
    
  } catch (error) {
    console.warn('スタイル適用エラー:', error);
  }
  
  // エディタにフォーカスを戻す
  editor.focus();
}

// 選択範囲を分析するヘルパー関数
function analyzeSelectionRange(range) {
  const commonAncestor = range.commonAncestorContainer;
  let targetSpan = null;
  let isCompleteSpan = false;
  let isPartialSpan = false;
  let isMultiSpan = false;
  let multiSpans = [];
  
  // テキストノードの場合は親要素をチェック
  if (commonAncestor.nodeType === Node.TEXT_NODE) {
    const parent = commonAncestor.parentElement;
    if (parent && parent.tagName === 'SPAN') {
      const spanText = parent.textContent;
      const selectedText = range.toString();
      
      if (selectedText === spanText) {
        // 完全一致：span要素全体が選択されている
        targetSpan = parent;
        isCompleteSpan = true;
      } else {
        // 部分一致：span要素の一部が選択されている
        targetSpan = parent;
        isPartialSpan = true;
      }
    }
  } else if (commonAncestor.tagName === 'SPAN') {
    targetSpan = commonAncestor;
    isCompleteSpan = true;
  } else {
    // 複数のspan要素にまたがる選択の可能性をチェック
    const selectedText = range.toString();
    const spans = Array.from(commonAncestor.querySelectorAll('span'));
    
    // 選択範囲に含まれるspan要素を検出
    const rangeSpans = spans.filter(span => {
      const spanRange = document.createRange();
      spanRange.selectNodeContents(span);
      
      // 選択範囲とspan要素が重複するかチェック
      return range.intersectsNode(span) && selectedText.includes(span.textContent);
    });
    
    if (rangeSpans.length > 1) {
      // 複数のspan要素にまたがる選択
      isMultiSpan = true;
      multiSpans = rangeSpans;
    }
  }
  
  return {
    targetSpan,
    isCompleteSpan,
    isPartialSpan,
    isMultiSpan,
    multiSpans,
    commonAncestor
  };
}

// span要素のスタイルを更新するヘルパー関数
function updateSpanStyle(span, styleProperty, styleValue, isDefault) {
  const currentStyle = span.getAttribute('style') || '';
  const styleMap = parseStyleString(currentStyle);
  
  // 新しいスタイルを適用または削除
  if (isDefault) {
    styleMap.delete(styleProperty);
  } else {
    const unit = styleProperty === 'font-size' && typeof styleValue === 'number' ? 'pt' : '';
    styleMap.set(styleProperty, styleValue + unit);
  }
  
  // スタイル属性を再構築
  if (styleMap.size === 0) {
    span.removeAttribute('style');
  } else {
    const newStyle = Array.from(styleMap.entries())
      .map(([prop, val]) => `${prop}: ${val}`)
      .join('; ');
    span.setAttribute('style', newStyle);
  }
  
  debugLog('span要素のスタイルを更新:', span.getAttribute('style') || '(スタイルなし)');
}

// 部分選択時のspan分割処理
function handlePartialSpanSelection(range, rangeInfo, styleProperty, styleValue, isDefault) {
  const targetSpan = rangeInfo.targetSpan;
  const selectedText = range.toString();

  // Textノード内の部分選択ならsplitTextで安全に分割（BR等の子要素を保持）
  if (range.startContainer === range.endContainer && range.startContainer.nodeType === Node.TEXT_NODE) {
    const textNode = range.startContainer;
    const start = range.startOffset;
    const end = range.endOffset;

    const fullText = textNode.nodeValue || '';
    const beforeText = fullText.slice(0, start);
    const midText = fullText.slice(start, end);
    const afterText = fullText.slice(end);

    // 既存Textノードを書き換え（before）
    textNode.nodeValue = beforeText;

    // 選択部分を新spanで挿入
    const selectedSpan = document.createElement('span');
    if (!isDefault) {
      const unit = styleProperty === 'font-size' && typeof styleValue === 'number' ? 'pt' : '';
      try { selectedSpan.style.setProperty(styleProperty, (typeof styleValue === 'number' ? String(styleValue) : styleValue) + unit); } catch {}
    } else {
      // 既定化の場合は、親のスタイルを打ち消したいが、ここでは明示的な解除はせずそのままテキストを包む
      // （フォントリセット処理は専用のフローに委ねる）
    }
    selectedSpan.appendChild(document.createTextNode(midText));

    // afterテキストノード
    const afterNode = document.createTextNode(afterText);

    // DOMに挿入: textNodeの直後にselectedSpan, その後にafterNode
    const parent = textNode.parentNode;
    if (textNode.nextSibling) {
      parent.insertBefore(selectedSpan, textNode.nextSibling);
      parent.insertBefore(afterNode, selectedSpan.nextSibling);
    } else {
      parent.appendChild(selectedSpan);
      parent.appendChild(afterNode);
    }

    // 新しい選択範囲を設定
    const selection = window.getSelection();
    selection.removeAllRanges();
    const newRange = document.createRange();
    newRange.selectNodeContents(selectedSpan);
    selection.addRange(newRange);
    return;
  }

  // 上記以外（複雑なノード境界含む）はsurroundContentsで試み、失敗時はフォールバック
  const overrideSpan = document.createElement('span');
  if (!isDefault) {
    const unit = styleProperty === 'font-size' && typeof styleValue === 'number' ? 'pt' : '';
    try { overrideSpan.style.setProperty(styleProperty, (typeof styleValue === 'number' ? String(styleValue) : styleValue) + unit); } catch {}
  }
  try {
    range.surroundContents(overrideSpan);
    const selection = window.getSelection();
    selection.removeAllRanges();
    const newRange = document.createRange();
    newRange.selectNodeContents(overrideSpan);
    selection.addRange(newRange);
  } catch (e) {
    debugLog('handlePartialSpanSelection: surroundContents失敗、旧ロジックにフォールバック', e);
    // 最低限のフォールバック: 旧ロジックはBRを失う可能性があるため、より安全なapplyStylePreservingBreaksを使用
    applyStylePreservingBreaks(range, styleProperty, styleValue, isDefault);
  }
}

// 複数span要素選択時の統合処理（改行保持版）
function handleMultiSpanSelection(range, rangeInfo, styleProperty, styleValue, isDefault) {
  const spans = rangeInfo.multiSpans;
  
  debugLog('複数span統合:', spans.map(s => s.outerHTML));
  
  // 複数span選択時は個別にスタイルを適用する方式に変更
  // これにより改行や他の要素を保持
  spans.forEach(span => {
    updateSpanStyle(span, styleProperty, styleValue, isDefault);
  });
  
  debugLog('個別スタイル適用完了');
  
  // 選択範囲を維持
  const selection = window.getSelection();
  if (selection.rangeCount > 0) {
    const currentRange = selection.getRangeAt(0);
    selection.removeAllRanges();
    selection.addRange(currentRange);
  }
}

// 新しいspan要素を作成するヘルパー関数
function createNewSpanForSelection(range, selectedText, styleProperty, styleValue, isDefault) {
  if (isDefault) return; // デフォルト値の場合は新しいspan要素を作成しない

  const newSpan = document.createElement('span');

  // スタイルを設定
  if (styleProperty === 'font-family') {
    newSpan.style.fontFamily = styleValue;
  } else if (styleProperty === 'font-size') {
    newSpan.style.fontSize = styleValue + (typeof styleValue === 'number' ? 'pt' : '');
  } else {
    // その他のプロパティにも対応
    try { newSpan.style.setProperty(styleProperty, typeof styleValue === 'number' ? String(styleValue) : styleValue); } catch {}
  }

  // surroundContentsで包める場合はそれを使う（構造を崩さない）
  try {
    range.surroundContents(newSpan);
    debugLog('createNewSpanForSelection: surroundContentsで適用');
    // ラップ直後に内側の同一プロパティを削除し、親の指定を有効化
    if (styleProperty === 'font-size' || styleProperty === 'font-family') {
      removeStyleFromDescendants(newSpan, styleProperty);
    }
  } catch (e) {
    // 部分的なノード選択などでsurroundContentsが失敗した場合は、抽出→ラップでフォールバック
    debugLog('createNewSpanForSelection: surroundContents失敗、フォールバック適用', e);
    const frag = range.extractContents();
    // 元の選択テキストだけでなく、抽出フラグメント全体をspanに入れることでBR等を保持
    newSpan.appendChild(frag);
    range.insertNode(newSpan);
  }

  // 新しいspan要素を選択状態にする
  const selection = window.getSelection();
  selection.removeAllRanges();
  const newRange = document.createRange();
  newRange.selectNodeContents(newSpan);
  selection.addRange(newRange);
}

// BRやゼロ幅スペースを保持しつつ、選択範囲内のテキストノードにだけスタイルを適用
function applyStylePreservingBreaks(range, styleProperty, styleValue, isDefault) {
  if (isDefault) {
    // デフォルト化要求はここでは何もしない（削除系は既存の処理で対応）
    debugLog('applyStylePreservingBreaks: isDefault=true のためスキップ');
    return;
  }

  const unit = styleProperty === 'font-size' && typeof styleValue === 'number' ? 'pt' : '';
  const valueStr = typeof styleValue === 'number' ? String(styleValue) + unit : String(styleValue);

  try {
    // まず試しにsurroundContentsを使う（選択が素直な場合はこれで十分、内部のBRも保持）
    const testSpan = document.createElement('span');
    try { testSpan.style.setProperty(styleProperty, valueStr); } catch {}
    range.surroundContents(testSpan);
    // 子孫spanにある同一プロパティは削除（親の指定を効かせる）
    if (styleProperty === 'font-size' || styleProperty === 'font-family') {
      removeStyleFromDescendants(testSpan, styleProperty);
    }
    debugLog('applyStylePreservingBreaks: surroundContents成功');
    return;
  } catch (_) {
    // 続行してフォールバックへ
  }

  // 抽出してフラグメントを加工
  const fragment = range.extractContents();

  const processNode = (node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.nodeValue && node.nodeValue.length > 0) {
        const span = document.createElement('span');
        try { span.style.setProperty(styleProperty, valueStr); } catch {}
        span.textContent = node.nodeValue;
        return span;
      }
      return document.createTextNode('');
    }
    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node;
      // BRはそのまま返す
      if (el.tagName === 'BR') return el;

      // 既存のSPANはスタイルを追記（上書き）しつつ子も処理
      if (el.tagName === 'SPAN') {
        // 子を先に処理してから自身のスタイルを更新
        const newSpan = document.createElement('span');
        // 既存スタイルを引き継ぎつつ、同一プロパティは削除（親で統一するため）
        const current = el.getAttribute('style') || '';
        if (current) {
          const map = parseStyleString(current);
          map.delete(styleProperty);
          const cleaned = Array.from(map.entries()).map(([k,v]) => `${k}: ${v}`).join('; ');
          if (cleaned) newSpan.setAttribute('style', cleaned);
        }
        try { newSpan.style.setProperty(styleProperty, valueStr); } catch {}
        // 子を再構築
        while (el.firstChild) {
          const child = el.firstChild;
          el.removeChild(child);
          newSpan.appendChild(processNode(child));
        }
        return newSpan;
      }

      // その他の要素は中の子に対してのみ適用
      const wrapper = document.createElement(el.tagName);
      // 属性コピー
      for (const attr of el.attributes) {
        wrapper.setAttribute(attr.name, attr.value);
      }
      while (el.firstChild) {
        const child = el.firstChild;
        el.removeChild(child);
        wrapper.appendChild(processNode(child));
      }
      return wrapper;
    }
    // それ以外のノードはそのまま
    return node;
  };

  // fragment直下の子を処理して新しいフラグメントに詰める
  const newFragment = document.createDocumentFragment();
  Array.from(fragment.childNodes).forEach(child => {
    newFragment.appendChild(processNode(child));
  });

  // 加工済みのフラグメントを挿入
  range.insertNode(newFragment);

  // 新しい選択範囲（適用部）をざっくり再選択（厳密なカーソル維持は後続操作で上書きされる）
  const selection = window.getSelection();
  selection.removeAllRanges();
  const newRange = document.createRange();
  // 直前に挿入した箇所を再選択するため、rangeのstartContainer付近を頼りにエディタ側での後処理に任せる
  // ここでは安全側で親ノード全体を再選択
  const anchor = range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
    ? range.commonAncestorContainer
    : range.commonAncestorContainer.parentNode;
  newRange.selectNodeContents(anchor);
  selection.addRange(newRange);
  debugLog('applyStylePreservingBreaks: フォールバック適用完了');
}

// フォントファミリーを選択範囲に適用（統合された関数を使用）
function applyFontFamilyToSelection(fontFamily, editor) {
  const isDefault = !fontFamily || fontFamily === '';
  
  if (isDefault) {
    // デフォルトフォントの場合は特別な処理
    applyDefaultFontToSelection(editor);
  } else {
    applyStyleToSelection('font-family', fontFamily, editor, false);
  }
}

// デフォルトフォントに戻す専用関数
function applyDefaultFontToSelection(editor) {
  const selection = window.getSelection();
  if (!selection.rangeCount || selection.isCollapsed) {
    return;
  }
  
  try {
    const range = selection.getRangeAt(0);
    const selectedText = range.toString();
    
    if (!selectedText) {
      return;
    }
    
    console.log(`デフォルトフォントに戻す: "${selectedText}"`);
    
    // 選択範囲に含まれるspan要素を直接検索
    const commonAncestor = range.commonAncestorContainer;
    let targetSpan = null;
    
    // 選択範囲が単一のspan要素内にある場合を検出
    if (commonAncestor.nodeType === Node.TEXT_NODE) {
      const parentElement = commonAncestor.parentElement;
      if (parentElement && parentElement.tagName === 'SPAN' && 
          parentElement.style.fontFamily) {
        targetSpan = parentElement;
      }
    } else if (commonAncestor.tagName === 'SPAN' && 
               commonAncestor.style.fontFamily) {
      targetSpan = commonAncestor;
    }
    
    if (targetSpan) {
      debugLog('対象span要素を発見:', targetSpan.outerHTML);
      
      // font-family以外のスタイルを保持
      const currentStyle = targetSpan.getAttribute('style') || '';
      const cleanStyle = currentStyle.split(';')
        .filter(rule => {
          const property = rule.trim().split(':')[0].trim().toLowerCase();
          return property && property !== 'font-family';
        })
        .join('; ');
      
      if (cleanStyle.trim()) {
        // 他のスタイルがある場合は、font-familyのみ削除
        targetSpan.setAttribute('style', cleanStyle);
        debugLog('font-familyスタイルを削除:', cleanStyle);
      } else {
        // スタイルがfont-familyのみの場合は、span要素を完全に削除
        const parent = targetSpan.parentNode;
        const textContent = targetSpan.textContent;
        
        if (parent) {
          const textNode = document.createTextNode(textContent);
          parent.replaceChild(textNode, targetSpan);
          debugLog('span要素を削除し、テキストノードに置換');
          
          // 新しいテキストノードを選択
          range.selectNode(textNode);
          selection.removeAllRanges();
          selection.addRange(range);
        }
      }
    } else {
      debugLog('対象のspan要素が見つからないため、通常処理を実行');
      // フォールバック: 元のapplyStyleToSelection関数を使用
      applyStyleToSelection('font-family', '', editor, true);
    }
    
    debugLog('デフォルトフォントに戻しました');
    debugLog('処理後のエディタHTML:', editor.innerHTML);
    
    // エディタ全体の空のspan要素を掃除
    cleanupEmptySpans(editor);
    
  } catch (error) {
    console.warn('デフォルトフォント適用エラー:', error);
    // フォールバック: シンプルな方法で処理
    applyStyleToSelection('font-family', '', editor, true);
  }
  
  // エディタにフォーカスを戻す
  editor.focus();
}

// 空のspan要素やネストしたspan要素を掃除（改良版）
function cleanupEmptySpans(editor) {
  console.log('cleanupEmptySpans開始');
  let removedCount = 0;
  
  try {
    // 最大5回実行して、深いネストも処理
    for (let round = 0; round < 5; round++) {
      const spans = Array.from(editor.querySelectorAll('span'));
      let currentRoundRemoved = 0;
      
      // 後ろから処理してインデックスのずれを防ぐ
      for (let i = spans.length - 1; i >= 0; i--) {
        const span = spans[i];
        if (!span.parentNode) continue; // 既に削除されている場合はスキップ
        
        const style = span.getAttribute('style') || '';
        const trimmedStyle = style.trim();
        const hasText = span.textContent.trim() !== '';
        
        // スタイルが空の場合
        if (!trimmedStyle) {
          if (hasText) {
            // テキストがある場合は内容を親に移動
            const parent = span.parentNode;
            while (span.firstChild) {
              parent.insertBefore(span.firstChild, span);
            }
            parent.removeChild(span);
            currentRoundRemoved++;
            console.log('空スタイルのspan削除:', span.textContent);
          } else {
            // テキストもない場合は単純に削除
            span.parentNode.removeChild(span);
            currentRoundRemoved++;
            console.log('空のspan削除');
          }
        } else {
          // 重複したネストの処理
          const parent = span.parentNode;
          if (parent && parent.tagName === 'SPAN') {
            const parentStyle = parent.getAttribute('style') || '';
            
            // 親と子で同じスタイルプロパティがある場合
            const childStyles = parseStyles(trimmedStyle);
            const parentStyles = parseStyles(parentStyle);
            
            let hasConflict = false;
            for (const [property, value] of childStyles) {
              if (parentStyles.has(property)) {
                hasConflict = true;
                break;
              }
            }
            
            if (hasConflict) {
              // 子要素のスタイルを優先し、親の重複スタイルを削除
              const mergedStyles = new Map([...parentStyles, ...childStyles]);
              const mergedStyleString = Array.from(mergedStyles.entries())
                .map(([prop, val]) => `${prop}: ${val}`)
                .join('; ');
              
              // 新しいspan要素を作成
              const newSpan = document.createElement('span');
              newSpan.setAttribute('style', mergedStyleString);
              
              // 子要素の内容をコピー
              while (span.firstChild) {
                newSpan.appendChild(span.firstChild);
              }
              
              // 親要素を新しいspan要素で置換
              parent.parentNode.replaceChild(newSpan, parent);
              currentRoundRemoved++;
              console.log('ネストspan統合:', mergedStyleString);
            }
          }
        }
      }
      
      removedCount += currentRoundRemoved;
      
      // この回で削除がなければ終了
      if (currentRoundRemoved === 0) {
        break;
      }
    }
    
    console.log(`cleanupEmptySpans完了: ${removedCount}個のspanを処理`);
  } catch (error) {
    console.warn('span要素掃除エラー:', error);
  }
}

// CSSスタイル文字列をMapに変換するヘルパー関数
function parseStyles(styleString) {
  const styles = new Map();
  if (styleString) {
    styleString.split(';').forEach(rule => {
      const [property, value] = rule.split(':').map(s => s.trim());
      if (property && value) {
        styles.set(property.toLowerCase(), value);
      }
    });
  }
  return styles;
}

// フォントサイズを選択範囲に適用（統合された関数を使用）
function applyFontSizeToSelection(fontSize, editor) {
  applyStyleToSelection('font-size', fontSize, editor, false);
}

// 選択範囲から既存のスタイルを収集する関数（指定されたプロパティを除外）
// 包括的なspan要素クリーンアップ関数（ネスト解決・改良版）
function cleanupEmptySpans(editor) {
  if (!editor) return;
  
  let deletedCount = 0;
  let processedInRound = 0;
  let roundCount = 0;
  const maxRounds = 10; // 無限ループ防止
  
  debugLog('cleanupEmptySpans: クリーンアップ開始');
  debugLog('処理前HTML:', editor.innerHTML);
  
  // 複数回実行して深いネストを解決
  do {
    processedInRound = 0;
    roundCount++;
    debugLog(`cleanupEmptySpans: ラウンド ${roundCount}`);
    
    // 1. 空のspan要素を削除
    const emptySpans = editor.querySelectorAll('span:empty');
    processedInRound += emptySpans.length;
    emptySpans.forEach(span => {
      debugLog('空のspan要素を削除:', span.outerHTML);
      span.remove();
      deletedCount++;
    });
    
    // 2. スタイルのないspan要素を削除
    const unstyledSpans = editor.querySelectorAll('span:not([style]), span[style=""]');
    unstyledSpans.forEach(span => {
      if (span.children.length === 0) { // テキストのみの場合
        debugLog('スタイルなしspan要素をアンラップ:', span.textContent);
        const parent = span.parentNode;
        while (span.firstChild) {
          parent.insertBefore(span.firstChild, span);
        }
        span.remove();
        processedInRound++;
        deletedCount++;
      }
    });
    
    // 3. ネストしたspan要素を統合
    const nestedSpans = editor.querySelectorAll('span span');
    nestedSpans.forEach(innerSpan => {
      const outerSpan = innerSpan.parentElement;
      if (outerSpan && outerSpan.tagName === 'SPAN') {
        debugLog('ネストspan検出:', {
          outer: outerSpan.outerHTML,
          inner: innerSpan.outerHTML
        });
        
        // 子要素が1つのspan要素のみの場合
        if (outerSpan.children.length === 1 && outerSpan.firstElementChild === innerSpan) {
          // スタイルを統合
          const mergedStyle = mergeSpanStyles(outerSpan, innerSpan);
          debugLog('統合されたスタイル:', mergedStyle);
          
          // 外側のspan要素のスタイルを更新
          if (mergedStyle) {
            outerSpan.setAttribute('style', mergedStyle);
          } else {
            outerSpan.removeAttribute('style');
          }
          
          // 内側のspan要素の内容を外側に移動
          while (innerSpan.firstChild) {
            outerSpan.insertBefore(innerSpan.firstChild, innerSpan);
          }
          innerSpan.remove();
          processedInRound++;
          deletedCount++;
          debugLog('ネストspan統合完了:', outerSpan.outerHTML);
        }
      }
    });
    
    // 4. 同じスタイルの隣接span要素を統合
    const spans = Array.from(editor.querySelectorAll('span[style]'));
    for (let i = 0; i < spans.length - 1; i++) {
      const currentSpan = spans[i];
      const nextSpan = spans[i + 1];
      
      if (nextSpan && currentSpan.nextSibling === nextSpan) {
        const currentStyle = normalizeStyle(currentSpan.getAttribute('style') || '');
        const nextStyle = normalizeStyle(nextSpan.getAttribute('style') || '');
        
        if (currentStyle === nextStyle) {
          debugLog('隣接span要素を統合:', currentSpan.textContent, '+', nextSpan.textContent);
          currentSpan.textContent += nextSpan.textContent;
          nextSpan.remove();
          processedInRound++;
          deletedCount++;
        }
      }
    }
    
    debugLog(`ラウンド ${roundCount} 完了: ${processedInRound}個処理`);
    
  } while (processedInRound > 0 && roundCount < maxRounds);
  
  debugLog(`cleanupEmptySpans完了: 合計${deletedCount}個のspan要素を処理/削除`);
  debugLog('処理後HTML:', editor.innerHTML);
  
  return deletedCount;
}

// span要素のスタイルを統合するヘルパー関数
function mergeSpanStyles(outerSpan, innerSpan) {
  const outerStyle = parseStyleString(outerSpan.getAttribute('style') || '');
  const innerStyle = parseStyleString(innerSpan.getAttribute('style') || '');
  
  // 内側のスタイルが優先される
  const mergedStyle = new Map([...outerStyle, ...innerStyle]);
  
  if (mergedStyle.size === 0) {
    return '';
  }
  
  return Array.from(mergedStyle.entries())
    .map(([prop, val]) => `${prop}: ${val}`)
    .join('; ');
}

// スタイル文字列を正規化するヘルパー関数
function normalizeStyle(styleString) {
  if (!styleString) return '';
  
  const styleMap = parseStyleString(styleString);
  return Array.from(styleMap.entries())
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([prop, val]) => `${prop}: ${val}`)
    .join('; ');
}

// スタイル文字列をMapに変換するヘルパー関数
function parseStyleString(styleString) {
  const styleMap = new Map();
  
  if (!styleString) return styleMap;
  
  styleString.split(';').forEach(rule => {
    const [property, value] = rule.split(':').map(s => s.trim());
    if (property && value) {
      styleMap.set(property.toLowerCase(), value);
    }
  });
  
  return styleMap;
}

// 指定した要素配下の全てのspanから、指定スタイルプロパティを取り除く
function removeStyleFromDescendants(rootEl, styleProperty) {
  if (!rootEl) return;
  try {
    const spans = rootEl.querySelectorAll('span[style]');
    spans.forEach(s => {
      const styleMap = parseStyleString(s.getAttribute('style') || '');
      if (styleMap.has(styleProperty)) {
        styleMap.delete(styleProperty);
        const newStyle = Array.from(styleMap.entries()).map(([k,v]) => `${k}: ${v}`).join('; ');
        if (newStyle) {
          s.setAttribute('style', newStyle);
        } else {
          s.removeAttribute('style');
        }
      }
    });
  } catch (e) {
    debugLog('removeStyleFromDescendantsエラー', e);
  }
}

// フォントサイズを選択範囲に適用（統合された関数を使用）
function applyFontSizeToSelection(fontSize, editor) {
  applyStyleToSelection('font-size', fontSize, editor, false);
}

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
function showSuccessMessage(message) {
  // 既存の通知があれば削除
  const existingNotification = document.querySelector('.success-notification');
  if (existingNotification) {
    existingNotification.remove();
  }
  
  const notification = document.createElement('div');
  notification.className = 'success-notification';
  notification.textContent = message;
  notification.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    background: #4CAF50;
    color: white;
    padding: 12px 24px;
    border-radius: 4px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    z-index: 10001;
    font-family: sans-serif;
    font-size: 14px;
    transition: opacity 0.3s ease;
  `;
  
  document.body.appendChild(notification);
  
  setTimeout(() => {
    notification.style.opacity = '0';
    setTimeout(() => {
      if (notification.parentNode) {
        notification.parentNode.removeChild(notification);
      }
    }, 300);
  }, 3000);
}

// ローディング表示ヘルパー
function showFontUploadProgress(show) {
  let progressDiv = document.getElementById('font-upload-progress');
  
  if (show) {
    if (!progressDiv) {
      progressDiv = document.createElement('div');
      progressDiv.id = 'font-upload-progress';
      progressDiv.style.cssText = `
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: rgba(0,0,0,0.8);
        color: white;
        padding: 20px;
        border-radius: 8px;
        z-index: 10002;
        font-family: sans-serif;
        text-align: center;
      `;
      progressDiv.innerHTML = `
        <div style="margin-bottom: 10px;">フォントをアップロード中...</div>
        <div style="width: 200px; height: 4px; background: #333; border-radius: 2px;">
          <div style="width: 100%; height: 100%; background: #4CAF50; border-radius: 2px; animation: pulse 1s infinite;"></div>
        </div>
      `;
      document.body.appendChild(progressDiv);
    }
  } else {
    if (progressDiv && progressDiv.parentNode) {
      progressDiv.parentNode.removeChild(progressDiv);
    }
  }
}

// CSS用のpulseアニメーションを追加
if (!document.getElementById('font-animations')) {
  const style = document.createElement('style');
  style.id = 'font-animations';
  style.textContent = `
    @keyframes pulse {
      0% { opacity: 1; }
      50% { opacity: 0.5; }
      100% { opacity: 1; }
    }
  `;
  document.head.appendChild(style);
}
