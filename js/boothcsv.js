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

// リアルタイム更新制御フラグ（customLabels.js に移動）
// let isEditingCustomLabel, pendingUpdateTimer は customLabels.js で定義

// デバッグログ用関数
function debugLog(...args) {
  if (DEBUG_MODE) {
    console.log(...args);
  }
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

// CustomLabelCalculator は customLabels.js に移動（エイリアスは後段で設定）

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

// ===== 再導入: カスタムラベル分離時に除外された補助関数 =====
// 注文画像表示行の表示/非表示を切り替え
function toggleOrderImageRow(enabled) {
  const row = document.getElementById('orderImageRow');
  if (row) row.style.display = enabled ? 'table-row' : 'none';
}

// ボタン有効状態等を更新（customLabels.js から参照される）
async function updateButtonStates() {
  const printBtnHeader = document.getElementById('print-btn'); // 固定ヘッダー
  // 生成済みシート数で印刷可否を判定
  const hasContent = document.querySelectorAll('.sheet').length > 0 || document.querySelectorAll('.label44').length > 0;
  if (printBtnHeader) printBtnHeader.disabled = !hasContent;
  // カスタムラベルサマリーも更新（存在する場合のみ）
  if (typeof updateCustomLabelsSummary === 'function') {
    try { await updateCustomLabelsSummary(); } catch (e) { console.warn('updateCustomLabelsSummary失敗:', e); }
  }
}
// ============================================================

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
        await StorageManager.clearQRImages();
        alert('全てのQR画像をクリアしました');
        location.reload();
      }
    };
  }

  // 全ての注文画像をクリアするボタンのイベントリスナーを追加
  const clearAllOrderImagesButton = document.getElementById('clearAllOrderImagesButton');
  if (clearAllOrderImagesButton) {
    clearAllOrderImagesButton.onclick = async () => {
      if (confirm('本当に全ての注文画像（グローバル画像と個別画像）をクリアしますか？')) {
        await StorageManager.clearOrderImages();
        alert('全ての注文画像をクリアしました');
        location.reload();
      }
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
// 注文番号のセクションにスクロール（固定ヘッダー分オフセット考慮）
// スクロール位置を保持・復元する（再描画でDOMが差し替わってもUXを維持）
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
    this.textContent = null;
    this.innerHTML = "<p>Paste QR image here!</p>";
  });
}

function createDropzone(div){
  const divDrop = createDiv('dropzone', 'Paste QR image here!');
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
      dropZone.innerHTML = `<p style="margin: 5px; font-size: 12px; color: #666;">${defaultMessage}</p>`;
    } else {
      const defaultContentElement = document.getElementById('dropZoneDefaultContent');
      const defaultContent = defaultContentElement ? defaultContentElement.innerHTML : defaultMessage;
      dropZone.innerHTML = defaultContent;
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
        dropZone.innerHTML = `<p style="margin: 5px; font-size: 12px; color: #666;">${defaultMessage}</p>`;
        await updateOrderImageDisplay(null);
      } else {
        const defaultContentElement = document.getElementById('dropZoneDefaultContent');
        const defaultContent = defaultContentElement ? defaultContentElement.innerHTML : defaultMessage;
        dropZone.innerHTML = defaultContent;
        // グローバル画像がクリアされた場合も全ての注文明細を更新
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
// カスタムラベル関連のクラス/関数は customLabels.js に移動しました
// ここでは互換のため主要関数を参照するエイリアスを設定します。
window.toggleCustomLabelRow = window.CustomLabels?.toggleCustomLabelRow;
window.initializeCustomLabels = window.CustomLabels?.initializeCustomLabels;
window.addCustomLabelItem = window.CustomLabels?.addCustomLabelItem;
window.removeCustomLabelItem = window.CustomLabels?.removeCustomLabelItem;
window.getCustomLabelsFromUI = window.CustomLabels?.getCustomLabelsFromUI;
window.saveCustomLabels = window.CustomLabels?.saveCustomLabels;
window.updateCustomLabelsSummary = window.CustomLabels?.updateCustomLabelsSummary;
window.adjustCustomLabelsForTotal = window.CustomLabels?.adjustCustomLabelsForTotal;
window.setupCustomLabelEvents = window.CustomLabels?.setupCustomLabelEvents;
window.clearAllCustomLabels = window.CustomLabels?.clearAllCustomLabels;
window.hasCustomLabelsWithContent = window.CustomLabels?.hasCustomLabelsWithContent;
window.hasEmptyEnabledCustomLabels = window.CustomLabels?.hasEmptyEnabledCustomLabels;
window.removeEmptyCustomLabels = window.CustomLabels?.removeEmptyCustomLabels;
window.highlightEmptyCustomLabels = window.CustomLabels?.highlightEmptyCustomLabels;
window.clearCustomLabelHighlights = window.CustomLabels?.clearCustomLabelHighlights;
window.reindexCustomLabelItems = window.CustomLabels?.reindexCustomLabelItems;
window.scheduleDelayedPreviewUpdate = window.CustomLabels?.scheduleDelayedPreviewUpdate;
