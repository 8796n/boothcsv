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

// IndexedDBを使用したフォント管理クラス
class FontDatabase {
  constructor() {
    this.dbName = 'BoothCSVFonts';
    this.version = 1;
    this.storeName = 'fonts';
    this.db = null;
  }

  // データベースを初期化
  async init() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, this.version);
      
      request.onerror = () => {
        console.error('IndexedDB初期化エラー:', request.error);
        reject(request.error);
      };
      
      request.onsuccess = () => {
        this.db = request.result;
        console.log('IndexedDB初期化完了');
        resolve(this.db);
      };
      
      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        
        // フォントストアを作成
        if (!db.objectStoreNames.contains(this.storeName)) {
          const store = db.createObjectStore(this.storeName, { keyPath: 'name' });
          store.createIndex('createdAt', 'createdAt', { unique: false });
          store.createIndex('size', 'size', { unique: false });
          console.log('フォントストア作成完了');
        }
      };
    });
  }

  // フォントを保存
  async saveFont(fontName, fontData, metadata = {}) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.storeName], 'readwrite');
      const store = transaction.objectStore(this.storeName);
      
      const fontObject = {
        name: fontName,
        data: fontData, // ArrayBufferを直接保存
        metadata: {
          type: metadata.type || 'font/ttf',
          originalName: metadata.originalName || fontName,
          size: fontData.byteLength || fontData.length,
          createdAt: Date.now()
        }
      };
      
      const request = store.put(fontObject);
      
      request.onsuccess = () => {
        console.log(`フォント保存完了: ${fontName}`);
        resolve(fontObject);
      };
      
      request.onerror = () => {
        console.error('フォント保存エラー:', request.error);
        reject(request.error);
      };
    });
  }

  // フォントを取得
  async getFont(fontName) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.storeName], 'readonly');
      const store = transaction.objectStore(this.storeName);
      const request = store.get(fontName);
      
      request.onsuccess = () => {
        resolve(request.result);
      };
      
      request.onerror = () => {
        console.error('フォント取得エラー:', request.error);
        reject(request.error);
      };
    });
  }

  // すべてのフォントを取得（オブジェクト形式で返す）
  async getAllFonts() {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.storeName], 'readonly');
      const store = transaction.objectStore(this.storeName);
      const request = store.getAll();
      
      request.onsuccess = () => {
        const fonts = request.result;
        const fontMap = {};
        
        // 配列をオブジェクトマップに変換
        fonts.forEach(font => {
          if (font && font.name) {
            fontMap[font.name] = {
              data: font.data,
              metadata: font.metadata || {}
            };
          }
        });
        
        resolve(fontMap);
      };
      
      request.onerror = () => {
        console.error('フォント一覧取得エラー:', request.error);
        reject(request.error);
      };
    });
  }

  // フォントを削除
  async deleteFont(fontName) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.storeName], 'readwrite');
      const store = transaction.objectStore(this.storeName);
      const request = store.delete(fontName);
      
      request.onsuccess = () => {
        console.log(`フォント削除完了: ${fontName}`);
        resolve();
      };
      
      request.onerror = () => {
        console.error('フォント削除エラー:', request.error);
        reject(request.error);
      };
    });
  }

  // すべてのフォントを削除
  async clearAllFonts() {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.storeName], 'readwrite');
      const store = transaction.objectStore(this.storeName);
      const request = store.clear();
      
      request.onsuccess = () => {
        console.log('全フォント削除完了');
        resolve();
      };
      
      request.onerror = () => {
        console.error('全フォント削除エラー:', request.error);
        reject(request.error);
      };
    });
  }

  // ストレージ使用量を取得
  async getStorageInfo() {
    const fonts = await this.getAllFonts();
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

// グローバルなフォントデータベースインスタンス
let fontDB = null;

// フォントデータベースを初期化
async function initializeFontDatabase() {
  try {
    fontDB = new FontDatabase();
    await fontDB.init();
    console.log('フォントデータベース初期化完了');
    return fontDB;
  } catch (error) {
    console.error('フォントデータベース初期化失敗:', error);
    // フォールバック: localStorageを使用
    alert('IndexedDBの初期化に失敗しました。localStorageを使用します。\n大容量のフォントファイルは制限される場合があります。');
    return null;
  }
}

// デバッグ機能：localStorageからIndexedDBへのデータ移行
class StorageManager {
  static KEYS = {
    ORDER_IMAGE_PREFIX: 'orderImage_',
    GLOBAL_ORDER_IMAGE: 'orderImage',
    LABEL_SETTING: 'labelyn',
    LABEL_SKIP: 'labelskip',
    SORT_BY_PAYMENT: 'sortByPaymentDate',
    CUSTOM_LABEL_ENABLE: 'customLabelEnable',
    CUSTOM_LABEL_TEXT: 'customLabelText',
    CUSTOM_LABEL_COUNT: 'customLabelCount',
    CUSTOM_LABELS: 'customLabels', // 複数のカスタムラベルを保存
    ORDER_IMAGE_ENABLE: 'orderImageEnable' // 注文画像表示の有効/無効
  };

  // 設定値の取得
  static getSettings() {
    return {
      labelyn: this.get(this.KEYS.LABEL_SETTING, 'true') !== 'false',
      labelskip: parseInt(this.get(this.KEYS.LABEL_SKIP, '0'), 10),
      sortByPaymentDate: this.get(this.KEYS.SORT_BY_PAYMENT) === 'true',
      customLabelEnable: this.get(this.KEYS.CUSTOM_LABEL_ENABLE) === 'true',
      customLabelText: this.get(this.KEYS.CUSTOM_LABEL_TEXT, ''),
      customLabelCount: parseInt(this.get(this.KEYS.CUSTOM_LABEL_COUNT, '1'), 10),
      customLabels: this.getCustomLabels(),
      orderImageEnable: this.get(this.KEYS.ORDER_IMAGE_ENABLE) === 'true'
    };
  }

  // 複数のカスタムラベルを取得
  static getCustomLabels() {
    const labelsData = this.get(this.KEYS.CUSTOM_LABELS);
    if (!labelsData) return [];
    
    try {
      return JSON.parse(labelsData);
    } catch (e) {
      console.warn('Custom labels data parsing failed:', e);
      return [];
    }
  }

  // 複数のカスタムラベルを保存
  static setCustomLabels(labels) {
    this.set(this.KEYS.CUSTOM_LABELS, JSON.stringify(labels));
  }

  // 設定値の保存
  static saveSettings(settings) {
    Object.entries(settings).forEach(([key, value]) => {
      if (this.KEYS[key.toUpperCase()]) {
        this.set(this.KEYS[key.toUpperCase()], String(value));
      }
    });
  }

  // 注文画像の取得
  static getOrderImage(orderNumber = null) {
    const key = orderNumber ? 
      `${this.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : 
      this.KEYS.GLOBAL_ORDER_IMAGE;
    return this.get(key);
  }

  // 注文画像の保存
  static setOrderImage(imageData, orderNumber = null) {
    const key = orderNumber ? 
      `${this.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : 
      this.KEYS.GLOBAL_ORDER_IMAGE;
    this.set(key, imageData);
  }

  // 注文画像の削除
  static removeOrderImage(orderNumber = null) {
    const key = orderNumber ? 
      `${this.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : 
      this.KEYS.GLOBAL_ORDER_IMAGE;
    this.remove(key);
  }

  // QR画像の一括削除
  static clearQRImages() {
    Object.keys(localStorage).forEach(key => {
      const value = localStorage.getItem(key);
      if (value?.includes('qrimage')) {
        this.remove(key);
      }
    });
  }

  // 注文画像の一括削除
  static clearOrderImages() {
    Object.keys(localStorage).forEach(key => {
      if (key === this.KEYS.GLOBAL_ORDER_IMAGE || 
          key.startsWith(this.KEYS.ORDER_IMAGE_PREFIX)) {
        this.remove(key);
      }
    });
  }

  // QRコードデータの取得
  static getQRData(orderNumber) {
    const data = this.get(orderNumber);
    if (!data) return null;
    
    try {
      return JSON.parse(data);
    } catch (e) {
      console.warn('QR data parsing failed:', e);
      return null;
    }
  }

  // QRコードデータの保存
  static setQRData(orderNumber, qrData) {
    this.set(orderNumber, JSON.stringify(qrData));
  }

  // QRコードの重複チェック
  static checkQRDuplicate(qrContent, currentOrderNumber) {
    const qrHash = this.generateQRHash(qrContent);
    const duplicates = [];
    
    Object.keys(localStorage).forEach(key => {
      if (key !== currentOrderNumber) {
        const data = this.get(key);
        if (data) {
          try {
            const parsedData = JSON.parse(data);
            if (parsedData.qrhash === qrHash) {
              duplicates.push(key);
            }
          } catch (e) {
            // 無視
          }
        }
      }
    });
    
    return duplicates;
  }

  // QRコンテンツのハッシュ化
  static generateQRHash(qrContent) {
    // シンプルなハッシュ関数（本格的な場合はCrypto APIを使用）
    let hash = 0;
    if (qrContent.length === 0) return hash;
    for (let i = 0; i < qrContent.length; i++) {
      const char = qrContent.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // 32bit整数に変換
    }
    return hash.toString();
  }

  // 基本的なlocalStorage操作
  static get(key, defaultValue = null) {
    const value = localStorage.getItem(key);
    return value !== null ? value : defaultValue;
  }

  static set(key, value) {
    localStorage.setItem(key, value);
  }

  static remove(key) {
    localStorage.removeItem(key);
  }
}

window.addEventListener("load", function(){
  const settings = StorageManager.getSettings();
  
  document.getElementById("labelyn").checked = settings.labelyn;
  document.getElementById("labelskipnum").value = settings.labelskip;
  document.getElementById("sortByPaymentDate").checked = settings.sortByPaymentDate;
  document.getElementById("customLabelEnable").checked = settings.customLabelEnable;
  document.getElementById("orderImageEnable").checked = settings.orderImageEnable;

  // カスタムラベル行の表示/非表示
  toggleCustomLabelRow(settings.customLabelEnable);

  // 注文画像行の表示/非表示
  toggleOrderImageRow(settings.orderImageEnable);

  // 複数のカスタムラベルを初期化
  initializeCustomLabels(settings.customLabels);

   // 画像ドロップゾーンの初期化
  const imageDropZoneElement = document.getElementById('imageDropZone');
  const imageDropZone = createOrderImageDropZone();
  imageDropZoneElement.appendChild(imageDropZone.element);
  window.orderImageDropZone = imageDropZone;

  // フォントドロップゾーンの初期化
  initializeFontDropZone();
  
  // カスタムフォントのCSS読み込み
  loadCustomFontsCSS();

  // 全ての画像をクリアするボタンのイベントリスナーを追加
  const clearAllButton = document.getElementById('clearAllButton');
  clearAllButton.onclick = () => {
    if (confirm('本当に全てのQR画像をクリアしますか？')) {
      StorageManager.clearQRImages();
      alert('全てのQR画像をクリアしました');
      location.reload();
    }
  };

  // 全ての注文画像をクリアするボタンのイベントリスナーを追加
  const clearAllOrderImagesButton = document.getElementById('clearAllOrderImagesButton');
  clearAllOrderImagesButton.onclick = () => {
    if (confirm('本当に全ての注文画像（グローバル画像と個別画像）をクリアしますか？')) {
      StorageManager.clearOrderImages();
      alert('全ての注文画像をクリアしました');
      location.reload();
    }
  };

  // 全てのカスタムフォントをクリアするボタンのイベントリスナーを追加
  const clearAllFontsButton = document.getElementById('clearAllFontsButton');
  if (clearAllFontsButton) {
    clearAllFontsButton.onclick = async () => {
      if (confirm('本当に全てのカスタムフォントをクリアしますか？')) {
        try {
          await fontDB.clearAllFonts();
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

   // チェックボックスの状態が変更されたときにlocalStorageに保存
   document.getElementById("sortByPaymentDate").addEventListener("change", function() {
     StorageManager.set(StorageManager.KEYS.SORT_BY_PAYMENT, this.checked);
   });

   // 注文画像表示機能のイベントリスナー
   document.getElementById("orderImageEnable").addEventListener("change", function() {
     StorageManager.set(StorageManager.KEYS.ORDER_IMAGE_ENABLE, this.checked);
     toggleOrderImageRow(this.checked);
     
     // 画像表示をリアルタイムで更新
     updateAllOrderImagesVisibility(this.checked);
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

}, false);

function clickstart() {
  // カスタムラベルのバリデーション
  if (!validateAndPromptCustomLabels()) {
    return; // バリデーション失敗時は処理を中断
  }
  
  clearPreviousResults();
  const config = getConfigFromUI();
  
  Papa.parse(config.file, {
    header: true,
    skipEmptyLines: true,
    complete: async function(results) {
      await processCSVResults(results, config);
    }
  });
}

function executeCustomLabelsOnly() {
  // カスタムラベルのバリデーション
  if (!validateAndPromptCustomLabels()) {
    return; // バリデーション失敗時は処理を中断
  }
  
  clearPreviousResults();
  const config = getConfigFromUI();
  
  // カスタムラベルが有効でない場合は警告
  if (!config.customLabelEnable) {
    alert('カスタムラベル機能を有効にしてください。');
    return;
  }
  
  // カスタムラベルの内容が空の場合は警告
  if (!config.customLabels || config.customLabels.length === 0) {
    alert('印刷する文字列を入力してください。');
    return;
  }
  
  // 有効なカスタムラベルがあるかチェック
  const validLabels = config.customLabels.filter(label => label.text.trim() !== '');
  if (validLabels.length === 0) {
    alert('印刷する文字列を入力してください。');
    return;
  }
  
  // カスタムラベルのみを処理
  processCustomLabelsOnly(config);
}

function clearPreviousResults() {
  for (let sheet of document.querySelectorAll('section')) {
    sheet.parentNode.removeChild(sheet);
  }
}

function getConfigFromUI() {
  const file = document.getElementById("file").files[0];
  const labelyn = document.getElementById("labelyn").checked;
  const labelskip = document.getElementById("labelskipnum").value;
  const sortByPaymentDate = document.getElementById("sortByPaymentDate").checked;
  const customLabelEnable = document.getElementById("customLabelEnable").checked;
  
  // 複数のカスタムラベルを取得（有効なもののみ）
  const allCustomLabels = getCustomLabelsFromUI();
  const customLabels = customLabelEnable ? allCustomLabels.filter(label => label.enabled) : [];
  
  StorageManager.set(StorageManager.KEYS.LABEL_SETTING, labelyn);
  StorageManager.set(StorageManager.KEYS.LABEL_SKIP, labelskip);
  StorageManager.set(StorageManager.KEYS.CUSTOM_LABEL_ENABLE, customLabelEnable);
  StorageManager.setCustomLabels(allCustomLabels); // 全てのラベルを保存（有効/無効問わず）
  
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
  // CSV行数を取得
  const csvRowCount = results.data.length;
  
  // 複数カスタムラベルの総面数を計算
  const totalCustomLabelCount = config.customLabels.reduce((sum, label) => sum + label.count, 0);
  
  // 複数シート対応：1シートの制限を撤廃
  // CSVデータとカスタムラベルの合計で必要なシート数を計算
  const skipCount = parseInt(config.labelskip, 10) || 0;
  const totalLabelsNeeded = skipCount + csvRowCount + totalCustomLabelCount;
  const requiredSheets = Math.ceil(totalLabelsNeeded / CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET);

  // データの並び替え
  if (config.sortByPaymentDate) {
    results.data.sort((a, b) => {
      const timeA = a[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      const timeB = b[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      return timeA.localeCompare(timeB);
    });
  }

  // 注文明細の生成
  generateOrderDetails(results.data, config.labelarr);
  
  // ラベル生成（注文分＋カスタムラベル）- 複数シート対応
  if (config.labelyn) {
    let totalLabelArray = [...config.labelarr];
    
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
      generateLabels(totalLabelArray);
    }
  }
  
  // 印刷枚数の表示（複数シート対応）
  showCSVWithCustomLabelPrintSummary(csvRowCount, totalCustomLabelCount, skipCount, requiredSheets);
  
  // CSV処理完了後のカスタムラベルサマリー更新（複数シート対応）
  await updateCustomLabelsSummary();
  
  // ボタンの状態を更新
  updateButtonStates();
}

function processCustomLabelsOnly(config) {
  // 複数カスタムラベルの総面数を計算
  const totalCustomLabelCount = config.customLabels.reduce((sum, label) => sum + label.count, 0);
  const labelskipNum = parseInt(config.labelskip, 10) || 0;
  
  // 有効なカスタムラベルがあるかチェック
  const validLabels = config.customLabels.filter(label => label.text.trim() !== '');
  if (validLabels.length === 0) {
    alert('印刷する文字列を入力してください。');
    return;
  }
  
  if (totalCustomLabelCount === 0) {
    alert('印刷する面数を1以上に設定してください。');
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
      generateLabels(labelarr);
    }
    
    // 使い切ったラベルを削除
    remainingLabels = remainingLabels.filter(label => label.count > 0);
    currentSkip = 0; // 2シート目以降はスキップなし
  }
  
  // 印刷枚数の表示（複数シート対応）
  showMultiSheetCustomLabelPrintSummary(totalCustomLabelCount, labelskipNum, sheetsInfo);
  
  // ボタンの状態を更新
  updateButtonStates();
}

function generateOrderDetails(data, labelarr) {
  const tOrder = document.querySelector('#注文明細');
  
  for (let row of data) {
    const cOrder = document.importNode(tOrder.content, true);
    let orderNumber = '';
    
    // 注文情報の設定
    orderNumber = setOrderInfo(cOrder, row, labelarr);
    
    // 個別画像ドロップゾーンの作成
    createIndividualImageDropZone(cOrder, orderNumber);
    
    // 商品項目の処理
    processProductItems(cOrder, row);
    
    // 画像表示の処理
    displayOrderImage(cOrder, orderNumber);
    
    document.body.appendChild(cOrder);
  }
}

function setOrderInfo(cOrder, row, labelarr) {
  let orderNumber = '';
  
  for (let c of Object.keys(row).filter(key => key != CONSTANTS.CSV.PRODUCT_COLUMN)) {
    const divc = cOrder.querySelector("." + c);
    if (divc) {
      if (c == CONSTANTS.CSV.ORDER_NUMBER_COLUMN) {
        orderNumber = OrderNumberManager.getFromCSVRow(row);
        const displayFormat = OrderNumberManager.createDisplayFormat(orderNumber);
        divc.textContent = displayFormat;
        labelarr.push(orderNumber);
      } else if (row[c]) {
        divc.textContent = row[c];
      }
    }
  }
  
  return orderNumber;
}

function createIndividualImageDropZone(cOrder, orderNumber) {
  debugLog(`個別画像ドロップゾーン作成開始 - 注文番号: "${orderNumber}"`);
  
  const individualDropZoneContainer = cOrder.querySelector('.individual-image-dropzone');
  const individualZone = cOrder.querySelector('.individual-order-image-zone');
  
  debugLog(`ドロップゾーンコンテナ発見: ${!!individualDropZoneContainer}`);
  debugLog(`個別ゾーン発見: ${!!individualZone}`);
  
  // 注文画像表示機能が無効の場合は個別画像ゾーン全体を非表示
  const settings = StorageManager.getSettings();
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
      const individualImageDropZone = createIndividualOrderImageDropZone(normalizedOrderNumber);
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

function displayOrderImage(cOrder, orderNumber) {
  // 注文画像表示機能が無効の場合は何もしない
  const settings = StorageManager.getSettings();
  if (!settings.orderImageEnable) {
    return;
  }

  let imageToShow = null;
  if (OrderNumberManager.isValid(orderNumber)) {
    // 注文番号を正規化
    const normalizedOrderNumber = OrderNumberManager.normalize(orderNumber);
    
    // 個別画像があるかチェック
    const individualImage = StorageManager.getOrderImage(normalizedOrderNumber);
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

function generateLabels(labelarr) {
  if (labelarr.length % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET) {
    for (let i = 0; i < labelarr.length % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET; i++) {
      labelarr.push("");
    }
  }
  
  const tL44 = document.querySelector('#L44');
  let cL44 = document.importNode(tL44.content, true);
  let tableL44 = cL44.querySelector("table");
  let tr = document.createElement("tr");
  let i = 0;
  
  for (let label of labelarr) {
    if (i > 0 && i % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET == 0) {
      tableL44.appendChild(tr);
      tr = document.createElement("tr");
      document.body.insertBefore(cL44, tL44);
      cL44 = document.importNode(tL44.content, true);
      tableL44 = cL44.querySelector("table");
      tr = document.createElement("tr");
    } else if (i > 0 && i % CONSTANTS.LABEL.LABELS_PER_ROW == 0) {
      tableL44.appendChild(tr);
      tr = document.createElement("tr");
    }
    tr.appendChild(createLabel(label));
    i++;
  }
  tableL44.appendChild(tr);
  document.body.insertBefore(cL44, tL44);
}

function showPrintSummary() {
  const labelTable = document.querySelectorAll(".label44");
  const pageDiv = document.querySelectorAll(".page");
  
  if (labelTable.length > 0 && pageDiv.length > 0) {
    alert(`印刷準備完了\nラベルシート: ${labelTable.length}枚\n注文明細: ${pageDiv.length}枚`);
  } else if (labelTable.length > 0) {
    alert(`印刷準備完了\nラベルシート: ${labelTable.length}枚`);
  } else if (pageDiv.length > 0) {
    alert(`印刷準備完了\n注文明細: ${pageDiv.length}枚`);
  }
}

function showCustomLabelPrintSummary(customLabelCount, skipCount) {
  const labelTable = document.querySelectorAll(".label44");
  
  if (labelTable.length > 0) {
    alert(`印刷準備完了\nラベルシート: ${labelTable.length}枚`);
  }
}

function showMultiSheetCustomLabelPrintSummary(totalCustomLabelCount, skipCount, sheetsInfo) {
  const labelTable = document.querySelectorAll(".label44");
  
  if (labelTable.length > 0) {
    alert(`印刷準備完了\nラベルシート: ${labelTable.length}枚`);
  }
}

function showCSVWithCustomLabelPrintSummary(csvRowCount, customLabelCount, skipCount, requiredSheets) {
  const labelTable = document.querySelectorAll(".label44");
  const pageDiv = document.querySelectorAll(".page");
  
  if (labelTable.length > 0 && pageDiv.length > 0) {
    alert(`印刷準備完了\nラベルシート: ${labelTable.length}枚\n注文明細: ${pageDiv.length}枚`);
  } else if (labelTable.length > 0) {
    alert(`印刷準備完了\nラベルシート: ${labelTable.length}枚`);
  } else if (pageDiv.length > 0) {
    alert(`印刷準備完了\n注文明細: ${pageDiv.length}枚`);
  }
}
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

function createLabel(labelData=""){
  const divQr = createDiv('qr');
  const divOrdernum = createDiv('ordernum');
  const divYamato = createDiv('yamato');

  // ラベルデータが文字列の場合（既存の注文番号）
  if (typeof labelData === 'string') {
    if (labelData) {
      addP(divOrdernum, labelData);
      createDropzone(divQr);
      const qr = StorageManager.getQRData(labelData);
      if(qr){
        const elImage = document.createElement('img');
        elImage.src = qr['qrimage'];
        divQr.insertBefore(elImage, divQr.firstChild)
        addP(divYamato, qr['receiptnum']);
        addP(divYamato, qr['receiptpassword']);
        addEventQrReset(elImage);
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
    elImage.addEventListener('click', function(event) {
      event.preventDefault();
      const elDrop = elImage.parentNode.querySelector("div");
      elDrop.classList.add('dropzone');
      elDrop.style.zIndex = 99;
      elImage.parentNode.removeChild(elImage);
    });
}

function showDropping(elDrop) {
        elDrop.classList.add('dropover');
}

function hideDropping(elDrop) {
        elDrop.classList.remove('dropover');
}

function readQR(elImage){
  try {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.src = elImage.src;
    
    img.onload = function() {
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
            const duplicates = StorageManager.checkQRDuplicate(barcode.data, ordernum);
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
            
            StorageManager.setQRData(ordernum, qrData);
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
function createBaseImageDropZone(options = {}) {
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

  // localStorageから保存された画像を読み込む
  const savedImage = localStorage.getItem(storageKey);
  if (savedImage) {
    debugLog(`保存された画像を復元: ${storageKey}`);
    updatePreview(savedImage);
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
  function updateAllOrderImages() {
    // 注文画像表示機能が無効の場合は何もしない
    const settings = StorageManager.getSettings();
    if (!settings.orderImageEnable) {
      return;
    }

    const allOrderSections = document.querySelectorAll('section');
    allOrderSections.forEach(orderSection => {
      const imageContainer = orderSection.querySelector('.order-image-container');
      if (!imageContainer) return;

      // 統一化された方法で注文番号を取得
      const orderNumber = OrderNumberManager.getFromOrderSection(orderSection);

      // 個別画像があるかチェック（個別画像を最優先）
      let imageToShow = null;
      if (orderNumber) {
        const individualImage = StorageManager.getOrderImage(orderNumber);
        const globalImage = StorageManager.getOrderImage(); // グローバル画像を取得
        
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
        const globalImage = StorageManager.getOrderImage(); // グローバル画像を取得
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
    });
  }

  function updatePreview(imageUrl) {
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
    localStorage.setItem(storageKey, imageUrl);

    // 個別画像の場合は即座に表示を更新
    if (isIndividual && orderNumber) {
      updateOrderImageDisplay(imageUrl);
    } else if (!isIndividual) {
      // グローバル画像の場合は全ての注文明細の画像を更新
      updateAllOrderImages();
    }

    // 画像クリックでリセット
    preview.addEventListener('click', (e) => {
      e.stopPropagation();
      localStorage.removeItem(storageKey);
      droppedImage = null;
      
      if (isIndividual) {
        dropZone.innerHTML = `<p style="margin: 5px; font-size: 12px; color: #666;">${defaultMessage}</p>`;
        updateOrderImageDisplay(null);
      } else {
        const defaultContentElement = document.getElementById('dropZoneDefaultContent');
        const defaultContent = defaultContentElement ? defaultContentElement.innerHTML : defaultMessage;
        dropZone.innerHTML = defaultContent;
        // グローバル画像がクリアされた場合も全ての注文明細を更新
        updateAllOrderImages();
      }
    });
  }

  // 個別画像用の表示更新関数
  function updateOrderImageDisplay(imageUrl) {
    // 注文画像表示機能が無効の場合は何もしない
    const settings = StorageManager.getSettings();
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

  return {
    element: dropZone,
    getImage: () => droppedImage
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
      reader.onload = (e) => {
        updatePreview(e.target.result);
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
          reader.onload = (e) => {
            debugLog(`画像読み込み完了 - サイズ: ${e.target.result.length} bytes`);
            updatePreview(e.target.result);
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

function createOrderImageDropZone() {
  return createBaseImageDropZone({
    storageKey: 'orderImage',
    isIndividual: false,
    containerClass: 'order-image-drop'
  });
}

// 個別注文用の画像ドロップゾーンを作成する関数（リファクタリング済み）
function createIndividualOrderImageDropZone(orderNumber) {
  return createBaseImageDropZone({
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
});

// ページロード時に説明を表示し、ユーザーにファイル選択を促す
window.addEventListener("load", function() {
  // 初期状態でボタンを設定
  updateButtonStates();

  // ファイル選択を促すメッセージを表示
  alert("CSVファイルを選択するか、カスタムラベル機能を使用してください。");
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
function updateAllOrderImagesVisibility(enabled) {
  debugLog(`画像表示機能が${enabled ? '有効' : '無効'}に変更されました`);
  
  const allOrderSections = document.querySelectorAll('section');
  allOrderSections.forEach(orderSection => {
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
          const individualImageDropZone = createIndividualOrderImageDropZone(orderNumber);
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
          const individualImage = StorageManager.getOrderImage(orderNumber);
          if (individualImage) {
            imageToShow = individualImage;
            debugLog(`個別画像を表示: ${orderNumber}`);
          } else {
            const globalImage = StorageManager.getOrderImage();
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
  });
}

// 複数カスタムラベルの初期化
function initializeCustomLabels(customLabels) {
  const container = document.getElementById('customLabelsContainer');
  container.innerHTML = '';
  
  // 説明文を一番上に追加
  const instructionDiv = document.createElement('div');
  instructionDiv.className = 'custom-labels-instructions';
  instructionDiv.innerHTML = `
    <div class="instructions-content">
      <strong>カスタムラベル設定について：</strong><br>
      • 実際のラベルサイズ: 48.3mm × 25.3mm<br>
      • テキストのみ入力可能<br>
      • 改行: Enterキー<br>
      • 書式設定: 右クリックでコンテキストメニューを表示（太字、斜体、フォントサイズ変更、フォントの変更）
    </div>
  `;
  container.appendChild(instructionDiv);
  
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
    
    editorElement.addEventListener('input', async function() {
      // 文字列が入力されたら強調表示をクリア
      const item = editorElement.closest('.custom-label-item');
      if (item && editorElement.textContent.trim() !== '') {
        item.classList.remove('error');
      }
      
      saveCustomLabels();
      updateButtonStates();
      await updateCustomLabelsSummary();
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
  // カスタムラベル有効化チェックボックス
  const customLabelEnable = document.getElementById('customLabelEnable');
  if (customLabelEnable) {
    customLabelEnable.addEventListener('change', function() {
      toggleCustomLabelRow(this.checked);
      StorageManager.set(StorageManager.KEYS.CUSTOM_LABEL_ENABLE, this.checked);
      updateButtonStates();
    });
  }

  // カスタムラベル追加ボタン
  const addButton = document.getElementById('addCustomLabelBtn');
  if (addButton) {
    addButton.addEventListener('click', function() {
      debugLog('ラベル追加ボタンがクリックされました'); // デバッグ用
      addCustomLabelItem('', 1, null, true);
      saveCustomLabels();
      updateButtonStates();
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
  const executeButton = document.getElementById("executeButton");
  const customLabelOnlyButton = document.getElementById("customLabelOnlyButton");
  const printButton = document.getElementById("printButton");
  const customLabelEnable = document.getElementById("customLabelEnable");

  // CSV処理実行ボタンの状態
  executeButton.disabled = fileInput.files.length === 0;

  // カスタムラベル専用実行ボタンの状態
  const hasValidCustomLabels = customLabelEnable.checked && hasCustomLabelsWithContent();
  customLabelOnlyButton.disabled = !hasValidCustomLabels;

  // 印刷ボタンの状態（何らかのコンテンツが生成されている場合に有効）
  const hasSheets = document.querySelectorAll('.sheet').length > 0;
  const hasLabels = document.querySelectorAll('.label44').length > 0;
  const hasContent = hasSheets || hasLabels;
  printButton.disabled = !hasContent;

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

// カスタムラベルのバリデーションと入力促進
function validateAndPromptCustomLabels() {
  const customLabelEnable = document.getElementById("customLabelEnable");
  if (!customLabelEnable.checked) {
    return true; // カスタムラベル無効の場合は問題なし
  }
  
  if (hasEmptyEnabledCustomLabels()) {
    // 未設定項目を強調表示
    highlightEmptyCustomLabels();
    
    const result = confirm(
      'カスタムラベルが有効になっていますが、文字列が未設定の項目があります。\n\n' +
      '選択肢：\n' +
      'OK: 入力画面に戻って文字列を設定する\n' +
      'キャンセル: 未設定のラベル項目を削除して続行する'
    );
    
    if (result) {
      // OKの場合は処理を中断して入力を促す
      alert('赤く強調表示された項目の文字列を入力してから再度実行してください。');
      return false;
    } else {
      // キャンセルの場合は未設定項目を削除して続行
      removeEmptyCustomLabels();
      clearCustomLabelHighlights(); // 強調表示をクリア
      return true;
    }
  }
  
  return true; // 問題なし
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
  
  // Enterキーでの改行処理を改善
  editor.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      // Enterキーの処理を改善 - 1回で改行
      e.preventDefault();
      
      const selection = window.getSelection();
      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        
        // 選択範囲がある場合は削除
        range.deleteContents();
        
        // 改行要素を作成
        const br = document.createElement('br');
        range.insertNode(br);
        
        // さらにもう一つのbrを挿入してカーソル位置を確保
        const br2 = document.createElement('br');
        range.setStartAfter(br);
        range.insertNode(br2);
        
        // カーソルを2番目のbrの前に配置
        range.setStartBefore(br2);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
      }
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
        if (menu.parentNode) {
          menu.parentNode.removeChild(menu);
        }
        document.removeEventListener('click', closeMenu);
      });
    }, 100);
  });
}

// フォントサイズ選択メニューを作成（非同期）
async function createFontSizeMenu(x, y, editor, hasSelection = true) {
  const menu = document.createElement('div');
  
  // 初期スタイル設定（位置は後で調整）
  menu.style.cssText = `
    position: fixed;
    background: white;
    border: 1px solid #ccc;
    border-radius: 4px;
    padding: 5px 0;
    box-shadow: 0 2px 10px rgba(0,0,0,0.2);
    z-index: 10000;
    font-family: sans-serif;
    min-width: 140px;
    visibility: hidden;
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
        if (menu.parentNode) {
          menu.parentNode.removeChild(menu);
        }
        
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
          if (menu.parentNode) {
            menu.parentNode.removeChild(menu);
          }
          
          saveCustomLabels();
        }, 10);
      });
      
      menu.appendChild(item);
    });

    // カスタムフォント選択オプション（IndexedDBベース）
    try {
      if (fontDB) {
        const customFonts = await fontDB.getAllFonts();
        if (Object.keys(customFonts).length > 0) {
          // フォント用区切り線
          const fontSeparator = document.createElement('div');
          fontSeparator.style.cssText = `
            height: 1px;
            background-color: #ddd;
            margin: 5px 0;
          `;
          menu.appendChild(fontSeparator);

          // フォントファミリーオプション
          const fontFamilyLabel = document.createElement('div');
          fontFamilyLabel.textContent = 'フォント';
          fontFamilyLabel.style.cssText = `
            padding: 5px 15px;
            font-size: 11px;
            color: #666;
            font-weight: bold;
          `;
          menu.appendChild(fontFamilyLabel);

          // デフォルトフォント
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
            
            if (menu.parentNode) {
              menu.parentNode.removeChild(menu);
            }
            
            saveCustomLabels();
          }, 10);
        });
        
        menu.appendChild(defaultFontItem);

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
                
                if (menu.parentNode) {
                  menu.parentNode.removeChild(menu);
                }
                
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
        }

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
              
              if (menu.parentNode) {
                menu.parentNode.removeChild(menu);
              }
              
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
  
  // 右端からはみ出る場合は左にずらす
  let adjustedX = x;
  if (x + menuRect.width > viewportWidth) {
    adjustedX = viewportWidth - menuRect.width - 5;
  }
  
  // 下端からはみ出る場合は上にずらす
  let adjustedY = y;
  if (y + menuRect.height > viewportHeight) {
    adjustedY = y - menuRect.height;
  }
  
  // 調整後の位置を設定して表示
  menu.style.left = `${adjustedX}px`;
  menu.style.top = `${adjustedY}px`;
  menu.style.visibility = 'visible';
  
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
  
  // 既存のonClickを保存
  const originalOnClick = printButton.onclick;
  
  // 新しい処理を設定
  printButton.onclick = function() {
    // まず印刷を実行
    window.print();
    
    // 印刷ダイアログが閉じた後に実行される
    setTimeout(() => {
      // 印刷完了の確認
      if (confirm('印刷が完了しましたか？完了した場合、次回のスキップ枚数を更新します。')) {
        updateSkipCount();
      }
    }, 100);
  };
});

// スキップ枚数を更新する関数
async function updateSkipCount() {
  // 現在のスキップ枚数を取得
  const currentSkip = parseInt(document.getElementById("labelskipnum").value, 10) || 0;
  
  // 使用したラベル枚数を計算
  let usedLabels = 0;
  
  // CSV処理による注文ラベル数を取得
  const orderPages = document.querySelectorAll(".page");
  if (orderPages.length > 0) {
    // 注文ページの数 = CSV行数
    usedLabels = orderPages.length;
  }
  
  // カスタムラベルが有効な場合、その枚数を追加
  if (document.getElementById("customLabelEnable").checked) {
    const customLabels = getCustomLabelsFromUI();
    const totalCustomCount = customLabels.reduce((sum, label) => sum + label.count, 0);
    usedLabels += totalCustomCount;
  }
  
  // 合計使用枚数を計算
  const totalUsed = currentSkip + usedLabels;
  
  // 44枚のシートサイズに合わせて余りを計算
  const newSkipValue = totalUsed % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET;
  
  // 新しいスキップ枚数を設定
  document.getElementById("labelskipnum").value = newSkipValue;
  StorageManager.set(StorageManager.KEYS.LABEL_SKIP, newSkipValue);
  
  // 更新完了メッセージ
  alert(`次回のスキップ枚数を ${newSkipValue} 枚に更新しました。`);
  
  // カスタムラベルの上限も更新
  await updateCustomLabelsSummary();
}

// カスタムフォント管理機能
function initializeFontDropZone() {
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
    // fontDBが初期化されていない場合
    if (!fontDB) {
      alert('フォントデータベースが初期化されていません。\nページを再読み込みしてから再試行してください。');
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
    
    // 重複チェック
    const existingFont = await fontDB.getFont(fontName);
    if (existingFont) {
      if (!confirm(`フォント "${fontName}" は既に存在します。上書きしますか？`)) {
        showFontUploadProgress(false);
        return;
      }
    }

    // IndexedDBに保存
    await fontDB.saveFont(fontName, arrayBuffer, {
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

// ローカルストレージ容量チェック機能
function checkStorageCapacity(requiredSize) {
  try {
    // 現在の使用量を計算
    let currentSize = 0;
    for (let key in localStorage) {
      if (localStorage.hasOwnProperty(key)) {
        currentSize += localStorage[key].length + key.length;
      }
    }
    
    // 概算の残り容量（通常5-10MB程度）
    const estimatedLimit = 10 * 1024 * 1024; // 10MB
    const remainingSize = estimatedLimit - currentSize;
    
    console.log(`ストレージ使用量: ${Math.round(currentSize / 1024)}KB / 推定制限: ${Math.round(estimatedLimit / 1024)}KB`);
    console.log(`必要サイズ: ${Math.round(requiredSize / 1024)}KB / 残り容量: ${Math.round(remainingSize / 1024)}KB`);
    
    return requiredSize < remainingSize;
  } catch (error) {
    console.warn('ストレージ容量チェックエラー:', error);
    return true; // エラーの場合は続行
  }
}

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
    // fontDBが初期化されていない場合は何もしない
    if (!fontDB) {
      console.warn('FontDatabaseが初期化されていません');
      return;
    }
    
    const fonts = await fontDB.getAllFonts();
    
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
    // fontDBが初期化されていない場合はメッセージを表示
    if (!fontDB) {
      fontListElement.innerHTML = '<div style="color: #999; text-align: center; padding: 10px;">フォントデータベースを初期化中...</div>';
      return;
    }
    
    const fonts = await fontDB.getAllFonts();
    
    if (Object.keys(fonts).length === 0) {
      fontListElement.innerHTML = '<div style="color: #999; text-align: center; padding: 10px;">フォントファイルをアップロードしてください</div>';
      return;
    }

    fontListElement.innerHTML = Object.entries(fonts).map(([fontName, fontData]) => {
      const metadata = fontData.metadata || {};
      const originalName = metadata.originalName || fontName;
      const createdAt = metadata.createdAt || Date.now();
      const sizeMB = (fontData.data.byteLength / 1024 / 1024).toFixed(2);
      
      return `
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px; border-bottom: 1px solid #eee; background: #f8f9fa; margin-bottom: 5px; border-radius: 4px;">
          <div style="flex: 1;">
            <div style="font-family: '${fontName}', sans-serif; font-size: 14px; font-weight: 500; color: #333;">${fontName}</div>
            <div style="font-size: 11px; color: #666; margin-top: 2px;">
              ${originalName} 
              <span style="color: #999;">• ${new Date(createdAt).toLocaleDateString()} • ${sizeMB}MB</span>
            </div>
          </div>
          <button onclick="removeFontFromList('${fontName}')" 
                  style="background: #dc3545; color: white; border: none; padding: 4px 8px; border-radius: 3px; cursor: pointer; font-size: 11px; margin-left: 10px; min-width: 50px;"
                  onmouseover="this.style.background='#c82333'" 
                  onmouseout="this.style.background='#dc3545'"
                  title="フォント '${fontName}' を削除">
            削除
          </button>
        </div>
      `;
    }).join('');
  } catch (error) {
    console.error('フォントリスト更新エラー:', error);
    fontListElement.innerHTML = '<div style="color: #d32f2f; text-align: center; padding: 10px;">フォントリストの読み込みに失敗しました</div>';
  } finally {
    // セクションの高さを調整
    setTimeout(adjustFontSectionHeight, 100);
  }
}

async function removeFontFromList(fontName) {
  try {
    // fontDBが初期化されていない場合
    if (!fontDB) {
      alert('フォントデータベースが初期化されていません。');
      return;
    }
    
    const fontData = await fontDB.getFont(fontName);
    
    if (!fontData) {
      alert('指定されたフォントが見つかりません。');
      return;
    }
    
    const originalName = fontData.metadata?.originalName || fontName;
    const sizeMB = (fontData.data.byteLength / 1024 / 1024).toFixed(2);
    
    const confirmMessage = `フォント "${fontName}" (${originalName}, ${sizeMB}MB) を削除しますか？\n\n削除すると、このフォントを使用しているカスタムラベルの表示が変わる可能性があります。`;
    
    if (confirm(confirmMessage)) {
      try {
        await fontDB.deleteFont(fontName);
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

// シンプルで確実なスタイル適用関数
function applyStyleToSelection(styleProperty, styleValue, editor, isDefault = false) {
  const selection = window.getSelection();
  
  if (selection.rangeCount === 0 || selection.isCollapsed) {
    return; // 選択範囲がない場合は何もしない
  }
  
  try {
    const range = selection.getRangeAt(0);
    console.log(`スタイル変更: "${styleProperty}: ${styleValue}" を適用中...`);
    
    // 選択範囲のテキスト内容を取得
    const selectedText = range.toString();
    console.log('選択されたテキスト:', selectedText);
    
    if (!selectedText) {
      console.log('選択されたテキストがありません');
      return;
    }
    
    // 既存のスタイルを簡単に収集（選択範囲の最初の要素から）
    let existingStyles = '';
    const startContainer = range.startContainer;
    let currentNode = startContainer.nodeType === Node.TEXT_NODE ? startContainer.parentNode : startContainer;
    
    // 親要素をたどってspan要素を探す（最大3層まで）
    for (let i = 0; i < 3 && currentNode && currentNode !== editor; i++) {
      if (currentNode.tagName === 'SPAN') {
        const style = currentNode.getAttribute('style') || '';
        console.log('発見したspan要素:', currentNode, 'スタイル:', style);
        
        if (style) {
          // 除外するプロパティ以外を収集
          const styleRules = style.split(';').filter(rule => {
            const trimmed = rule.trim();
            if (!trimmed) return false;
            const property = trimmed.split(':')[0].trim().toLowerCase();
            return property !== styleProperty;
          });
          
          existingStyles = styleRules.join('; ');
          console.log('既存スタイル（除外後）:', existingStyles);
        }
        break; // 最初に見つかったspan要素を使用
      }
      currentNode = currentNode.parentNode;
    }
    
    // 選択範囲の内容を削除
    range.deleteContents();
    
    // デフォルト値でない場合、または既存スタイルがある場合はspanで囲む
    if (!isDefault || existingStyles) {
      console.log('新しいspan要素を作成します');
      
      const newSpan = document.createElement('span');
      
      // 既存のスタイルを適用
      if (existingStyles) {
        newSpan.setAttribute('style', existingStyles);
      }
      
      // 新しいスタイルを設定（デフォルト値でない場合のみ）
      if (!isDefault) {
        if (styleProperty === 'font-family') {
          newSpan.style.fontFamily = styleValue;
        } else if (styleProperty === 'font-size') {
          newSpan.style.fontSize = styleValue + (typeof styleValue === 'number' ? 'pt' : '');
        }
      }
      
      newSpan.textContent = selectedText;
      range.insertNode(newSpan);
      
      // 新しいspan要素を選択状態にする
      range.selectNodeContents(newSpan);
      selection.removeAllRanges();
      selection.addRange(range);
      
      console.log(`スタイル "${styleProperty}: ${styleValue}" を適用しました`);
    } else {
      // デフォルト値で既存スタイルもない場合はプレーンテキストを挿入
      console.log('プレーンテキストを挿入します');
      
      const textNode = document.createTextNode(selectedText);
      range.insertNode(textNode);
      
      range.setStartBefore(textNode);
      range.setEndAfter(textNode);
      selection.removeAllRanges();
      selection.addRange(range);
      
      console.log('デフォルト値に戻しました');
    }
    
    // DOM構造の確認用ログ
    console.log('処理後のエディタHTML:', editor.innerHTML);
    
    // エディタ全体の空のspan要素を掃除
    cleanupEmptySpans(editor);
    
    console.log('スタイル変更完了');
    
  } catch (error) {
    console.warn('スタイル適用エラー:', error);
  }
  
  // エディタにフォーカスを戻す
  editor.focus();
}

// フォントファミリーを選択範囲に適用（統合された関数を使用）
function applyFontFamilyToSelection(fontFamily, editor) {
  const isDefault = !fontFamily || fontFamily === '';
  applyStyleToSelection('font-family', fontFamily, editor, isDefault);
}

// フォントサイズを選択範囲に適用（統合された関数を使用）
function applyFontSizeToSelection(fontSize, editor) {
  applyStyleToSelection('font-size', fontSize, editor, false);
}

// 選択範囲から既存のスタイルを収集する関数（指定されたプロパティを除外）
// 空のspan要素を掃除する関数（シンプル版）
function cleanupEmptySpans(editor) {
  console.log('cleanupEmptySpans開始');
  
  let removedCount = 0;
  const spans = editor.querySelectorAll('span');
  
  // 後ろから処理して、インデックスのずれを防ぐ
  Array.from(spans).reverse().forEach(span => {
    // 空のspanを削除
    if (span.innerHTML.trim() === '' || span.textContent.trim() === '') {
      console.log('空のspanを削除:', span);
      span.remove();
      removedCount++;
      return;
    }
    
    // スタイル属性が空のspanは子要素を親に移動して削除
    const style = span.getAttribute('style') || '';
    if (style.trim() === '') {
      console.log('スタイルが空のspanを削除:', span);
      const parent = span.parentNode;
      if (parent) {
        while (span.firstChild) {
          parent.insertBefore(span.firstChild, span);
        }
        span.remove();
        removedCount++;
      }
    }
  });
  
  console.log(`cleanupEmptySpans完了: ${removedCount}個のspanを削除`);
}

// ===========================================
// IndexedDBフォント機能の初期化とヘルパー関数
// ===========================================

// フォントセクションの折りたたみ機能
function toggleFontSection() {
  const content = document.getElementById('fontSectionContent');
  const arrow = document.getElementById('fontSectionArrow');
  
  if (content.style.maxHeight && content.style.maxHeight !== '0px') {
    // 折りたたむ
    content.style.maxHeight = '0px';
    arrow.style.transform = 'rotate(-90deg)';
    localStorage.setItem('fontSectionCollapsed', 'true');
  } else {
    // 展開する
    content.style.maxHeight = content.scrollHeight + 'px';
    arrow.style.transform = 'rotate(0deg)';
    localStorage.setItem('fontSectionCollapsed', 'false');
  }
}

// フォントセクションの初期状態を設定
function initializeFontSection() {
  const content = document.getElementById('fontSectionContent');
  const arrow = document.getElementById('fontSectionArrow');
  const isCollapsed = localStorage.getItem('fontSectionCollapsed');
  
  // 初期状態は折りたたみ（明示的に展開が設定されている場合のみ展開）
  if (isCollapsed === 'false') {
    // 展開状態
    setTimeout(() => {
      content.style.maxHeight = content.scrollHeight + 'px';
    }, 100);
    arrow.style.transform = 'rotate(0deg)';
  } else {
    // 折りたたみ状態（デフォルト）
    content.style.maxHeight = '0px';
    arrow.style.transform = 'rotate(-90deg)';
  }
}

// フォントリスト更新時にセクション高さを調整
function adjustFontSectionHeight() {
  const content = document.getElementById('fontSectionContent');
  if (content && content.style.maxHeight !== '0px') {
    content.style.maxHeight = content.scrollHeight + 'px';
  }
}

// フォント機能の初期化
document.addEventListener('DOMContentLoaded', async function() {
  try {
    // フォントセクションの初期化
    initializeFontSection();
    
    // FontDatabaseインスタンスを作成して初期化
    fontDB = new FontDatabase();
    await fontDB.init();
    console.log('FontDatabase初期化完了');
    
    // 既存のデータをチェックして、不正なデータがあればクリア
    const fonts = await fontDB.getAllFonts();
    let hasInvalidData = false;
    
    for (const [fontName, fontData] of Object.entries(fonts)) {
      if (!fontData || !fontData.data || !fontData.metadata || !fontName || fontName === '' || fontName === '0') {
        console.warn(`不正なフォントデータを発見: "${fontName}"`, fontData);
        hasInvalidData = true;
        break;
      }
    }
    
    if (hasInvalidData) {
      console.log('不正なフォントデータが見つかりました。データベースをクリアします。');
      await fontDB.clearAllFonts();
    }
    
    // IndexedDBのフォントをCSSに読み込み
    await loadCustomFontsCSS();
    
  } catch (error) {
    console.error('フォント機能初期化エラー:', error);
  }
});

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
