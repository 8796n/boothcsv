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

// ===== バイナリデータ変換ユーティリティ =====

// Base64文字列をArrayBufferに変換
function base64ToArrayBuffer(base64) {
  try {
    // Data URLの場合はプレフィックスを除去
    const base64Data = base64.includes(',') ? base64.split(',')[1] : base64;
    const binaryString = atob(base64Data);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  } catch (error) {
    console.error('Base64変換エラー:', error);
    return null;
  }
}

// ArrayBufferをBase64文字列に変換
function arrayBufferToBase64(buffer) {
  try {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  } catch (error) {
    console.error('ArrayBuffer変換エラー:', error);
    return null;
  }
}

// データがBase64文字列かどうかを判定
function isBase64String(data) {
  if (typeof data !== 'string') return false;
  
  // Data URLの場合
  if (data.startsWith('data:')) return true;
  
  // Base64文字列の場合（最低限のチェック）
  const base64Regex = /^[A-Za-z0-9+/]+=*$/;
  return base64Regex.test(data) && data.length % 4 === 0 && data.length > 20;
}

// MIMEタイプをData URLから抽出
function extractMimeType(dataUrl) {
  if (!dataUrl.startsWith('data:')) return 'application/octet-stream';
  const match = dataUrl.match(/^data:([^;]+)/);
  return match ? match[1] : 'application/octet-stream';
}

// IndexedDBを使用した統合ストレージ管理クラス（破壊的移行版）
class UnifiedDatabase {
  constructor() {
    this.dbName = 'BoothCSVStorage';
    this.version = 3; // バージョンアップ（破壊的変更）
    this.fontStoreName = 'fonts';
    this.settingsStoreName = 'settings';
    this.imagesStoreName = 'images';
    this.qrDataStoreName = 'qrData';
    this.db = null;
    this.connectionLogged = false;
  }

  // データベース初期化
  async init() {
    if (this.db) {
      return this.db;
    }
    
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, this.version);
      
      request.onerror = () => {
        console.error('IndexedDB初期化エラー:', request.error);
        reject(request.error);
      };
      
      request.onsuccess = () => {
        this.db = request.result;
        if (!this.connectionLogged) {
          console.log(`📂 IndexedDB "${this.dbName}" v${this.version} に接続しました`);
          this.connectionLogged = true;
        }
        resolve(this.db);
      };
      
      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        this.createObjectStores(db);
      };
    });
  }

  // オブジェクトストア作成
  createObjectStores(db) {
    // 既存のストアを削除（破壊的変更）
    const existingStores = Array.from(db.objectStoreNames);
    existingStores.forEach(storeName => {
      db.deleteObjectStore(storeName);
      console.log(`🗑️ 既存ストア削除: ${storeName}`);
    });

    // フォントストアを作成
    const fontStore = db.createObjectStore(this.fontStoreName, { keyPath: 'name' });
    fontStore.createIndex('createdAt', 'createdAt', { unique: false });
    fontStore.createIndex('size', 'size', { unique: false });

    // 設定ストアを作成
    const settingsStore = db.createObjectStore(this.settingsStoreName, { keyPath: 'key' });

    // 画像ストアを作成（バイナリ対応）
    const imagesStore = db.createObjectStore(this.imagesStoreName, { keyPath: 'key' });
    imagesStore.createIndex('type', 'type', { unique: false });
    imagesStore.createIndex('orderNumber', 'orderNumber', { unique: false });
    imagesStore.createIndex('createdAt', 'createdAt', { unique: false });

    // QRデータストアを作成（バイナリ対応）
    const qrStore = db.createObjectStore(this.qrDataStoreName, { keyPath: 'orderNumber' });
    qrStore.createIndex('qrhash', 'qrhash', { unique: false });
    qrStore.createIndex('createdAt', 'createdAt', { unique: false });

    console.log('🆕 新しいデータストアを作成しました');
  }

  // 破壊的移行処理
  async performDestructiveMigration() {
    try {
      console.log('🚨 破壊的移行を開始します...');
      
      // localStorage使用量を確認
      const usage = this.analyzeLocalStorageUsage();
      
      if (usage.totalItems > 0) {
        console.log(`📊 削除対象: ${usage.totalItems}項目 (${usage.totalSizeMB}MB)`);
        
        // ユーザーに通知
        const userConfirm = confirm(
          `🔄 システム移行のお知らせ\n\n` +
          `より高速で安定したデータ保存システムに移行します。\n` +
          `既存の設定・データ（${usage.totalItems}項目）は削除され、\n` +
          `改めて設定が必要になります。\n\n` +
          `移行を実行しますか？\n\n` +
          `※この操作は取り消せません`
        );
        
        if (!userConfirm) {
          alert('移行がキャンセルされました。\nアプリケーションは従来の方式で動作します。');
          return false;
        }
        
        // localStorage完全削除
        await this.clearAllLocalStorage();
        
        // 移行完了通知
        alert(
          `✅ システム移行完了\n\n` +
          `新しいデータ保存システムが有効になりました。\n` +
          `設定・フォント・画像データを改めて登録してください。\n\n` +
          `今後はより高速で大容量のデータ保存が可能です。`
        );
      }
      
      console.log('✅ 破壊的移行が完了しました');
      return true;
      
    } catch (error) {
      console.error('❌ 破壊的移行エラー:', error);
      return false;
    }
  }

  // localStorage完全削除（アプリケーション固有のキーのみ）
  async clearAllLocalStorage() {
    // このアプリケーションで使用するキーのパターン
    const appKeyPatterns = [
      'labelyn',
      'labelskip', 
      'sortByPaymentDate',
      'customLabelEnable',
      'customLabelText',
      'customLabelCount',
      'customLabels',
      'orderImageEnable',
      'fontSectionCollapsed', // IndexedDBに移行済み
      'orderImage',
      'orderImage_',
      'customFont_',
      'migrationCompleted'
    ];
    
    const itemsToRemove = [];
    
    // 全てのキーをチェックして、アプリケーション固有のもののみを収集
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key) {
        const isAppKey = appKeyPatterns.some(pattern => {
          if (pattern.endsWith('_')) {
            return key.startsWith(pattern);
          } else {
            return key === pattern;
          }
        });
        
        // QRコードデータの可能性もチェック（JSON形式でqrhashを含む）
        if (!isAppKey) {
          try {
            const value = localStorage.getItem(key);
            const parsed = JSON.parse(value);
            if (parsed && parsed.qrhash) {
              // QRコードデータと判定
              isAppKey = true;
            }
          } catch (e) {
            // JSON以外は無視
          }
        }
        
        if (isAppKey) {
          itemsToRemove.push(key);
        }
      }
    }
    
    // アプリケーション固有のキーのみ削除
    itemsToRemove.forEach(key => {
      localStorage.removeItem(key);
      console.log(`🗑️ アプリデータ削除: ${key}`);
    });
    
    const otherKeysCount = localStorage.length;
    console.log(`🧹 アプリデータ削除完了: ${itemsToRemove.length}項目`);
    if (otherKeysCount > 0) {
      console.log(`ℹ️ 他サービスのデータ保持: ${otherKeysCount}項目`);
    }
  }

  // localStorage使用量分析（アプリケーション固有のキーのみ）
  analyzeLocalStorageUsage() {
    let totalSize = 0;
    let totalItems = 0;
    const categories = {
      fonts: 0,
      settings: 0,
      images: 0,
      qrData: 0,
      other: 0
    };
    
    // このアプリケーションで使用するキーのパターン
    const appKeyPatterns = [
      'labelyn',
      'labelskip', 
      'sortByPaymentDate',
      'customLabelEnable',
      'customLabelText',
      'customLabelCount',
      'customLabels',
      'orderImageEnable',
      'fontSectionCollapsed', // IndexedDBに移行済み
      'orderImage',
      'orderImage_',
      'customFont_',
      'migrationCompleted'
    ];
    
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key) {
        let isAppKey = false;
        
        // アプリケーション固有のキーかチェック
        isAppKey = appKeyPatterns.some(pattern => {
          if (pattern.endsWith('_')) {
            return key.startsWith(pattern);
          } else {
            return key === pattern;
          }
        });
        
        // QRコードデータの可能性もチェック
        if (!isAppKey) {
          try {
            const value = localStorage.getItem(key);
            const parsed = JSON.parse(value);
            if (parsed && parsed.qrhash) {
              isAppKey = true;
            }
          } catch (e) {
            // JSON以外は無視
          }
        }
        
        if (isAppKey) {
          const value = localStorage.getItem(key);
          const size = new Blob([value || '']).size;
          totalSize += size;
          totalItems++;
          
          // カテゴリ分類
          if (key.startsWith('customFont_')) {
            categories.fonts++;
          } else if (['labelyn', 'labelskip', 'sortByPaymentDate', 'customLabelEnable', 'orderImageEnable'].includes(key)) {
            categories.settings++;
          } else if (key.startsWith('orderImage')) {
            categories.images++;
          } else if (key.includes('qr') || key.includes('receipt')) {
            categories.qrData++;
          } else {
            categories.other++;
          }
        }
      }
    }
    
    return {
      totalItems,
      totalSizeKB: Math.round(totalSize / 1024 * 100) / 100,
      totalSizeMB: Math.round(totalSize / 1024 / 1024 * 100) / 100,
      categories
    };
  }

  // === バイナリ対応メソッド ===

  // 画像保存（バイナリ優先）
  async setImage(key, imageData, type = 'unknown', orderNumber = null) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.imagesStoreName], 'readwrite');
      const store = transaction.objectStore(this.imagesStoreName);
      
      // データの形式を最適化
      let optimizedData = imageData;
      let mimeType = type;
      let isBinary = false;
      
      if (isBase64String(imageData)) {
        // Base64をバイナリに変換
        const arrayBuffer = base64ToArrayBuffer(imageData);
        if (arrayBuffer) {
          optimizedData = arrayBuffer;
          mimeType = extractMimeType(imageData);
          isBinary = true;
          console.log(`🔄 画像をバイナリ最適化: ${key}`);
        }
      } else if (imageData instanceof ArrayBuffer) {
        isBinary = true;
      }
      
      const imageObject = {
        key: key,
        data: optimizedData,
        type: mimeType,
        orderNumber: orderNumber,
        createdAt: Date.now(),
        isBinary: isBinary
      };
      
      const request = store.put(imageObject);
      
      request.onsuccess = () => resolve();
      request.onerror = () => {
        console.error('画像保存エラー:', request.error);
        reject(request.error);
      };
    });
  }

  // 画像取得（自動変換）
  async getImage(key) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.imagesStoreName], 'readonly');
      const store = transaction.objectStore(this.imagesStoreName);
      const request = store.get(key);
      
      request.onsuccess = () => {
        const result = request.result;
        if (!result) {
          resolve(null);
          return;
        }
        
        // バイナリデータをData URLに変換
        if (result.isBinary && result.data instanceof ArrayBuffer) {
          const base64Data = arrayBufferToBase64(result.data);
          if (base64Data) {
            const dataUrl = `data:${result.type || 'image/png'};base64,${base64Data}`;
            resolve(dataUrl);
          } else {
            resolve(result.data);
          }
        } else {
          resolve(result.data);
        }
      };
      
      request.onerror = () => {
        console.error('画像取得エラー:', request.error);
        resolve(null);
      };
    });
  }

  // 設定管理
  async setSetting(key, value) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.settingsStoreName], 'readwrite');
      const store = transaction.objectStore(this.settingsStoreName);
      
      const settingObject = {
        key: key,
        value: value,
        updatedAt: Date.now()
      };
      
      const request = store.put(settingObject);
      
      request.onsuccess = () => resolve();
      request.onerror = () => {
        console.error('設定保存エラー:', request.error);
        reject(request.error);
      };
    });
  }

  async getSetting(key) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.settingsStoreName], 'readonly');
      const store = transaction.objectStore(this.settingsStoreName);
      const request = store.get(key);
      
      request.onsuccess = () => {
        const result = request.result;
        resolve(result ? result.value : null);
      };
      
      request.onerror = () => {
        console.error('設定取得エラー:', request.error);
        resolve(null);
      };
    });
  }

  // フォント管理（バイナリ対応）
  async setFont(fontName, fontData) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.fontStoreName], 'readwrite');
      const store = transaction.objectStore(this.fontStoreName);
      
      // フォントデータの最適化
      let optimizedData = fontData;
      let isBinary = false;
      
      if (isBase64String(fontData)) {
        const arrayBuffer = base64ToArrayBuffer(fontData);
        if (arrayBuffer) {
          optimizedData = arrayBuffer;
          isBinary = true;
          console.log(`🔄 フォントをバイナリ最適化: ${fontName}`);
        }
      } else if (fontData instanceof ArrayBuffer) {
        isBinary = true;
      }
      
      const fontObject = {
        name: fontName,
        data: optimizedData,
        metadata: {
          type: isBinary ? 'font/ttf' : 'text/plain',
          size: optimizedData.byteLength || optimizedData.length
        },
        createdAt: Date.now(),
        isBinary: isBinary
      };
      
      const request = store.put(fontObject);
      
      request.onsuccess = () => resolve();
      request.onerror = () => {
        console.error('フォント保存エラー:', request.error);
        reject(request.error);
      };
    });
  }

  async getAllFonts() {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.fontStoreName], 'readonly');
      const store = transaction.objectStore(this.fontStoreName);
      const request = store.getAll();
      
      request.onsuccess = () => {
        const fonts = {};
        request.result.forEach(font => {
          fonts[font.name] = font;
        });
        resolve(fonts);
      };
      
      request.onerror = () => {
        console.error('フォント一覧取得エラー:', request.error);
        reject(request.error);
      };
    });
  }

  // QRデータ管理（バイナリ対応）
  async setQRData(orderNumber, qrData) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.qrDataStoreName], 'readwrite');
      const store = transaction.objectStore(this.qrDataStoreName);
      
      // qrDataがnullの場合は削除処理
      if (qrData === null || qrData === undefined) {
        const deleteRequest = store.delete(orderNumber);
        
        deleteRequest.onsuccess = () => {
          debugLog(`🗑️ QRデータを削除しました: ${orderNumber}`);
          resolve();
        };
        
        deleteRequest.onerror = () => {
          console.error('QRデータ削除エラー:', deleteRequest.error);
          reject(deleteRequest.error);
        };
        
        return;
      }
      
      // QR画像のバイナリ最適化
      let optimizedQRImage = qrData.qrimage;
      let isBinary = false;
      
      if (qrData.qrimage && isBase64String(qrData.qrimage)) {
        const arrayBuffer = base64ToArrayBuffer(qrData.qrimage);
        if (arrayBuffer) {
          optimizedQRImage = arrayBuffer;
          isBinary = true;
        }
      } else if (qrData.qrimage instanceof ArrayBuffer) {
        isBinary = true;
      }
      
      const qrObject = {
        orderNumber: orderNumber,
        receiptnum: qrData.receiptnum,
        receiptpassword: qrData.receiptpassword,
        qrimage: optimizedQRImage,
        qrimageType: qrData.qrimageType || 'image/png',
        qrhash: qrData.qrhash,
        createdAt: Date.now(),
        isBinary: isBinary
      };
      
      const request = store.put(qrObject);
      
      request.onsuccess = () => resolve();
      request.onerror = () => {
        console.error('QRデータ保存エラー:', request.error);
        reject(request.error);
      };
    });
  }

  async getQRData(orderNumber) {
    if (!this.db) await this.init();
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.qrDataStoreName], 'readonly');
      const store = transaction.objectStore(this.qrDataStoreName);
      const request = store.get(orderNumber);
      
      request.onsuccess = () => {
        const result = request.result;
        if (!result) {
          resolve(null);
          return;
        }
        
        // バイナリQR画像をBase64に変換
        if (result.isBinary && result.qrimage instanceof ArrayBuffer) {
          const base64Data = arrayBufferToBase64(result.qrimage);
          if (base64Data) {
            result.qrimage = `data:${result.qrimageType || 'image/png'};base64,${base64Data}`;
          }
        }
        
        resolve(result);
      };
      
      request.onerror = () => {
        console.error('QRデータ取得エラー:', request.error);
        resolve(null);
      };
    });
  }

  // QRハッシュ生成
  generateQRHash(qrContent) {
    let hash = 0;
    if (qrContent.length === 0) return hash;
    for (let i = 0; i < qrContent.length; i++) {
      const char = qrContent.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash;
    }
    return hash.toString();
  }

  // 重複チェック
  async checkQRDuplicate(qrContent, currentOrderNumber) {
    if (!this.db) await this.init();
    
    const qrHash = this.generateQRHash(qrContent);
    
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([this.qrDataStoreName], 'readonly');
      const store = transaction.objectStore(this.qrDataStoreName);
      const index = store.index('qrhash');
      const request = index.getAll(qrHash);
      
      request.onsuccess = () => {
        const results = request.result.filter(item => item.orderNumber !== currentOrderNumber);
        const duplicates = results.map(item => item.orderNumber);
        resolve(duplicates);
      };
      
      request.onerror = () => {
        console.error('QR重複チェックエラー:', request.error);
        resolve([]);
      };
    });
  }
}

// === グローバル初期化 ===

let unifiedDB = null;

async function initializeUnifiedDatabase() {
  try {
    console.log('🚀 統合データベースの初期化を開始します...');
    
    unifiedDB = new UnifiedDatabase();
    await unifiedDB.init();
    
    // 破壊的移行を実行
    const migrationSuccess = await unifiedDB.performDestructiveMigration();
    
    if (migrationSuccess) {
      console.log('📂 統合データベース初期化完了');
    } else {
      console.log('⚠️ 移行がキャンセルされました');
    }
    
    return unifiedDB;
  } catch (error) {
    console.error('❌ 統合データベース初期化失敗:', error);
    alert('データベースの初期化に失敗しました。\nページを再読み込みしてください。');
    return null;
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

// 統合ストレージ管理クラス（UnifiedDatabaseのラッパー）
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
    CUSTOM_LABELS: 'customLabels',
    ORDER_IMAGE_ENABLE: 'orderImageEnable',
    FONT_SECTION_COLLAPSED: 'fontSectionCollapsed'
  };

  // UnifiedDatabaseの初期化確認
  static async ensureDatabase() {
    if (!unifiedDB) {
      unifiedDB = await initializeUnifiedDatabase();
    }
    return unifiedDB;
  }

  // デフォルト設定値
  static getDefaultSettings() {
    return {
      labelyn: true,
      labelskip: 0,
      sortByPaymentDate: false,
      customLabelEnable: false,
      customLabelText: '',
      customLabelCount: 1,
      customLabels: [],
      orderImageEnable: false
    };
  }

  // 設定値の取得（非同期版）
  static async getSettingsAsync() {
    const db = await StorageManager.ensureDatabase();
    if (!db) {
      return StorageManager.getDefaultSettings();
    }

    try {
      const settings = {};
      for (const [key, storageKey] of Object.entries(StorageManager.KEYS)) {
        const value = await db.getSetting(storageKey);
        settings[storageKey] = value;
      }

      const result = {
        labelyn: settings.labelyn !== null ? settings.labelyn : true,
        labelskip: settings.labelskip !== null ? parseInt(settings.labelskip, 10) : 0,
        sortByPaymentDate: settings.sortByPaymentDate !== null ? settings.sortByPaymentDate : false,
        customLabelEnable: settings.customLabelEnable !== null ? settings.customLabelEnable : false,
        customLabelText: settings.customLabelText || '',
        customLabelCount: settings.customLabelCount !== null ? parseInt(settings.customLabelCount, 10) : 1,
        customLabels: await StorageManager.getCustomLabels(),
        orderImageEnable: settings.orderImageEnable !== null ? settings.orderImageEnable : false
      };

      return result;
    } catch (error) {
      console.error('設定取得エラー:', error);
      return StorageManager.getDefaultSettings();
    }
  }

  // 同期版設定取得（後方互換性のため）
  static getSettings() {
    // 警告を表示して非同期版の使用を促す
    console.warn('StorageManager.getSettings() は非推奨です。StorageManager.getSettingsAsync() を使用してください。');
    
    // フォールバック用のデフォルト設定
    return StorageManager.getDefaultSettings();
  }

  // 設定保存
  static async set(key, value) {
    const db = await StorageManager.ensureDatabase();
    if (db) {
      await db.setSetting(key, value);
    } else {
      localStorage.setItem(key, value);
    }
  }

  // 設定取得
  static async get(key, defaultValue = null) {
    const db = await StorageManager.ensureDatabase();
    if (db) {
      const value = await db.getSetting(key);
      return value !== null ? value : defaultValue;
    } else {
      return localStorage.getItem(key) || defaultValue;
    }
  }

  // カスタムラベル取得
  static async getCustomLabels() {
    const labelsData = await StorageManager.get(StorageManager.KEYS.CUSTOM_LABELS);
    if (!labelsData) return [];
    
    try {
      return JSON.parse(labelsData);
    } catch (e) {
      console.warn('Custom labels data parsing failed:', e);
      return [];
    }
  }

  // カスタムラベル保存
  static async setCustomLabels(labels) {
    await StorageManager.set(StorageManager.KEYS.CUSTOM_LABELS, JSON.stringify(labels));
  }

  // 注文画像取得
  static async getOrderImage(orderNumber = null) {
    const key = orderNumber ? 
      `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : 
      StorageManager.KEYS.GLOBAL_ORDER_IMAGE;
    
    const db = await StorageManager.ensureDatabase();
    if (db) {
      return await db.getImage(key);
    } else {
      return localStorage.getItem(key);
    }
  }

  // 注文画像保存
  static async setOrderImage(imageData, orderNumber = null) {
    const key = orderNumber ? 
      `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : 
      StorageManager.KEYS.GLOBAL_ORDER_IMAGE;
    
    const db = await StorageManager.ensureDatabase();
    if (db) {
      const type = orderNumber ? 'order' : 'global';
      await db.setImage(key, imageData, type, orderNumber);
    } else {
      localStorage.setItem(key, imageData);
    }
  }

  // 注文画像削除
  static async removeOrderImage(orderNumber = null) {
    const key = orderNumber ? 
      `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : 
      StorageManager.KEYS.GLOBAL_ORDER_IMAGE;
    
    const db = await StorageManager.ensureDatabase();
    if (db) {
      try {
        const transaction = db.db.transaction(['images'], 'readwrite');
        const store = transaction.objectStore('images');
        await new Promise((resolve, reject) => {
          const request = store.delete(key);
          request.onsuccess = () => resolve();
          request.onerror = () => reject(request.error);
        });
        console.log(`✅ 画像削除完了: ${key}`);
      } catch (error) {
        console.error(`❌ 画像削除エラー: ${key}`, error);
      }
    } else {
      localStorage.removeItem(key);
    }
  }

  // QR画像一括削除
  static async clearQRImages() {
    const db = await StorageManager.ensureDatabase();
    if (db) {
      // IndexedDBでのQR画像一括削除実装が必要
      console.log('QR画像一括削除');
    } else {
      Object.keys(localStorage).forEach(key => {
        const value = localStorage.getItem(key);
        if (value?.includes('qrimage')) {
          localStorage.removeItem(key);
        }
      });
    }
  }

  // 注文画像一括削除
  static async clearOrderImages() {
    const db = await StorageManager.ensureDatabase();
    if (db) {
      // IndexedDBでの注文画像一括削除実装が必要
      console.log('注文画像一括削除');
    } else {
      Object.keys(localStorage).forEach(key => {
        if (key === StorageManager.KEYS.GLOBAL_ORDER_IMAGE || 
            key.startsWith(StorageManager.KEYS.ORDER_IMAGE_PREFIX)) {
          localStorage.removeItem(key);
        }
      });
    }
  }

  // QRデータ取得
  static async getQRData(orderNumber) {
    const db = await StorageManager.ensureDatabase();
    if (db) {
      return await db.getQRData(orderNumber);
    } else {
      const data = localStorage.getItem(orderNumber);
      if (!data) return null;
      
      try {
        return JSON.parse(data);
      } catch (e) {
        console.warn('QR data parsing failed:', e);
        return null;
      }
    }
  }

  // QRデータ保存
  static async setQRData(orderNumber, qrData) {
    const db = await StorageManager.ensureDatabase();
    if (db) {
      await db.setQRData(orderNumber, qrData);
    } else {
      localStorage.setItem(orderNumber, JSON.stringify(qrData));
    }
  }

  // QR重複チェック
  static async checkQRDuplicate(qrContent, currentOrderNumber) {
    const db = await StorageManager.ensureDatabase();
    if (db) {
      return await db.checkQRDuplicate(qrContent, currentOrderNumber);
    } else {
      // localStorage版の重複チェック
      const qrHash = StorageManager.generateQRHash(qrContent);
      const duplicates = [];
      
      Object.keys(localStorage).forEach(key => {
        if (key !== currentOrderNumber) {
          const data = localStorage.getItem(key);
          if (data) {
            try {
              const parsedData = JSON.parse(data);
              if (parsedData && parsedData.qrhash === qrHash) {
                duplicates.push(key);
              }
            } catch (e) {
              // JSON以外のデータは無視
            }
          }
        }
      });
      
      return duplicates;
    }
  }

  // QRハッシュ生成
  static generateQRHash(qrContent) {
    let hash = 0;
    if (qrContent.length === 0) return hash;
    for (let i = 0; i < qrContent.length; i++) {
      const char = qrContent.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash;
    }
    return hash.toString();
  }

  // 下位互換性のための同期メソッド（非推奨）
  static remove(key) {
    console.warn(`StorageManager.remove("${key}") は非推奨です。`);
    localStorage.removeItem(key);
  }

  // UI状態管理メソッド
  static async setUIState(key, value) {
    await StorageManager.set(key, value);
  }

  static async getUIState(key, defaultValue = null) {
    return await StorageManager.get(key, defaultValue);
  }

  // フォントセクション折りたたみ状態の管理
  static async setFontSectionCollapsed(collapsed) {
    await StorageManager.setUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED, collapsed);
  }

  static async getFontSectionCollapsed() {
    const value = await StorageManager.getUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED, false);
    return value === true || value === 'true';
  }
}

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
    document.getElementById("sortByPaymentDate").checked = settings.sortByPaymentDate;
    document.getElementById("customLabelEnable").checked = settings.customLabelEnable;
    document.getElementById("orderImageEnable").checked = settings.orderImageEnable;

    toggleCustomLabelRow(settings.customLabelEnable);
    toggleOrderImageRow(settings.orderImageEnable);
  }

  // 複数のカスタムラベルを初期化
  initializeCustomLabels(settings.customLabels);

   // 画像ドロップゾーンの初期化
  const imageDropZoneElement = document.getElementById('imageDropZone');
  const imageDropZone = await createOrderImageDropZone();
  imageDropZoneElement.appendChild(imageDropZone.element);
  window.orderImageDropZone = imageDropZone;

  // フォントドロップゾーンの初期化
  initializeFontDropZone();
  
  // フォントセクションの初期状態設定
  await initializeFontSection();
  
  // カスタムフォントのCSS読み込み（非同期で実行）
  setTimeout(async () => {
    try {
      await loadCustomFontsCSS();
    } catch (error) {
      console.warn('フォントCSS読み込みエラー:', error);
    }
  }, 100); // 少し遅らせて確実にfontDBが初期化されるのを待つ

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

   // チェックボックスの状態が変更されたときにStorageManagerに保存 + 自動再処理
   document.getElementById("labelyn").addEventListener("change", async function() {
     await StorageManager.set(StorageManager.KEYS.LABEL_SETTING, this.checked);
     await autoProcessCSV(); // 設定変更時に自動再処理
   });

   document.getElementById("labelskipnum").addEventListener("change", async function() {
     await StorageManager.set(StorageManager.KEYS.LABEL_SKIP, parseInt(this.value, 10) || 0);
     await autoProcessCSV(); // 設定変更時に自動再処理
   });

   document.getElementById("sortByPaymentDate").addEventListener("change", async function() {
     await StorageManager.set(StorageManager.KEYS.SORT_BY_PAYMENT, this.checked);
     await autoProcessCSV(); // 設定変更時に自動再処理
   });

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

}, false);

async function clickstart() {
  // カスタムラベルのバリデーション
  if (!validateAndPromptCustomLabels()) {
    return; // バリデーション失敗時は処理を中断
  }
  
  clearPreviousResults();
  const config = await getConfigFromUI();
  
  Papa.parse(config.file, {
    header: true,
    skipEmptyLines: true,
    complete: async function(results) {
      await processCSVResults(results, config);
    }
  });
}

async function executeCustomLabelsOnly() {
  // カスタムラベルのバリデーション
  if (!validateAndPromptCustomLabels()) {
    return; // バリデーション失敗時は処理を中断
  }
  
  clearPreviousResults();
  const config = await getConfigFromUI();
  
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
  await processCustomLabelsOnly(config);
}

// 自動処理関数（ファイル選択時や設定変更時に呼ばれる）
async function autoProcessCSV() {
  try {
    const fileInput = document.getElementById("file");
    if (!fileInput.files || fileInput.files.length === 0) {
      console.log('ファイルが選択されていません。自動処理をスキップします。');
      clearPrintCountDisplay(); // ファイルが未選択の場合は印刷枚数をクリア
      return;
    }

    // カスタムラベルのバリデーション（エラー表示なし）
    if (!validateCustomLabelsQuiet()) {
      console.log('カスタムラベルにエラーがあります。手動実行が必要です。');
      return;
    }
    
    console.log('自動CSV処理を開始します...');
    clearPreviousResults();
    const config = await getConfigFromUI();
    
    Papa.parse(config.file, {
      header: true,
      skipEmptyLines: true,
      complete: async function(results) {
        await processCSVResults(results, config);
        console.log('自動CSV処理が完了しました。');
      }
    });
  } catch (error) {
    console.error('自動処理中にエラーが発生しました:', error);
  }
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
  await generateOrderDetails(results.data, config.labelarr);
  
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
      await generateLabels(totalLabelArray);
    }
  }
  
  // 印刷枚数の表示（複数シート対応）
  // showCSVWithCustomLabelPrintSummary(csvRowCount, totalCustomLabelCount, skipCount, requiredSheets);
  
  // ヘッダーの印刷枚数表示を更新
  updatePrintCountDisplay(csvRowCount, requiredSheets, totalCustomLabelCount);
  
  // CSV処理完了後のカスタムラベルサマリー更新（複数シート対応）
  await updateCustomLabelsSummary();
  
  // ボタンの状態を更新
  updateButtonStates();
}

async function processCustomLabelsOnly(config) {
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
      await generateLabels(labelarr);
    }
    
    // 使い切ったラベルを削除
    remainingLabels = remainingLabels.filter(label => label.count > 0);
    currentSkip = 0; // 2シート目以降はスキップなし
  }
  
  // 印刷枚数の表示（複数シート対応）
  // showMultiSheetCustomLabelPrintSummary(totalCustomLabelCount, labelskipNum, sheetsInfo);
  
  // ヘッダーの印刷枚数表示を更新（カスタムラベルのみ）
  updatePrintCountDisplay(0, sheetsInfo.length, totalCustomLabelCount);
  
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
  
  if (!displayElement) return;
  
  // 値を更新
  if (orderCountElement) orderCountElement.textContent = orderSheetCount;
  if (labelCountElement) labelCountElement.textContent = labelSheetCount;
  if (customLabelCountElement) customLabelCountElement.textContent = customLabelCount;
  
  // カスタムラベルの表示/非表示を制御
  if (customLabelItem) {
    customLabelItem.style.display = customLabelCount > 0 ? 'flex' : 'none';
  }
  
  // 全体の表示/非表示を制御
  const hasAnyCount = orderSheetCount > 0 || labelSheetCount > 0 || customLabelCount > 0;
  displayElement.style.display = hasAnyCount ? 'flex' : 'none';
  
  console.log(`印刷枚数更新: ラベル:${labelSheetCount}枚, 普通紙:${orderSheetCount}枚, カスタム:${customLabelCount}面`);
}

// 印刷枚数をクリアする関数
function clearPrintCountDisplay() {
  updatePrintCountDisplay(0, 0, 0);
}

async function generateOrderDetails(data, labelarr) {
  const tOrder = document.querySelector('#注文明細');
  
  for (let row of data) {
    const cOrder = document.importNode(tOrder.content, true);
    let orderNumber = '';
    
    // 注文情報の設定
    orderNumber = setOrderInfo(cOrder, row, labelarr);
    
    // 個別画像ドロップゾーンの作成
    await createIndividualImageDropZone(cOrder, orderNumber);
    
    // 商品項目の処理
    processProductItems(cOrder, row);
    
    // 画像表示の処理
    await displayOrderImage(cOrder, orderNumber);
    
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

async function generateLabels(labelarr) {
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
    tr.appendChild(await createLabel(label));
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
      
      // テキスト変更時に自動再処理
      await autoProcessCSV();
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
  const executeButton = document.getElementById("executeButton");
  const customLabelOnlyButton = document.getElementById("customLabelOnlyButton");
  const printButton = document.getElementById("printButton");
  
  // 固定ヘッダーのボタン要素も取得
  const executeButtonCompact = document.getElementById("executeButtonCompact");
  const customLabelOnlyButtonCompact = document.getElementById("customLabelOnlyButtonCompact");
  const printButtonCompact = document.getElementById("printButtonCompact");
  
  const customLabelEnable = document.getElementById("customLabelEnable");

  // CSV処理実行ボタンの状態
  const executeDisabled = fileInput.files.length === 0;
  if (executeButton) executeButton.disabled = executeDisabled;
  if (executeButtonCompact) executeButtonCompact.disabled = executeDisabled;

  // カスタムラベル専用実行ボタンの状態
  const hasValidCustomLabels = customLabelEnable.checked && hasCustomLabelsWithContent();
  if (customLabelOnlyButton) customLabelOnlyButton.disabled = !hasValidCustomLabels;
  if (customLabelOnlyButtonCompact) customLabelOnlyButtonCompact.disabled = !hasValidCustomLabels;

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

    // カスタムフォント選択オプション（IndexedDBベース）
    try {
      if (fontDB) {
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
        const customFonts = await fontDB.getAllFonts();

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
      // まず印刷を実行
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
      // まず印刷を実行
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
      // まず印刷を実行
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
  // FontDatabaseを初期化（まだ初期化されていない場合）
  if (!fontDB) {
    fontDB = new FontDatabase();
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
      
    } else if (rangeInfo.isPartialSpan) {
      debugLog('部分選択でspan分割処理');
      handlePartialSpanSelection(range, rangeInfo, styleProperty, styleValue, isDefault);
      
    } else if (rangeInfo.isMultiSpan) {
      debugLog('複数span要素を統合処理:', rangeInfo.multiSpans.length + '個');
      handleMultiSpanSelection(range, rangeInfo, styleProperty, styleValue, isDefault);
      
    } else {
      debugLog('新しいspan要素を作成');
      createNewSpanForSelection(range, selectedText, styleProperty, styleValue, isDefault);
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
  
  // 元のspan要素のスタイルを取得
  const originalStyle = targetSpan.getAttribute('style') || '';
  
  // 選択範囲の前後でspan要素を分割
  const beforeText = targetSpan.textContent.substring(0, targetSpan.textContent.indexOf(selectedText));
  const afterText = targetSpan.textContent.substring(targetSpan.textContent.indexOf(selectedText) + selectedText.length);
  
  // 親要素を取得
  const parent = targetSpan.parentNode;
  const nextSibling = targetSpan.nextSibling;
  
  // 元のspan要素を削除
  targetSpan.remove();
  
  // 新しい要素を作成（順番通りに）
  const elements = [];
  
  // 1. 前の部分があれば元のスタイルで作成
  if (beforeText) {
    const beforeSpan = document.createElement('span');
    beforeSpan.setAttribute('style', originalStyle);
    beforeSpan.textContent = beforeText;
    elements.push(beforeSpan);
  }
  
  // 2. 選択部分に新しいスタイルを適用
  const selectedSpan = document.createElement('span');
  const styleMap = parseStyleString(originalStyle);
  
  if (isDefault) {
    styleMap.delete(styleProperty);
  } else {
    const unit = styleProperty === 'font-size' && typeof styleValue === 'number' ? 'pt' : '';
    styleMap.set(styleProperty, styleValue + unit);
  }
  
  if (styleMap.size > 0) {
    const newStyle = Array.from(styleMap.entries())
      .map(([prop, val]) => `${prop}: ${val}`)
      .join('; ');
    selectedSpan.setAttribute('style', newStyle);
  }
  selectedSpan.textContent = selectedText;
  elements.push(selectedSpan);
  
  // 3. 後の部分があれば元のスタイルで作成
  if (afterText) {
    const afterSpan = document.createElement('span');
    afterSpan.setAttribute('style', originalStyle);
    afterSpan.textContent = afterText;
    elements.push(afterSpan);
  }
  
  // 正しい順序で全ての要素を挿入
  elements.forEach(element => {
    if (nextSibling) {
      parent.insertBefore(element, nextSibling);
    } else {
      parent.appendChild(element);
    }
  });
  
  // 新しい選択範囲を設定
  range.selectNodeContents(selectedSpan);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
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
  }
  
  // 選択範囲の内容を削除してspan要素を挿入
  range.deleteContents();
  newSpan.textContent = selectedText;
  range.insertNode(newSpan);
  
  // 新しいspan要素を選択状態にする
  range.selectNodeContents(newSpan);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
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
