(function(){
  'use strict';

  // ----- 内部ユーティリティ（モジュール内限定） -----
  function toArrayBufferFromBase64(base64) {
    try {
      const base64Data = base64.includes(',') ? base64.split(',')[1] : base64;
      const binaryString = atob(base64Data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) bytes[i] = binaryString.charCodeAt(i);
      return bytes.buffer;
    } catch {
      return null;
    }
  }

  function toBase64FromArrayBuffer(buffer) {
    try {
      const bytes = new Uint8Array(buffer);
      const chunk = 8192;
      let binary = '';
      for (let i = 0; i < bytes.byteLength; i += chunk) {
        const slice = bytes.subarray(i, i + chunk);
        binary += String.fromCharCode.apply(null, slice);
      }
      return btoa(binary);
    } catch {
      return null;
    }
  }

  function isBase64Like(data) {
    if (typeof data !== 'string') return false;
    if (data.startsWith('data:')) return true;
    const base64Regex = /^[A-Za-z0-9+/]+=*$/;
    return base64Regex.test(data) && data.length % 4 === 0 && data.length > 20;
  }

  function getMimeFromDataUrl(dataUrl) {
    if (!dataUrl.startsWith('data:')) return 'application/octet-stream';
    const m = dataUrl.match(/^data:([^;]+)/);
    return m ? m[1] : 'application/octet-stream';
  }

  // ----- IndexedDB 統合クラス -----
  class UnifiedDatabase {
    constructor() {
      this.dbName = 'BoothCSVStorage';
      this.version = 4;
      this.fontStoreName = 'fonts';
      this.settingsStoreName = 'settings';
      this.imagesStoreName = 'images';
      this.qrDataStoreName = 'qrData';
      this.db = null;
      this.connectionLogged = false;
    }

    async init() {
      if (this.db) return this.db;
      return new Promise((resolve, reject) => {
        const request = indexedDB.open(this.dbName, this.version);
        request.onerror = () => reject(request.error);
        request.onsuccess = () => {
          this.db = request.result;
          if (!this.connectionLogged) {
            console.log(`📂 IndexedDB "${this.dbName}" v${this.version} に接続しました`);
            this.connectionLogged = true;
          }
          resolve(this.db);
        };
        request.onupgradeneeded = (event) => {
          this.createObjectStores(event.target.result);
        };
      });
    }

    createObjectStores(db) {
      const requiredStores = [
        { name: this.fontStoreName, options: { keyPath: 'name' }, indexes: [
          { name: 'createdAt', key: 'createdAt' },
          { name: 'size', key: 'size' }
        ] },
        { name: this.settingsStoreName, options: { keyPath: 'key' }, indexes: [] },
        { name: this.imagesStoreName, options: { keyPath: 'key' }, indexes: [
          { name: 'type', key: 'type' },
          { name: 'orderNumber', key: 'orderNumber' },
          { name: 'createdAt', key: 'createdAt' }
        ] },
        { name: this.qrDataStoreName, options: { keyPath: 'orderNumber' }, indexes: [
          { name: 'qrhash', key: 'qrhash' },
          { name: 'createdAt', key: 'createdAt' }
        ] },
        { name: 'orders', options: { keyPath: 'orderNumber' }, indexes: [
          { name: 'printedAt', key: 'printedAt' },
          { name: 'createdAt', key: 'createdAt' }
        ] }
      ];

      requiredStores.forEach(store => {
        if (!db.objectStoreNames.contains(store.name)) {
          const objStore = db.createObjectStore(store.name, store.options);
          store.indexes.forEach(idx => objStore.createIndex(idx.name, idx.key, { unique: false }));
          console.log(`🆕 ストア追加: ${store.name}`);
        }
      });
    }

    // Orders
    async saveOrder(order) {
      if (!this.db) await this.init();
      if (order && order.orderNumber) order.orderNumber = String(order.orderNumber);
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['orders'], 'readwrite');
        const store = tx.objectStore('orders');
        const req = store.put(order);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
      });
    }
    async getAllOrders() {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['orders'], 'readonly');
        const store = tx.objectStore('orders');
        const req = store.getAll();
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    async getOrder(orderNumber) {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['orders'], 'readonly');
        const store = tx.objectStore('orders');
        const req = store.get(orderNumber);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    async setPrintedAt(orderNumber, printedAt) {
      const order = await this.getOrder(orderNumber);
      if (!order) return false;
      order.printedAt = printedAt;
      return await this.saveOrder(order);
    }
    async getPrintedOrderNumbers() {
      const all = await this.getAllOrders();
      return all.filter(o => !!o.printedAt).map(o => o.orderNumber);
    }

    // Settings
    async setSetting(key, value) {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['settings'], 'readwrite');
        const store = tx.objectStore('settings');
        const req = store.put({ key, value, updatedAt: Date.now() });
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
      });
    }
    async getSetting(key) {
      if (!this.db) await this.init();
      return new Promise((resolve) => {
        const tx = this.db.transaction(['settings'], 'readonly');
        const store = tx.objectStore('settings');
        const req = store.get(key);
        req.onsuccess = () => resolve(req.result ? req.result.value : null);
        req.onerror = () => resolve(null);
      });
    }

    // Images
    async setImage(key, imageData, type = 'unknown', orderNumber = null) {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['images'], 'readwrite');
        const store = tx.objectStore('images');

        let optimizedData = imageData;
        let mimeType = type;
        let isBinary = false;
        if (isBase64Like(imageData)) {
          const ab = toArrayBufferFromBase64(imageData);
          if (ab) { optimizedData = ab; mimeType = getMimeFromDataUrl(imageData); isBinary = true; }
        } else if (imageData instanceof ArrayBuffer) {
          isBinary = true;
        }

        const imageObject = { key, data: optimizedData, type: mimeType, orderNumber, createdAt: Date.now(), isBinary };
        const req = store.put(imageObject);
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
      });
    }
    async getImage(key) {
      if (!this.db) await this.init();
      return new Promise((resolve) => {
        const tx = this.db.transaction(['images'], 'readonly');
        const store = tx.objectStore('images');
        const req = store.get(key);
        req.onsuccess = () => {
          const result = req.result;
          if (!result) return resolve(null);
          if (result.isBinary && result.data instanceof ArrayBuffer) {
            const b64 = toBase64FromArrayBuffer(result.data);
            resolve(b64 ? `data:${result.type || 'image/png'};base64,${b64}` : result.data);
          } else {
            resolve(result.data);
          }
        };
        req.onerror = () => resolve(null);
      });
    }

    // Fonts
    async setFont(fontName, fontObject) {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['fonts'], 'readwrite');
        const store = tx.objectStore('fonts');
        const obj = { name: fontName, data: fontObject.data, type: fontObject.type, originalName: fontObject.originalName, size: fontObject.size, createdAt: fontObject.createdAt };
        const req = store.put(obj);
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
      });
    }
    async getFont(fontName) {
      if (!this.db) await this.init();
      return new Promise((resolve) => {
        const tx = this.db.transaction(['fonts'], 'readonly');
        const store = tx.objectStore('fonts');
        const req = store.get(fontName);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
      });
    }
    async getAllFonts() {
      if (!this.db) await this.init();
      return new Promise((resolve) => {
        const tx = this.db.transaction(['fonts'], 'readonly');
        const store = tx.objectStore('fonts');
        const req = store.getAll();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => resolve([]);
      });
    }
    async deleteFont(fontName) {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['fonts'], 'readwrite');
        const store = tx.objectStore('fonts');
        const req = store.delete(fontName);
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
      });
    }
    async clearAllFonts() {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['fonts'], 'readwrite');
        const store = tx.objectStore('fonts');
        const req = store.clear();
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
      });
    }

    // QR
    async setQRData(orderNumber, qrData) {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['qrData'], 'readwrite');
        const store = tx.objectStore('qrData');
        if (qrData === null || qrData === undefined) {
          const delReq = store.delete(orderNumber);
          delReq.onsuccess = () => resolve();
          delReq.onerror = () => reject(delReq.error);
          return;
        }
        let optimizedQRImage = qrData.qrimage;
        let isBinary = false;
        if (qrData.qrimage && isBase64Like(qrData.qrimage)) {
          const ab = toArrayBufferFromBase64(qrData.qrimage);
          if (ab) { optimizedQRImage = ab; isBinary = true; }
        } else if (qrData.qrimage instanceof ArrayBuffer) {
          isBinary = true;
        }
        const obj = {
          orderNumber,
          receiptnum: qrData.receiptnum,
          receiptpassword: qrData.receiptpassword,
          qrimage: optimizedQRImage,
          qrimageType: qrData.qrimageType || 'image/png',
          qrhash: qrData.qrhash,
          createdAt: Date.now(),
          isBinary
        };
        const req = store.put(obj);
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
      });
    }
    async getQRData(orderNumber) {
      if (!this.db) await this.init();
      return new Promise((resolve) => {
        const tx = this.db.transaction(['qrData'], 'readonly');
        const store = tx.objectStore('qrData');
        const req = store.get(orderNumber);
        req.onsuccess = () => {
          const result = req.result;
          if (!result) return resolve(null);
          if (result.isBinary && result.qrimage instanceof ArrayBuffer) {
            const b64 = toBase64FromArrayBuffer(result.qrimage);
            result.qrimage = b64 ? `data:${result.qrimageType || 'image/png'};base64,${b64}` : result.qrimage;
          }
          resolve(result);
        };
        req.onerror = () => resolve(null);
      });
    }
    generateQRHash(qrContent) {
      let hash = 0;
      if (!qrContent || qrContent.length === 0) return hash.toString();
      for (let i = 0; i < qrContent.length; i++) {
        const char = qrContent.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash;
      }
      return hash.toString();
    }
    async checkQRDuplicate(qrContent, currentOrderNumber) {
      if (!this.db) await this.init();
      const qrHash = this.generateQRHash(qrContent);
      return new Promise((resolve) => {
        const tx = this.db.transaction(['qrData'], 'readonly');
        const store = tx.objectStore('qrData');
        const index = store.index('qrhash');
        const req = index.getAll(qrHash);
        req.onsuccess = () => {
          const results = req.result.filter(item => item.orderNumber !== currentOrderNumber);
          resolve(results.map(item => item.orderNumber));
        };
        req.onerror = () => resolve([]);
      });
    }

    // 破壊的移行
    async performDestructiveMigration() {
      try {
        console.log('🚨 破壊的移行を開始します...');
        const usage = this.analyzeLocalStorageUsage();
        if (usage.totalItems > 0) {
          console.log(`📊 削除対象: ${usage.totalItems}項目 (${usage.totalSizeMB}MB)`);
          const userConfirm = confirm(
            `🔄 システム移行のお知らせ\n\n` +
            `より高速で安定したデータ保存システムに移行します。\n` +
            `既存の設定・データ（${usage.totalItems}項目）は削除され、\n` +
            `改めて設定が必要になります。\n\n` +
            `移行を実行しますか？\n\n` +
            `※この操作は取り消せません`
          );
          if (!userConfirm) return false;
          await this.clearAllLocalStorage();
          alert(
            `✅ システム移行完了\n\n` +
            `新しいデータ保存システムが有効になりました。\n` +
            `設定・フォント・画像データを改めて登録してください。\n\n` +
            `今後はより高速で大容量のデータ保存が可能です。`
          );
        }
        console.log('✅ 破壊的移行が完了しました');
        return true;
      } catch (e) {
        console.error('❌ 破壊的移行エラー:', e);
        return false;
      }
    }
    async clearAllLocalStorage() {
      const appKeyPatterns = [
        'labelyn','labelskip','sortByPaymentDate','customLabelEnable','customLabelText','customLabelCount','customLabels','orderImageEnable','fontSectionCollapsed','orderImage','orderImage_','customFont_','migrationCompleted'
      ];
      const itemsToRemove = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (!key) continue;
        let isAppKey = appKeyPatterns.some(p => p.endsWith('_') ? key.startsWith(p) : key === p);
        if (!isAppKey) {
          try {
            const value = localStorage.getItem(key);
            const parsed = JSON.parse(value);
            if (parsed && parsed.qrhash) isAppKey = true;
          } catch {}
        }
        if (isAppKey) itemsToRemove.push(key);
      }
      itemsToRemove.forEach(key => { localStorage.removeItem(key); console.log(`🗑️ アプリデータ削除: ${key}`); });
      console.log(`🧹 アプリデータ削除完了: ${itemsToRemove.length}項目`);
    }
    analyzeLocalStorageUsage() {
      let totalSize = 0, totalItems = 0;
      const categories = { fonts: 0, settings: 0, images: 0, qrData: 0, other: 0 };
      const appKeyPatterns = [
        'labelyn','labelskip','sortByPaymentDate','customLabelEnable','customLabelText','customLabelCount','customLabels','orderImageEnable','fontSectionCollapsed','orderImage','orderImage_','customFont_','migrationCompleted'
      ];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i); if (!key) continue;
        let isAppKey = appKeyPatterns.some(p => p.endsWith('_') ? key.startsWith(p) : key === p);
        if (!isAppKey) {
          try { const value = localStorage.getItem(key); const parsed = JSON.parse(value); if (parsed && parsed.qrhash) isAppKey = true; } catch {}
        }
        if (isAppKey) {
          const value = localStorage.getItem(key);
          const size = new Blob([value || '']).size; totalSize += size; totalItems++;
          if (key.startsWith('customFont_')) categories.fonts++; else if (['labelyn','labelskip','sortByPaymentDate','customLabelEnable','orderImageEnable'].includes(key)) categories.settings++;
          else if (key.startsWith('orderImage')) categories.images++; else if (key.includes('qr') || key.includes('receipt')) categories.qrData++; else categories.other++;
        }
      }
      return { totalItems, totalSizeKB: Math.round(totalSize/1024*100)/100, totalSizeMB: Math.round(totalSize/1024/1024*100)/100, categories };
    }
  }

  // 旧フォントDB掃除
  async function cleanupOldFontDatabase() {
    try {
      await new Promise((resolve) => {
        const del = indexedDB.deleteDatabase('BoothCSVFonts');
        del.onerror = () => resolve();
        del.onsuccess = () => resolve();
        del.onblocked = () => resolve();
      });
      console.log('✅ 旧フォントDBクリーンアップ完了');
    } catch (e) {
      console.warn('旧フォントDBクリーンアップエラー:', e);
    }
  }

  // グローバル共有インスタンス
  let unifiedDB = null;
  async function initializeUnifiedDatabase() {
    try {
      console.log('🚀 統合データベースの初期化を開始します...');
      unifiedDB = new UnifiedDatabase();
      await unifiedDB.init();
      const migrated = await unifiedDB.performDestructiveMigration();
      if (migrated) await cleanupOldFontDatabase();
      return unifiedDB;
    } catch (e) {
      console.error('❌ 統合データベース初期化失敗:', e);
      alert('データベースの初期化に失敗しました。ページを再読み込みしてください。');
      return null;
    }
  }

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

    static async ensureDatabase() {
      if (!unifiedDB) unifiedDB = await initializeUnifiedDatabase();
      return unifiedDB;
    }

    static getDefaultSettings() {
      return { labelyn: true, labelskip: 0, sortByPaymentDate: false, customLabelEnable: false, customLabelText: '', customLabelCount: 1, customLabels: [], orderImageEnable: false };
    }

    static async getSettingsAsync() {
      const db = await StorageManager.ensureDatabase();
      if (!db) return StorageManager.getDefaultSettings();
      try {
        const settings = {};
        for (const [_, storageKey] of Object.entries(StorageManager.KEYS)) {
          const value = await db.getSetting(storageKey);
          settings[storageKey] = value;
        }
        return {
          labelyn: settings.labelyn !== null ? settings.labelyn : true,
          labelskip: settings.labelskip !== null ? parseInt(settings.labelskip, 10) : 0,
          sortByPaymentDate: settings.sortByPaymentDate !== null ? settings.sortByPaymentDate : false,
          customLabelEnable: settings.customLabelEnable !== null ? settings.customLabelEnable : false,
          customLabelText: settings.customLabelText || '',
          customLabelCount: settings.customLabelCount !== null ? parseInt(settings.customLabelCount, 10) : 1,
          customLabels: await StorageManager.getCustomLabels(),
          orderImageEnable: settings.orderImageEnable !== null ? settings.orderImageEnable : false
        };
      } catch (e) {
        console.error('設定取得エラー:', e);
        return StorageManager.getDefaultSettings();
      }
    }

    static getSettings() {
      console.warn('StorageManager.getSettings() は非推奨です。StorageManager.getSettingsAsync() を使用してください。');
      return StorageManager.getDefaultSettings();
    }

    static async set(key, value) {
      const db = await StorageManager.ensureDatabase();
      if (db) await db.setSetting(key, value); else localStorage.setItem(key, value);
    }
    static async get(key, defaultValue = null) {
      const db = await StorageManager.ensureDatabase();
      if (db) { const v = await db.getSetting(key); return v !== null ? v : defaultValue; }
      return localStorage.getItem(key) || defaultValue;
    }

    static async getCustomLabels() {
      const data = await StorageManager.get(StorageManager.KEYS.CUSTOM_LABELS);
      if (!data) return [];
      try { return JSON.parse(data); } catch { return []; }
    }
    static async setCustomLabels(labels) {
      await StorageManager.set(StorageManager.KEYS.CUSTOM_LABELS, JSON.stringify(labels));
    }

    static async getOrderImage(orderNumber = null) {
      const key = orderNumber ? `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : StorageManager.KEYS.GLOBAL_ORDER_IMAGE;
      const db = await StorageManager.ensureDatabase();
      if (db) return await db.getImage(key);
      return localStorage.getItem(key);
    }
    static async setOrderImage(imageData, orderNumber = null) {
      const key = orderNumber ? `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : StorageManager.KEYS.GLOBAL_ORDER_IMAGE;
      const db = await StorageManager.ensureDatabase();
      if (db) { const type = orderNumber ? 'order' : 'global'; await db.setImage(key, imageData, type, orderNumber); }
      else localStorage.setItem(key, imageData);
    }
    static async removeOrderImage(orderNumber = null) {
      const key = orderNumber ? `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : StorageManager.KEYS.GLOBAL_ORDER_IMAGE;
      const db = await StorageManager.ensureDatabase();
      if (db) {
        try {
          const tx = db.db.transaction(['images'], 'readwrite');
          const store = tx.objectStore('images');
          await new Promise((resolve, reject) => { const req = store.delete(key); req.onsuccess = () => resolve(); req.onerror = () => reject(req.error); });
        } catch (e) { console.error('画像削除エラー:', e); }
      } else {
        localStorage.removeItem(key);
      }
    }

    static async clearQRImages() {
      const db = await StorageManager.ensureDatabase();
      if (db) { console.log('QR画像一括削除'); /* TODO: implement if needed */ }
      else {
        Object.keys(localStorage).forEach(key => { const value = localStorage.getItem(key); if (value?.includes('qrimage')) localStorage.removeItem(key); });
      }
    }
    static async clearOrderImages() {
      const db = await StorageManager.ensureDatabase();
      if (db) { console.log('注文画像一括削除'); /* TODO */ }
      else {
        Object.keys(localStorage).forEach(key => { if (key === StorageManager.KEYS.GLOBAL_ORDER_IMAGE || key.startsWith(StorageManager.KEYS.ORDER_IMAGE_PREFIX)) localStorage.removeItem(key); });
      }
    }

    static async getQRData(orderNumber) {
      const db = await StorageManager.ensureDatabase();
      if (db) return await db.getQRData(orderNumber);
      const data = localStorage.getItem(orderNumber); if (!data) return null; try { return JSON.parse(data); } catch { return null; }
    }
    static async setQRData(orderNumber, qrData) {
      const db = await StorageManager.ensureDatabase();
      if (db) await db.setQRData(orderNumber, qrData); else localStorage.setItem(orderNumber, JSON.stringify(qrData));
    }
    static async checkQRDuplicate(qrContent, currentOrderNumber) {
      const db = await StorageManager.ensureDatabase();
      if (db) return await db.checkQRDuplicate(qrContent, currentOrderNumber);
      const qrHash = StorageManager.generateQRHash(qrContent); const duplicates = [];
      Object.keys(localStorage).forEach(key => {
        if (key !== currentOrderNumber) {
          const data = localStorage.getItem(key);
          if (data) { try { const parsed = JSON.parse(data); if (parsed && parsed.qrhash === qrHash) duplicates.push(key); } catch {} }
        }
      });
      return duplicates;
    }
    static generateQRHash(qrContent) {
      let hash = 0; if (qrContent.length === 0) return hash.toString();
      for (let i = 0; i < qrContent.length; i++) { const char = qrContent.charCodeAt(i); hash = ((hash << 5) - hash) + char; hash = hash & hash; }
      return hash.toString();
    }
    static remove(key) { console.warn(`StorageManager.remove("${key}") は非推奨です。`); localStorage.removeItem(key); }
    static async setUIState(key, value) { await StorageManager.set(key, value); }
    static async getUIState(key, defaultValue = null) { return await StorageManager.get(key, defaultValue); }
    static async setFontSectionCollapsed(collapsed) { await StorageManager.setUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED, collapsed); }
    static async getFontSectionCollapsed() { const v = await StorageManager.getUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED, false); return v === true || v === 'true'; }
  }

  // グローバルへ公開
  window.UnifiedDatabase = UnifiedDatabase;
  window.StorageManager = StorageManager;
  window.initializeUnifiedDatabase = initializeUnifiedDatabase;
  window.cleanupOldFontDatabase = cleanupOldFontDatabase;
  Object.defineProperty(window, 'unifiedDB', {
    get() { return unifiedDB; },
    set(v) { unifiedDB = v; }
  });
})();
