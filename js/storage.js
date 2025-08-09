(function(){
  'use strict';

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
  async setImage(key, imageData, mimeType = null, category = 'unknown', orderNumber = null) {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['images'], 'readwrite');
        const store = tx.objectStore('images');
        if (!(imageData instanceof ArrayBuffer)) {
          console.warn('setImage: ArrayBuffer 以外は非推奨です (無視される可能性)');
        }
        const categoryValue = category;
        const isBinary = imageData instanceof ArrayBuffer;
    const imageObject = { key, data: imageData, mimeType: mimeType || null, category: categoryValue, orderNumber, createdAt: Date.now(), isBinary };
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
            let mimeType = result.mimeType || null;
            if (!mimeType) {
              // 簡易MIME推測 (SVG判定)
              try {
                const head = new Uint8Array(result.data.slice(0, 64));
                const text = new TextDecoder('utf-8', { fatal: false }).decode(head);
                if (/^\s*<svg[\s>]/i.test(text)) {
                  mimeType = 'image/svg+xml';
                }
              } catch {}
              if (!mimeType) mimeType = 'image/png'; // デフォルト
            }
            try {
              const blob = new Blob([result.data], { type: mimeType });
              const url = URL.createObjectURL(blob);
              resolve(url);
            } catch (e) {
              console.error('getImage: Blob URL 生成失敗', e);
              resolve(null);
            }
          } else {
            resolve(result.data || null);
          }
        };
        req.onerror = () => resolve(null);
      });
    }
    async clearAllOrderImages() {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction(['images'], 'readwrite');
        const store = tx.objectStore('images');
        const countReq = store.count();
        countReq.onsuccess = () => {
          const total = countReq.result || 0;
          const clearReq = store.clear();
          clearReq.onsuccess = () => resolve(total);
          clearReq.onerror = () => reject(clearReq.error);
        };
        countReq.onerror = () => reject(countReq.error);
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
        // QR画像は ArrayBuffer のみをサポート（Base64文字列は非対応化）
        let optimizedQRImage = qrData.qrimage;
        const isBinary = qrData.qrimage instanceof ArrayBuffer;
        if (!isBinary && qrData.qrimage) {
          console.warn('QR画像はArrayBufferのみ対応です (無視される可能性)', qrData.qrimage);
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
    async clearAllQRImages() {
      if (!this.db) await this.init();
      return new Promise((resolve, reject) => {
        let deleted = 0;
        const tx = this.db.transaction(['qrData'], 'readwrite');
        const store = tx.objectStore('qrData');
        const req = store.openCursor();
        req.onsuccess = () => {
          const cursor = req.result;
          if (cursor) {
            const record = cursor.value;
            if (record && record.qrimage) {
              // 画像付きレコードを物理削除
              cursor.delete();
              deleted++;
            }
            cursor.continue();
          } else {
            resolve(deleted);
          }
        };
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
          // ここでは Base64 化せずそのまま返す（呼び出し側で Blob URL 生成）
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

  // localStorage 移行コードは完全削除（IndexedDB 専用化）
  }

  // 旧フォントDB掃除処理は不要になったため削除

  // グローバル共有インスタンス
  let unifiedDB = null;
  async function initializeUnifiedDatabase() {
    try {
      console.log('🚀 IndexedDB 初期化...');
      unifiedDB = new UnifiedDatabase();
      await unifiedDB.init();
      return unifiedDB;
    } catch (e) {
      console.error('❌ IndexedDB 初期化失敗:', e);
      alert('お使いの環境では必要なデータベース機能(IndexedDB)を利用できません。\nブラウザ設定(プライベートモード/ストレージ無効化等)を確認してください。');
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
      if (!db) throw new Error('IndexedDB 未初期化のため設定を保存できません');
      await db.setSetting(key, value);
    }
    static async get(key, defaultValue = null) {
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のため設定を取得できません');
      const v = await db.getSetting(key); return v !== null ? v : defaultValue;
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
      if (!db) throw new Error('IndexedDB 未初期化のため画像を取得できません');
      return await db.getImage(key);
    }
    static async setOrderImage(imageData, orderNumber = null, mimeType = null) {
      const key = orderNumber ? `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : StorageManager.KEYS.GLOBAL_ORDER_IMAGE;
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のため画像を保存できません');
  const category = orderNumber ? 'order' : 'global';
  await db.setImage(key, imageData, mimeType, category, orderNumber);
    }
    static async removeOrderImage(orderNumber = null) {
      const key = orderNumber ? `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : StorageManager.KEYS.GLOBAL_ORDER_IMAGE;
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のため画像を削除できません');
      try {
        const tx = db.db.transaction(['images'], 'readwrite');
        const store = tx.objectStore('images');
        await new Promise((resolve, reject) => { const req = store.delete(key); req.onsuccess = () => resolve(); req.onerror = () => reject(req.error); });
      } catch (e) { console.error('画像削除エラー:', e); }
    }

    static async clearQRImages() {
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のためQR画像を削除できません');
      try {
        const count = await db.clearAllQRImages();
        console.log(`🧹 QR画像クリア: ${count}件`);
        return count;
      } catch (e) {
        console.error('QR画像一括削除エラー:', e);
        throw e;
      }
    }
    static async clearOrderImages() {
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のため注文画像を削除できません');
      try {
        const count = await db.clearAllOrderImages();
        console.log(`🧹 注文画像クリア: ${count}件`);
        return count;
      } catch (e) {
        console.error('注文画像一括削除エラー:', e);
        throw e;
      }
    }

    static async getQRData(orderNumber) {
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のためQRデータを取得できません');
      return await db.getQRData(orderNumber);
    }
    static async setQRData(orderNumber, qrData) {
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のためQRデータを保存できません');
      await db.setQRData(orderNumber, qrData);
    }
    static async checkQRDuplicate(qrContent, currentOrderNumber) {
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のため重複チェックができません');
      return await db.checkQRDuplicate(qrContent, currentOrderNumber);
    }
    static generateQRHash(qrContent) {
      let hash = 0; if (qrContent.length === 0) return hash.toString();
      for (let i = 0; i < qrContent.length; i++) { const char = qrContent.charCodeAt(i); hash = ((hash << 5) - hash) + char; hash = hash & hash; }
      return hash.toString();
    }
  // remove は localStorage 廃止に伴い削除
    static async setUIState(key, value) { await StorageManager.set(key, value); }
    static async getUIState(key, defaultValue = null) { return await StorageManager.get(key, defaultValue); }
    static async setFontSectionCollapsed(collapsed) { await StorageManager.setUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED, collapsed); }
    static async getFontSectionCollapsed() { const v = await StorageManager.getUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED, false); return v === true || v === 'true'; }

    // --- Backup / Restore ---
    static async exportAllData() {
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のためバックアップできません');
      // 各ストアの全件を取得
      const exportStore = async (storeName) => new Promise((resolve, reject) => {
        try {
          const tx = db.db.transaction([storeName], 'readonly');
          const store = tx.objectStore(storeName);
          const req = store.getAll();
          req.onsuccess = () => resolve(req.result || []);
          req.onerror = () => reject(req.error);
        } catch (e) { reject(e); }
      });
      const [fonts, settings, images, qrData, orders] = await Promise.all([
        exportStore('fonts'),
        exportStore('settings'),
        exportStore('images'),
        exportStore('qrData'),
        exportStore('orders')
      ]);
      // ArrayBuffer を Uint8Array 数値配列へシリアライズ
      const encodeBuffer = (obj, fieldNames) => {
        fieldNames.forEach(f => {
          if (obj && obj[f] instanceof ArrayBuffer) {
            const u8 = new Uint8Array(obj[f]);
            obj[f] = { __type: 'u8', mime: obj.type || obj.qrimageType, data: Array.from(u8) };
          }
        });
      };
      images.forEach(img => encodeBuffer(img, ['data']));
      qrData.forEach(qr => encodeBuffer(qr, ['qrimage']));
      // フォント data も ArrayBuffer の可能性
      fonts.forEach(f => encodeBuffer(f, ['data']));
      const payload = { version: 1, exportedAt: new Date().toISOString(), fonts, settings, images, qrData, orders };
      return payload;
    }
    static async importAllData(json, { clearExisting = true } = {}) {
      const db = await StorageManager.ensureDatabase();
      if (!db) throw new Error('IndexedDB 未初期化のためリストアできません');
      if (!json || typeof json !== 'object') throw new Error('無効なバックアップデータ');
      const decodeBuffer = (obj, field) => {
        const v = obj[field];
        if (v && typeof v === 'object' && v.__type === 'u8' && Array.isArray(v.data)) {
          const u8 = new Uint8Array(v.data);
          obj[field] = u8.buffer;
          if (v.mime) obj[field + 'Type'] = v.mime;
        }
      };
      const { fonts = [], settings = [], images = [], qrData = [], orders = [] } = json;
      // 既存データクリア
      if (clearExisting) {
        const clearStore = (storeName) => new Promise((resolve, reject) => {
          const tx = db.db.transaction([storeName], 'readwrite');
          const store = tx.objectStore(storeName); const req = store.clear();
          req.onsuccess = () => resolve(); req.onerror = () => reject(req.error);
        });
        await Promise.all(['fonts','settings','images','qrData','orders'].map(clearStore));
      }
      // デコード
      images.forEach(img => decodeBuffer(img, 'data'));
      qrData.forEach(qr => decodeBuffer(qr, 'qrimage'));
      fonts.forEach(f => decodeBuffer(f, 'data'));
      // 挿入ヘルパ
      const putAll = (storeName, items) => new Promise((resolve, reject) => {
        const tx = db.db.transaction([storeName], 'readwrite');
        const store = tx.objectStore(storeName);
        items.forEach(item => store.put(item));
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
      await putAll('fonts', fonts);
      await putAll('settings', settings);
      await putAll('images', images);
      await putAll('qrData', qrData);
      await putAll('orders', orders);
    }
  }

  // グローバルへ公開
  window.UnifiedDatabase = UnifiedDatabase;
  window.StorageManager = StorageManager;
  window.initializeUnifiedDatabase = initializeUnifiedDatabase;
  Object.defineProperty(window, 'unifiedDB', {
    get() { return unifiedDB; },
    set(v) { unifiedDB = v; }
  });
})();
