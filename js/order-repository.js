(function(){
  'use strict';
  class OrderRepository {
    constructor(db){
      this.db = db; // UnifiedDatabase インスタンス
      this.cache = new Map(); // key: orderNumber -> record
      this.listeners = new Set();
      this.initialized = false;
    }
    static normalize(num){
  if(num == null) return '';
  return String(num).trim();
    }
    async init(){
      if(this.initialized) return;
      const all = await this.db.getAllOrders();
      for(const rec of all){
        if(rec && rec.orderNumber){
          const key = OrderRepository.normalize(rec.orderNumber);
          rec.orderNumber = key;
          this.cache.set(key, rec);
        }
      }
      this.initialized = true;
    }
    getAll(){ return Array.from(this.cache.values()); }
    get(orderNumber){ return this.cache.get(OrderRepository.normalize(orderNumber)) || null; }
    async bulkUpsert(csvRows){
      let changed = 0;
      for(const row of csvRows){
        // 現行実装では OrderNumberManager 廃止済みのため CSV ヘッダー名から取得
        let raw;
        if (row) {
          if (typeof OrderNumberManager !== 'undefined' && OrderNumberManager.getFromCSVRow) {
            raw = OrderNumberManager.getFromCSVRow(row);
          } else {
            // ヘッダー付きパース: 注文番号列名でアクセス
            if (typeof CONSTANTS !== 'undefined' && CONSTANTS.CSV && CONSTANTS.CSV.ORDER_NUMBER_COLUMN) {
              raw = row[CONSTANTS.CSV.ORDER_NUMBER_COLUMN];
            } else {
              raw = row[0]; // フォールバック（想定外）
            }
          }
        }
        if (DEBUG_MODE) {
          debugLog('[repo] bulkUpsert row', { raw, sampleKeys: row ? Object.keys(row).slice(0,5) : null });
        }
        const key = OrderRepository.normalize(raw);
        if(!key) continue;
        const existing = this.cache.get(key);
        if(!existing){
          const rec = { orderNumber: key, row, createdAt: new Date().toISOString(), printedAt: null };
          await this.db.saveOrder(rec);
          this.cache.set(key, rec);
          changed++;
        } else {
          // 既存の row が変わった場合のみ更新（軽量）
          if(existing.row !== row){
            existing.row = row;
            await this.db.saveOrder(existing);
          }
        }
      }
      if (DEBUG_MODE) {
        debugLog('[repo] bulkUpsert summary', { inputRows: csvRows.length, newRecords: changed, totalCache: this.cache.size });
      }
      if(changed>0) this.emit();
      return changed;
    }
  // --- QR データ統合 (メソッド名をより明確化) ---
  getOrderQRData(orderNumber){
      const rec = this.get(orderNumber);
      if(!rec) return null;
      return rec.qr || null;
    }
  async setOrderQRData(orderNumber, qrData){
      const key = OrderRepository.normalize(orderNumber);
      let rec = this.cache.get(key);
      if(!rec){
        // 存在しない注文に直接QRを設定するケースは通常無いが、必要なら作成
        rec = { orderNumber:key, row:null, createdAt:new Date().toISOString(), printedAt:null };
        this.cache.set(key, rec);
      }
      if(!qrData){
        delete rec.qr;
      } else {
        // 期待プロパティのみコピー
        const { receiptnum, receiptpassword, qrimage, qrimageType, qrhash, isBinary } = qrData;
        rec.qr = { receiptnum, receiptpassword, qrimage, qrimageType, qrhash, isBinary, updatedAt:Date.now() };
      }
      await this.db.saveOrder(rec);
      this.emit();
      return true;
    }
  async clearOrderQRData(orderNumber){ return this.setOrderQRData(orderNumber, null); }
    async checkQRDuplicate(content, current){
      if(!content) return [];
      // 旧 StorageManager.generateQRHash と同等の軽量ハッシュ
      let hash=0; for(let i=0;i<content.length;i++){ const ch=content.charCodeAt(i); hash=((hash<<5)-hash)+ch; hash&=hash; }
      const list=[];
      for(const rec of this.cache.values()){
        if(rec.orderNumber === current) continue;
        if(rec.qr && rec.qr.qrhash === String(hash)) list.push(rec.orderNumber);
      }
      return list;
    }
    // --- 画像データ統合 (1注文1画像) ---
    getOrderImage(orderNumber){ const rec=this.get(orderNumber); return rec && rec.image ? rec.image : null; }
    async setOrderImage(orderNumber, image){
      const key=OrderRepository.normalize(orderNumber);
      let rec=this.cache.get(key);
      if(!rec){ rec={ orderNumber:key, row:null, createdAt:new Date().toISOString(), printedAt:null }; this.cache.set(key, rec); }
      if(!image){ delete rec.image; }
      else {
        const { data, mimeType } = image; // data: ArrayBuffer, mimeType:string
        rec.image = { data, mimeType: mimeType || 'image/png', updatedAt: Date.now(), isBinary: data instanceof ArrayBuffer };
      }
      await this.db.saveOrder(rec); this.emit(); return true;
    }
    async clearOrderImage(orderNumber){ return this.setOrderImage(orderNumber, null); }
    async markPrinted(orderNumber, printedAt = new Date().toISOString()){
      const key = OrderRepository.normalize(orderNumber); const rec = this.cache.get(key); if(!rec) return false;
      rec.printedAt = printedAt; await this.db.saveOrder(rec); this.emit(); return true;
    }
    async clearPrinted(orderNumber){
      const key = OrderRepository.normalize(orderNumber); const rec = this.cache.get(key); if(!rec) return false;
      rec.printedAt = null; await this.db.saveOrder(rec); this.emit(); return true;
    }
    onUpdate(fn){ this.listeners.add(fn); return () => this.listeners.delete(fn); }
    emit(){ for(const fn of this.listeners) try{ fn(this.getAll()); }catch(e){ console.warn('OrderRepository listener error', e); } }
  }
  window.OrderRepository = OrderRepository;
})();
