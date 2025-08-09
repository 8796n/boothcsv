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
        const raw = (typeof OrderNumberManager !== 'undefined' && OrderNumberManager.getFromCSVRow) ? OrderNumberManager.getFromCSVRow(row) : (row && row[0]);
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
      if(changed>0) this.emit();
      return changed;
    }
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
