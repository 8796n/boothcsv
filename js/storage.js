(function(){
  'use strict';

  function getMimeFromDataUrl(dataUrl) {
    if (!dataUrl || !dataUrl.startsWith('data:')) return 'application/octet-stream';
    const m = dataUrl.match(/^data:([^;]+)/); return m ? m[1] : 'application/octet-stream';
  }

  class UnifiedDatabase {
    constructor() {
    this.dbName = 'BoothCSVStorage';
  this.version = 7; // v6: images ストア削除, グローバル画像は settings にバイナリ格納 / v7: 残存 qrData ストア物理削除
      this.fontStoreName = 'fonts';
      this.settingsStoreName = 'settings';
      this.db = null; this.connectionLogged = false; this._migrationDone = false;
    }
    async init() {
      if (this.db) return this.db;
      return new Promise((resolve, reject) => {
        const req = indexedDB.open(this.dbName, this.version);
        req.onerror = () => reject(req.error);
        req.onupgradeneeded = (e) => this.createObjectStores(e.target.result);
        req.onsuccess = async () => {
          this.db = req.result;
          if (!this.connectionLogged) { console.log(`📂 IndexedDB "${this.dbName}" v${this.version} に接続しました`); this.connectionLogged = true; }
          try { await this.migrateLegacyImageMetadata(); } catch(e){ console.warn('画像メタ移行失敗:', e); }
          resolve(this.db);
        };
      });
    }
    createObjectStores(db) {
      const stores = [
        { name: this.fontStoreName, options: { keyPath: 'name' }, idx: [ ['createdAt','createdAt'], ['size','size'] ] },
        { name: this.settingsStoreName, options: { keyPath: 'key' }, idx: [] },
        // v6: images ストア廃止（注文画像は orders.image, グローバル画像は settings:globalOrderImageBin）
        { name: 'orders', options: { keyPath: 'orderNumber' }, idx: [ ['printedAt','printedAt'], ['createdAt','createdAt'] ] }
      ];
      // 旧 images / qrData ストアが存在する場合は削除 (バージョンアップ後も残存する可能性があるため明示削除)
      if (db.objectStoreNames.contains('images')) {
        try { db.deleteObjectStore('images'); console.log('🗑 旧 images ストア削除'); } catch(e){ console.warn('images ストア削除失敗', e); }
      }
      if (db.objectStoreNames.contains('qrData')) {
        try { db.deleteObjectStore('qrData'); console.log('🗑 旧 qrData ストア削除'); } catch(e){ console.warn('qrData ストア削除失敗', e); }
      }
      stores.forEach(s => { if (!db.objectStoreNames.contains(s.name)) { const os = db.createObjectStore(s.name, s.options); s.idx.forEach(([n,k]) => os.createIndex(n,k,{unique:false})); console.log(`🆕 ストア追加: ${s.name}`); } });
    }
    // Orders
    async saveOrder(order){ if (!this.db) await this.init(); if (order && order.orderNumber) order.orderNumber = String(order.orderNumber); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readwrite'); const st=tx.objectStore('orders'); const r=st.put(order); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
    async getAllOrders(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readonly'); const st=tx.objectStore('orders'); const r=st.getAll(); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
    async getOrder(num){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readonly'); const st=tx.objectStore('orders'); const r=st.get(num); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
    async setPrintedAt(orderNumber, printedAt){ const o=await this.getOrder(orderNumber); if(!o)return false; o.printedAt=printedAt; return await this.saveOrder(o); }
    async getPrintedOrderNumbers(){ const all=await this.getAllOrders(); return all.filter(o=>!!o.printedAt).map(o=>o.orderNumber); }
  async clearAllOrders(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readwrite'); const st=tx.objectStore('orders'); const cnt=st.count(); cnt.onsuccess=()=>{ const total=cnt.result||0; const clr=st.clear(); clr.onsuccess=()=>res(total); clr.onerror=()=>rej(clr.error); }; cnt.onerror=()=>rej(cnt.error); }); }
    // Settings
    async setSetting(key,value){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['settings'],'readwrite'); const st=tx.objectStore('settings'); const r=st.put({key,value,updatedAt:Date.now()}); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    async getSetting(key){ if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['settings'],'readonly'); const st=tx.objectStore('settings'); const r=st.get(key); r.onsuccess=()=>res(r.result? r.result.value:null); r.onerror=()=>res(null); }); }
  // v6: Images ストア削除 → 注文画像は orders.image に内包, グローバル画像は settings
    // Fonts
    async setFont(name,fontObject){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['fonts'],'readwrite'); const st=tx.objectStore('fonts'); const obj={ name, data:fontObject.data, type:fontObject.type, originalName:fontObject.originalName, size:fontObject.size, createdAt:fontObject.createdAt }; const r=st.put(obj); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    async getFont(name){ if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['fonts'],'readonly'); const st=tx.objectStore('fonts'); const r=st.get(name); r.onsuccess=()=>res(r.result); r.onerror=()=>res(null); }); }
    async getAllFonts(){ if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['fonts'],'readonly'); const st=tx.objectStore('fonts'); const r=st.getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>res([]); }); }
    async deleteFont(name){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['fonts'],'readwrite'); const st=tx.objectStore('fonts'); const r=st.delete(name); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    async clearAllFonts(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['fonts'],'readwrite'); const st=tx.objectStore('fonts'); const r=st.clear(); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
  // v5: 旧 QR 専用ストア API 削除 (OrderRepository 統合)
    // Helpers
    static sniffMime(ab){ if(!(ab instanceof ArrayBuffer)) return 'application/octet-stream'; const u8=new Uint8Array(ab.slice(0,32)); if(u8.length>=8 && u8[0]==0x89&&u8[1]==0x50&&u8[2]==0x4E&&u8[3]==0x47) return 'image/png'; if(u8.length>=3 && u8[0]==0xFF&&u8[1]==0xD8&&u8[2]==0xFF) return 'image/jpeg'; if(u8.length>=6 && u8[0]==0x47&&u8[1]==0x49&&u8[2]==0x46&&u8[3]==0x38) return 'image/gif'; if(u8.length>=12 && u8[0]==0x52&&u8[1]==0x49&&u8[2]==0x46&&u8[3]==0x46&&u8[8]==0x57&&u8[9]==0x45&&u8[10]==0x42&&u8[11]==0x50) return 'image/webp'; try{ const txt=new TextDecoder('utf-8').decode(u8); if(/^\s*<svg[\s>]/i.test(txt)) return 'image/svg+xml'; }catch{} return 'image/png'; }
    static extractDimensions(ab,mime){ if(!(ab instanceof ArrayBuffer)) return null; const u8=new Uint8Array(ab); try{ if(mime==='image/png'){ if(u8.length>=24){ const dv=new DataView(ab); const w=dv.getUint32(16); const h=dv.getUint32(20); if(w>0&&h>0) return {width:w,height:h}; } } else if(mime==='image/gif'){ if(u8.length>=10){ const w=u8[6]+(u8[7]<<8); const h=u8[8]+(u8[9]<<8); if(w>0&&h>0) return {width:w,height:h}; } } else if(mime==='image/jpeg'){ let off=2; while(off<u8.length){ if(u8[off]!==0xFF){ off++; continue;} const marker=u8[off+1]; if(marker>=0xC0&&marker<=0xCF && ![0xC4,0xC8,0xCC].includes(marker)){ const h=(u8[off+5]<<8)+u8[off+6]; const w=(u8[off+7]<<8)+u8[off+8]; if(w>0&&h>0) return {width:w,height:h}; break; } else { const len=(u8[off+2]<<8)+u8[off+3]; off+=2+len; } } } }catch{} return null; }
    async migrateLegacyImageMetadata(){ return 0; }
  }

  let unifiedDB=null;
  async function initializeUnifiedDatabase(){ try{ console.log('🚀 IndexedDB 初期化...'); unifiedDB=new UnifiedDatabase(); await unifiedDB.init(); return unifiedDB; }catch(e){ console.error('❌ IndexedDB 初期化失敗:', e); alert('お使いの環境では必要なデータベース機能(IndexedDB)を利用できません。\nブラウザ設定(プライベートモード/ストレージ無効化等)を確認してください。'); return null; } }

  class StorageManager {
  static KEYS = { LABEL_SETTING:'labelyn', LABEL_SKIP:'labelskip', SORT_BY_PAYMENT:'sortByPaymentDate', CUSTOM_LABEL_ENABLE:'customLabelEnable', CUSTOM_LABEL_TEXT:'customLabelText', CUSTOM_LABEL_COUNT:'customLabelCount', CUSTOM_LABELS:'customLabels', ORDER_IMAGE_ENABLE:'orderImageEnable', FONT_SECTION_COLLAPSED:'fontSectionCollapsed', GLOBAL_ORDER_IMAGE_BIN:'globalOrderImageBin' };
    static async ensureDatabase(){ if(!unifiedDB) unifiedDB=await initializeUnifiedDatabase(); return unifiedDB; }
    static getDefaultSettings(){ return { labelyn:true, labelskip:0, sortByPaymentDate:false, customLabelEnable:false, customLabelText:'', customLabelCount:1, customLabels:[], orderImageEnable:false }; }
    static async getSettingsAsync(){ const db=await StorageManager.ensureDatabase(); if(!db) return StorageManager.getDefaultSettings(); try{ const s={}; for(const [_,k] of Object.entries(StorageManager.KEYS)){ const v=await db.getSetting(k); s[k]=v; } return { labelyn: s.labelyn!==null ? s.labelyn : true, labelskip: s.labelskip!==null ? parseInt(s.labelskip,10):0, sortByPaymentDate: s.sortByPaymentDate!==null ? s.sortByPaymentDate:false, customLabelEnable: s.customLabelEnable!==null ? s.customLabelEnable:false, customLabelText: s.customLabelText||'', customLabelCount: s.customLabelCount!==null ? parseInt(s.customLabelCount,10):1, customLabels: await StorageManager.getCustomLabels(), orderImageEnable: s.orderImageEnable!==null ? s.orderImageEnable:false }; }catch(e){ console.error('設定取得エラー:', e); return StorageManager.getDefaultSettings(); } }
    static getSettings(){ console.warn('StorageManager.getSettings() は非推奨です。StorageManager.getSettingsAsync() を使用してください。'); return StorageManager.getDefaultSettings(); }
    static async set(key,val){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため設定を保存できません'); await db.setSetting(key,val); }
    static async get(key,def=null){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため設定を取得できません'); const v=await db.getSetting(key); return v!==null? v: def; }
    static async getCustomLabels(){ const d=await StorageManager.get(StorageManager.KEYS.CUSTOM_LABELS); if(!d) return []; try{ return JSON.parse(d);}catch{return[];} }
    static async setCustomLabels(labels){ await StorageManager.set(StorageManager.KEYS.CUSTOM_LABELS, JSON.stringify(labels)); }
  // v6: 個別注文画像 API 削除 (OrderRepository を使用) / グローバル画像バイナリ
    static async setGlobalOrderImageBinary(arrayBuffer,mimeType='image/png'){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化'); if(!(arrayBuffer instanceof ArrayBuffer)) throw new Error('ArrayBuffer 必須'); const value={ data: arrayBuffer, mimeType, updatedAt: Date.now() }; await db.setSetting(StorageManager.KEYS.GLOBAL_ORDER_IMAGE_BIN, value); }
    static async getGlobalOrderImageBinary(){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化'); const v=await db.getSetting(StorageManager.KEYS.GLOBAL_ORDER_IMAGE_BIN); if(!v||!v.data) return null; return v; }
  static async clearAllOrders(){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため注文データを削除できません'); try{ const c=await db.clearAllOrders(); console.log(`🧹 注文データクリア: ${c}件`); return c; }catch(e){ console.error('注文データ一括削除エラー:', e); throw e; } }
  // v5: 旧 QR helper 削除 (重複チェック/ハッシュは repository 側)
    static async setUIState(key,val){ await StorageManager.set(key,val); }
    static async getUIState(key,def=null){ return await StorageManager.get(key,def); }
    static async setFontSectionCollapsed(col){ await StorageManager.setUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED,col); }
    static async getFontSectionCollapsed(){ const v=await StorageManager.getUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED,false); return v===true||v==='true'; }
  static async exportAllData(){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のためバックアップできません'); const exportStore=async(name)=>new Promise((res,rej)=>{ try{ const tx=db.db.transaction([name],'readonly'); const st=tx.objectStore(name); const r=st.getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>rej(r.error);}catch(e){rej(e);} }); const [fonts,settings,orders]=await Promise.all(['fonts','settings','orders'].map(exportStore)); const encodeAB=(container, fieldPathArr)=>{ // fieldPathArr e.g. ['qr','qrimage'] or ['image','data']
    if(!container) return; let target=container; for(let i=0;i<fieldPathArr.length-1;i++){ target=target[fieldPathArr[i]]; if(!target) return; } const last=fieldPathArr[fieldPathArr.length-1]; const val=target[last]; if(val instanceof ArrayBuffer){ const u8=new Uint8Array(val); target[last]={ __type:'u8', data:Array.from(u8) }; }
  };
    fonts.forEach(f=>{ if(f && f.data instanceof ArrayBuffer){ const u8=new Uint8Array(f.data); f.data={ __type:'u8', data:Array.from(u8) }; } });
    orders.forEach(o=>{ if(o){ if(o.qr) encodeAB(o,['qr','qrimage']); if(o.image) encodeAB(o,['image','data']); } });
    settings.forEach(s=>{ if(s && s.value && s.value.data instanceof ArrayBuffer){ const u8=new Uint8Array(s.value.data); s.value.data={ __type:'u8', data:Array.from(u8) }; } });
    return { version:4, exportedAt:new Date().toISOString(), fonts, settings, orders };
  }
  static async importAllData(json,{clearExisting=true}={}){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のためリストアできません'); if(!json||typeof json!=='object') throw new Error('無効なバックアップデータ'); const { fonts=[], settings=[], orders=[] }=json; if(clearExisting){ const clear=(name)=>new Promise((res,rej)=>{ const tx=db.db.transaction([name],'readwrite'); const st=tx.objectStore(name); const r=st.clear(); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); await Promise.all(['fonts','settings','orders'].map(clear)); }
    const decodeAB=(obj, pathArr)=>{ let target=obj; for(let i=0;i<pathArr.length-1;i++){ target=target[pathArr[i]]; if(!target) return; } const last=pathArr[pathArr.length-1]; const v=target[last]; if(v && v.__type==='u8' && Array.isArray(v.data)){ const u8=new Uint8Array(v.data); target[last]=u8.buffer; } };
    fonts.forEach(f=>{ if(f && f.data && f.data.__type==='u8'){ decodeAB(f,['data']); } });
    orders.forEach(o=>{ if(o){ if(o.qr && o.qr.qrimage && o.qr.qrimage.__type==='u8') decodeAB(o,['qr','qrimage']); if(o.image && o.image.data && o.image.data.__type==='u8') decodeAB(o,['image','data']); } });
    settings.forEach(s=>{ if(s && s.value && s.value.data && s.value.data.__type==='u8'){ decodeAB(s,['value','data']); } });
    const putAll=(name,items)=>new Promise((res,rej)=>{ const tx=db.db.transaction([name],'readwrite'); const st=tx.objectStore(name); items.forEach(it=>st.put(it)); tx.oncomplete=()=>res(); tx.onerror=()=>rej(tx.error); });
    await putAll('fonts',fonts); await putAll('settings',settings); await putAll('orders',orders);
  }
  }

  window.UnifiedDatabase=UnifiedDatabase; window.StorageManager=StorageManager; window.initializeUnifiedDatabase=initializeUnifiedDatabase; Object.defineProperty(window,'unifiedDB',{ get(){ return unifiedDB; }, set(v){ unifiedDB=v; } });
})();
// end of storage.js (metadata migration enhanced)
