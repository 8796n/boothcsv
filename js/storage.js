(function(){
  'use strict';

  function getMimeFromDataUrl(dataUrl) {
    if (!dataUrl || !dataUrl.startsWith('data:')) return 'application/octet-stream';
    const m = dataUrl.match(/^data:([^;]+)/); return m ? m[1] : 'application/octet-stream';
  }

  function arrayBufferToBase64(arrayBuffer) {
    if (!(arrayBuffer instanceof ArrayBuffer)) return '';
    const bytes = new Uint8Array(arrayBuffer);
    const chunkSize = 0x8000;
    let binary = '';
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }

  function base64ToArrayBuffer(base64) {
    if (typeof base64 !== 'string' || !base64) return null;
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes.buffer;
  }

  class UnifiedDatabase {
    constructor() {
    this.dbName = 'BoothCSVStorage';
  this.version = 11; // v6: images ストア削除, グローバル画像は settings にバイナリ格納 / v7: 残存 qrData ストア物理削除 / v8: customLabels 独立ストア追加 / v9: 旧 v8 環境で customLabels が欠落している DB を強制アップグレードで修復 / v10: orders に shippedAt インデックス追加 / v11: productImages ストア追加
      this.fontStoreName = 'fonts';
      this.settingsStoreName = 'settings';
      this.db = null; this.connectionLogged = false; this._migrationDone = false;
      // customLabels のキー(createdAt)生成用の単調増加カウンタ
      this._lastCustomLabelTs = 0;
    }
    async init() {
      if (this.db) return this.db;
      return new Promise((resolve, reject) => {
        const req = indexedDB.open(this.dbName, this.version);
        req.onerror = () => reject(req.error);
        req.onupgradeneeded = (e) => this.createObjectStores(e.target.result, e.target.transaction);
        req.onsuccess = async () => {
          this.db = req.result;
          if (!this.connectionLogged) { console.log(`📂 IndexedDB "${this.dbName}" v${this.version} に接続しました`); this.connectionLogged = true; }
          try { await this.migrateLegacyImageMetadata(); } catch(e){ console.warn('画像メタ移行失敗:', e); }
          // Phase1: カスタムラベルの新ストアへの一度きり移行（後方互換のため settings 側は削除しない）
          try { await this.migrateCustomLabelsToStoreOnce(); } catch(e){ console.warn('customLabels 移行失敗:', e); }
          // 既存データから最後の createdAt を取得して初期化
          try { await this._syncLastCustomLabelTs(); } catch(e){ /* 初期化失敗は致命的ではない */ }
          resolve(this.db);
        };
      });
    }
    async _syncLastCustomLabelTs(){
      if(!this.db || !this.db.objectStoreNames.contains('customLabels')){ this._lastCustomLabelTs = Date.now(); return; }
      return new Promise((res)=>{
        try{
          const tx=this.db.transaction(['customLabels'],'readonly');
          const st=tx.objectStore('customLabels');
          // キーは createdAt。末尾(最大キー)から 1 件取得
          const c=st.openKeyCursor(null,'prev');
          c.onsuccess=()=>{
            const cursor=c.result;
            if(cursor){ this._lastCustomLabelTs = typeof cursor.key==='number'? cursor.key: Date.now(); }
            else { this._lastCustomLabelTs = Date.now()-1; }
            res(this._lastCustomLabelTs);
          };
          c.onerror=()=>{ this._lastCustomLabelTs = Date.now()-1; res(this._lastCustomLabelTs); };
        }catch{ this._lastCustomLabelTs = Date.now()-1; res(this._lastCustomLabelTs); }
      });
    }
    createObjectStores(db, transaction) {
      const stores = [
        { name: this.fontStoreName, options: { keyPath: 'name' }, idx: [ ['createdAt','createdAt'], ['size','size'] ] },
        { name: this.settingsStoreName, options: { keyPath: 'key' }, idx: [] },
        // v6: images ストア廃止（注文画像は orders.image, グローバル画像は settings:globalOrderImageBin）
        { name: 'orders', options: { keyPath: 'orderNumber' }, idx: [ ['printedAt','printedAt'], ['shippedAt','shippedAt'], ['createdAt','createdAt'] ] },
        // v8: customLabels 独立ストア（キーは UNIX ms タイムスタンプ）
        { name: 'customLabels', options: { keyPath: 'createdAt' }, idx: [ ['updatedAt','updatedAt'], ['order','order'] ] },
        { name: 'productImages', options: { keyPath: 'productId' }, idx: [ ['updatedAt','updatedAt'], ['productName','productName'] ] }
      ];
      // 旧 images / qrData ストアが存在する場合は削除 (バージョンアップ後も残存する可能性があるため明示削除)
      if (db.objectStoreNames.contains('images')) {
        try { db.deleteObjectStore('images'); console.log('🗑 旧 images ストア削除'); } catch(e){ console.warn('images ストア削除失敗', e); }
      }
      if (db.objectStoreNames.contains('qrData')) {
        try { db.deleteObjectStore('qrData'); console.log('🗑 旧 qrData ストア削除'); } catch(e){ console.warn('qrData ストア削除失敗', e); }
      }
      stores.forEach(s => {
        let os = null;
        if (!db.objectStoreNames.contains(s.name)) {
          os = db.createObjectStore(s.name, s.options);
          console.log(`🆕 ストア追加: ${s.name}`);
        } else if (transaction) {
          try {
            os = transaction.objectStore(s.name);
          } catch (e) {
            console.warn(`objectStore取得失敗: ${s.name}`, e);
          }
        }
        if (!os) return;
        s.idx.forEach(([n,k]) => {
          if (!os.indexNames.contains(n)) {
            os.createIndex(n,k,{unique:false});
            console.log(`🆕 インデックス追加: ${s.name}.${n}`);
          }
        });
      });
    }
    // Orders
    async saveOrder(order){ if (!this.db) await this.init(); if (order && order.orderNumber) order.orderNumber = String(order.orderNumber); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readwrite'); const st=tx.objectStore('orders'); const r=st.put(order); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
    async getAllOrders(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readonly'); const st=tx.objectStore('orders'); const r=st.getAll(); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
    async getOrder(num){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readonly'); const st=tx.objectStore('orders'); const r=st.get(num); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
    async setPrintedAt(orderNumber, printedAt){ const o=await this.getOrder(orderNumber); if(!o)return false; o.printedAt=printedAt; return await this.saveOrder(o); }
    async setShippedAt(orderNumber, shippedAt){ const o=await this.getOrder(orderNumber); if(!o)return false; o.shippedAt=shippedAt; return await this.saveOrder(o); }
    async getPrintedOrderNumbers(){ const all=await this.getAllOrders(); return all.filter(o=>!!o.printedAt).map(o=>o.orderNumber); }
  async clearAllOrders(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readwrite'); const st=tx.objectStore('orders'); const cnt=st.count(); cnt.onsuccess=()=>{ const total=cnt.result||0; const clr=st.clear(); clr.onsuccess=()=>res(total); clr.onerror=()=>rej(clr.error); }; cnt.onerror=()=>rej(cnt.error); }); }
    async deleteOrders(orderNumbers){ if(!this.db) await this.init(); if(!Array.isArray(orderNumbers)||orderNumbers.length===0) return 0; const normalized=orderNumbers.map(o=>o==null?"":String(o)).filter(Boolean); if(normalized.length===0) return 0; return new Promise((res,rej)=>{ try{ const tx=this.db.transaction(['orders'],'readwrite'); const st=tx.objectStore('orders'); let removed=0; normalized.forEach(key=>{ const req=st.delete(key); req.onsuccess=()=>{ removed++; }; req.onerror=()=>{ console.warn('orders delete failed', key, req.error); }; }); tx.oncomplete=()=>res(removed); tx.onerror=()=>rej(tx.error); }catch(e){ rej(e); } }); }
    // Settings
    async setSetting(key,value){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['settings'],'readwrite'); const st=tx.objectStore('settings'); const r=st.put({key,value,updatedAt:Date.now()}); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    async getSetting(key){ if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['settings'],'readonly'); const st=tx.objectStore('settings'); const r=st.get(key); r.onsuccess=()=>res(r.result? r.result.value:null); r.onerror=()=>res(null); }); }
  // v6: Images ストア削除 → 注文画像は orders.image に内包, グローバル画像は settings
    // Product images
    async saveProductImage(productImage){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['productImages'],'readwrite'); const st=tx.objectStore('productImages'); const r=st.put(productImage); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
    async getProductImage(productId){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['productImages'],'readonly'); const st=tx.objectStore('productImages'); const r=st.get(productId); r.onsuccess=()=>res(r.result||null); r.onerror=()=>rej(r.error); }); }
    async getAllProductImages(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['productImages'],'readonly'); const st=tx.objectStore('productImages'); const r=st.getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>rej(r.error); }); }
    async deleteProductImage(productId){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['productImages'],'readwrite'); const st=tx.objectStore('productImages'); const r=st.delete(productId); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
    async clearAllProductImages(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['productImages'],'readwrite'); const st=tx.objectStore('productImages'); const r=st.clear(); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
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

    // Phase1: customLabels を settings から独立ストアへ一度だけ移行
  async migrateCustomLabelsToStoreOnce(){
      if(!this.db) await this.init();
      try {
        // レガシー settings.customLabels を削除する共通関数（存在しなくてもOK）
        const purgeLegacySetting = () => new Promise((res)=>{
          try{
            const tx=this.db.transaction(['settings'],'readwrite');
            const st=tx.objectStore('settings');
            st.delete('customLabels');
            tx.oncomplete=()=>res(true);
            tx.onerror=()=>res(false);
          }catch{ res(false); }
        });
        // すでに移行済みのフラグ
        const migrated = await this.getSetting('customLabelsMigratedV1');
        if(migrated){ await purgeLegacySetting(); return 0; }
        // 新ストアが存在しない場合は何もしない（バージョンが古い）
        if(!this.db.objectStoreNames.contains('customLabels')){ return 0; }
        // 既にデータがある場合は二重移行しない
        const hasData = await new Promise((res,rej)=>{ const tx=this.db.transaction(['customLabels'],'readonly'); const st=tx.objectStore('customLabels'); const c=st.count(); c.onsuccess=()=>res((c.result||0)>0); c.onerror=()=>rej(c.error); });
        if(hasData){ await this.setSetting('customLabelsMigratedV1', true); await purgeLegacySetting(); return 0; }
        // 旧 settings の customLabels を取得（JSON 文字列）
        const legacy = await this.getSetting('customLabels');
        if(!legacy){ await this.setSetting('customLabelsMigratedV1', true); await purgeLegacySetting(); return 0; }
        let arr=[]; try{ arr=JSON.parse(legacy)||[]; }catch{ arr=[]; }
        if(!Array.isArray(arr) || arr.length===0){ await this.setSetting('customLabelsMigratedV1', true); await purgeLegacySetting(); return 0; }
        // 追加処理（キー衝突は考慮しない前提。念のため +index）
        const base=Date.now();
  await new Promise((res,rej)=>{
          const tx=this.db.transaction(['customLabels'],'readwrite');
          const st=tx.objectStore('customLabels');
          arr.forEach((l,i)=>{
            let ts=base+i; // createdAt を一意化
            const rec=()=>({
              createdAt: ts,
              updatedAt: ts,
              enabled: l.enabled!==false,
              count: (parseInt(l.count,10)||1),
              html: (l.html||l.text||'').trim(),
              fontSize: (l.fontSize||'12pt'),
              order: ts
            });
            // put は上書きなので衝突しにくいが、念のため +1ms 前進
            try { st.put(rec()); }
            catch(e){ try{ ts+=1; st.put(rec()); }catch(_){} }
          });
          tx.oncomplete=()=>res();
          tx.onerror=()=>rej(tx.error);
        });
        // 後方互換のため settings 側は削除しない
  await this.setSetting('customLabelsMigratedV1', true);
  // 直近の ts を更新
  this._lastCustomLabelTs = Math.max(this._lastCustomLabelTs, base + arr.length - 1);
  // レガシー設定を最後に削除
  await purgeLegacySetting();
  return arr.length;
      } catch(e){ console.warn('migrateCustomLabelsToStoreOnce error:', e); return 0; }
    }

    // Phase1: customLabels CRUD
  async addCustomLabel(label){
    if(!this.db) await this.init();
    // 新しいトランザクションを試行ごとに作る（エラーでアボート後の再使用を避ける）
    const attemptAdd = (ts, attempt=0) => new Promise((res,rej)=>{
      try{
        const tx=this.db.transaction(['customLabels'],'readwrite');
        const st=tx.objectStore('customLabels');
        const rec={
          createdAt: ts,
          updatedAt: ts,
          enabled: label?.enabled!==false,
          count: (parseInt(label?.count,10)||1),
          html: (label?.html||label?.text||'').trim(),
          fontSize: (label?.fontSize||'12pt'),
          order: (label?.order??ts)
        };
        const r=st.add(rec);
        r.onsuccess=()=>{ if(ts>this._lastCustomLabelTs) this._lastCustomLabelTs = ts; res(rec); };
        r.onerror=()=>{
          const err=r.error;
          if(err && err.name==='ConstraintError' && attempt<20){ attemptAdd(ts+1, attempt+1).then(res).catch(rej); }
          else { rej(err); }
        };
      }catch(e){ rej(e); }
    });
    const seed = Math.max(Date.now(), this._lastCustomLabelTs + 1);
    return await attemptAdd(seed, 0);
  }
    async getAllCustomLabels(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ try{ const tx=this.db.transaction(['customLabels'],'readonly'); const st=tx.objectStore('customLabels'); const r=st.getAll(); r.onsuccess=()=>{ const list=(r.result||[]).sort((a,b)=>a.createdAt-b.createdAt); res(list); }; r.onerror=()=>rej(r.error);}catch(e){rej(e);} }); }
    async getCustomLabel(ts){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ try{ const tx=this.db.transaction(['customLabels'],'readonly'); const st=tx.objectStore('customLabels'); const r=st.get(ts); r.onsuccess=()=>res(r.result||null); r.onerror=()=>rej(r.error);}catch(e){rej(e);} }); }
    async updateCustomLabel(ts, patch){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ try{ const tx=this.db.transaction(['customLabels'],'readwrite'); const st=tx.objectStore('customLabels'); const g=st.get(ts); g.onsuccess=()=>{ const cur=g.result; if(!cur){ res(null); return; } const upd=Object.assign({},cur,patch||{}, { createdAt: ts, updatedAt: Date.now() }); const p=st.put(upd); p.onsuccess=()=>res(upd); p.onerror=()=>rej(p.error); }; g.onerror=()=>rej(g.error);}catch(e){rej(e);} }); }
    async deleteCustomLabel(ts){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ try{ const tx=this.db.transaction(['customLabels'],'readwrite'); const st=tx.objectStore('customLabels'); const r=st.delete(ts); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error);}catch(e){rej(e);} }); }
    async clearAllCustomLabels(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ try{ const tx=this.db.transaction(['customLabels'],'readwrite'); const st=tx.objectStore('customLabels'); const r=st.clear(); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error);}catch(e){rej(e);} }); }
  async saveAllCustomLabels(labels){
    if(!this.db) await this.init();
    if(!Array.isArray(labels)) return [];
    // 事前にタイムスタンプを単調増加で正規化して衝突を回避
    const now = Date.now();
    let prevTs = Math.max(this._lastCustomLabelTs, now - 1);
    const normalized = labels.map((l,i)=>{
      const baseTs = (l && typeof l.createdAt==='number') ? l.createdAt : (now + i);
      const ts = Math.max(baseTs, prevTs + 1);
      prevTs = ts;
      return {
        createdAt: ts,
        updatedAt: Date.now(),
        enabled: l && l.enabled!==false,
        count: (parseInt(l && l.count,10)||1),
        html: ((l && (l.html||l.text))||'').trim(),
        fontSize: (l && l.fontSize) || '12pt',
        order: i
      };
    });
    return new Promise((res,rej)=>{
      try{
        const tx=this.db.transaction(['customLabels'],'readwrite');
        const st=tx.objectStore('customLabels');
        const clr=st.clear();
        clr.onsuccess=()=>{
          normalized.forEach(rec=>{ st.add(rec); });
        };
        clr.onerror=()=>rej(clr.error);
        tx.oncomplete=()=>{ if(normalized.length){ this._lastCustomLabelTs = Math.max(this._lastCustomLabelTs, normalized[normalized.length-1].createdAt); } res(true); };
        tx.onerror=()=>rej(tx.error);
      }catch(e){ rej(e); }
    });
  }
  }

  let unifiedDB=null;
  async function initializeUnifiedDatabase(){ try{ console.log('🚀 IndexedDB 初期化...'); unifiedDB=new UnifiedDatabase(); await unifiedDB.init(); return unifiedDB; }catch(e){ console.error('❌ IndexedDB 初期化失敗:', e); alert('お使いの環境では必要なデータベース機能(IndexedDB)を利用できません。\nブラウザ設定(プライベートモード/ストレージ無効化等)を確認してください。'); return null; } }

  class StorageManager {
  static KEYS = { LABEL_SETTING:'labelyn', LABEL_SKIP:'labelskip', SORT_BY_PAYMENT:'sortByPaymentDate', CUSTOM_LABEL_ENABLE:'customLabelEnable', CUSTOM_LABEL_TEXT:'customLabelText', CUSTOM_LABEL_COUNT:'customLabelCount', CUSTOM_LABELS:'customLabels', ORDER_IMAGE_ENABLE:'orderImageEnable', FONT_SECTION_COLLAPSED:'fontSectionCollapsed', GLOBAL_ORDER_IMAGE_BIN:'globalOrderImageBin', CUSTOM_LABELS_HELP_OPEN:'customLabelsHelpOpen', SIDEBAR_DOCKED:'sidebarDocked', SHIPPING_MESSAGE_TEMPLATE:'shippingMessageTemplate', YAMATO_PACKAGE_SIZE_ID:'yamatoPackageSizeId', YAMATO_CODE_TYPE:'yamatoCodeType', YAMATO_DESCRIPTION:'yamatoDescription', YAMATO_DESCRIPTION_HISTORY:'yamatoDescriptionHistory', YAMATO_INCLUDE_ORDER_NUMBER:'yamatoIncludeOrderNumber', YAMATO_HANDLING_CODES:'yamatoHandlingCodes' };
    static async ensureDatabase(){ if(!unifiedDB) unifiedDB=await initializeUnifiedDatabase(); return unifiedDB; }
    static getDefaultSettings(){ return { labelyn:true, labelskip:0, sortByPaymentDate:false, customLabelEnable:false, customLabelText:'', customLabelCount:1, customLabels:[], orderImageEnable:false, shippingMessageTemplate:'', yamatoPackageSizeId:'16', yamatoCodeType:'0', yamatoDescription:'', yamatoDescriptionHistory:[], yamatoIncludeOrderNumber:true, yamatoHandlingCodes:[] }; }
  static async getSettingsAsync(){ const db=await StorageManager.ensureDatabase(); if(!db) return StorageManager.getDefaultSettings(); try{ const s={}; for(const [_,k] of Object.entries(StorageManager.KEYS)){ const v=await db.getSetting(k); s[k]=v; } return { labelyn: s.labelyn!==null ? s.labelyn : true, labelskip: s.labelskip!==null ? parseInt(s.labelskip,10):0, sortByPaymentDate: s.sortByPaymentDate!==null ? s.sortByPaymentDate:false, customLabelEnable: s.customLabelEnable!==null ? s.customLabelEnable:false, customLabelText: s.customLabelText||'', customLabelCount: s.customLabelCount!==null ? parseInt(s.customLabelCount,10):1, customLabels: await StorageManager.getCustomLabels(), orderImageEnable: s.orderImageEnable!==null ? s.orderImageEnable:false, shippingMessageTemplate: typeof s.shippingMessageTemplate==='string'? s.shippingMessageTemplate : '', yamatoPackageSizeId: typeof s.yamatoPackageSizeId==='string' && s.yamatoPackageSizeId ? s.yamatoPackageSizeId : '16', yamatoCodeType: typeof s.yamatoCodeType==='string' && s.yamatoCodeType ? s.yamatoCodeType : '0', yamatoDescription: typeof s.yamatoDescription==='string' ? s.yamatoDescription : '', yamatoDescriptionHistory: Array.isArray(s.yamatoDescriptionHistory) ? s.yamatoDescriptionHistory : [], yamatoIncludeOrderNumber: s.yamatoIncludeOrderNumber!==null ? !!s.yamatoIncludeOrderNumber : true, yamatoHandlingCodes: Array.isArray(s.yamatoHandlingCodes) ? s.yamatoHandlingCodes : [] }; }catch(e){ console.error('設定取得エラー:', e); return StorageManager.getDefaultSettings(); } }
    static getSettings(){ console.warn('StorageManager.getSettings() は非推奨です。StorageManager.getSettingsAsync() を使用してください。'); return StorageManager.getDefaultSettings(); }
    static async set(key,val){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため設定を保存できません'); await db.setSetting(key,val); }
    static async get(key,def=null){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため設定を取得できません'); const v=await db.getSetting(key); return v!==null? v: def; }
  // Phase1: 新ストア customLabels に統合
  static async getCustomLabels(){ const db=await StorageManager.ensureDatabase(); if(!db) return []; try{ const list=await db.getAllCustomLabels(); return (list||[]).map(x=>({ text:x.html, html:x.html, count:x.count, fontSize:x.fontSize, enabled:x.enabled, createdAt:x.createdAt, updatedAt:x.updatedAt })); }catch(e){ console.warn('getCustomLabels error, fallback to settings JSON', e); const d=await StorageManager.get(StorageManager.KEYS.CUSTOM_LABELS); if(!d) return []; try{ return JSON.parse(d);}catch{return[];} }
  }
  static async setCustomLabels(labels){ const db=await StorageManager.ensureDatabase(); if(!db) return; try{ await db.saveAllCustomLabels(Array.isArray(labels)? labels: []); }catch(e){ console.error('setCustomLabels error:', e); }
  }
  static async addCustomLabel(label){ const db=await StorageManager.ensureDatabase(); if(!db) return null; try{ return await db.addCustomLabel(label); }catch(e){ console.error('addCustomLabel error:', e); return null; } }
  static async updateCustomLabel(createdAt, patch){ const db=await StorageManager.ensureDatabase(); if(!db) return null; try{ return await db.updateCustomLabel(createdAt, patch); }catch(e){ console.error('updateCustomLabel error:', e); return null; } }
  static async deleteCustomLabel(createdAt){ const db=await StorageManager.ensureDatabase(); if(!db) return false; try{ await db.deleteCustomLabel(createdAt); return true; }catch(e){ console.error('deleteCustomLabel error:', e); return false; } }
  static async clearAllCustomLabels(){ const db=await StorageManager.ensureDatabase(); if(!db) return; try{ await db.clearAllCustomLabels(); }catch(e){ console.error('clearAllCustomLabels error:', e); }
  }
  // v6: 個別注文画像 API 削除 (OrderRepository を使用) / グローバル画像バイナリ
    static async setGlobalOrderImageBinary(arrayBuffer,mimeType='image/png'){
      const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化');
      if(!(arrayBuffer instanceof ArrayBuffer)) throw new Error('ArrayBuffer 必須');
      // 0バイトは保存せずクリア扱い
      if(arrayBuffer.byteLength===0){ await db.setSetting(StorageManager.KEYS.GLOBAL_ORDER_IMAGE_BIN, null); return; }
      const value={ data: arrayBuffer, mimeType, updatedAt: Date.now() };
      await db.setSetting(StorageManager.KEYS.GLOBAL_ORDER_IMAGE_BIN, value);
    }
    static async clearGlobalOrderImageBinary(){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化'); await db.setSetting(StorageManager.KEYS.GLOBAL_ORDER_IMAGE_BIN, null); }
    static async getGlobalOrderImageBinary(){
      const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化');
      const v=await db.getSetting(StorageManager.KEYS.GLOBAL_ORDER_IMAGE_BIN);
      // null または data フィールドが無い／0バイトは未設定扱い
      if(!v||!v.data||!(v.data instanceof ArrayBuffer)||v.data.byteLength===0) return null;
      return v;
    }
    static async setProductOrderImage(productId, image){
      const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化');
      const normalizedId = productId == null ? '' : String(productId).trim();
      if(!normalizedId) throw new Error('商品IDが空です');
      if(!image || !(image.data instanceof ArrayBuffer)) throw new Error('画像データが不正です');
      const record = {
        productId: normalizedId,
        productName: typeof image.productName === 'string' ? image.productName : '',
        data: image.data,
        mimeType: image.mimeType || 'image/png',
        updatedAt: Date.now()
      };
      await db.saveProductImage(record);
      return record;
    }
    static async getProductOrderImage(productId){
      const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化');
      const normalizedId = productId == null ? '' : String(productId).trim();
      if(!normalizedId) return null;
      const rec = await db.getProductImage(normalizedId);
      if(!rec || !(rec.data instanceof ArrayBuffer) || rec.data.byteLength===0) return null;
      return rec;
    }
    static async getAllProductOrderImages(){
      const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化');
      const list = await db.getAllProductImages();
      return (list || []).filter(rec => rec && rec.productId && rec.data instanceof ArrayBuffer && rec.data.byteLength > 0);
    }
    static async clearProductOrderImage(productId){
      const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化');
      const normalizedId = productId == null ? '' : String(productId).trim();
      if(!normalizedId) return false;
      return await db.deleteProductImage(normalizedId);
    }
  static async clearAllOrders(){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため注文データを削除できません'); try{ const c=await db.clearAllOrders(); console.log(`🧹 注文データクリア: ${c}件`); return c; }catch(e){ console.error('注文データ一括削除エラー:', e); throw e; } }
  static async deleteOrders(orderNumbers){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため注文データを削除できません'); try{ return await db.deleteOrders(orderNumbers); }catch(e){ console.error('注文データ削除エラー:', e); throw e; } }
  // v5: 旧 QR helper 削除 (重複チェック/ハッシュは repository 側)
    static async setUIState(key,val){ await StorageManager.set(key,val); }
    static async getUIState(key,def=null){ return await StorageManager.get(key,def); }
    static async setFontSectionCollapsed(col){ await StorageManager.setUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED,col); }
    static async getFontSectionCollapsed(){ const v=await StorageManager.getUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED,false); return v===true||v==='true'; }
  static async setCustomLabelsHelpOpen(isOpen){ await StorageManager.setUIState(StorageManager.KEYS.CUSTOM_LABELS_HELP_OPEN, !!isOpen); }
  static async getCustomLabelsHelpOpen(){ const v=await StorageManager.getUIState(StorageManager.KEYS.CUSTOM_LABELS_HELP_OPEN,false); return v===true||v==='true'; }
    static async setSidebarDocked(isDocked){ await StorageManager.setUIState(StorageManager.KEYS.SIDEBAR_DOCKED, !!isDocked); }
    static async getSidebarDocked(){ const v=await StorageManager.getUIState(StorageManager.KEYS.SIDEBAR_DOCKED,null); if(v===null||typeof v==='undefined') return null; return v===true||v==='true'||v===1||v==='1'; }
  static async exportAllData(){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のためバックアップできません'); const exportStore=async(name)=>new Promise((res,rej)=>{ try{ const tx=db.db.transaction([name],'readonly'); const st=tx.objectStore(name); const r=st.getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>rej(r.error);}catch(e){rej(e);} }); const [fonts,settings,orders,customLabels,productImages]=await Promise.all(['fonts','settings','orders','customLabels','productImages'].map(exportStore)); const encodeAB=(container, fieldPathArr)=>{ // fieldPathArr e.g. ['qr','qrimage'] or ['image','data']
    if(!container) return; let target=container; for(let i=0;i<fieldPathArr.length-1;i++){ target=target[fieldPathArr[i]]; if(!target) return; } const last=fieldPathArr[fieldPathArr.length-1]; const val=target[last]; if(val instanceof ArrayBuffer){ target[last]={ __type:'base64', data:arrayBufferToBase64(val) }; }
  };
    fonts.forEach(f=>{ if(f && f.data instanceof ArrayBuffer){ f.data={ __type:'base64', data:arrayBufferToBase64(f.data) }; } });
    orders.forEach(o=>{ if(o){ if(o.qr) encodeAB(o,['qr','qrimage']); if(o.image) encodeAB(o,['image','data']); } });
    productImages.forEach(p=>{ if(p && p.data instanceof ArrayBuffer){ p.data={ __type:'base64', data:arrayBufferToBase64(p.data) }; } });
    settings.forEach(s=>{ if(s && s.value && s.value.data instanceof ArrayBuffer){ s.value.data={ __type:'base64', data:arrayBufferToBase64(s.value.data) }; } });
    return { version:7, exportedAt:new Date().toISOString(), fonts, settings, orders, customLabels, productImages };
  }
  static async importAllData(json,{clearExisting=true}={}){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のためリストアできません'); if(!json||typeof json!=='object') throw new Error('無効なバックアップデータ'); const { fonts=[], settings=[], orders=[], customLabels=[], productImages=[] }=json; if(clearExisting){ const clear=(name)=>new Promise((res,rej)=>{ const tx=db.db.transaction([name],'readwrite'); const st=tx.objectStore(name); const r=st.clear(); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); await Promise.all(['fonts','settings','orders','customLabels','productImages'].map(clear)); }
    const decodeAB=(obj, pathArr)=>{ let target=obj; for(let i=0;i<pathArr.length-1;i++){ target=target[pathArr[i]]; if(!target) return; } const last=pathArr[pathArr.length-1]; const v=target[last]; if(v && v.__type==='u8' && Array.isArray(v.data)){ const u8=new Uint8Array(v.data); target[last]=u8.buffer; return; } if(v && v.__type==='base64' && typeof v.data==='string'){ target[last]=base64ToArrayBuffer(v.data); } };
    fonts.forEach(f=>{ if(f && f.data && (f.data.__type==='u8' || f.data.__type==='base64')){ decodeAB(f,['data']); } });
    orders.forEach(o=>{ if(o){ if(o.qr && o.qr.qrimage && (o.qr.qrimage.__type==='u8' || o.qr.qrimage.__type==='base64')) decodeAB(o,['qr','qrimage']); if(o.image && o.image.data && (o.image.data.__type==='u8' || o.image.data.__type==='base64')) decodeAB(o,['image','data']); } });
    productImages.forEach(p=>{ if(p && p.data && (p.data.__type==='u8' || p.data.__type==='base64')){ decodeAB(p,['data']); } });
    settings.forEach(s=>{ if(s && s.value && s.value.data && (s.value.data.__type==='u8' || s.value.data.__type==='base64')){ decodeAB(s,['value','data']); } });
  const putAll=(name,items)=>new Promise((res,rej)=>{ const tx=db.db.transaction([name],'readwrite'); const st=tx.objectStore(name); items.forEach(it=>st.put(it)); tx.oncomplete=()=>res(); tx.onerror=()=>rej(tx.error); });
  await putAll('fonts',fonts); await putAll('settings',settings); await putAll('orders',orders); if(db.db.objectStoreNames.contains('customLabels')){ await putAll('customLabels',customLabels); } if(db.db.objectStoreNames.contains('productImages')){ await putAll('productImages',productImages); }
  }
  }

  window.UnifiedDatabase=UnifiedDatabase; window.StorageManager=StorageManager; window.initializeUnifiedDatabase=initializeUnifiedDatabase; Object.defineProperty(window,'unifiedDB',{ get(){ return unifiedDB; }, set(v){ unifiedDB=v; } });
})();
// end of storage.js (metadata migration enhanced)
