(function(){
  'use strict';

  function getMimeFromDataUrl(dataUrl) {
    if (!dataUrl || !dataUrl.startsWith('data:')) return 'application/octet-stream';
    const m = dataUrl.match(/^data:([^;]+)/); return m ? m[1] : 'application/octet-stream';
  }

  class UnifiedDatabase {
    constructor() {
      this.dbName = 'BoothCSVStorage';
      this.version = 4; // schema version (no structural change to stores)
      this.fontStoreName = 'fonts';
      this.settingsStoreName = 'settings';
      this.imagesStoreName = 'images';
      this.qrDataStoreName = 'qrData';
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
        { name: this.imagesStoreName, options: { keyPath: 'key' }, idx: [ ['type','type'], ['orderNumber','orderNumber'], ['createdAt','createdAt'] ] },
        { name: this.qrDataStoreName, options: { keyPath: 'orderNumber' }, idx: [ ['qrhash','qrhash'], ['createdAt','createdAt'] ] },
        { name: 'orders', options: { keyPath: 'orderNumber' }, idx: [ ['printedAt','printedAt'], ['createdAt','createdAt'] ] }
      ];
      stores.forEach(s => { if (!db.objectStoreNames.contains(s.name)) { const os = db.createObjectStore(s.name, s.options); s.idx.forEach(([n,k]) => os.createIndex(n,k,{unique:false})); console.log(`🆕 ストア追加: ${s.name}`); } });
    }
    // Orders
    async saveOrder(order){ if (!this.db) await this.init(); if (order && order.orderNumber) order.orderNumber = String(order.orderNumber); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readwrite'); const st=tx.objectStore('orders'); const r=st.put(order); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
    async getAllOrders(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readonly'); const st=tx.objectStore('orders'); const r=st.getAll(); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
    async getOrder(num){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['orders'],'readonly'); const st=tx.objectStore('orders'); const r=st.get(num); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
    async setPrintedAt(orderNumber, printedAt){ const o=await this.getOrder(orderNumber); if(!o)return false; o.printedAt=printedAt; return await this.saveOrder(o); }
    async getPrintedOrderNumbers(){ const all=await this.getAllOrders(); return all.filter(o=>!!o.printedAt).map(o=>o.orderNumber); }
    // Settings
    async setSetting(key,value){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['settings'],'readwrite'); const st=tx.objectStore('settings'); const r=st.put({key,value,updatedAt:Date.now()}); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    async getSetting(key){ if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['settings'],'readonly'); const st=tx.objectStore('settings'); const r=st.get(key); r.onsuccess=()=>res(r.result? r.result.value:null); r.onerror=()=>res(null); }); }
    // Images
    async setImage(key,imageData,mimeType=null,category='unknown',orderNumber=null){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['images'],'readwrite'); const st=tx.objectStore('images'); if(!(imageData instanceof ArrayBuffer)) console.warn('setImage: ArrayBuffer 以外は非推奨です'); const isBinary=imageData instanceof ArrayBuffer; let finalMime=mimeType; let width=null,height=null; if(isBinary){ if(!finalMime){ try{ finalMime=UnifiedDatabase.sniffMime(imageData);}catch{} } try{ const dim=UnifiedDatabase.extractDimensions(imageData,finalMime); if(dim){ width=dim.width;height=dim.height; } }catch{} } const obj={ key,data:imageData,mimeType:finalMime||null,category,orderNumber,createdAt:Date.now(),isBinary,width,height }; const r=st.put(obj); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    async getImage(key){ if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['images'],'readonly'); const st=tx.objectStore('images'); const r=st.get(key); r.onsuccess=()=>{ const result=r.result; if(!result) return res(null); if(result.isBinary && result.data instanceof ArrayBuffer){ let mt=result.mimeType||UnifiedDatabase.sniffMime(result.data); try{ const blob=new Blob([result.data],{type:mt}); const url=URL.createObjectURL(blob); res(url);}catch(e){ console.error('getImage: Blob URL 生成失敗',e); res(null);} } else res(result.data||null); }; r.onerror=()=>res(null); }); }
    async clearAllOrderImages(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['images'],'readwrite'); const st=tx.objectStore('images'); const c=st.count(); c.onsuccess=()=>{ const total=c.result||0; const clr=st.clear(); clr.onsuccess=()=>res(total); clr.onerror=()=>rej(clr.error); }; c.onerror=()=>rej(c.error); }); }
    // Fonts
    async setFont(name,fontObject){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['fonts'],'readwrite'); const st=tx.objectStore('fonts'); const obj={ name, data:fontObject.data, type:fontObject.type, originalName:fontObject.originalName, size:fontObject.size, createdAt:fontObject.createdAt }; const r=st.put(obj); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    async getFont(name){ if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['fonts'],'readonly'); const st=tx.objectStore('fonts'); const r=st.get(name); r.onsuccess=()=>res(r.result); r.onerror=()=>res(null); }); }
    async getAllFonts(){ if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['fonts'],'readonly'); const st=tx.objectStore('fonts'); const r=st.getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>res([]); }); }
    async deleteFont(name){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['fonts'],'readwrite'); const st=tx.objectStore('fonts'); const r=st.delete(name); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    async clearAllFonts(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['fonts'],'readwrite'); const st=tx.objectStore('fonts'); const r=st.clear(); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    // QR
    async setQRData(orderNumber,qrData){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ const tx=this.db.transaction(['qrData'],'readwrite'); const st=tx.objectStore('qrData'); if(qrData==null){ const d=st.delete(orderNumber); d.onsuccess=()=>res(); d.onerror=()=>rej(d.error); return; } const isBinary=qrData.qrimage instanceof ArrayBuffer; if(!isBinary && qrData.qrimage) console.warn('QR画像はArrayBufferのみ対応です'); const obj={ orderNumber, receiptnum:qrData.receiptnum, receiptpassword:qrData.receiptpassword, qrimage:qrData.qrimage, qrimageType: qrData.qrimageType || 'image/png', qrhash: qrData.qrhash, createdAt: Date.now(), isBinary }; const r=st.put(obj); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
    async clearAllQRImages(){ if(!this.db) await this.init(); return new Promise((res,rej)=>{ let deleted=0; const tx=this.db.transaction(['qrData'],'readwrite'); const st=tx.objectStore('qrData'); const cur=st.openCursor(); cur.onsuccess=()=>{ const c=cur.result; if(c){ const v=c.value; if(v && v.qrimage){ c.delete(); deleted++; } c.continue(); } else res(deleted); }; cur.onerror=()=>rej(cur.error); }); }
    async getQRData(orderNumber){ if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['qrData'],'readonly'); const st=tx.objectStore('qrData'); const r=st.get(orderNumber); r.onsuccess=()=>res(r.result||null); r.onerror=()=>res(null); }); }
    generateQRHash(content){ let hash=0; if(!content) return hash.toString(); for(let i=0;i<content.length;i++){ const ch=content.charCodeAt(i); hash=((hash<<5)-hash)+ch; hash&=hash; } return hash.toString(); }
    async checkQRDuplicate(content,current){ if(!this.db) await this.init(); const qrHash=this.generateQRHash(content); return new Promise((res)=>{ const tx=this.db.transaction(['qrData'],'readonly'); const st=tx.objectStore('qrData'); const idx=st.index('qrhash'); const r=idx.getAll(qrHash); r.onsuccess=()=>{ const results=r.result.filter(item=>item.orderNumber!==current); res(results.map(i=>i.orderNumber)); }; r.onerror=()=>res([]); }); }
    // Helpers
    static sniffMime(ab){ if(!(ab instanceof ArrayBuffer)) return 'application/octet-stream'; const u8=new Uint8Array(ab.slice(0,32)); if(u8.length>=8 && u8[0]==0x89&&u8[1]==0x50&&u8[2]==0x4E&&u8[3]==0x47) return 'image/png'; if(u8.length>=3 && u8[0]==0xFF&&u8[1]==0xD8&&u8[2]==0xFF) return 'image/jpeg'; if(u8.length>=6 && u8[0]==0x47&&u8[1]==0x49&&u8[2]==0x46&&u8[3]==0x38) return 'image/gif'; if(u8.length>=12 && u8[0]==0x52&&u8[1]==0x49&&u8[2]==0x46&&u8[3]==0x46&&u8[8]==0x57&&u8[9]==0x45&&u8[10]==0x42&&u8[11]==0x50) return 'image/webp'; try{ const txt=new TextDecoder('utf-8').decode(u8); if(/^\s*<svg[\s>]/i.test(txt)) return 'image/svg+xml'; }catch{} return 'image/png'; }
    static extractDimensions(ab,mime){ if(!(ab instanceof ArrayBuffer)) return null; const u8=new Uint8Array(ab); try{ if(mime==='image/png'){ if(u8.length>=24){ const dv=new DataView(ab); const w=dv.getUint32(16); const h=dv.getUint32(20); if(w>0&&h>0) return {width:w,height:h}; } } else if(mime==='image/gif'){ if(u8.length>=10){ const w=u8[6]+(u8[7]<<8); const h=u8[8]+(u8[9]<<8); if(w>0&&h>0) return {width:w,height:h}; } } else if(mime==='image/jpeg'){ let off=2; while(off<u8.length){ if(u8[off]!==0xFF){ off++; continue;} const marker=u8[off+1]; if(marker>=0xC0&&marker<=0xCF && ![0xC4,0xC8,0xCC].includes(marker)){ const h=(u8[off+5]<<8)+u8[off+6]; const w=(u8[off+7]<<8)+u8[off+8]; if(w>0&&h>0) return {width:w,height:h}; break; } else { const len=(u8[off+2]<<8)+u8[off+3]; off+=2+len; } } } }catch{} return null; }
    async migrateLegacyImageMetadata(){ if(this._migrationDone) return 0; this._migrationDone=true; if(!this.db) await this.init(); return new Promise((res)=>{ const tx=this.db.transaction(['images'],'readwrite'); const st=tx.objectStore('images'); const cur=st.openCursor(); let updated=0; cur.onsuccess=()=>{ const c=cur.result; if(c){ const v=c.value; if(v && v.isBinary && v.data instanceof ArrayBuffer){ let need=false; if(!v.mimeType){ v.mimeType=UnifiedDatabase.sniffMime(v.data); need=true; } if(v.width==null||v.height==null){ const dim=UnifiedDatabase.extractDimensions(v.data,v.mimeType); if(dim){ v.width=dim.width; v.height=dim.height; need=true; } } if(need){ c.update(v); updated++; } } c.continue(); } else { if(updated>0) console.log(`🛠 旧画像メタデータ移行: ${updated}件 更新`); res(updated); } }; cur.onerror=()=>res(0); }); }
  }

  let unifiedDB=null;
  async function initializeUnifiedDatabase(){ try{ console.log('🚀 IndexedDB 初期化...'); unifiedDB=new UnifiedDatabase(); await unifiedDB.init(); return unifiedDB; }catch(e){ console.error('❌ IndexedDB 初期化失敗:', e); alert('お使いの環境では必要なデータベース機能(IndexedDB)を利用できません。\nブラウザ設定(プライベートモード/ストレージ無効化等)を確認してください。'); return null; } }

  class StorageManager {
    static KEYS = { ORDER_IMAGE_PREFIX:'orderImage_', GLOBAL_ORDER_IMAGE:'orderImage', LABEL_SETTING:'labelyn', LABEL_SKIP:'labelskip', SORT_BY_PAYMENT:'sortByPaymentDate', CUSTOM_LABEL_ENABLE:'customLabelEnable', CUSTOM_LABEL_TEXT:'customLabelText', CUSTOM_LABEL_COUNT:'customLabelCount', CUSTOM_LABELS:'customLabels', ORDER_IMAGE_ENABLE:'orderImageEnable', FONT_SECTION_COLLAPSED:'fontSectionCollapsed' };
    static async ensureDatabase(){ if(!unifiedDB) unifiedDB=await initializeUnifiedDatabase(); return unifiedDB; }
    static getDefaultSettings(){ return { labelyn:true, labelskip:0, sortByPaymentDate:false, customLabelEnable:false, customLabelText:'', customLabelCount:1, customLabels:[], orderImageEnable:false }; }
    static async getSettingsAsync(){ const db=await StorageManager.ensureDatabase(); if(!db) return StorageManager.getDefaultSettings(); try{ const s={}; for(const [_,k] of Object.entries(StorageManager.KEYS)){ const v=await db.getSetting(k); s[k]=v; } return { labelyn: s.labelyn!==null ? s.labelyn : true, labelskip: s.labelskip!==null ? parseInt(s.labelskip,10):0, sortByPaymentDate: s.sortByPaymentDate!==null ? s.sortByPaymentDate:false, customLabelEnable: s.customLabelEnable!==null ? s.customLabelEnable:false, customLabelText: s.customLabelText||'', customLabelCount: s.customLabelCount!==null ? parseInt(s.customLabelCount,10):1, customLabels: await StorageManager.getCustomLabels(), orderImageEnable: s.orderImageEnable!==null ? s.orderImageEnable:false }; }catch(e){ console.error('設定取得エラー:', e); return StorageManager.getDefaultSettings(); } }
    static getSettings(){ console.warn('StorageManager.getSettings() は非推奨です。StorageManager.getSettingsAsync() を使用してください。'); return StorageManager.getDefaultSettings(); }
    static async set(key,val){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため設定を保存できません'); await db.setSetting(key,val); }
    static async get(key,def=null){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため設定を取得できません'); const v=await db.getSetting(key); return v!==null? v: def; }
    static async getCustomLabels(){ const d=await StorageManager.get(StorageManager.KEYS.CUSTOM_LABELS); if(!d) return []; try{ return JSON.parse(d);}catch{return[];} }
    static async setCustomLabels(labels){ await StorageManager.set(StorageManager.KEYS.CUSTOM_LABELS, JSON.stringify(labels)); }
    static async getOrderImage(orderNumber=null){ const key=orderNumber? `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}`: StorageManager.KEYS.GLOBAL_ORDER_IMAGE; const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため画像を取得できません'); return await db.getImage(key); }
    static async setOrderImage(imageData,orderNumber=null,mimeType=null){ const key=orderNumber? `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}`: StorageManager.KEYS.GLOBAL_ORDER_IMAGE; const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため画像を保存できません'); const category=orderNumber? 'order':'global'; await db.setImage(key,imageData,mimeType,category,orderNumber); }
    static async removeOrderImage(orderNumber=null){ const key=orderNumber? `${StorageManager.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}`: StorageManager.KEYS.GLOBAL_ORDER_IMAGE; const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため画像を削除できません'); try{ const tx=db.db.transaction(['images'],'readwrite'); const st=tx.objectStore('images'); await new Promise((res,rej)=>{ const r=st.delete(key); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }catch(e){ console.error('画像削除エラー:', e);} }
    static async clearQRImages(){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のためQR画像を削除できません'); try{ const c=await db.clearAllQRImages(); console.log(`🧹 QR画像クリア: ${c}件`); return c; }catch(e){ console.error('QR画像一括削除エラー:', e); throw e; } }
    static async clearOrderImages(){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため注文画像を削除できません'); try{ const c=await db.clearAllOrderImages(); console.log(`🧹 注文画像クリア: ${c}件`); return c; }catch(e){ console.error('注文画像一括削除エラー:', e); throw e; } }
    static async getQRData(orderNumber){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のためQRデータを取得できません'); return await db.getQRData(orderNumber); }
    static async setQRData(orderNumber,qr){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のためQRデータを保存できません'); await db.setQRData(orderNumber,qr); }
    static async checkQRDuplicate(content,current){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のため重複チェックができません'); return await db.checkQRDuplicate(content,current); }
    static generateQRHash(c){ let h=0; if(c.length===0) return h.toString(); for(let i=0;i<c.length;i++){ const ch=c.charCodeAt(i); h=((h<<5)-h)+ch; h&=h;} return h.toString(); }
    static async setUIState(key,val){ await StorageManager.set(key,val); }
    static async getUIState(key,def=null){ return await StorageManager.get(key,def); }
    static async setFontSectionCollapsed(col){ await StorageManager.setUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED,col); }
    static async getFontSectionCollapsed(){ const v=await StorageManager.getUIState(StorageManager.KEYS.FONT_SECTION_COLLAPSED,false); return v===true||v==='true'; }
    static async exportAllData(){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のためバックアップできません'); const exportStore=async(name)=>new Promise((res,rej)=>{ try{ const tx=db.db.transaction([name],'readonly'); const st=tx.objectStore(name); const r=st.getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>rej(r.error);}catch(e){rej(e);} }); const [fonts,settings,images,qrData,orders]=await Promise.all(['fonts','settings','images','qrData','orders'].map(exportStore)); const encode=(obj,fields)=>{ fields.forEach(f=>{ if(obj && obj[f] instanceof ArrayBuffer){ const u8=new Uint8Array(obj[f]); obj[f]={ __type:'u8', mime: obj.mimeType||obj.type||obj.qrimageType, data:Array.from(u8)}; } }); }; images.forEach(i=>encode(i,['data'])); qrData.forEach(q=>encode(q,['qrimage'])); fonts.forEach(f=>encode(f,['data'])); return { version:2, exportedAt:new Date().toISOString(), fonts, settings, images, qrData, orders }; }
    static async importAllData(json,{clearExisting=true}={}){ const db=await StorageManager.ensureDatabase(); if(!db) throw new Error('IndexedDB 未初期化のためリストアできません'); if(!json||typeof json!=='object') throw new Error('無効なバックアップデータ'); const decode=(obj,field)=>{ const v=obj[field]; if(v&&typeof v==='object'&&v.__type==='u8'&&Array.isArray(v.data)){ const u8=new Uint8Array(v.data); obj[field]=u8.buffer; if(v.mime){ if(field==='qrimage') obj.qrimageType=v.mime; else obj.mimeType=v.mime; } } }; const { fonts=[], settings=[], images=[], qrData=[], orders=[] }=json; if(clearExisting){ const clear=(name)=>new Promise((res,rej)=>{ const tx=db.db.transaction([name],'readwrite'); const st=tx.objectStore(name); const r=st.clear(); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); await Promise.all(['fonts','settings','images','qrData','orders'].map(clear)); } images.forEach(i=>decode(i,'data')); qrData.forEach(q=>decode(q,'qrimage')); fonts.forEach(f=>decode(f,'data')); const putAll=(name,items)=>new Promise((res,rej)=>{ const tx=db.db.transaction([name],'readwrite'); const st=tx.objectStore(name); items.forEach(it=>st.put(it)); tx.oncomplete=()=>res(); tx.onerror=()=>rej(tx.error); }); await putAll('fonts',fonts); await putAll('settings',settings); await putAll('images',images); await putAll('qrData',qrData); await putAll('orders',orders); try{ await db.migrateLegacyImageMetadata(); }catch(e){ console.warn('インポート後メタ移行失敗:', e);} }
  }

  window.UnifiedDatabase=UnifiedDatabase; window.StorageManager=StorageManager; window.initializeUnifiedDatabase=initializeUnifiedDatabase; Object.defineProperty(window,'unifiedDB',{ get(){ return unifiedDB; }, set(v){ unifiedDB=v; } });
})();
// end of storage.js (metadata migration enhanced)
