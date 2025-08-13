// custom-labels-font.js
// カスタムラベルエディタ & フォント管理モジュール（暫定抽出版）
// NOTE: 現状 boothcsv.js にも同等実装が残っているため重複。次段階で本体側の重複除去を行う。
// 先に分離構造だけ用意し、以後はこちらを正とする。

(function(){
  if(window.CustomLabelFont){ return; }

  // 依存（遅延評価）: CONSTANTS / StorageManager / debugLog などは boothcsv.js 側で定義
  function dlog(...a){ if(typeof debugLog === 'function') debugLog('font', ...a); }

  // ===================== フォント管理 =====================
  let fontManager = null;
  class FontManager {
    constructor(unifiedDB){ this.unifiedDB = unifiedDB; }
    async saveFont(fontName, fontData, metadata = {}) {
      const fontObject = { name: fontName, data: fontData, type: metadata.type || 'font/ttf', originalName: metadata.originalName || fontName, size: fontData.byteLength || fontData.length, createdAt: Date.now() };
      await this.unifiedDB.setFont(fontName, fontObject); return fontObject;
    }
    async getFont(n){ return await this.unifiedDB.getFont(n); }
    async getAllFonts(){ const arr = await this.unifiedDB.getAllFonts(); const map={}; arr.forEach(f=>{ if(f&&f.name){ map[f.name]={ data:f.data, metadata:{ type:f.type, originalName:f.originalName, size:f.size, createdAt:f.createdAt } }; }}); return map; }
    async deleteFont(n){ await this.unifiedDB.deleteFont(n); }
    async clearAllFonts(){ await this.unifiedDB.clearAllFonts(); }
  }
  async function initializeFontManager(){
    if(fontManager) return fontManager;
    try {
      if(!window.unifiedDB && window.initializeUnifiedDatabase){ await window.initializeUnifiedDatabase(); }
      if(!window.unifiedDB){ throw new Error('unifiedDB 未初期化'); }
      fontManager = new FontManager(window.unifiedDB);
      // 互換: 既存コードが参照する可能性
      window.fontManager = fontManager;
      dlog('FontManager 初期化完了');
      return fontManager;
    } catch(e){ console.error('FontManager 初期化失敗', e); return null; }
  }

  // ========== UI ==========
  async function loadCustomFontsCSS(){
    if(!fontManager) return;
    try {
      const fonts = await fontManager.getAllFonts();
      if(typeof FontFace === 'undefined'){ console.error('FontFace API 未対応'); return; }
      // 旧 style クリア
      const old = document.getElementById('custom-fonts-style'); if(old) old.remove();
      const promises=[];
      for(const [name,data] of Object.entries(fonts)){
        try { const face = new FontFace(name, data.data, { display:'swap' }); promises.push(face.load().then(f=>document.fonts.add(f))); }
        catch(e){ console.warn('FontFace 初期化失敗', name, e); }
      }
      await Promise.all(promises);
      dlog('フォントロード完了', Object.keys(fonts));
    } catch(e){ console.error('フォントCSS生成エラー', e); }
  }

  async function updateFontList(){
    const el = document.getElementById('fontList'); if(!el) return;
    try {
      if(!fontManager){ el.innerHTML='<div class="font-list-placeholder">フォントマネージャ初期化中...</div>'; return; }
      const fonts = await fontManager.getAllFonts();
      const entries = Object.entries(fonts);
      if(entries.length===0){ el.innerHTML='<div class="font-list-placeholder">フォントファイルをアップロードしてください</div>'; return; }
      el.textContent='';
      const tpl = document.getElementById('fontListItemTemplate');
      entries.forEach(([fontName,fontData])=>{
        const meta = fontData.metadata||{};
        const node = tpl ? tpl.content.firstElementChild.cloneNode(true) : document.createElement('div');
        node.dataset.fontName = fontName;
        const nameEl = node.querySelector('.font-name') || node;
        nameEl.textContent = fontName; nameEl.style.fontFamily=`'${fontName}', sans-serif`;
        const metaEl = node.querySelector('.font-meta');
        if(metaEl){ const sizeMB = ( (fontData.data.byteLength||0) / 1024 / 1024).toFixed(2); metaEl.innerHTML = `${escapeHTML(meta.originalName||fontName)} <span>• ${new Date(meta.createdAt||Date.now()).toLocaleDateString()} • ${sizeMB}MB</span>`; }
        const btn = node.querySelector('.font-remove-btn');
        if(btn && !btn._bound){ btn.addEventListener('click', () => removeFontFromList(fontName)); btn._bound=true; }
        el.appendChild(node);
      });
      adjustFontSectionHeight();
    } catch(e){ console.error('フォントリスト更新失敗', e); el.innerHTML='<div class="font-list-error">読み込み失敗</div>'; }
  }

  async function handleFontFile(file){
    try {
      if(!file) return; if(!/\.(ttf|otf|woff2?|)$/i.test(file.name)){ alert('対応していないフォント形式です'); return; }
      if(!fontManager) await initializeFontManager(); if(!fontManager) return;
      showFontUploadProgress(true);
      const arrayBuffer = await file.arrayBuffer();
      const baseName = file.name.replace(/\.[^/.]+$/, '');
      const existing = await fontManager.getFont(baseName);
      if(existing && !confirm(`フォント "${baseName}" は既に存在します。上書きしますか？`)){ showFontUploadProgress(false); return; }
      await fontManager.saveFont(baseName, arrayBuffer, { type: file.type || 'font/ttf', originalName: file.name });
      await loadCustomFontsCSS();
      await updateFontList();
      showSuccessMessage(`フォント "${baseName}" をアップロードしました`);
    } catch(e){ console.error('フォント処理エラー', e); alert('フォント処理中にエラー: '+ e.message); }
    finally { showFontUploadProgress(false); }
  }

  function initializeFontDropZone(){
    const dzWrapper = document.getElementById('fontDropZone'); if(!dzWrapper) return;
    const zone = document.createElement('div');
  zone.className='dropzone-card font-drop-zone';
  zone.innerHTML='<p class="dz-message">フォントファイルをここにドロップ<br><small>TTF / OTF / WOFF / WOFF2 対応</small></p>';
  zone.addEventListener('dragover',e=>{e.preventDefault(); zone.classList.add('dragover');});
  zone.addEventListener('dragleave',e=>{e.preventDefault(); zone.classList.remove('dragover');});
  zone.addEventListener('drop',e=>{e.preventDefault(); zone.classList.remove('dragover'); const files=[...e.dataTransfer.files]; files.forEach(f=>handleFontFile(f)); });
    zone.addEventListener('click',()=>{ const inp=document.createElement('input'); inp.type='file'; inp.accept='.ttf,.otf,.woff,.woff2'; inp.multiple=true; inp.addEventListener('change',()=>{ [...inp.files].forEach(f=>handleFontFile(f)); }); inp.click(); });
    dzWrapper.appendChild(zone);
  }

  async function removeFontFromList(fontName){
    try { if(!fontManager) await initializeFontManager(); if(!fontManager) return; const f= await fontManager.getFont(fontName); if(!f){ alert('フォントが見つかりません'); return; } if(!confirm(`フォント "${fontName}" を削除しますか？`)) return; await fontManager.deleteFont(fontName); await updateFontList(); await loadCustomFontsCSS(); showSuccessMessage(`フォント "${fontName}" を削除しました`); } catch(e){ console.error('削除失敗', e); alert('削除中にエラー'); }
  }

  function showSuccessMessage(msg){
    const old=document.querySelector('.success-notification'); if(old) old.remove();
    const div=document.createElement('div'); div.className='success-notification'; div.textContent=msg; div.style.cssText='position:fixed;top:20px;right:20px;background:#4CAF50;color:#fff;padding:12px 24px;border-radius:4px;z-index:10001;font-family:sans-serif;font-size:14px;box-shadow:0 2px 8px rgba(0,0,0,.2);';
    document.body.appendChild(div); setTimeout(()=>{ div.style.transition='opacity .3s'; div.style.opacity='0'; setTimeout(()=>div.remove(),300); },3000);
  }
  function showFontUploadProgress(show){
    let p=document.getElementById('font-upload-progress');
    if(show){ if(!p){ p=document.createElement('div'); p.id='font-upload-progress'; p.style.cssText='position:fixed;top:50%;left:50%;transform:translate(-50%, -50%);background:rgba(0,0,0,.8);color:#fff;padding:20px;border-radius:8px;z-index:10002;font-family:sans-serif;text-align:center;'; p.innerHTML='<div style="margin-bottom:10px;">フォントをアップロード中...</div><div style="width:200px;height:4px;background:#333;border-radius:2px;"><div style="width:100%;height:100%;background:#4CAF50;animation:pulse 1s infinite;"></div></div>'; document.body.appendChild(p);} }
    else if(p){ p.remove(); }
  }
  if(!document.getElementById('font-animations')){ const s=document.createElement('style'); s.id='font-animations'; s.textContent='@keyframes pulse{0%{opacity:1}50%{opacity:.5}100%{opacity:1}}'; document.head.appendChild(s); }

  // 折りたたみ
  async function toggleFontSection(){ const content=document.getElementById('fontSectionContent'); const arrow=document.getElementById('fontSectionArrow'); if(!content||!arrow) return; const collapsed = content.style.maxHeight && content.style.maxHeight!=='0px'; if(collapsed){ content.style.maxHeight='0px'; arrow.style.transform='rotate(-90deg)'; if(window.StorageManager?.setFontSectionCollapsed) StorageManager.setFontSectionCollapsed(true); } else { content.style.transition='none'; content.style.maxHeight=content.scrollHeight+'px'; setTimeout(()=>content.style.transition='max-height .3s ease-out',10); arrow.style.transform='rotate(0deg)'; if(window.StorageManager?.setFontSectionCollapsed) StorageManager.setFontSectionCollapsed(false); } }
  async function initializeFontSection(){ const content=document.getElementById('fontSectionContent'); const arrow=document.getElementById('fontSectionArrow'); if(!content||!arrow) return; let isCollapsed=true; try{ if(window.StorageManager?.getFontSectionCollapsed) isCollapsed = await StorageManager.getFontSectionCollapsed(); }catch{} if(!isCollapsed){ setTimeout(()=>{ content.style.maxHeight=content.scrollHeight+'px'; arrow.style.transform='rotate(0deg)'; },100); } else { content.style.maxHeight='0px'; arrow.style.transform='rotate(-90deg)'; } }
  function adjustFontSectionHeight(){ const c=document.getElementById('fontSectionContent'); if(c && c.style.maxHeight && c.style.maxHeight!=='0px'){ c.style.maxHeight = c.scrollHeight+'px'; } }

  function escapeHTML(str){ if(str==null) return ''; return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }

  // 公開
  window.CustomLabelFont = {
    initializeFontManager,
    loadCustomFontsCSS,
    updateFontList,
    initializeFontDropZone,
    removeFontFromList,
    toggleFontSection,
    initializeFontSection,
    adjustFontSectionHeight,
    get fontManager(){ return fontManager; }
  };
  // 後方互換（既存 HTML onclick 等）
  window.toggleFontSection = toggleFontSection;
})();
