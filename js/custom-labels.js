// custom-labels.js
// カスタムラベル管理モジュール（boothcsv.js から分離）
// CustomLabelCalculator もここに統合
class CustomLabelCalculator {
  static calculateMultiSheetDistribution(totalLabels, skipCount){
    const sheetsInfo=[]; let remainingLabels=totalLabels; let currentSkip=skipCount; let sheetNumber=1;
    while(remainingLabels>0){
      const availableInSheet = (CONSTANTS?.LABEL?.TOTAL_LABELS_PER_SHEET||44) - currentSkip;
      const labelsInThisSheet = Math.min(remainingLabels, availableInSheet);
      const remainingInSheet = availableInSheet - labelsInThisSheet;
      sheetsInfo.push({ sheetNumber, skipCount: currentSkip, labelCount: labelsInThisSheet, remainingCount: remainingInSheet, totalInSheet: currentSkip + labelsInThisSheet });
      remainingLabels -= labelsInThisSheet; currentSkip=0; sheetNumber++;
    }
    return sheetsInfo.length? sheetsInfo : [{ sheetNumber:1, skipCount:currentSkip, labelCount:0, remainingCount:(CONSTANTS?.LABEL?.TOTAL_LABELS_PER_SHEET||44)-currentSkip, totalInSheet:currentSkip }];
  }
}
window.CustomLabelCalculator = window.CustomLabelCalculator || CustomLabelCalculator;
(function(){
  if(window.CustomLabels) return;
  const dlog = (...a)=>{ if(typeof debugLog==='function') debugLog('[customLabel]', ...a); };

  let isEditingCustomLabel = false; // boothcsv.js 内の同名フラグと同期用に window へも露出
  Object.defineProperty(window, 'isEditingCustomLabel', { get: ()=>isEditingCustomLabel, set:v=>{ isEditingCustomLabel=v; } });

  function getContainer(){ return document.getElementById('customLabelsContainer'); }

  // =============================
  // Phase5 (Model/Diff Save): 内部モデル & 差分保存最適化
  //  - __model: UI 行と1:1対応するソースオブジェクト配列
  //  - getFromUI(): 既存 API を内部モデル経由に差し替え（DOM走査削減）
  //  - save(): JSONスナップショット比較で StorageManager への書込最小化
  //  - 各イベントでモデル部分更新 -> dirty フラグ
  // 後方互換: 外部 API シグネチャ/戻り値は不変
  // =============================
  const __model = []; // { enabled, count, html, fontSize, createdAt }
  let __modelDirty = false;
  let __lastSavedSnapshot = '';
  let __saveTimer = null; // debounced save timer

  function __ensureModelIndex(i){
    while(__model.length <= i){ __model.push({ enabled:true, count:1, html:'', fontSize:'12pt', createdAt: Date.now() }); }
  }
  function __updateModel(i, patch){
    __ensureModelIndex(i);
    Object.assign(__model[i], patch);
    __modelDirty = true;
  }
  function __rebuildModelFromDom(){
    const container = getContainer(); if(!container) return;
    __model.length = 0;
    container.querySelectorAll('.custom-label-item').forEach((item,i)=>{
      const editor = item.querySelector('.rich-text-editor');
      const countInput = item.querySelector('input[type="number"]');
      const enabledCb = item.querySelector('.custom-label-enabled');
      const html = (editor?.innerHTML||'').trim();
      const count = parseInt(countInput?.value||'1',10)||1;
      const enabled = enabledCb? enabledCb.checked : true;
      const fontSize = (editor? window.getComputedStyle(editor).fontSize : '12pt') || '12pt';
      const tsStr = item?.dataset?.timestamp;
      const ts = tsStr? parseInt(tsStr,10): NaN;
      __model.push({ enabled, count, html, fontSize, createdAt: (!isNaN(ts) && ts>0)? ts: Date.now() });
    });
    __modelDirty = true; // 再構築直後は未保存状態とみなす
  }
  function __exportModel(){
    return __model.map(m=>({
      text: m.html, html: m.html, count: m.count, fontSize: m.fontSize, enabled: m.enabled, createdAt: m.createdAt
    }));
  }
  function __serializeForSave(list){
    try { return JSON.stringify(list); } catch { return ''; }
  }
  function scheduleSave(delay=140){
    if(__saveTimer){ clearTimeout(__saveTimer); }
    __saveTimer = setTimeout(()=>{ try{ save(); } finally { __saveTimer=null; } }, delay);
  }
  function __markEditorDirty(editor, flushFast=false){
    if(!editor) return; const idx = parseInt(editor.dataset.index,10);
    if(!isNaN(idx)){
      __updateModel(idx, { html: (editor.innerHTML||'').trim(), fontSize: (window.getComputedStyle(editor).fontSize)||'12pt' });
      scheduleSave(flushFast? 50: 140);
    }
  }

  function initialize(customLabels){
    const container = getContainer(); if(!container) return;
    container.innerHTML='';
    const helpNode = cloneTemplate('customLabelsInstructionTemplate');
    try {
      // 開閉状態の復元
      const details = helpNode && helpNode.querySelector ? helpNode.querySelector('#customLabelsHelp') : null;
      if(details && window.StorageManager && typeof StorageManager.getCustomLabelsHelpOpen==='function'){
        StorageManager.getCustomLabelsHelpOpen().then(isOpen=>{ if(details.open !== !!isOpen) details.open = !!isOpen; }).catch(()=>{});
      }
      // トグル時に保存
      setTimeout(()=>{
        const d = document.getElementById('customLabelsHelp');
        if(d){
          d.addEventListener('toggle', ()=>{
            if(window.StorageManager && typeof StorageManager.setCustomLabelsHelpOpen==='function'){
              StorageManager.setCustomLabelsHelpOpen(!!d.open);
            }
          });
        }
      }, 0);
    } catch {}
    container.appendChild(helpNode);
    if(customLabels && customLabels.length){
      customLabels.forEach((label,index)=>{
        addItem(label.html || label.text, label.count, index, label.enabled !== false, label.fontSize, label.createdAt);
      });
    } else {
      addItem('',1,0,true);
    }
    // 初期読み込み時に内部モデルを DOM から再構築（addItem で登録済みでも念のため整合）
    __rebuildModelFromDom();
    updateSummary();
  }

  // =============================
  // Phase5(Partial): addItem の責務分割
  //  - buildLabelItemFragment: DOMフラグメント生成
  //  - initEditor: RichTextManager 初期化 + エディタイベント
  //  - wireLabelItemEvents: チェックボックス/枚数入力イベント
  //  - registerEditorPlugins: 将来のプラグイン拡張ポイント
  // 既存外部API (addItem) の挙動は維持
  // =============================

  function buildLabelItemFragment(state){
    const { text='', count=1, index=null, enabled=true, fontSize, createdAt } = state||{};
    const template = document.getElementById('customLabelItem'); if(!template) return null;
    const fragment = template.content.cloneNode(true);
    const itemDiv = fragment.querySelector('.custom-label-item');
    return { fragment, itemDiv, text, count, index, enabled, fontSize, createdAt };
  }

  function initializeLabelItemElements(container, ctx){
    const itemIndex = ctx.index!==null ? ctx.index : container.querySelectorAll('.custom-label-item').length;
    const checkbox = ctx.fragment.querySelector('.custom-label-enabled');
    const labelEl = ctx.fragment.querySelector('.custom-label-item-title');
    const editor = ctx.fragment.querySelector('.rich-text-editor');
    const countInput = ctx.fragment.querySelector('input[type="number"]');
    const removeBtn = ctx.fragment.querySelector('.btn-remove');

    ctx.itemDiv.dataset.index = itemIndex;
    if(ctx.createdAt){ ctx.itemDiv.dataset.timestamp = String(ctx.createdAt); }
    // checkbox 初期化
    checkbox.id = `customLabel_${itemIndex}_enabled`;
    checkbox.dataset.index = itemIndex;
    checkbox.checked = ctx.enabled;
    labelEl.setAttribute('for', checkbox.id);
    // editor 初期化
    editor.dataset.index = itemIndex;
    if(ctx.text && ctx.text.trim()) {
      const sanitize = window.sanitizeCustomLabelHTML || ((value) => String(value || ''));
      editor.innerHTML = sanitize(ctx.text);
    }
    if(ctx.fontSize){ editor.style.fontSize = ctx.fontSize.toString().includes('pt')? ctx.fontSize : ctx.fontSize + 'pt'; }
    // count
    countInput.value = ctx.count;
    countInput.dataset.index = itemIndex;
    // remove
    removeBtn.onclick = ()=> removeItem(itemIndex);

    return { itemIndex, checkbox, editor, countInput, removeBtn };
  }

  function initEditor(editorElement, itemIndex){
    if(!editorElement) return;
    if(window.initRichTextManager){
      try { window.initRichTextManager(editorElement); } catch(e){ console.error('initRichTextManager error', e); }
    } else {
      if(typeof setupRichTextFormatting==='function'){ try{ setupRichTextFormatting(editorElement);}catch(e){ console.error('setupRichTextFormatting early error',e);} }
      else { (window.__pendingEditorInit = window.__pendingEditorInit || []).push(editorElement); }
      if(typeof setupTextOnlyEditor==='function'){ try{ setupTextOnlyEditor(editorElement);}catch(e){ console.error('setupTextOnlyEditor early error',e);} }
    }
    // エディタ固有イベント
    editorElement.addEventListener('focus',()=>{ dlog('edit start', itemIndex); isEditingCustomLabel=true; });
    editorElement.addEventListener('blur',()=>{ dlog('edit end', itemIndex); isEditingCustomLabel=false; CustomLabels.schedulePreview(300); __markEditorDirty(editorElement,true); });
    editorElement.addEventListener('input', async ()=>{
      const item = editorElement.closest('.custom-label-item');
      if(item && editorElement.textContent.trim()!=='') item.classList.remove('error');
      __markEditorDirty(editorElement,false);
      updateButtonStates(); await updateSummary();
      if(!isEditingCustomLabel){ await autoProcessCSV?.(); } else { CustomLabels.schedulePreview(1000); }
    });
  }

  function wireLabelItemEvents(container, itemIndex){
    // checkbox change
    const enabledCb = container.querySelector(`.custom-label-enabled[data-index="${itemIndex}"]`);
  enabledCb?.addEventListener('change', async ()=>{ __updateModel(itemIndex,{ enabled: !!enabledCb.checked }); scheduleSave(); await updateSummary(); await autoProcessCSV?.(); });
    // count input
    const countEl = container.querySelector(`input[type="number"][data-index="${itemIndex}"]`);
  countEl?.addEventListener('input', async ()=>{ const val=parseInt(countEl.value||'1',10)||1; __updateModel(itemIndex,{ count: val }); scheduleSave(); updateButtonStates(); await updateSummary(); await autoProcessCSV?.(); });
  }

  function registerEditorPlugins(editorElement){
    // 将来的なプラグイン（例: context menu manager, spellchecker 等）を一括登録する拡張ポイント
    // 現段階では特別な追加処理なし（コンテキストメニューは RichTextManager の標準フックで動作）
  }

  function addItem(text='', count=1, index=null, enabled=true, fontSize, createdAt){
    dlog('addItem', {text,count,index,enabled,createdAt});
    const container = getContainer(); if(!container) return;
    const ts = (typeof createdAt==='number' && createdAt>0)? createdAt : Date.now();
    const ctx = buildLabelItemFragment({ text, count, index, enabled, fontSize, createdAt: ts }); if(!ctx) return;
    const { itemIndex, editor } = initializeLabelItemElements(container, ctx);
    // DOM 挿入
    container.appendChild(ctx.fragment);
    // エディタ初期化 + プラグイン
    initEditor(editor, itemIndex);
    registerEditorPlugins(editor);
    // コントロールのイベント
    wireLabelItemEvents(container, itemIndex);
    // モデル反映
    const baseFontSize = (editor? window.getComputedStyle(editor).fontSize : (fontSize||'12pt')) || '12pt';
    __ensureModelIndex(itemIndex);
    __model[itemIndex] = { enabled, count, html: (text||'').trim(), fontSize: baseFontSize, createdAt: ts };
    __modelDirty = true;
    return ctx.itemDiv;
  }

  function removeItem(index){
    const container = getContainer(); if(!container) return;
    const items = container.querySelectorAll('.custom-label-item');
    if(items.length<=1){ alert('最低1つのカスタムラベルは必要です。'); return; }
    const target = container.querySelector(`.custom-label-item[data-index="${index}"]`);
    if(target){
      // RichTextManager destroy (メモリリーク予防)
      try { const ed=target.querySelector('.rich-text-editor'); if(ed && ed.__rtm && typeof ed.__rtm.destroy==='function'){ ed.__rtm.destroy(); } } catch{}
      target.remove();
    }
  // モデルからも削除
  if(index>=0 && index < __model.length){ __model.splice(index,1); __modelDirty = true; }
    reindex(); save(); updateSummary(); updateButtonStates(); autoProcessCSV?.();
  }

  function reindex(){
    const container = getContainer(); if(!container) return;
    const items = container.querySelectorAll('.custom-label-item');
    items.forEach((item,i)=>{ item.dataset.index=i; const del = item.querySelector('.btn-danger'); if(del) del.setAttribute('onclick', `removeCustomLabelItem(${i})`); });
    // DOM順にモデルも並び替え（splice削除後は DOM と整合している前提）
    if(items.length !== __model.length){ __rebuildModelFromDom(); }
  }

  function getFromUI(){
    // モデルが DOM と著しく不整合なら再構築
    const container = getContainer();
    if(!container) return [];
    const domCount = container.querySelectorAll('.custom-label-item').length;
    if(domCount !== __model.length){ __rebuildModelFromDom(); }
    return __exportModel();
  }

  function save(){
    const data = __exportModel();
    const serialized = __serializeForSave(data);
    if(serialized && serialized !== __lastSavedSnapshot){
      StorageManager?.setCustomLabels(data);
      __lastSavedSnapshot = serialized;
      __modelDirty = false;
      dlog('saved (diff)');
    }
  }

  async function updateSummary(){
    const labels = getFromUI();
    const enabledLabels = labels.filter(l=>l.enabled);
    const customCount = enabledLabels.reduce((s,l)=>s+l.count,0);
    const skipCount = parseInt(document.getElementById('labelskipnum')?.value||'0',10)||0;
    const fileInput = document.getElementById('file');
    const summary = document.getElementById('customLabelsSummary'); if(!summary) return;

    let csvRows = 0;
    if(fileInput?.files?.length>0){
      try { const info = await CSVAnalyzer.getFileInfo(fileInput.files[0]); csvRows = parseInt(info.rowCount||0,10)||0; }
      catch { /* 解析失敗時は CSV を 0 件扱い（簡素表示優先） */ }
    }
    const totalLabels = csvRows + customCount;
    const sheetsInfo = CustomLabelCalculator.calculateMultiSheetDistribution(totalLabels, skipCount);
    const last = sheetsInfo[sheetsInfo.length-1];
    const remaining = Math.max(0, last?.remainingCount ?? 0);

    summary.textContent = `最終シートの残り面数: ${remaining}面`;
    summary.style.color='#666'; summary.style.fontWeight='normal';
  }

  async function adjustForTotal(customLabels, maxTotalLabels){
    let remaining = maxTotalLabels;
    for(let i=0;i<customLabels.length;i++){
      if(remaining<=0) customLabels[i].count=0; else if(customLabels[i].count>remaining) customLabels[i].count=remaining; remaining -= customLabels[i].count;
    }
    const container=getContainer(); const items=container.querySelectorAll('.custom-label-item');
    items.forEach((item,i)=>{ const countInput=item.querySelector('input[type="number"]'); if(customLabels[i]) countInput.value=customLabels[i].count; });
    save(); await updateSummary();
  }

  function setupEvents(){
    // HTML直置きに移行のため、移動ロジックは不要
    const initialEditor=document.getElementById('initialCustomLabelEditor');
    if(initialEditor){
      // Phase2: initialEditor も RichTextManager で統合初期化
      if(window.initRichTextManager && !initialEditor.__rtm){
        try { window.initRichTextManager(initialEditor); } catch(e){ console.error('initialEditor initRichTextManager error', e); }
      }
      initialEditor.addEventListener('focus',()=>{ isEditingCustomLabel=true; });
  initialEditor.addEventListener('blur',()=>{ isEditingCustomLabel=false; CustomLabels.schedulePreview(300); });
    }
    const enableCb=document.getElementById('customLabelEnable');
    if (enableCb) {
      enableCb.addEventListener('change', async function(){
        toggleCustomLabelRow?.(this.checked);
        await StorageManager.set(StorageManager.KEYS.CUSTOM_LABEL_ENABLE,this.checked);
        settingsCache.customLabelEnable=this.checked;
        updateButtonStates();
        await autoProcessCSV?.();
      });
      // 初期値の反映（Storageからload済のsettingsCacheに追従）
      StorageManager.getSettingsAsync()?.then(s=>{
        try { enableCb.checked = !!s.customLabelEnable; toggleCustomLabelRow?.(enableCb.checked); } catch {}
      }).catch(()=>{});
    }
    const addBtn=document.getElementById('addCustomLabelBtn');
    addBtn?.addEventListener('click', async ()=>{ addItem('',1,null,true); save(); updateButtonStates(); await autoProcessCSV?.(); });
    const clearBtn=document.getElementById('clearCustomLabelsBtn');
    clearBtn?.addEventListener('click', async ()=>{ if(confirm('本当に全てのカスタムラベルを削除しますか？')){ await clearAll(); await autoProcessCSV?.(); } });
  }

  async function clearAll(){
    const container = getContainer(); if(!container) return;
    container.querySelectorAll('.custom-label-item').forEach(n=>n.remove());
    addItem('',1,0,true); save(); await updateSummary(); updateButtonStates();
  }

  function hasContent(){ const labels=getFromUI(); return labels.length>0 && labels.some(l=>l.text.trim()!==''); }
  function hasEmptyEnabled(){ const enableRoot=document.getElementById('customLabelEnable'); if(enableRoot && !enableRoot.checked) return false; const labels=getFromUI(); return labels.filter(l=>l.enabled).some(l=>l.text.trim()===''); }

  // 静かなバリデーション（boothcsv.js から移動予定）
  function validateQuiet(){
    const enableEl=document.getElementById('customLabelEnable');
    if(enableEl && !enableEl.checked) return true; // 無効なら常にOK
    const labels=getFromUI();
    const enabled=labels.filter(l=>l.enabled);
    if(enabled.length===0) return false;
    for(const l of enabled){
      if(!l.text || l.text.trim()==='') return false;
      if(!l.count || l.count<=0) return false;
    }
    return true;
  }

  function removeEmpty(){ const labels=getFromUI(); const container=getContainer(); const items=container.querySelectorAll('.custom-label-item'); let removed=0; for(let i=labels.length-1;i>=0;i--){ const l=labels[i]; if(l.enabled && l.text.trim()===''){ const item=items[i]; if(item){ item.remove(); removed++; } } } reindex(); save(); updateSummary(); if(container.querySelectorAll('.custom-label-item').length===0) addItem('',1,0,true); if(removed>0) alert(`未設定のカスタムラベル ${removed} 項目を削除しました。設定済みのラベルのみで処理を続行します。`); }

  function highlightEmpty(){ const labels=getFromUI(); const items=getContainer().querySelectorAll('.custom-label-item'); items.forEach((item,i)=>{ if(labels[i] && labels[i].enabled && labels[i].text.trim()===''){ item.classList.add('error'); } else { item.classList.remove('error'); } }); }
  function clearHighlights(){ getContainer().querySelectorAll('.custom-label-item').forEach(i=>i.classList.remove('error')); }

  async function updateButtonStates(){ // 既存ロジック簡略委譲（完全移行時に整理）
    const printButton = document.getElementById('printButton');
    const printButtonCompact = document.getElementById('printButtonCompact');
    const hasSheets = document.querySelectorAll('.sheet').length > 0;
    const hasLabels = document.querySelectorAll('.label44').length > 0;
    const hasContent = hasSheets || hasLabels;
    if(printButton) printButton.disabled=!hasContent;
    if(printButtonCompact) printButtonCompact.disabled=!hasContent;
    await updateSummary();
  }

  // カスタムラベル専用 遅延プレビュー更新 (boothcsv.js から移動)
  let __pendingPreviewTimer=null;
  async function schedulePreview(delay=500){
    if(__pendingPreviewTimer){ clearTimeout(__pendingPreviewTimer); }
    __pendingPreviewTimer = setTimeout(async ()=>{
      try { debugLog?.('遅延プレビュー更新を実行 (CustomLabels.schedulePreview)'); await window.updateCustomLabelsPreview?.(); }
      finally { __pendingPreviewTimer=null; }
    }, delay);
  }

  // 公開 API
  window.CustomLabels = {
    initialize, addItem, removeItem, reindex, getFromUI, save, updateSummary,
  adjustForTotal, setupEvents, clearAll, hasContent, hasEmptyEnabled, removeEmpty, validateQuiet,
  highlightEmpty, clearHighlights, updateButtonStates, schedulePreview,
  __markEditorDirty // 内部用 (Phase5) – 将来的に非公開化予定
  };

  // Phase2 完了マーカー: RichTextManager 導入・旧初期化関数委譲・コンテキストメニュー委譲済み
  window.CustomLabels.__phase2Completed = true;

  // 後方互換グローバルは整理済み（直接呼び出しは CustomLabels.* を利用してください）
})();

// =============================
// スタイル適用ロジック（分離）
// =============================
(function(){
  const log = (...a)=>{ if(typeof debugLog==='function') debugLog('[labelCore]', ...a); };

  // =============================
  // Phase1 抽出: StyleHelper (span スタイル/クリーンアップ共通化)
  // =============================
  class StyleHelper {
    static parseStyle(styleStr){
      const map=new Map();
      if(!styleStr) return map;
      styleStr.split(';').forEach(rule=>{
        const [p,v]=rule.split(':').map(s=>s&&s.trim()).filter(Boolean);
        if(p&&v) map.set(p.toLowerCase(), v);
      });
      return map;
    }
    static mapToStyle(map){ return map.size? Array.from(map.entries()).map(([k,v])=>`${k}: ${v}`).join('; '): ''; }
    static updateSpanStyle(span, prop, val, isDefault){
      if(!span) return;
      const map=this.parseStyle(span.getAttribute('style')||'');
      if(isDefault){ map.delete(prop); } else { map.set(prop, val); }
      const styleStr=this.mapToStyle(map);
      if(styleStr) span.setAttribute('style', styleStr); else span.removeAttribute('style');
    }
    static mergeAdjacentSpans(editor){
      if(!editor) return;
      Array.from(editor.querySelectorAll('span[style]')).forEach(sp=>{
        const next=sp.nextSibling;
        if(next && next.nodeType===Node.ELEMENT_NODE && next.tagName==='SPAN' && next.getAttribute('style')===sp.getAttribute('style')){
          // textContent 連結: ZWSP 保護
          const nextText=next.textContent;
          if(nextText && !/\u200B$/.test(sp.textContent)) sp.textContent += nextText; else sp.textContent += nextText;
          next.remove();
        }
      });
    }
    static mergeNestedSpans(editor){
      if(!editor) return false;
      let changed=false;
      Array.from(editor.querySelectorAll('span[style]')).forEach(parentSpan=>{
        if(parentSpan.childElementCount!==1) return;
        const child=parentSpan.firstElementChild;
        if(!child || child.tagName!=='SPAN') return;
        const textNodes=Array.from(parentSpan.childNodes).filter(n=>n.nodeType===Node.TEXT_NODE && (n.nodeValue||'').trim()!=='');
        if(textNodes.length) return;
        const parentMap=this.parseStyle(parentSpan.getAttribute('style')||'');
        const childMap=this.parseStyle(child.getAttribute('style')||'');
        const merged=new Map([...parentMap.entries(), ...childMap.entries()]);
        const mergedStyle=this.mapToStyle(merged);
        if(mergedStyle) child.setAttribute('style', mergedStyle); else child.removeAttribute('style');
        parentSpan.parentNode?.insertBefore(child, parentSpan);
        parentSpan.remove();
        changed=true;
      });
      return changed;
    }
    static cleanupSpans(editor){
      if(!editor) return;
      let changed=true, loops=0;
      while(changed && loops<5){
        changed=false; loops++;
        editor.querySelectorAll('span').forEach(span=>{
          const style=(span.getAttribute('style')||'').trim();
            if(!style){
              if(span.childElementCount===0){
                const parent=span.parentNode; if(!parent) return;
                while(span.firstChild) parent.insertBefore(span.firstChild, span);
                parent.removeChild(span); changed=true;
              }
            }
        });
        if(this.mergeNestedSpans(editor)) changed=true;
        this.mergeAdjacentSpans(editor);
      }
    }
  }
  window.StyleHelper = window.StyleHelper || StyleHelper; // デバッグ / 将来利用用

  function parseStyleString(styleString){
  // Phase1: StyleHelper へ委譲
  return StyleHelper.parseStyle(styleString);
  }

  function updateSpanStyle(span, prop, val, isDefault){
  // Phase1: StyleHelper へ委譲
  StyleHelper.updateSpanStyle(span, prop, val, isDefault);
  }

  function removeStyleFromDescendants(rootEl, prop){
    if(!rootEl) return; try {
      rootEl.querySelectorAll('span[style]')?.forEach(s=>{
        const map=parseStyleString(s.getAttribute('style')||''); if(map.has(prop)){ map.delete(prop); const ns=Array.from(map.entries()).map(([k,v])=>`${k}: ${v}`).join('; '); if(ns) s.setAttribute('style', ns); else s.removeAttribute('style'); }
      });
    } catch(e){ log('removeStyleFromDescendants error', e); }
  }

  function analyzeSelectionRange(range){
    const ca = range.commonAncestorContainer;
    let targetSpan=null,isCompleteSpan=false,isPartialSpan=false,isMultiSpan=false,multiSpans=[];
    if(ca.nodeType===Node.TEXT_NODE){ const p=ca.parentElement; if(p&&p.tagName==='SPAN'){ const sel=range.toString(); if(sel===p.textContent){ targetSpan=p; isCompleteSpan=true; } else { targetSpan=p; isPartialSpan=true; } } }
    else if(ca.nodeType===Node.ELEMENT_NODE && ca.tagName==='SPAN'){ targetSpan=ca; isCompleteSpan=true; }
    else { const root=(ca.nodeType===Node.ELEMENT_NODE? ca: ca.parentElement); if(root){ const sel=range.toString(); const spans=Array.from(root.querySelectorAll('span')); const hit=spans.filter(s=>range.intersectsNode(s) && sel.includes(s.textContent)); if(hit.length>1){ isMultiSpan=true; multiSpans=hit; } } }
    return { targetSpan,isCompleteSpan,isPartialSpan,isMultiSpan,multiSpans, commonAncestor: ca };
  }

  function unwrapElement(el){
    if(!el?.parentNode) return;
    const parent=el.parentNode;
    while(el.firstChild) parent.insertBefore(el.firstChild, el);
    parent.removeChild(el);
  }

  function getEditorFromRange(range, editor){
    if(editor) return editor;
    const base = range?.commonAncestorContainer;
    const element = base?.nodeType===Node.ELEMENT_NODE ? base : base?.parentElement;
    return element?.closest?.('.rich-text-editor') || null;
  }

  function isRangeInsideEditor(range, editor){
    if(!range || !editor) return false;
    const nodes=[range.startContainer, range.endContainer, range.commonAncestorContainer];
    return nodes.every(node=>{
      if(!node) return false;
      if(node===editor) return true;
      const base=node.nodeType===Node.ELEMENT_NODE ? node : node.parentNode;
      return !!(base && (base===editor || editor.contains(base)));
    });
  }

  function replaceExtractedSelection(range, content){
    const sel=window.getSelection();
    const fragment=document.createDocumentFragment();
    const startMarker=document.createComment('sel-start');
    const endMarker=document.createComment('sel-end');
    fragment.appendChild(startMarker);
    if(content){
      if(content.nodeType===Node.DOCUMENT_FRAGMENT_NODE){ fragment.appendChild(content); }
      else { fragment.appendChild(content); }
    }
    fragment.appendChild(endMarker);
    range.insertNode(fragment);
    const nextRange=document.createRange();
    nextRange.setStartAfter(startMarker);
    nextRange.setEndBefore(endMarker);
    startMarker.parentNode?.removeChild(startMarker);
    endMarker.parentNode?.removeChild(endMarker);
    sel?.removeAllRanges();
    sel?.addRange(nextRange);
    return nextRange;
  }

  function fragmentToContainer(fragment){
    const container=document.createElement('div');
    container.appendChild(fragment);
    return container;
  }

  function containerToFragment(container){
    const fragment=document.createDocumentFragment();
    while(container.firstChild) fragment.appendChild(container.firstChild);
    return fragment;
  }

  function stripStylePropertyFromFragment(fragment, prop){
    const container=fragmentToContainer(fragment);
    container.querySelectorAll('[style]').forEach(el=>{
      const tagName=el.tagName;
      if(tagName!=='SPAN' && tagName!=='STRONG' && tagName!=='EM' && tagName!=='U') return;
      const map=parseStyleString(el.getAttribute('style')||'');
      map.delete(prop);
      if(prop==='font-size'){
        map.delete('line-height');
      }
      const styleStr=StyleHelper.mapToStyle(map);
      if(styleStr) el.setAttribute('style', styleStr); else el.removeAttribute('style');
      if(tagName==='SPAN' && !(el.getAttribute('style')||'').trim()) unwrapElement(el);
    });
    return containerToFragment(container);
  }

  function wrapFragmentWithStyle(fragment, prop, value){
    const span=document.createElement('span');
    try{
      if(prop==='font-family') span.style.fontFamily=value;
      else span.style.setProperty(prop, value);
      if(prop==='font-size') span.style.lineHeight='1.2';
    } catch{}
    span.appendChild(fragment);
    return span;
  }

  function applyStyleToFormattedDescendants(fragment, prop, value, isDefault=false){
    const container=fragmentToContainer(fragment);
    container.querySelectorAll('strong, em, u').forEach(el=>{
      const map=parseStyleString(el.getAttribute('style')||'');
      if(isDefault){
        map.delete(prop);
      } else {
        map.set(prop, value);
      }
      const styleStr=StyleHelper.mapToStyle(map);
      if(styleStr) el.setAttribute('style', styleStr); else el.removeAttribute('style');
    });
    return containerToFragment(container);
  }

  function getSelectedTextNodes(range, editor){
    if(!range) return [];
    const root=getEditorFromRange(range, editor) || (range.commonAncestorContainer.nodeType===Node.ELEMENT_NODE ? range.commonAncestorContainer : range.commonAncestorContainer.parentNode);
    if(!root) return [];
    const nodes=[];
    const walker=document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(node){
        if(!node?.nodeValue || node.nodeValue.length===0) return NodeFilter.FILTER_REJECT;
        const parent=node.parentNode;
        if(editor && parent && !(parent===editor || editor.contains(parent))) return NodeFilter.FILTER_REJECT;
        try { return range.intersectsNode(node) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT; }
        catch { return NodeFilter.FILTER_REJECT; }
      }
    });
    while(walker.nextNode()) nodes.push(walker.currentNode);
    return nodes;
  }

  function getUniformSelectedStyles(range, editor, props=[]){
    const textNodes=getSelectedTextNodes(range, editor).filter(node=>node.nodeValue.trim()!=='');
    if(!textNodes.length || !props.length) return {};
    const result={};
    props.forEach(prop=>{
      let uniformValue=null;
      let isFirst=true;
      for(const node of textNodes){
        const base=node.parentElement || node.parentNode;
        const value=base ? window.getComputedStyle(base).getPropertyValue(prop) : '';
        if(isFirst){
          uniformValue=value;
          isFirst=false;
        } else if(uniformValue !== value){
          uniformValue=null;
          break;
        }
      }
      if(uniformValue) result[prop]=uniformValue;
    });
    return result;
  }

  function findClosestAncestorTag(node, tag, editor){
    let current=node?.nodeType===Node.TEXT_NODE ? node.parentNode : node;
    while(current && current!==editor){
      if(current.nodeType===Node.ELEMENT_NODE && current.tagName===tag) return current;
      current=current.parentNode;
    }
    return null;
  }

  function splitTextBoundaries(range){
    if(!range) return range;
    if(range.endContainer?.nodeType===Node.TEXT_NODE){
      const endNode=range.endContainer;
      const endOffset=range.endOffset;
      if(endOffset>0 && endOffset<endNode.nodeValue.length){
        endNode.splitText(endOffset);
        range.setEnd(endNode, endNode.nodeValue.length);
      }
    }
    if(range.startContainer?.nodeType===Node.TEXT_NODE){
      const startNode=range.startContainer;
      const startOffset=range.startOffset;
      if(startOffset>0 && startOffset<startNode.nodeValue.length){
        const endWasSameNode=range.endContainer===startNode;
        const endOffsetBeforeSplit=range.endOffset;
        const afterNode=startNode.splitText(startOffset);
        range.setStart(afterNode, 0);
        if(endWasSameNode){
          range.setEnd(afterNode, Math.max(0, endOffsetBeforeSplit - startOffset));
        }
      }
    }
    return range;
  }

  function findBoundaryReferenceNode(container, offset, atStart){
    if(container?.nodeType===Node.TEXT_NODE) return container;
    if(!container?.childNodes?.length) return container;
    if(atStart){
      return container.childNodes[offset] || container.childNodes[offset-1] || container;
    }
    return container.childNodes[offset-1] || container.childNodes[offset] || container;
  }

  function isTagBoundaryEdge(range, target, atStart){
    const edgeRange=document.createRange();
    edgeRange.selectNodeContents(target);
    try {
      if(atStart){
        edgeRange.setEnd(range.startContainer, range.startOffset);
      } else {
        edgeRange.setStart(range.endContainer, range.endOffset);
      }
    } catch {
      return false;
    }
    return edgeRange.collapsed;
  }

  function splitTagAncestorAtBoundary(range, tag, editor, atStart){
    const container=atStart ? range.startContainer : range.endContainer;
    const offset=atStart ? range.startOffset : range.endOffset;
    const refNode=findBoundaryReferenceNode(container, offset, atStart);
    const target=findClosestAncestorTag(refNode, tag, editor);
    if(!target || !target.parentNode) return;
    if(isTagBoundaryEdge(range, target, atStart)) return;
    const tailRange=document.createRange();
    tailRange.selectNodeContents(target);
    try {
      if(atStart){
        tailRange.setStart(range.startContainer, range.startOffset);
        if(tailRange.collapsed) return;
      } else {
        tailRange.setStart(range.endContainer, range.endOffset);
        if(tailRange.collapsed) return;
      }
    } catch {
      return;
    }
    const fragment=tailRange.extractContents();
    const clone=target.cloneNode(false);
    clone.appendChild(fragment);
    target.parentNode.insertBefore(clone, target.nextSibling);
    if(atStart){
      range.setStart(clone, 0);
    } else {
      range.setEnd(target, target.childNodes.length);
    }
  }

  function unwrapTagFromFragment(fragment, tag){
    const container=fragmentToContainer(fragment);
    container.querySelectorAll(tag.toLowerCase()).forEach(el=> unwrapElement(el));
    return containerToFragment(container);
  }

  function fragmentHasNodes(fragment){
    return !!fragment && Array.from(fragment.childNodes).some(node=>{
      if(node.nodeType===Node.TEXT_NODE) return node.nodeValue.length>0;
      return true;
    });
  }

  function removeFormatWithinSingleAncestor(range, tag, editor){
    const startTag=findClosestAncestorTag(range.startContainer, tag, editor);
    const endTag=findClosestAncestorTag(range.endContainer, tag, editor);
    if(!startTag || startTag!==endTag || !startTag.parentNode) return false;
    const target=startTag;
    const beforeRange=document.createRange();
    beforeRange.selectNodeContents(target);
    beforeRange.setEnd(range.startContainer, range.startOffset);
    const beforeFragment=beforeRange.cloneContents();

    const selectedFragment=unwrapTagFromFragment(range.cloneContents(), tag);

    const afterRange=document.createRange();
    afterRange.selectNodeContents(target);
    afterRange.setStart(range.endContainer, range.endOffset);
    const afterFragment=afterRange.cloneContents();

    const replacement=document.createDocumentFragment();
    const startMarker=document.createComment('sel-start');
    const endMarker=document.createComment('sel-end');

    if(fragmentHasNodes(beforeFragment)){
      const beforeWrapper=target.cloneNode(false);
      beforeWrapper.appendChild(beforeFragment);
      replacement.appendChild(beforeWrapper);
    }

    replacement.appendChild(startMarker);
    replacement.appendChild(selectedFragment);
    replacement.appendChild(endMarker);

    if(fragmentHasNodes(afterFragment)){
      const afterWrapper=target.cloneNode(false);
      afterWrapper.appendChild(afterFragment);
      replacement.appendChild(afterWrapper);
    }

    target.parentNode.replaceChild(replacement, target);

    const sel=window.getSelection();
    const nextRange=document.createRange();
    nextRange.setStartAfter(startMarker);
    nextRange.setEndBefore(endMarker);
    startMarker.parentNode?.removeChild(startMarker);
    endMarker.parentNode?.removeChild(endMarker);
    sel?.removeAllRanges();
    sel?.addRange(nextRange);
    return true;
  }

  function applyStylePreservingBreaks(range, prop, value, editor, isDefault=false){
    const scopedEditor=getEditorFromRange(range, editor);
    const unit = (prop==='font-size' && /^(\d+)$/.test(String(value)))? 'pt': '';
    const explicitValue = isDefault && prop==='font-family'
      ? (window.getComputedStyle(scopedEditor || document.body).fontFamily || 'sans-serif')
      : String(value)+unit;
    const extracted=range.extractContents();
    const cleanedBase=stripStylePropertyFromFragment(extracted, prop);
    const cleaned=applyStyleToFormattedDescendants(cleanedBase, prop, explicitValue, isDefault);
    const wrapper=wrapFragmentWithStyle(cleaned, prop, explicitValue);
    return replaceExtractedSelection(range, wrapper);
  }

  function cleanupEmptySpans(editor){
  // Phase1: StyleHelper へ委譲
  StyleHelper.cleanupSpans(editor);
  }

  function applyStyleToSelection(prop, value, editor, isDefault=false){
    const sel=window.getSelection(); if(!sel.rangeCount || sel.isCollapsed) return; const range=sel.getRangeAt(0); const text=range.toString(); if(!text) return;
    if(!isRangeInsideEditor(range, editor)) return;
    try {
      applyStylePreservingBreaks(range, prop, value, editor, isDefault);
      cleanupEmptySpans(editor);
    } catch(e){ log('applyStyleToSelection error', e); }
    editor?.focus();
  }

  function applyFontFamilyToSelection(fontFamily, editor){ const isDefault=!fontFamily; if(isDefault){ applyDefaultFontToSelection(editor); } else { applyStyleToSelection('font-family', fontFamily, editor, false); } }
  function applyFontSizeToSelection(fontSize, editor){ applyStyleToSelection('font-size', fontSize, editor, false); }

  function applyDefaultFontToSelection(editor){
    const sel=window.getSelection(); if(!sel.rangeCount || sel.isCollapsed) return;
    try {
      const range=sel.getRangeAt(0);
      if(!isRangeInsideEditor(range, editor)) return;
      applyStylePreservingBreaks(range, 'font-family', '', editor, true);
      StyleHelper.cleanupSpans(editor);
    } catch(e){ log('applyDefaultFontToSelection error', e); applyStyleToSelection('font-family','',editor,true); }
    editor?.focus();
  }

  // Facade撤去: 最小コアAPIを __LabelCore として公開
  window.__LabelCore = {
    applyStyleToSelection,
    applyFontFamilyToSelection,
    applyFontSizeToSelection,
    applyDefaultFontToSelection,
    applyFormatToSelection,
    applyFormatToSelectionFallback,
    clearAllContent,
    isSelectionFormatted,
    isRangeInsideEditor,
    getTargetTagName,
    analyzeSelectionRange,
    cleanupEmptySpans,
    StyleHelper
  };
  if(window.CustomLabels){ Object.assign(window.CustomLabels, { analyzeSelectionRange, cleanupEmptySpans }); }
  // 書式（bold/italic/underline）関連を追加
  function getTargetTagName(command){
    switch(command){
      case 'bold': return 'STRONG';
      case 'italic': return 'EM';
      case 'underline': return 'U';
      default: return null;
    }
  }
  function isSelectionFormatted(range, command, editor){
    const tag = getTargetTagName(command);
    if(!tag) return false;
    const scopedEditor=getEditorFromRange(range, editor);
    const textNodes=getSelectedTextNodes(range, scopedEditor).filter(node=>node.nodeValue.trim()!=='');
    if(!textNodes.length) return false;
    return textNodes.every(node=>!!findClosestAncestorTag(node, tag, scopedEditor));
  }
  function applyFormatToRange(range, command, editor){
    const scopedEditor=getEditorFromRange(range, editor);
    const inheritedStyles=getUniformSelectedStyles(range, scopedEditor, ['font-family','font-size','line-height']);
    const frag=range.extractContents(); let el; switch(command){ case 'bold': el=document.createElement('strong'); break; case 'italic': el=document.createElement('em'); break; case 'underline': el=document.createElement('u'); break; default: return; }
    Object.entries(inheritedStyles).forEach(([prop, val])=>{ try { el.style.setProperty(prop, val); } catch {} });
    el.appendChild(frag);
    replaceExtractedSelection(range, el);
  }
  function removeFormatFromSelection(range, command, editor){
    const tag=getTargetTagName(command);
    const scopedEditor=getEditorFromRange(range, editor);
    if(!tag || !scopedEditor || !isRangeInsideEditor(range, scopedEditor)) return;
    if(removeFormatWithinSingleAncestor(range, tag, scopedEditor)){
      normalizeInlineFormatting(scopedEditor);
      return;
    }
    splitTextBoundaries(range);
    splitTagAncestorAtBoundary(range, tag, scopedEditor, false);
    splitTagAncestorAtBoundary(range, tag, scopedEditor, true);
    const extracted=range.extractContents();
    const cleaned=unwrapTagFromFragment(extracted, tag);
    replaceExtractedSelection(range, cleaned);
    normalizeInlineFormatting(scopedEditor);
  }
  function applyFormatToSelectionFallback(command, editor){ const sel=window.getSelection(); if(!sel.rangeCount||sel.isCollapsed) return; const range=sel.getRangeAt(0); if(!isRangeInsideEditor(range, editor)) return; if(isSelectionFormatted(range,command,editor)){ removeFormatFromSelection(range,command,editor); } else { applyFormatToRange(range,command,editor); } }
  function clearAllContent(editor){ if(!editor) return; if(confirm('このカスタムラベルの内容と書式をすべてクリアしますか？')){ editor.innerHTML=''; editor.style.fontSize='12pt'; editor.style.lineHeight='1.2'; editor.style.textAlign='center'; editor.focus(); window.CustomLabels?.save(); } }
  function normalizeInlineFormatting(editor){
    if(!editor) return;
    // b -> strong, i -> em (semantic)
    editor.querySelectorAll('b').forEach(b=>{ const strong=document.createElement('strong'); while(b.firstChild) strong.appendChild(b.firstChild); b.parentNode.replaceChild(strong,b); });
    editor.querySelectorAll('i').forEach(i=>{ const em=document.createElement('em'); while(i.firstChild) em.appendChild(i.firstChild); i.parentNode.replaceChild(em,i); });
    ['strong','em','u'].forEach(tagName=>{
      editor.querySelectorAll(`${tagName} ${tagName}`).forEach(el=> unwrapElement(el));
      editor.querySelectorAll(tagName).forEach(el=>{
        const next=el.nextSibling;
        if(next && next.nodeType===Node.ELEMENT_NODE && next.tagName===tagName.toUpperCase()){
          while(next.firstChild) el.appendChild(next.firstChild);
          next.remove();
        }
      });
    });
  }
  function applyFormatToSelection(command, editor){
    if(!editor) return;
    if(command==='clear'){ clearAllContent(editor); return; }
    // 常に手動トグル: execCommand を排除して安定化
    applyFormatToSelectionFallback(command, editor);
    normalizeInlineFormatting(editor);
    try { window.CustomLabels?.__markEditorDirty?.(editor,false); } catch {}
    editor.focus();
  }
  // 後方互換用 CustomLabelStyle グローバルを削除告知だけのプロキシに (存在していれば保持)
  try {
    Object.defineProperty(window, 'CustomLabelStyle', { get(){ console.error('[CustomLabelStyle REMOVED] 旧APIは削除されました。__LabelCore または RichTextManager を利用してください。'); return undefined; }, configurable:true });
  } catch{}
})();

// (Removed: Phase7 Step2 duplicate facade)

// =============================
// コンテキストメニュー（フォント/書式）生成ロジック（boothcsv.js から移動）
// =============================
(function(){
  if(window.ContextMenuManager){ return; }
  const clog = (...a)=>{ if(typeof debugLog==='function') debugLog('[ContextMenuManager]', ...a); };

  class ContextMenuManager {
  constructor(){ this.currentMenu=null; this.fontCache=null; this.fontCacheTime=0; this.fontCacheTTL=5*60*1000; this.fontLoadError=false; }
    invalidateFontCache(){
      this.fontCache=null;
      this.fontCacheTime=0;
      this.fontLoadError=false;
    }
    rememberSelection(editor){
      const sel=window.getSelection();
      if(!editor || !sel?.rangeCount || sel.isCollapsed) return false;
      const range=sel.getRangeAt(0);
      if(!window.__LabelCore?.isRangeInsideEditor?.(range, editor)) return false;
      this.savedSelectionRange=range.cloneRange();
      this.savedSelectionEditor=editor;
      return true;
    }
    hasRememberedSelection(editor){
      if(!this.savedSelectionRange || (editor && this.savedSelectionEditor!==editor)) return false;
      try { return !this.savedSelectionRange.collapsed && this.savedSelectionRange.toString().length>0; }
      catch { return false; }
    }
    restoreSelection(editor=this.savedSelectionEditor){
      if(!editor || !this.savedSelectionRange) return false;
      try {
        const sel=window.getSelection();
        sel?.removeAllRanges();
        sel?.addRange(this.savedSelectionRange.cloneRange());
        editor.focus();
        return true;
      } catch(err){
        clog('restoreSelection err', err);
        return false;
      }
    }
    clearSelectionSnapshot(){
      this.savedSelectionRange=null;
      this.savedSelectionEditor=null;
    }
    close(menu){
      const targets = menu? [menu] : Array.from(document.querySelectorAll('.custom-label-context-menu'));
      targets.forEach(el=>{ if(el && el.parentNode){ el.style.opacity='0'; setTimeout(()=>{ try{ el.parentNode && el.parentNode.removeChild(el); }catch{} },150);} });
      if(!menu) this.currentMenu=null; else if(this.currentMenu===menu) this.currentMenu=null;
    }
    async show(x,y,editor,hasSelection=true){
      try { this.close(); } catch{}
      const remembered=this.rememberSelection(editor);
      if(!hasSelection && !remembered && this.hasRememberedSelection(editor)) hasSelection=true;
      const menu=document.createElement('div');
      menu.className='custom-label-context-menu';
      menu.style.cssText=`position:fixed;background:#fff;border:1px solid #ccc;border-radius:6px;padding:8px 0;box-shadow:0 4px 20px rgba(0,0,0,.15);z-index:10000;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;min-width:160px;max-width:250px;max-height:400px;overflow-y:auto;visibility:hidden;opacity:0;transition:opacity .2s ease;`;
      document.body.appendChild(menu); this.currentMenu=menu;
      const appendItem=(opt)=>{
        const item=document.createElement('div'); item.textContent=opt.label; item.style.cssText=`padding:8px 15px;cursor:pointer;font-size:12px;transition:background-color .2s;${opt.style||''}`;
        item.addEventListener('mouseenter',function(){ this.style.backgroundColor='#f0f0f0'; });
        item.addEventListener('mouseleave',function(){ this.style.backgroundColor='transparent'; });
        item.addEventListener('mousedown',e=>{ e.preventDefault(); e.stopPropagation(); });
        item.addEventListener('click', e=>{ e.preventDefault(); e.stopPropagation(); setTimeout(()=>{ try{ this.restoreSelection(editor); opt.onClick?.(); }catch(err){clog('item click err',err);} this.close(menu); try{ window.CustomLabels?.save(); }catch{} },10); });
        menu.appendChild(item);
      };
      const sep=()=>{ const d=document.createElement('div'); d.style.cssText='height:1px;background:#ddd;margin:5px 0;'; menu.appendChild(d); };
      const header=(text)=>{ const h=document.createElement('div'); h.textContent=text; h.style.cssText='padding:5px 15px;font-size:11px;color:#666;font-weight:bold;'; menu.appendChild(h); };

      // Format options
      const formatDefs=[
        {label:'太字', cmd:'bold', style:'font-weight:bold;', needsSel:true},
        {label:'斜体', cmd:'italic', style:'font-style:italic;', needsSel:true},
        {label:'下線', cmd:'underline', style:'text-decoration:underline;', needsSel:true},
        {label:'すべてクリア', cmd:'clear', style:'color:#dc3545;font-weight:bold;', needsSel:false}
      ];
      formatDefs.filter(d=>hasSelection || d.cmd==='clear').forEach(d=>{
        appendItem({ label:d.label, style:d.style, onClick:()=>{
          if(d.cmd==='clear'){ editor.__rtm?.clearAll(); }
          else { editor.__rtm?.applyFormat(d.cmd); }
        }});
      });

      if(hasSelection){
        sep(); header('フォントサイズ');
  [6,8,10,12,14,16,18,20,24,28].forEach(size=> appendItem({ label:`${size}pt`, style:'padding-left:20px;font-size:11px;', onClick:()=>{ editor.__rtm?.applyStyle('font-size', size); } }));
        sep(); header('フォント');
        appendItem({ label:'デフォルトフォント（システムフォント）', style:'font-size:11px;font-family:sans-serif;font-weight:bold;color:#333;border-bottom:1px solid #eee;', onClick:()=>{
          try { const sel=window.getSelection(); if(sel.rangeCount>0 && !sel.isCollapsed){ editor.__rtm?.applyStyle('font-family','',true);} else if(editor.style.fontFamily){ editor.style.fontFamily=''; } }
          catch(err){ clog('defaultFont err',err); editor.__rtm?.applyStyle('font-family','',true);} }
        });

        const systemFonts=[
          { name:'ゴシック（sans-serif）', family:'sans-serif' },
            { name:'明朝（serif）', family:'serif' },
            { name:'等幅（monospace）', family:'monospace' },
            { name:'Arial', family:'Arial, sans-serif' },
            { name:'Times New Roman', family:'Times New Roman, serif' },
            { name:'メイリオ', family:'Meiryo, sans-serif' },
            { name:'ヒラギノ角ゴ', family:'Hiragino Kaku Gothic Pro, sans-serif' }
        ];
        const addFontItem=(name,family,custom=false)=>{
          appendItem({ label:name, style:`font-size:11px;font-family:${family};padding-left:20px;`, onClick:()=>{ editor.__rtm?.applyStyle('font-family', family); } });
        };
        systemFonts.forEach(f=>addFontItem(f.name, f.family));

        // Custom fonts (async load / cache)
    const now=Date.now(); if(this.fontCache && (now - this.fontCacheTime > this.fontCacheTTL)){ this.fontCache=null; }
    if(this.fontCache===null && !this.fontLoadError){
          // placeholder while loading
          const placeholder=document.createElement('div'); placeholder.textContent='(カスタムフォント取得中...)'; placeholder.style.cssText='padding:4px 15px;font-size:10px;color:#888;'; menu.appendChild(placeholder);
          try {
            if(window.fontManager){
      const fonts= await window.fontManager.getAllFonts(); this.fontCache=fonts; this.fontCacheTime=Date.now(); placeholder.textContent='カスタムフォント'; placeholder.style.fontWeight='bold'; placeholder.style.color='#666'; placeholder.style.borderTop='1px solid #eee'; placeholder.style.marginTop='4px';
              Object.keys(fonts).forEach(name=> addFontItem(name, `"${name}", sans-serif`, true));
            } else { this.fontCache={}; placeholder.remove(); }
          } catch(err){ this.fontLoadError=true; placeholder.textContent='(カスタムフォント取得失敗)'; clog('custom font load error', err); }
        } else if(this.fontCache && Object.keys(this.fontCache).length){
          const lbl=document.createElement('div'); lbl.textContent='カスタムフォント'; lbl.style.cssText='padding:5px 15px;font-size:10px;color:#666;font-weight:bold;border-top:1px solid #eee;border-bottom:1px solid #eee;'; menu.appendChild(lbl);
          Object.keys(this.fontCache).forEach(name=> addFontItem(name, `"${name}", sans-serif`, true));
        }
      }

      // Positioning
      const rect=menu.getBoundingClientRect();
      const vw=window.innerWidth, vh=window.innerHeight, margin=10;
      let ax=x, ay=y;

      const maxX = Math.max(margin, vw - rect.width - margin);
      const maxY = Math.max(margin, vh - rect.height - margin);
      ax = Math.min(ax, maxX);
      ay = Math.min(ay, maxY);
      ax = Math.max(ax, margin);
      ay = Math.max(ay, margin);

      menu.style.left=ax+'px';
      menu.style.top=ay+'px';
      menu.style.visibility='visible';
      requestAnimationFrame(()=>{ menu.style.opacity='1'; });
      return menu;
    }
  }

  window.ContextMenuManager = new ContextMenuManager();
  document.addEventListener('custom-label-fonts-updated', ()=>{ try { window.ContextMenuManager?.invalidateFontCache?.(); } catch{} });
  // 互換APIラッパ
  function createFontSizeMenu(x,y,editor,hasSelection){ return window.ContextMenuManager.show(x,y,editor,hasSelection); }
  function closeContextMenu(menu){ return window.ContextMenuManager.close(menu); }
  window.CustomLabelContextMenu = { createFontSizeMenu, closeContextMenu }; // レガシー互換
  window.createFontSizeMenu = window.createFontSizeMenu || createFontSizeMenu;
  window.closeContextMenu = window.closeContextMenu || closeContextMenu;
  // グローバル閉じるイベント (一度だけ登録)
  if(!window.__customLabelContextMenuGlobalHandlers){
    window.__customLabelContextMenuGlobalHandlers = true;
    document.addEventListener('mousedown', e=>{
      const open=document.querySelector('.custom-label-context-menu'); if(!open) return;
      if(!open.contains(e.target)) window.ContextMenuManager.close(open);
    });
    window.addEventListener('keydown', e=>{ if(e.key==='Escape'){ window.ContextMenuManager.close(); }});
    window.addEventListener('scroll', ()=>{ window.ContextMenuManager.close(); }, { passive:true });
  }
  if(window.CustomLabels){ window.CustomLabels.__phase4Completed = true; }
})();

// =============================
// プレーンテキスト入力制限 & サニタイズ (setupTextOnlyEditor 移設)
// =============================
(function(){
  if(window.CustomLabelEditor && window.CustomLabelEditor.setupTextOnlyEditor && window.CustomLabelEditor.setupRichTextFormatting) return; // 冪等
  // =============================
  // Phase2: RichTextManager 導入 (新規エディタは統合初期化)
  // =============================
  (function(){
    if(window.RichTextManager) return;
    class RichTextManager {
      constructor(editor){
        this.editor=editor; this.isComposing=false; this._bind();
      }
      syncSelectionSnapshot(){
        try { window.ContextMenuManager?.rememberSelection?.(this.editor); } catch {}
      }
      ensureSelectionForOperation(){
        if(this.hasActiveSelection()) return true;
        try { window.ContextMenuManager?.restoreSelection?.(this.editor); } catch {}
        return this.hasActiveSelection();
      }
      hasActiveSelection(){
        if(!this.editor) return false;
        const sel=window.getSelection();
        if(!sel?.rangeCount || sel.isCollapsed) return false;
        const range=sel.getRangeAt(0);
        return window.__LabelCore?.isRangeInsideEditor?.(range, this.editor) || false;
      }
      applyStyle(prop, value, isDefault=false){
        if(!this.editor) return;
        try {
            if(!this.ensureSelectionForOperation()) return;
            const disabled = window.__DISABLE_DEPRECATED_CUSTOM_LABEL_STYLE;
            if(prop==='font-size'){
              window.__LabelCore?.applyFontSizeToSelection(value, this.editor);
            } else if(prop==='font-family') {
              if(isDefault) window.__LabelCore?.applyFontFamilyToSelection('', this.editor);
              else window.__LabelCore?.applyFontFamilyToSelection(value, this.editor);
            } else {
              window.__LabelCore?.applyStyleToSelection(prop, value, this.editor, isDefault);
            }
            this.syncSelectionSnapshot();
        } catch(e){ console.error('RichTextManager.applyStyle error', e); }
      }
      applyFormat(command){
        if(!this.editor) return;
          try { if(!this.ensureSelectionForOperation()) return; window.__LabelCore?.applyFormatToSelection(command, this.editor); this.syncSelectionSnapshot(); }
        catch(e){ console.error('RichTextManager.applyFormat error', e); }
      }
  clearAll(){ try { window.__LabelCore?.clearAllContent(this.editor); } catch(e){ console.error('RichTextManager.clearAll error', e);} }
      _bind(){
        const ed=this.editor; if(!ed) return;
        ed.__rtm = this; // マーカー
        this._selectionChangeHandler = ()=>{ this.syncSelectionSnapshot(); };
        document.addEventListener('selectionchange', this._selectionChangeHandler);
        ed.addEventListener('compositionstart',()=>{ this.isComposing=true; });
        ed.addEventListener('compositionend',()=>{ this.isComposing=false; });
        // ペースト/ドロップ（プレーンテキスト化）
        const insertPlain=(text)=>{
          if(!text) return;
          // 1) HTMLタグ除去 2) CRLF -> LF 正規化 3) 末尾の単一改行は格納時に不要なので除去
          let clean=text.replace(/<[^>]*>/g,'');
          clean=clean.replace(/\r\n?|\u2028|\u2029/g,'\n');
          // 末尾に一つだけある改行は視覚的に不要なので削る (複数連続改行は意図とみなして1つ残す案もあるがまず単一ケースに限定)
          if(/[^\n]\n$/.test(clean)) clean=clean.replace(/\n$/,'');
          const lines=clean.split('\n');
          const sel=window.getSelection(); if(!sel?.rangeCount) return; const range=sel.getRangeAt(0); range.deleteContents();
          lines.forEach((line,i)=>{
            if(i>0){ const br=document.createElement('br'); range.insertNode(br); range.setStartAfter(br); }
            if(line.length){ const tn=document.createTextNode(line); range.insertNode(tn); range.setStartAfter(tn); }
          });
          range.collapse(true); sel.removeAllRanges(); sel.addRange(range);
          // モデル即時反映 (貼付けのみでキーイベントが発生しないケースの空行ズレ防止)
          try { window.CustomLabels?.__markEditorDirty?.(ed,true); } catch {}
        };
        ed.addEventListener('paste', e=>{ try{ e.preventDefault(); }catch{} let t=''; try{ t=(e.clipboardData||window.clipboardData).getData('text/plain'); }catch{} insertPlain(t); });
        ed.addEventListener('drop', e=>{ try{ e.preventDefault(); }catch{} if(e.dataTransfer?.files?.length) return false; const t=e.dataTransfer?.getData('text/plain'); insertPlain(t); });
        ed.addEventListener('dragover', e=>{ if(e.dataTransfer?.types?.includes('Files')){ e.dataTransfer.dropEffect='none'; return false;} e.preventDefault(); });
        ed.addEventListener('mouseup', ()=>{ setTimeout(()=>this.syncSelectionSnapshot(), 0); });
        ed.addEventListener('keyup', ()=>{ setTimeout(()=>this.syncSelectionSnapshot(), 0); });
        ed.addEventListener('focus', ()=>{ setTimeout(()=>this.syncSelectionSnapshot(), 0); });
        // Enter -> <br>
        const supportsBeforeInput=('onbeforeinput' in ed);
        const insertBr=(src)=>{ const sel=window.getSelection(); if(!sel?.rangeCount) return; const range=sel.getRangeAt(0); let atEnd=false; try{ const tail=document.createRange(); tail.selectNodeContents(ed); tail.setStart(range.endContainer, range.endOffset); atEnd=tail.toString().length===0; }catch{} range.deleteContents(); const br=document.createElement('br'); range.insertNode(br); if(atEnd){ const zw=document.createTextNode('\u200B'); if(!(br.nextSibling && br.nextSibling.nodeType===Node.TEXT_NODE && br.nextSibling.nodeValue.startsWith('\u200B'))){ br.parentNode.insertBefore(zw, br.nextSibling); range.setStartAfter(zw);} } else { range.setStartAfter(br);} range.collapse(true); sel.removeAllRanges(); sel.addRange(range); };
        if(supportsBeforeInput){ ed.addEventListener('beforeinput', e=>{ if(e.inputType==='insertParagraph'){ if(this.isComposing||e.isComposing) return; e.preventDefault(); insertBr('beforeinput'); }}); }
        ed.addEventListener('keydown', e=>{
          const key=(e.key||'').toLowerCase();
          const isShortcut=(e.ctrlKey || e.metaKey) && !e.altKey;
          if(isShortcut && (key==='b' || key==='i' || key==='u')){
            if(this.isComposing||e.isComposing) return;
            if(!this.hasActiveSelection()) return;
            e.preventDefault();
            const commandMap={ b:'bold', i:'italic', u:'underline' };
            this.applyFormat(commandMap[key]);
            return;
          }
          if(e.key==='Enter'){ if(supportsBeforeInput) return; if(this.isComposing||e.isComposing) return; e.preventDefault(); insertBr('keydown'); }
        });
        // コンテキストメニュー
  ed.addEventListener('contextmenu', async e=>{ try{ e.preventDefault(); const sel=window.getSelection(); const hasSel=!!sel && sel.toString().length>0; if(window.ContextMenuManager){ await window.ContextMenuManager.show(e.clientX,e.clientY,ed,hasSel); } else { (window.createFontSizeMenu||window.CustomLabelContextMenu?.createFontSizeMenu)?.(e.clientX,e.clientY,ed,hasSel); } }catch(err){ console.error('rtm contextmenu',err);} });
        // 変な要素除去
        const observer=new MutationObserver(muts=>{ muts.forEach(m=>{ if(m.type==='childList'){ m.addedNodes.forEach(n=>{ if(n.nodeType===Node.ELEMENT_NODE){ const tg=n.tagName; if(tg==='IMG'||tg==='VIDEO'||tg==='AUDIO'){ n.remove(); } } }); } }); }); observer.observe(ed,{childList:true,subtree:true}); this._observer=observer;
      }
      destroy(){ try{ this._observer?.disconnect(); }catch{} try{ if(this._selectionChangeHandler) document.removeEventListener('selectionchange', this._selectionChangeHandler); }catch{} }
    }
    window.RichTextManager = RichTextManager;
    window.initRichTextManager = function(editor){ if(!editor || editor.__rtm) return; try { new RichTextManager(editor); } catch(e){ console.error('initRichTextManager error', e); } };
  })();

  function setupTextOnlyEditor(editor){
    // Phase2: 旧実装を廃し RichTextManager へ委譲する薄いラッパ
    if(!editor) return;
    if(editor.dataset.plainOnlyInit) return;
    editor.dataset.plainOnlyInit='1';
    if(!editor.__rtm){
      if(window.initRichTextManager){
        try { window.initRichTextManager(editor); } catch(e){ console.error('setupTextOnlyEditor delegate error', e);} }
    }
  }
  function setupRichTextFormatting(editor){
    // Phase2: 旧実装を廃し RichTextManager へ委譲する薄いラッパ
    if(!editor) return;
    if(editor.dataset.richFormattingInit) return;
    editor.dataset.richFormattingInit='1';
    if(!editor.__rtm){
      if(window.initRichTextManager){
        try { window.initRichTextManager(editor); } catch(e){ console.error('setupRichTextFormatting delegate error', e);} }
    }
  }
  window.CustomLabelEditor = { setupTextOnlyEditor, setupRichTextFormatting };
  window.setupTextOnlyEditor = window.setupTextOnlyEditor || setupTextOnlyEditor;
  window.setupRichTextFormatting = window.setupRichTextFormatting || setupRichTextFormatting;
  if(window.CustomLabels){ Object.assign(window.CustomLabels, { setupTextOnlyEditor, setupRichTextFormatting }); }
  if(window.__pendingEditorInit){ window.__pendingEditorInit.forEach(ed=>{ try { setupRichTextFormatting(ed); setupTextOnlyEditor(ed); } catch(e){ console.error('delayed init error', e);} }); delete window.__pendingEditorInit; }
})();
