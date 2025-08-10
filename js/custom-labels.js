// custom-labels.js
// カスタムラベル管理モジュール（boothcsv.js から分離）
(function(){
  if(window.CustomLabels) return;
  const dlog = (...a)=>{ if(typeof debugLog==='function') debugLog('[customLabel]', ...a); };

  let isEditingCustomLabel = false; // boothcsv.js 内の同名フラグと同期用に window へも露出
  Object.defineProperty(window, 'isEditingCustomLabel', { get: ()=>isEditingCustomLabel, set:v=>{ isEditingCustomLabel=v; } });

  function getContainer(){ return document.getElementById('customLabelsContainer'); }

  function initialize(customLabels){
    const container = getContainer(); if(!container) return;
    container.innerHTML='';
    container.appendChild(cloneTemplate('customLabelsInstructionTemplate'));
    if(customLabels && customLabels.length){
      customLabels.forEach((label,index)=>{
        addItem(label.html || label.text, label.count, index, label.enabled !== false, label.fontSize);
      });
    } else {
      addItem('',1,0,true);
    }
    updateSummary();
  }

  function addItem(text='', count=1, index=null, enabled=true, fontSize){
    dlog('addItem', {text,count,index,enabled});
    const container = getContainer(); if(!container) return;
    const template = document.getElementById('customLabelItem'); if(!template) return;
    const fragment = template.content.cloneNode(true);
    const itemDiv = fragment.querySelector('.custom-label-item');
    const itemIndex = index!==null? index : container.querySelectorAll('.custom-label-item').length;
    itemDiv.dataset.index = itemIndex;
    // checkbox
    const checkbox = fragment.querySelector('.custom-label-enabled');
    checkbox.id = `customLabel_${itemIndex}_enabled`; checkbox.dataset.index=itemIndex; checkbox.checked = enabled;
    const labelEl = fragment.querySelector('.custom-label-item-title');
    labelEl.setAttribute('for', checkbox.id);
    const editor = fragment.querySelector('.rich-text-editor');
    editor.dataset.index=itemIndex;
    if(text && text.trim()) editor.innerHTML = text;
    if(fontSize){ editor.style.fontSize = fontSize.toString().includes('pt')? fontSize: fontSize+'pt'; }
    const countInput = fragment.querySelector('input[type="number"]'); countInput.value = count; countInput.dataset.index=itemIndex;
    const removeBtn = fragment.querySelector('.btn-remove'); removeBtn.onclick = ()=> removeItem(itemIndex);
    container.appendChild(fragment);
    const editorElement = container.querySelector(`.rich-text-editor[data-index="${itemIndex}"]`);
    if(editorElement){
      // リッチテキスト初期化 + プレーンテキスト制約（定義前ならキュー）
      if(typeof setupRichTextFormatting==='function'){ try{ setupRichTextFormatting(editorElement);}catch(e){ console.error('setupRichTextFormatting early error',e);} }
      else { (window.__pendingEditorInit = window.__pendingEditorInit || []).push(editorElement); }
      if(typeof setupTextOnlyEditor==='function'){ try{ setupTextOnlyEditor(editorElement);}catch(e){ console.error('setupTextOnlyEditor early error',e);} }
      editorElement.addEventListener('focus',()=>{ dlog('edit start'); isEditingCustomLabel=true; });
      editorElement.addEventListener('blur',()=>{ dlog('edit end'); isEditingCustomLabel=false; scheduleDelayedPreviewUpdate?.(300); });
      editorElement.addEventListener('input', async ()=>{
        const item = editorElement.closest('.custom-label-item');
        if(item && editorElement.textContent.trim()!=='') item.classList.remove('error');
        save(); updateButtonStates(); await updateSummary();
        if(!isEditingCustomLabel){ await autoProcessCSV?.(); } else { scheduleDelayedPreviewUpdate?.(1000); }
      });
    }
    // checkbox change
    const enabledCb = container.querySelector(`.custom-label-enabled[data-index="${itemIndex}"]`);
    enabledCb?.addEventListener('change', async ()=>{ save(); await updateSummary(); await autoProcessCSV?.(); });
    // count input
    const countEl = container.querySelector(`input[type="number"][data-index="${itemIndex}"]`);
    countEl?.addEventListener('input', async ()=>{ save(); updateButtonStates(); await updateSummary(); await autoProcessCSV?.(); });
  }

  function removeItem(index){
    const container = getContainer(); if(!container) return;
    const items = container.querySelectorAll('.custom-label-item');
    if(items.length<=1){ alert('最低1つのカスタムラベルは必要です。'); return; }
    const target = container.querySelector(`.custom-label-item[data-index="${index}"]`); if(target) target.remove();
    reindex(); save(); updateSummary(); updateButtonStates(); autoProcessCSV?.();
  }

  function reindex(){
    const container = getContainer(); if(!container) return;
    const items = container.querySelectorAll('.custom-label-item');
    items.forEach((item,i)=>{ item.dataset.index=i; const del = item.querySelector('.btn-danger'); if(del) del.setAttribute('onclick', `removeCustomLabelItem(${i})`); });
  }

  function getFromUI(){
    const container = getContainer(); if(!container) return [];
    const items = container.querySelectorAll('.custom-label-item');
    const labels=[];
    items.forEach(item=>{
      const editor = item.querySelector('.rich-text-editor');
      const countInput = item.querySelector('input[type="number"]');
      const enabledCb = item.querySelector('.custom-label-enabled');
      const text = (editor?.innerHTML||'').trim();
      const count = parseInt(countInput?.value||'1',10)||1;
      const enabled = enabledCb? enabledCb.checked : true;
      const fontSize = window.getComputedStyle(editor).fontSize || '12pt';
      labels.push({ text, count, fontSize, html:text, enabled });
    });
    return labels;
  }

  function save(){ StorageManager?.setCustomLabels(getFromUI()); }

  async function updateSummary(){
    const labels = getFromUI();
    const enabledLabels = labels.filter(l=>l.enabled);
    const totalCount = enabledLabels.reduce((s,l)=>s+l.count,0);
    const skipCount = parseInt(document.getElementById('labelskipnum')?.value||'0',10)||0;
    const fileInput = document.getElementById('file');
    const summary = document.getElementById('customLabelsSummary'); if(!summary) return;
    if(totalCount===0){
      if(fileInput?.files?.length>0){
        try { const info = await CSVAnalyzer.getFileInfo(fileInput.files[0]); const remaining = CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET - skipCount - info.rowCount; summary.innerHTML=`カスタムラベルなし。<br>44面シート中 スキップ${skipCount} + CSV${info.rowCount} = ${skipCount+info.rowCount}面使用済み。<br>残り${Math.max(0,remaining)}面設定可能。`; }
        catch { summary.innerHTML='カスタムラベルなし。<br>CSVファイル選択済み（行数解析中...）。'; }
      } else { summary.innerHTML=`カスタムラベルなし。<br>44面シート中 スキップ${skipCount}面使用済み。`; }
      summary.style.color='#666'; summary.style.fontWeight='normal'; return;
    }
    if(fileInput?.files?.length>0){
      try {
        const info = await CSVAnalyzer.getFileInfo(fileInput.files[0]);
        const totalLabels = info.rowCount + totalCount;
        const sheetsInfo = CustomLabelCalculator.calculateMultiSheetDistribution(totalLabels, skipCount);
        const last = sheetsInfo[sheetsInfo.length-1];
        if(sheetsInfo.length===1){ summary.innerHTML=`合計 ${totalCount}面のカスタムラベル。<br>1シート使用: スキップ${skipCount} + CSV${info.rowCount} + カスタム${totalCount} = ${skipCount+info.rowCount+totalCount}面<br>最終シート残り${last.remainingCount}面。`; }
        else { summary.innerHTML=`合計 ${totalCount}面のカスタムラベル。<br>${sheetsInfo.length}シート使用: CSV${info.rowCount} + カスタム${totalCount} = ${info.rowCount+totalCount}面<br>最終シート残り${last.remainingCount}面。`; }
        summary.style.color='#666'; summary.style.fontWeight='normal';
      } catch (e) {
        summary.innerHTML=`合計 ${totalCount}面のカスタムラベル。<br>CSVファイル選択済み（行数解析エラー）<br>CSV処理実行後に最終配置が決定されます。`;
        summary.style.color='#ffc107'; summary.style.fontWeight='normal';
      }
    } else {
      const sheetsInfo = CustomLabelCalculator.calculateMultiSheetDistribution(totalCount, skipCount);
      const last = sheetsInfo[sheetsInfo.length-1];
      if(sheetsInfo.length===1) summary.innerHTML=`合計 ${totalCount}面のカスタムラベル。<br>1シート使用: スキップ${skipCount} + カスタム${totalCount} = ${skipCount+totalCount}面<br>最終シート残り${last.remainingCount}面。`;
      else summary.innerHTML=`合計 ${totalCount}面のカスタムラベル。<br>${sheetsInfo.length}シート使用: カスタム${totalCount}面<br>最終シート残り${last.remainingCount}面。`;
      summary.style.color='#666'; summary.style.fontWeight='normal';
    }
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
    const initialEditor=document.getElementById('initialCustomLabelEditor');
    if(initialEditor){
      initialEditor.addEventListener('focus',()=>{ isEditingCustomLabel=true; });
      initialEditor.addEventListener('blur',()=>{ isEditingCustomLabel=false; scheduleDelayedPreviewUpdate?.(300); });
    }
    const enableCb=document.getElementById('customLabelEnable');
    enableCb?.addEventListener('change', async function(){ toggleCustomLabelRow?.(this.checked); await StorageManager.set(StorageManager.KEYS.CUSTOM_LABEL_ENABLE,this.checked); settingsCache.customLabelEnable=this.checked; updateButtonStates(); await autoProcessCSV?.(); });
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

  // 公開 API
  window.CustomLabels = {
    initialize, addItem, removeItem, reindex, getFromUI, save, updateSummary,
  adjustForTotal, setupEvents, clearAll, hasContent, hasEmptyEnabled, removeEmpty, validateQuiet,
  highlightEmpty, clearHighlights, updateButtonStates
  };

  // 後方互換グローバルは整理済み（直接呼び出しは CustomLabels.* を利用してください）
})();

// =============================
// スタイル適用ロジック（分離）
// =============================
(function(){
  // 既に定義済みなら再定義しない
  if(window.CustomLabelStyle) return;
  const log = (...a)=>{ if(typeof debugLog==='function') debugLog('[customLabelStyle]', ...a); };

  function parseStyleString(styleString){
    const styleMap=new Map(); if(!styleString) return styleMap;
    styleString.split(';').forEach(rule=>{ const [p,v]=rule.split(':').map(s=>s.trim()); if(p&&v) styleMap.set(p.toLowerCase(), v); });
    return styleMap;
  }

  function updateSpanStyle(span, prop, val, isDefault){
    const map = parseStyleString(span.getAttribute('style')||'');
    if(isDefault){ map.delete(prop); } else { map.set(prop, val); }
    if(map.size===0) span.removeAttribute('style'); else span.setAttribute('style', Array.from(map.entries()).map(([k,v])=>`${k}: ${v}`).join('; '));
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

  function applyStylePreservingBreaks(range, prop, value){
    const unit = (prop==='font-size' && /^(\d+)$/.test(String(value)))? 'pt': '';
    const valStr = String(value)+unit;
    const frag=range.extractContents();
    const walker=document.createTreeWalker(frag, NodeFilter.SHOW_TEXT, null);
    const nodes=[]; while(walker.nextNode()) { const n=walker.currentNode; if(n.textContent.trim()!=='') nodes.push(n); }
    nodes.forEach(tn=>{ const span=document.createElement('span'); try{ if(prop==='font-family') span.style.fontFamily=valStr; else span.style.setProperty(prop, valStr); }catch{} tn.parentNode.replaceChild(span, tn); span.appendChild(tn); });
    range.insertNode(frag);
  }

  function cleanupEmptySpans(editor){
    if(!editor) return; let changed=true, loops=0; while(changed && loops<5){ changed=false; loops++;
      // 空 / スタイル無し unwrap
      editor.querySelectorAll('span').forEach(span=>{
        const style = (span.getAttribute('style')||'').trim();
        if(!style){ if(span.childElementCount===0){ // unwrap
          const parent=span.parentNode; while(span.firstChild) parent.insertBefore(span.firstChild, span); parent.removeChild(span); changed=true; }
        }
      });
      // 隣接同スタイル結合
      Array.from(editor.querySelectorAll('span[style]')).forEach(sp=>{
        const next=sp.nextSibling; if(next && next.nodeType===Node.ELEMENT_NODE && next.tagName==='SPAN' && next.getAttribute('style')===sp.getAttribute('style')){ sp.textContent += next.textContent; next.remove(); changed=true; }
      });
    }
  }

  function applyStyleToSelection(prop, value, editor, isDefault=false){
    const sel=window.getSelection(); if(!sel.rangeCount || sel.isCollapsed) return; const range=sel.getRangeAt(0); const text=range.toString(); if(!text) return;
    try {
      const info=analyzeSelectionRange(range);
      const unit = (prop==='font-size' && /^(\d+)$/.test(String(value)))? 'pt': '';
      const valStr = String(value)+unit;
      if(info.isCompleteSpan){ updateSpanStyle(info.targetSpan, prop, valStr, isDefault); if(prop==='font-size'||prop==='font-family') removeStyleFromDescendants(info.targetSpan, prop); }
      else if(info.isPartialSpan){ // テキストノード内部分選択
        if(range.startContainer===range.endContainer && range.startContainer.nodeType===Node.TEXT_NODE){
          const tn=range.startContainer; const full=tn.nodeValue||''; const before=full.slice(0,range.startOffset); const mid=full.slice(range.startOffset, range.endOffset); const after=full.slice(range.endOffset);
          tn.nodeValue=before; const span=document.createElement('span'); if(!isDefault){ try{ if(prop==='font-family') span.style.fontFamily=valStr; else span.style.setProperty(prop, valStr); }catch{} } span.textContent=mid; const afterNode=document.createTextNode(after); const parent=tn.parentNode; parent.insertBefore(span, tn.nextSibling); parent.insertBefore(afterNode, span.nextSibling);
          sel.removeAllRanges(); const nr=document.createRange(); nr.selectNodeContents(span); sel.addRange(nr);
        } else { applyStylePreservingBreaks(range, prop, value); }
      }
      else if(info.isMultiSpan){ info.multiSpans.forEach(s=> updateSpanStyle(s, prop, valStr, isDefault)); if(prop==='font-size'||prop==='font-family') info.multiSpans.forEach(s=> removeStyleFromDescendants(s, prop)); }
      else { applyStylePreservingBreaks(range, prop, value); }
      cleanupEmptySpans(editor);
    } catch(e){ log('applyStyleToSelection error', e); }
    editor?.focus();
  }

  function applyFontFamilyToSelection(fontFamily, editor){ const isDefault=!fontFamily; if(isDefault){ applyDefaultFontToSelection(editor); } else { applyStyleToSelection('font-family', fontFamily, editor, false); } }
  function applyFontSizeToSelection(fontSize, editor){ applyStyleToSelection('font-size', fontSize, editor, false); }

  function applyDefaultFontToSelection(editor){
    const sel=window.getSelection(); if(!sel.rangeCount || sel.isCollapsed) return; try { const range=sel.getRangeAt(0); const text=range.toString(); if(!text) return; const ca=range.commonAncestorContainer; let target=null; if(ca.nodeType===Node.TEXT_NODE){ const p=ca.parentElement; if(p&&p.tagName==='SPAN'&&p.style.fontFamily) target=p; } else if(ca.tagName==='SPAN' && ca.style.fontFamily) target=ca; if(target){ const style=(target.getAttribute('style')||''); const clean=style.split(';').filter(r=>{ const prop=r.trim().split(':')[0]?.trim().toLowerCase(); return prop && prop!=='font-family'; }).join('; '); if(clean.trim()) target.setAttribute('style', clean); else { const parent=target.parentNode; const textContent=target.textContent; const tn=document.createTextNode(textContent); parent.replaceChild(tn, target); range.selectNode(tn); sel.removeAllRanges(); sel.addRange(range); } } else { applyStyleToSelection('font-family','',editor,true); } cleanupEmptySpans(editor); } catch(e){ log('applyDefaultFontToSelection error', e); applyStyleToSelection('font-family','',editor,true); } editor?.focus(); }

  // 公開 (CustomLabels に統合)
  window.CustomLabelStyle = { applyStyleToSelection, applyFontFamilyToSelection, applyFontSizeToSelection, applyDefaultFontToSelection, analyzeSelectionRange, cleanupEmptySpans };
  if(window.CustomLabels){ Object.assign(window.CustomLabels, window.CustomLabelStyle); }
  // 書式（bold/italic/underline）関連を追加
  function getTargetTagName(command){
    switch(command){
      case 'bold': return 'STRONG';
      case 'italic': return 'EM';
      case 'underline': return 'U';
      default: return null;
    }
  }
  function isSelectionFormatted(range, command){
    const tag = getTargetTagName(command); if(!tag) return false; let node=range.startContainer; if(node.nodeType===Node.TEXT_NODE) node=node.parentNode; while(node && node.closest && node.closest('.rich-text-editor')){ if(node.tagName===tag) return true; node=node.parentNode; } return false;
  }
  function applyFormatToRange(range, command){
    const frag=range.extractContents(); let el; switch(command){ case 'bold': el=document.createElement('strong'); break; case 'italic': el=document.createElement('em'); break; case 'underline': el=document.createElement('u'); break; default: return; }
    el.appendChild(frag); range.insertNode(el); range.selectNodeContents(el); const sel=window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
  }
  function removeFormatFromSelection(range, command){
    const tag=getTargetTagName(command); if(!tag) return; let node=range.startContainer; if(node.nodeType===Node.TEXT_NODE) node=node.parentNode; let target=null; while(node && node.closest && node.closest('.rich-text-editor')){ if(node.tagName===tag){ target=node; break; } node=node.parentNode; }
    if(!target) return; const parent=target.parentNode; const editor=target.closest('.rich-text-editor'); const frag=document.createDocumentFragment(); const children=Array.from(target.childNodes); children.forEach(c=>frag.appendChild(c)); parent.replaceChild(frag, target);
    const sel=window.getSelection(); sel.removeAllRanges(); if(editor && children.length){ try { const nr=document.createRange(); const first=children[0]; const last=children[children.length-1]; if(first.nodeType===Node.TEXT_NODE) nr.setStart(first,0); else nr.setStartBefore(first); if(last.nodeType===Node.TEXT_NODE) nr.setEnd(last,last.textContent.length); else nr.setEndAfter(last); sel.addRange(nr); } catch { try { const nr=document.createRange(); nr.selectNodeContents(editor); nr.collapse(false); sel.addRange(nr); } catch {} } }
  }
  function applyFormatToSelectionFallback(command, editor){ const sel=window.getSelection(); if(!sel.rangeCount||sel.isCollapsed) return; const range=sel.getRangeAt(0); if(isSelectionFormatted(range,command)){ removeFormatFromSelection(range,command); } else { applyFormatToRange(range,command); } }
  function clearAllContent(editor){ if(!editor) return; if(confirm('このカスタムラベルの内容と書式をすべてクリアしますか？')){ editor.innerHTML=''; editor.style.fontSize='12pt'; editor.style.lineHeight='1.2'; editor.style.textAlign='center'; editor.focus(); window.CustomLabels?.save(); } }
  function applyFormatToSelection(command, editor){ if(command==='clear'){ clearAllContent(editor); return; } try { let exec; switch(command){ case 'bold': exec='bold'; break; case 'italic': exec='italic'; break; case 'underline': exec='underline'; break; default: return; } document.execCommand(exec,false,null); } catch(e){ applyFormatToSelectionFallback(command, editor); } editor?.focus(); }
  // 公開 / マージ (後方互換グローバル不要)
  window.CustomLabelStyle = { ...window.CustomLabelStyle, applyFormatToSelection, applyFormatToSelectionFallback, isSelectionFormatted, getTargetTagName, applyFormatToRange, removeFormatFromSelection, clearAllContent };
  if(window.CustomLabels){ Object.assign(window.CustomLabels, { applyFormatToSelection, clearAllContent }); }
})();

// =============================
// コンテキストメニュー（フォント/書式）生成ロジック（boothcsv.js から移動）
// =============================
(function(){
  if(window.CustomLabelContextMenu) return; // 冪等性確保
  const clog = (...a)=>{ if(typeof debugLog==='function') debugLog('[customLabelContextMenu]', ...a); };

  function closeContextMenu(menu) {
    const targets = menu ? [menu] : Array.from(document.querySelectorAll('.custom-label-context-menu'));
    targets.forEach(el=>{
      if (el && el.parentNode) {
        el.style.opacity = '0';
        setTimeout(() => { try { el.parentNode && el.parentNode.removeChild(el); } catch{} }, 150);
      }
    });
  }

  async function createFontSizeMenu(x, y, editor, hasSelection = true) {
    // 既存のメニュー（複数右クリック時の多重表示防止）をすべて即時削除
    try {
      document.querySelectorAll('.custom-label-context-menu').forEach(m=>{ try { m.parentNode && m.parentNode.removeChild(m);} catch{} });
    } catch {}
    const menu = document.createElement('div');
    menu.classList.add('custom-label-context-menu');
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
    document.body.appendChild(menu);

    const formatOptions = [
      { label: '太字', command: 'bold', style: 'font-weight: bold;' },
      { label: '斜体', command: 'italic', style: 'font-style: italic;' },
      { label: '下線', command: 'underline', style: 'text-decoration: underline;' },
      { label: 'すべてクリア', command: 'clear', style: 'color: #dc3545; font-weight: bold;' }
    ];
    const availableOptions = hasSelection ? formatOptions : [formatOptions[3]];
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
      item.addEventListener('mouseenter', function(){ this.style.backgroundColor = '#f0f0f0'; });
      item.addEventListener('mouseleave', function(){ this.style.backgroundColor = 'transparent'; });
      item.addEventListener('mousedown', e=>{ e.preventDefault(); e.stopPropagation(); });
      item.addEventListener('click', e => {
        e.preventDefault(); e.stopPropagation();
        setTimeout(()=>{
          try { window.CustomLabelStyle?.applyFormatToSelection(option.command, editor); } catch(err){ clog('applyFormat error', err); }
          closeContextMenu(menu);
          try { window.CustomLabels?.save(); } catch{}
        },10);
      });
      menu.appendChild(item);
    });

    if (hasSelection) {
      const separator = document.createElement('div');
      separator.style.cssText = 'height:1px;background-color:#ddd;margin:5px 0;';
      menu.appendChild(separator);

      const fontSizeLabel = document.createElement('div');
      fontSizeLabel.textContent = 'フォントサイズ';
      fontSizeLabel.style.cssText = 'padding:5px 15px;font-size:11px;color:#666;font-weight:bold;';
      menu.appendChild(fontSizeLabel);

      const fontSizes = [6,8,10,12,14,16,18,20,24,28];
      fontSizes.forEach(size => {
        const item = document.createElement('div');
        item.textContent = `${size}pt`;
        item.style.cssText = 'padding:6px 20px;cursor:pointer;font-size:11px;transition:background-color .2s;';
        item.addEventListener('mouseenter', function(){ this.style.backgroundColor='#f0f0f0'; });
        item.addEventListener('mouseleave', function(){ this.style.backgroundColor='transparent'; });
        item.addEventListener('mousedown', e=>{ e.preventDefault(); e.stopPropagation(); });
        item.addEventListener('click', e => {
          e.preventDefault(); e.stopPropagation();
            setTimeout(()=>{
              try { window.CustomLabelStyle?.applyFontSizeToSelection(size, editor); } catch(err){ clog('applyFontSize error', err); }
              closeContextMenu(menu); try { window.CustomLabels?.save(); } catch{}
            },10);
        });
        menu.appendChild(item);
      });

      try {
        // フォント区切り線
        const fontSeparator = document.createElement('div');
        fontSeparator.style.cssText = 'height:1px;background-color:#ddd;margin:5px 0;';
        menu.appendChild(fontSeparator);
        const fontFamilyLabel = document.createElement('div');
        fontFamilyLabel.textContent = 'フォント';
        fontFamilyLabel.style.cssText = 'padding:5px 15px;font-size:11px;color:#666;font-weight:bold;';
        menu.appendChild(fontFamilyLabel);

        const defaultFontItem = document.createElement('div');
        defaultFontItem.textContent = 'デフォルトフォント（システムフォント）';
        defaultFontItem.style.cssText = 'padding:6px 20px;cursor:pointer;font-size:11px;transition:background-color .2s;font-family:sans-serif;border-bottom:1px solid #eee;font-weight:bold;color:#333;';
        ['mouseenter','mouseleave'].forEach(ev=> defaultFontItem.addEventListener(ev, function(){ this.style.backgroundColor = (ev==='mouseenter')? '#f0f0f0':'transparent'; }));
        defaultFontItem.addEventListener('mousedown', e=>{ e.preventDefault(); e.stopPropagation(); });
        defaultFontItem.addEventListener('click', e=>{
          e.preventDefault(); e.stopPropagation();
          setTimeout(()=>{
            try {
              const selection = window.getSelection();
              if (selection.rangeCount>0 && !selection.isCollapsed) {
                window.CustomLabelStyle?.applyFontFamilyToSelection('', editor);
              } else if (editor?.style?.fontFamily) {
                editor.style.fontFamily='';
              }
            } catch(err){ clog('defaultFont error', err); window.CustomLabelStyle?.applyFontFamilyToSelection('', editor); }
            closeContextMenu(menu); try { window.CustomLabels?.save(); } catch{}
          },10);
        });
        menu.appendChild(defaultFontItem);

        let customFonts = {};
        try { if(window.fontManager) customFonts = await window.fontManager.getAllFonts(); } catch(err){ clog('getAllFonts error', err); }
        const systemFonts = [
          { name: 'ゴシック（sans-serif）', family: 'sans-serif' },
          { name: '明朝（serif）', family: 'serif' },
          { name: '等幅（monospace）', family: 'monospace' },
          { name: 'Arial', family: 'Arial, sans-serif' },
          { name: 'Times New Roman', family: 'Times New Roman, serif' },
          { name: 'メイリオ', family: 'Meiryo, sans-serif' },
          { name: 'ヒラギノ角ゴ', family: 'Hiragino Kaku Gothic Pro, sans-serif' }
        ];
        if(systemFonts.length){
          const systemFontLabel=document.createElement('div');
          systemFontLabel.textContent='システムフォント';
            systemFontLabel.style.cssText='padding:5px 15px;font-size:10px;color:#666;font-weight:bold;border-bottom:1px solid #eee;';
          menu.appendChild(systemFontLabel);
          systemFonts.forEach(font=>{
            const fontItem=document.createElement('div');
            fontItem.textContent=font.name;
            fontItem.style.cssText=`padding:6px 20px;cursor:pointer;font-size:11px;transition:background-color .2s;font-family:${font.family};`;
            fontItem.addEventListener('mouseenter',function(){ this.style.backgroundColor='#f0f0f0'; });
            fontItem.addEventListener('mouseleave',function(){ this.style.backgroundColor='transparent'; });
            fontItem.addEventListener('mousedown',e=>{ e.preventDefault(); e.stopPropagation(); });
            fontItem.addEventListener('click',e=>{ e.preventDefault(); e.stopPropagation(); setTimeout(()=>{ try { window.CustomLabelStyle?.applyFontFamilyToSelection(font.family, editor); } catch(err){ clog('fontFamily error', err);} closeContextMenu(menu); try { window.CustomLabels?.save(); } catch{} },10); });
            menu.appendChild(fontItem);
          });
        }
        if(Object.keys(customFonts).length){
          const customFontLabel=document.createElement('div');
          customFontLabel.textContent='カスタムフォント';
          customFontLabel.style.cssText='padding:5px 15px;font-size:10px;color:#666;font-weight:bold;border-top:1px solid #eee;border-bottom:1px solid #eee;';
          menu.appendChild(customFontLabel);
          Object.keys(customFonts).forEach(fontName=>{
            const fontItem=document.createElement('div');
            fontItem.textContent=fontName;
            fontItem.style.cssText=`padding:6px 20px;cursor:pointer;font-size:11px;transition:background-color .2s;font-family:"${fontName}", sans-serif;`;
            fontItem.addEventListener('mouseenter',function(){ this.style.backgroundColor='#f0f0f0'; });
            fontItem.addEventListener('mouseleave',function(){ this.style.backgroundColor='transparent'; });
            fontItem.addEventListener('mousedown',e=>{ e.preventDefault(); e.stopPropagation(); });
            fontItem.addEventListener('click',e=>{ e.preventDefault(); e.stopPropagation(); setTimeout(()=>{ try { window.CustomLabelStyle?.applyFontFamilyToSelection(fontName, editor); } catch(err){ clog('customFont error', err);} closeContextMenu(menu); try { window.CustomLabels?.save(); } catch{} },10); });
            menu.appendChild(fontItem);
          });
        }
      } catch(err){ clog('custom font section error', err); }
    }

    // 位置調整
    const menuRect = menu.getBoundingClientRect();
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;
    const scrollY = window.scrollY || window.pageYOffset;
    const margin = 10;
    let adjustedX = x;
    if (x + menuRect.width > viewportWidth) adjustedX = viewportWidth - menuRect.width - margin;
    if (adjustedX < margin) adjustedX = margin;
    let adjustedY = y;
    if (y + menuRect.height > viewportHeight) adjustedY = y - menuRect.height - margin;
    if (adjustedY < scrollY + margin) adjustedY = (y + menuRect.height <= viewportHeight)? y : (scrollY + margin);
    adjustedX = Math.max(margin, Math.min(adjustedX, viewportWidth - menuRect.width - margin));
    adjustedY = Math.max(scrollY + margin, Math.min(adjustedY, scrollY + viewportHeight - menuRect.height - margin));
    menu.style.left = adjustedX + 'px';
    menu.style.top = adjustedY + 'px';
    menu.style.visibility='visible';
    setTimeout(()=>{ menu.style.opacity='1'; },10);
    return menu;
  }

  window.CustomLabelContextMenu = { createFontSizeMenu, closeContextMenu };
  window.createFontSizeMenu = window.createFontSizeMenu || createFontSizeMenu; // 後方互換
  window.closeContextMenu = window.closeContextMenu || closeContextMenu;       // 後方互換
})();

// =============================
// プレーンテキスト入力制限 & サニタイズ (setupTextOnlyEditor 移設)
// =============================
(function(){
  if(window.CustomLabelEditor && window.CustomLabelEditor.setupTextOnlyEditor && window.CustomLabelEditor.setupRichTextFormatting) return; // 冪等
  function setupTextOnlyEditor(editor){
    if(!editor || editor.dataset.plainOnlyInit) return; // 多重適用防止
    editor.dataset.plainOnlyInit = '1';
    try { debugLog?.('setupTextOnlyEditor(init)', editor); } catch{}
    editor.addEventListener('paste', e=>{
      try { e.preventDefault(); } catch{}
      let text='';
      try { text=(e.clipboardData||window.clipboardData).getData('text/plain'); } catch{}
      const clean=(text||'').replace(/<[^>]*>/g,'');
      const sel=window.getSelection(); if(!sel?.rangeCount) return;
      const range=sel.getRangeAt(0); range.deleteContents();
      clean.split('\n').forEach((line,i)=>{ if(i>0){ const br=document.createElement('br'); range.insertNode(br); range.setStartAfter(br);} if(line.length){ const tn=document.createTextNode(line); range.insertNode(tn); range.setStartAfter(tn);} });
      range.collapse(true); sel.removeAllRanges(); sel.addRange(range);
    });
    editor.addEventListener('drop', e=>{
      try { e.preventDefault(); } catch{}
      if(e.dataTransfer?.files?.length) return false; // ファイル不可
      const text=e.dataTransfer?.getData('text/plain'); if(!text) return;
      const clean=text.replace(/<[^>]*>/g,'');
      const sel=window.getSelection(); if(!sel?.rangeCount) return; const range=sel.getRangeAt(0); range.deleteContents();
      clean.split('\n').forEach((line,i)=>{ if(i>0){ const br=document.createElement('br'); range.insertNode(br); range.setStartAfter(br);} if(line.length){ const tn=document.createTextNode(line); range.insertNode(tn); range.setStartAfter(tn);} });
      range.collapse(true); sel.removeAllRanges(); sel.addRange(range);
    });
    editor.addEventListener('dragover', e=>{ if(e.dataTransfer?.types?.includes('Files')){ e.dataTransfer.dropEffect='none'; return false;} e.preventDefault(); });
    const observer=new MutationObserver(muts=>{
      muts.forEach(m=>{ if(m.type==='childList'){ m.addedNodes.forEach(n=>{ if(n.nodeType===Node.ELEMENT_NODE){ const tag=n.tagName; if(tag==='IMG'||tag==='VIDEO'||tag==='AUDIO'){ n.remove(); } } }); } });
    });
    observer.observe(editor,{childList:true,subtree:true});
  }
  function setupRichTextFormatting(editor){
    if(!editor || editor.dataset.richFormattingInit) return;
    editor.dataset.richFormattingInit='1';
    let isComposing=false;
    editor.addEventListener('compositionstart',()=>{ isComposing=true; debugLog?.('[editor] compositionstart'); });
    editor.addEventListener('compositionend',()=>{ isComposing=false; debugLog?.('[editor] compositionend'); });
    const insertLineBreak=(source='unknown')=>{
      const sel=window.getSelection(); if(!sel?.rangeCount) return; const range=sel.getRangeAt(0); const before=editor.innerHTML;
      let atEnd=false; try { const tail=document.createRange(); tail.selectNodeContents(editor); tail.setStart(range.endContainer, range.endOffset); atEnd = tail.toString().length===0; } catch{}
      range.deleteContents(); const br=document.createElement('br'); range.insertNode(br);
      if(atEnd){ const zw=document.createTextNode('\u200B'); if(br.parentNode){ if(!(br.nextSibling && br.nextSibling.nodeType===Node.TEXT_NODE && br.nextSibling.nodeValue.startsWith('\u200B'))){ br.parentNode.insertBefore(zw, br.nextSibling); range.setStartAfter(zw); } } }
      else { range.setStartAfter(br); }
      range.collapse(true); sel.removeAllRanges(); sel.addRange(range);
      debugLog?.(`[editor] insertLineBreak via ${source}`, { beforeLen: before.length, afterLen: editor.innerHTML.length, atEnd });
    };
    const supportsBeforeInput=('onbeforeinput' in editor);
    if(supportsBeforeInput){ editor.addEventListener('beforeinput', e=>{ if(e.inputType==='insertParagraph'){ if(isComposing||e.isComposing) return; e.preventDefault(); insertLineBreak('beforeinput'); }}); }
    editor.addEventListener('keydown', e=>{ if(e.key==='Enter'){ if(supportsBeforeInput) return; if(isComposing||e.isComposing) return; e.preventDefault(); insertLineBreak('keydown'); }});
    editor.addEventListener('contextmenu', async e=>{ try { e.preventDefault();
      // 既存メニューを先に閉じる（素早い連続右クリック対応）
      try { (window.closeContextMenu||window.CustomLabelContextMenu?.closeContextMenu)?.(); } catch{}
      const sel=window.getSelection(); const hasSel=!!sel && sel.toString().length>0;
      const menu=await (window.createFontSizeMenu? window.createFontSizeMenu(e.clientX,e.clientY,editor,hasSel): window.CustomLabelContextMenu?.createFontSizeMenu(e.clientX,e.clientY,editor,hasSel));
      if(menu) document.body.appendChild(menu);
      // 外側クリックで閉じる（キャプチャしすぎないよう1回性リスナー）
      setTimeout(()=>{ const closer=(ev)=>{ try { if(!menu.contains(ev.target)) (window.closeContextMenu||window.CustomLabelContextMenu?.closeContextMenu)?.(menu);} catch{} document.removeEventListener('mousedown', closer); document.removeEventListener('scroll', closer, true); window.removeEventListener('resize', closer); }; document.addEventListener('mousedown', closer); document.addEventListener('scroll', closer, true); window.addEventListener('resize', closer); },10);
    } catch(err){ console.error('contextmenu error', err);} });
  }
  window.CustomLabelEditor = { setupTextOnlyEditor, setupRichTextFormatting };
  window.setupTextOnlyEditor = window.setupTextOnlyEditor || setupTextOnlyEditor;
  window.setupRichTextFormatting = window.setupRichTextFormatting || setupRichTextFormatting;
  if(window.CustomLabels){ Object.assign(window.CustomLabels, { setupTextOnlyEditor, setupRichTextFormatting }); }
  if(window.__pendingEditorInit){ window.__pendingEditorInit.forEach(ed=>{ try { setupRichTextFormatting(ed); setupTextOnlyEditor(ed); } catch(e){ console.error('delayed init error', e);} }); delete window.__pendingEditorInit; }
})();
