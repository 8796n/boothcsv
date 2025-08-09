// カスタムラベル関連モジュール抽出
// このファイルは boothcsv.js から保守性向上のため分離されたコードです。
// 依存: CONSTANTS, debugLog, StorageManager, CSVAnalyzer など（boothcsv.js 内定義）

let isEditingCustomLabel = false; // 編集状態フラグ
let pendingUpdateTimer = null;    // 遅延プレビュータイマー

class CustomLabelCalculator {
  static calculateMultiSheetDistribution(totalLabels, skipCount) {
    const sheetsInfo = [];
    let remainingLabels = totalLabels;
    let currentSkip = skipCount;
    let sheetNumber = 1;
    while (remainingLabels > 0) {
      const availableInSheet = CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET - currentSkip;
      const labelsInThisSheet = Math.min(remainingLabels, availableInSheet);
      const remainingInSheet = availableInSheet - labelsInThisSheet;
      sheetsInfo.push({ sheetNumber, skipCount: currentSkip, labelCount: labelsInThisSheet, remainingCount: remainingInSheet, totalInSheet: currentSkip + labelsInThisSheet });
      remainingLabels -= labelsInThisSheet; currentSkip = 0; sheetNumber++;
    }
    return sheetsInfo;
  }
  static getLastSheetInfo(totalLabels, skipCount) { const sheetsInfo = this.calculateMultiSheetDistribution(totalLabels, skipCount); return sheetsInfo[sheetsInfo.length - 1] || null; }
  static calculateTotalSheets(totalLabels, skipCount) { return this.calculateMultiSheetDistribution(totalLabels, skipCount).length; }
}

function toggleCustomLabelRow(enabled) { const row = document.getElementById('customLabelRow'); if (row) row.style.display = enabled ? 'table-row' : 'none'; }

function initializeCustomLabels(customLabels) {
  const container = document.getElementById('customLabelsContainer'); if (!container) return;
  container.innerHTML = '';
  const instructionDiv = document.createElement('div'); instructionDiv.className = 'custom-labels-instructions'; instructionDiv.innerHTML = `
    <div class="instructions-content">
      <strong>カスタムラベル設定について：</strong><br>
      • 実際のラベルサイズ: 48.3mm × 25.3mm<br>
      • テキストのみ入力可能<br>
      • 改行: Enterキー<br>
      • 書式設定: 右クリックでコンテキストメニューを表示（太字、斜体、フォントサイズ変更、フォントの変更）
    </div>`; container.appendChild(instructionDiv);
  if (customLabels && customLabels.length > 0) {
    customLabels.forEach((label, index) => {
      addCustomLabelItem(label.html || label.text, label.count, index, label.enabled !== false);
      if (label.fontSize) { const item = container.children[index]; const editor = item?.querySelector('.rich-text-editor'); if (editor) { const fontSize = label.fontSize.toString().includes('pt') ? label.fontSize : label.fontSize + 'pt'; editor.style.fontSize = fontSize; } }
    });
  } else { addCustomLabelItem('', 1, 0, true); }
  updateCustomLabelsSummary().catch(console.error);
}

function addCustomLabelItem(text = '', count = 1, index = null, enabled = true) {
  debugLog && debugLog('addCustomLabelItem', { text, count, index });
  const container = document.getElementById('customLabelsContainer'); if (!container) return;
  const template = document.getElementById('customLabelItem'); if (!template) return;
  const item = template.content.cloneNode(true); const itemDiv = item.querySelector('.custom-label-item');
  const itemIndex = index !== null ? index : container.children.length; itemDiv.dataset.index = itemIndex;
  const checkbox = item.querySelector('.custom-label-enabled'); checkbox.id = `customLabel_${itemIndex}_enabled`; checkbox.dataset.index = itemIndex; checkbox.checked = enabled;
  const label = item.querySelector('.custom-label-item-title'); label.setAttribute('for', `customLabel_${itemIndex}_enabled`);
  const editor = item.querySelector('.rich-text-editor'); editor.dataset.index = itemIndex;
  const countInput = item.querySelector('input[type="number"]'); countInput.value = count; countInput.dataset.index = itemIndex;
  const removeBtn = item.querySelector('.btn-remove'); removeBtn.onclick = () => removeCustomLabelItem(itemIndex);
  container.appendChild(item);
  const editorElement = container.querySelector(`[data-index="${itemIndex}"].rich-text-editor`);
  if (editorElement) {
    if (text && text.trim() !== '') editorElement.innerHTML = text;
    setupRichTextFormatting(editorElement); setupTextOnlyEditor(editorElement);
    editorElement.addEventListener('focus', () => { isEditingCustomLabel = true; });
    editorElement.addEventListener('blur', () => { isEditingCustomLabel = false; scheduleDelayedPreviewUpdate(300); });
    editorElement.addEventListener('input', async () => {
      const item = editorElement.closest('.custom-label-item'); if (item && editorElement.textContent.trim() !== '') item.classList.remove('error');
      saveCustomLabels(); updateButtonStates && updateButtonStates(); await updateCustomLabelsSummary();
      if (!isEditingCustomLabel) { autoProcessCSV && autoProcessCSV(); } else { scheduleDelayedPreviewUpdate(1000); }
    });
  }
  const enabledCheckbox = container.querySelector(`[data-index="${itemIndex}"].custom-label-enabled`);
  if (enabledCheckbox) enabledCheckbox.addEventListener('change', async function(){ saveCustomLabels(); await updateCustomLabelsSummary(); autoProcessCSV && autoProcessCSV(); });
  const countInputElement = container.querySelector(`[data-index="${itemIndex}"][type="number"]`);
  if (countInputElement) countInputElement.addEventListener('input', async function(){ saveCustomLabels(); updateButtonStates && updateButtonStates(); await updateCustomLabelsSummary(); autoProcessCSV && autoProcessCSV(); });
  updateCustomLabelsSummary().catch(console.error);
}

function removeCustomLabelItem(index) {
  const container = document.getElementById('customLabelsContainer'); if (!container) return;
  const items = container.querySelectorAll('.custom-label-item'); if (items.length <= 1) { alert('最低1つのカスタムラベルは必要です。'); return; }
  const itemToRemove = container.querySelector(`[data-index="${index}"]`); if (itemToRemove) itemToRemove.remove();
  reindexCustomLabelItems(); saveCustomLabels(); updateCustomLabelsSummary().catch(console.error); updateButtonStates && updateButtonStates(); autoProcessCSV && autoProcessCSV();
}

function getCustomLabelsFromUI() {
  const container = document.getElementById('customLabelsContainer'); if (!container) return [];
  const items = container.querySelectorAll('.custom-label-item'); const labels = [];
  items.forEach(item => { const editor = item.querySelector('.rich-text-editor'); const countInput = item.querySelector('input[type="number"]'); const enabledCheckbox = item.querySelector('.custom-label-enabled'); const text = editor.innerHTML.trim(); const count = parseInt(countInput.value, 10) || 1; const enabled = enabledCheckbox ? enabledCheckbox.checked : true; const computedStyle = window.getComputedStyle(editor); const fontSize = computedStyle.fontSize || '12pt'; labels.push({ text, count, fontSize, html: text, enabled }); });
  return labels;
}

function saveCustomLabels() { const labels = getCustomLabelsFromUI(); StorageManager.setCustomLabels(labels); }

async function updateCustomLabelsSummary() {
  const labels = getCustomLabelsFromUI(); const enabledLabels = labels.filter(l => l.enabled); const totalCount = enabledLabels.reduce((s,l)=>s+l.count,0); const skipCount = parseInt(document.getElementById('labelskipnum')?.value,10)||0; const fileInput = document.getElementById('file'); const summary = document.getElementById('customLabelsSummary'); if(!summary) return;
  if (totalCount === 0) { if (fileInput?.files.length > 0) { try { const csvInfo = await CSVAnalyzer.getFileInfo(fileInput.files[0]); const remaining = CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET - skipCount - csvInfo.rowCount; summary.innerHTML = `カスタムラベルなし。<br>44面シート中 スキップ${skipCount} + CSV${csvInfo.rowCount} = ${skipCount + csvInfo.rowCount}面使用済み。<br>残り${Math.max(0, remaining)}面設定可能。`; } catch { summary.innerHTML = 'カスタムラベルなし。<br>CSVファイル選択済み（行数解析中...）。'; } } else { summary.innerHTML = `カスタムラベルなし。<br>44面シート中 スキップ${skipCount}面使用済み。`; } summary.style.color='#666'; summary.style.fontWeight='normal'; return; }
  if (fileInput?.files.length > 0) { try { const csvInfo = await CSVAnalyzer.getFileInfo(fileInput.files[0]); const csvRowCount = csvInfo.rowCount; const totalLabels = csvRowCount + totalCount; const sheetsInfo = CustomLabelCalculator.calculateMultiSheetDistribution(totalLabels, skipCount); const totalSheets = sheetsInfo.length; const lastSheet = sheetsInfo[sheetsInfo.length - 1]; if (totalSheets === 1) { summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>1シート使用: スキップ${skipCount} + CSV${csvRowCount} + カスタム${totalCount} = ${skipCount + csvRowCount + totalCount}面<br>最終シート残り${lastSheet.remainingCount}面。`; } else { summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>${totalSheets}シート使用: CSV${csvRowCount} + カスタム${totalCount} = ${csvRowCount + totalCount}面<br>最終シート残り${lastSheet.remainingCount}面。`; } summary.style.color='#666'; summary.style.fontWeight='normal'; } catch (e) { console.error('CSV解析エラー:', e); summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>CSVファイル選択済み（行数解析エラー）<br>CSV処理実行後に最終配置が決定されます。`; summary.style.color='#ffc107'; summary.style.fontWeight='normal'; } }
  else { const sheetsInfo = CustomLabelCalculator.calculateMultiSheetDistribution(totalCount, skipCount); const totalSheets = sheetsInfo.length; const lastSheet = sheetsInfo[sheetsInfo.length - 1]; if (totalSheets === 1) { summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>1シート使用: スキップ${skipCount} + カスタム${totalCount} = ${skipCount + totalCount}面<br>最終シート残り${lastSheet.remainingCount}面。`; } else { summary.innerHTML = `合計 ${totalCount}面のカスタムラベル。<br>${totalSheets}シート使用: カスタム${totalCount}面<br>最終シート残り${lastSheet.remainingCount}面。`; } summary.style.color='#666'; summary.style.fontWeight='normal'; }
}

async function adjustCustomLabelsForTotal(customLabels, maxTotalLabels) { let remaining = maxTotalLabels; for (let i=0;i<customLabels.length;i++){ if(remaining<=0) customLabels[i].count=0; else if(customLabels[i].count>remaining) customLabels[i].count=remaining; remaining -= customLabels[i].count; } const container=document.getElementById('customLabelsContainer'); const items=container?.querySelectorAll('.custom-label-item')||[]; items.forEach((item,index)=>{ const countInput=item.querySelector('input[type="number"]'); if(customLabels[index]) countInput.value=customLabels[index].count; }); saveCustomLabels(); await updateCustomLabelsSummary(); }

function setupCustomLabelEvents(){ const initialEditor=document.getElementById('initialCustomLabelEditor'); if(initialEditor){ initialEditor.addEventListener('focus',()=>{isEditingCustomLabel=true;}); initialEditor.addEventListener('blur',()=>{isEditingCustomLabel=false; scheduleDelayedPreviewUpdate(300);}); } const customLabelEnable=document.getElementById('customLabelEnable'); if(customLabelEnable){ customLabelEnable.addEventListener('change', async function(){ toggleCustomLabelRow(this.checked); await StorageManager.set(StorageManager.KEYS.CUSTOM_LABEL_ENABLE,this.checked); updateButtonStates && updateButtonStates(); autoProcessCSV && autoProcessCSV(); }); } const addButton=document.getElementById('addCustomLabelBtn'); if(addButton){ addButton.addEventListener('click', async ()=>{ addCustomLabelItem('',1,null,true); saveCustomLabels(); updateButtonStates && updateButtonStates(); autoProcessCSV && autoProcessCSV(); }); } const clearButton=document.getElementById('clearCustomLabelsBtn'); if(clearButton){ clearButton.addEventListener('click', async ()=>{ if(confirm('本当に全てのカスタムラベルを削除しますか？')){ await clearAllCustomLabels(); autoProcessCSV && autoProcessCSV(); } }); } }

async function clearAllCustomLabels(){ const container=document.getElementById('customLabelsContainer'); const items=container?.querySelectorAll('.custom-label-item')||[]; items.forEach(item=>item.remove()); addCustomLabelItem('',1,0,true); saveCustomLabels(); await updateCustomLabelsSummary(); updateButtonStates && updateButtonStates(); }

function hasCustomLabelsWithContent(){ const labels=getCustomLabelsFromUI(); return labels.length>0 && labels.some(l=>l.text.trim()!==''); }
function hasEmptyEnabledCustomLabels(){ const enable=document.getElementById('customLabelEnable'); if(!enable?.checked) return false; const labels=getCustomLabelsFromUI(); return labels.filter(l=>l.enabled).some(l=>l.text.trim()===''); }
function removeEmptyCustomLabels(){ const labels=getCustomLabelsFromUI(); const container=document.getElementById('customLabelsContainer'); if(!container) return; const items=container.querySelectorAll('.custom-label-item'); let removed=0; for(let i=labels.length-1;i>=0;i--){ const label=labels[i]; if(label.enabled && label.text.trim()===''){ const item=items[i]; if(item){ item.remove(); removed++; } } } reindexCustomLabelItems(); saveCustomLabels(); updateCustomLabelsSummary().catch(console.error); if(container.querySelectorAll('.custom-label-item').length===0) addCustomLabelItem('',1,0,true); if(removed>0) alert(`未設定のカスタムラベル ${removed} 項目を削除しました。`); }
function highlightEmptyCustomLabels(){ const labels=getCustomLabelsFromUI(); const container=document.getElementById('customLabelsContainer'); if(!container) return; const items=container.querySelectorAll('.custom-label-item'); items.forEach((item,i)=>{ if(labels[i] && labels[i].enabled && labels[i].text.trim()==='') item.classList.add('error'); else item.classList.remove('error'); }); }
function clearCustomLabelHighlights(){ const container=document.getElementById('customLabelsContainer'); if(!container) return; container.querySelectorAll('.custom-label-item').forEach(i=>i.classList.remove('error')); }
function reindexCustomLabelItems(){ const container=document.getElementById('customLabelsContainer'); if(!container) return; const items=container.querySelectorAll('.custom-label-item'); items.forEach((item,index)=>{ item.dataset.index=index; const del=item.querySelector('.btn-danger'); if(del) del.setAttribute('onclick',`removeCustomLabelItem(${index})`); }); }

function setupRichTextFormatting(editor){ if(!editor) return; let isComposing=false; editor.addEventListener('compositionstart',()=>{isComposing=true;}); editor.addEventListener('compositionend',()=>{isComposing=false;}); const insertLineBreak=()=>{ const selection=window.getSelection(); if(!selection||selection.rangeCount===0) return; const range=selection.getRangeAt(0); let atEnd=false; try{ const tail=document.createRange(); tail.selectNodeContents(editor); tail.setStart(range.endContainer, range.endOffset); const remaining=tail.toString(); atEnd=remaining.length===0; }catch{} range.deleteContents(); const br=document.createElement('br'); range.insertNode(br); if(atEnd){ const zwsp=document.createTextNode('\u200B'); if(br.parentNode){ if(!(br.nextSibling && br.nextSibling.nodeType===Node.TEXT_NODE && br.nextSibling.nodeValue.startsWith('\u200B'))){ br.parentNode.insertBefore(zwsp, br.nextSibling); range.setStartAfter(zwsp);} } } else { range.setStartAfter(br); } range.collapse(true); selection.removeAllRanges(); selection.addRange(range); };
  const supportsBeforeInput='onbeforeinput' in editor; if(supportsBeforeInput){ editor.addEventListener('beforeinput',e=>{ if(e.inputType==='insertParagraph'){ if(isComposing||e.isComposing) return; e.preventDefault(); insertLineBreak(); } }); }
  editor.addEventListener('keydown',e=>{ if(e.key==='Enter'){ if(supportsBeforeInput) return; if(isComposing||e.isComposing) return; e.preventDefault(); insertLineBreak(); } });
  editor.addEventListener('contextmenu', async e=>{ e.preventDefault(); const selection=window.getSelection(); const hasSel=selection.toString().length>0; const menu=await createFontSizeMenu(e.clientX,e.clientY,editor,hasSel); document.body.appendChild(menu); setTimeout(()=>{ document.addEventListener('click',function close(){ closeContextMenu(menu); document.removeEventListener('click',close);});},100); }); }
function closeContextMenu(menu){ if(menu&&menu.parentNode){ menu.style.opacity='0'; setTimeout(()=>{ if(menu.parentNode) menu.parentNode.removeChild(menu); },200); } }
async function createFontSizeMenu(x,y,editor,hasSelection=true){ const menu=document.createElement('div'); menu.style.cssText='position:fixed;background:#fff;border:1px solid #ccc;border-radius:6px;padding:8px 0;box-shadow:0 4px 20px rgba(0,0,0,0.15);z-index:10000;min-width:160px;max-width:250px;max-height:400px;overflow-y:auto;visibility:hidden;opacity:0;transition:opacity .2s ease;'; document.body.appendChild(menu); const format=[{label:'太字',command:'bold',style:'font-weight:bold;'},{label:'斜体',command:'italic',style:'font-style:italic;'},{label:'下線',command:'underline',style:'text-decoration:underline;'},{label:'すべてクリア',command:'clear',style:'color:#dc3545;font-weight:bold;'}]; const opts=hasSelection?format:[format[3]]; opts.forEach(op=>{ const d=document.createElement('div'); d.textContent=op.label; d.style.cssText=`padding:8px 15px;cursor:pointer;font-size:12px;transition:background-color .2s;${op.style}`; d.addEventListener('mouseenter',function(){this.style.backgroundColor='#f0f0f0';}); d.addEventListener('mouseleave',function(){this.style.backgroundColor='transparent';}); d.addEventListener('mousedown',e=>{e.preventDefault();e.stopPropagation();}); d.addEventListener('click',e=>{e.preventDefault();e.stopPropagation(); setTimeout(()=>{applyFormatToSelection(op.command,editor); closeContextMenu(menu);},0);}); menu.appendChild(d); }); const sizes=['8pt','9pt','10pt','11pt','12pt','13pt','14pt','15pt','16pt','18pt','20pt']; if(hasSelection){ const sep=document.createElement('div'); sep.style.cssText='margin:4px 0;border-top:1px solid #e0e0e0;'; menu.appendChild(sep); sizes.forEach(size=>{ const d=document.createElement('div'); d.textContent=size; d.style.cssText='padding:6px 15px;cursor:pointer;font-size:12px;display:flex;justify-content:space-between;'; d.addEventListener('mouseenter',function(){this.style.backgroundColor='#f0f0f0';}); d.addEventListener('mouseleave',function(){this.style.backgroundColor='transparent';}); d.addEventListener('mousedown',e=>{e.preventDefault();e.stopPropagation();}); d.addEventListener('click',e=>{e.preventDefault();e.stopPropagation(); setTimeout(()=>{applyFontSizeToSelection(size,editor); closeContextMenu(menu);},0);}); menu.appendChild(d); }); }
  const rect=menu.getBoundingClientRect(); let ax=x, ay=y; const margin=8, vw=window.innerWidth, vh=window.innerHeight, sy=window.scrollY||window.pageYOffset; if(x+rect.width>vw) ax=vw-rect.width-margin; if(y+rect.height>vh) ay=y-rect.height-margin; if(ay<sy+margin) ay=sy+margin; ax=Math.max(margin,Math.min(ax,vw-rect.width-margin)); ay=Math.max(sy+margin,Math.min(ay,sy+vh-rect.height-margin)); menu.style.left=`${ax}px`; menu.style.top=`${ay}px`; menu.style.visibility='visible'; setTimeout(()=>{menu.style.opacity='1';},10); return menu; }
function applyFormatToSelection(command,editor){ if(command==='clear'){ clearAllContent(editor); return; } try{ let c; switch(command){ case 'bold': c='bold'; break; case 'italic': c='italic'; break; case 'underline': c='underline'; break; default: return; } document.execCommand(c,false,null); }catch{ applyFormatToSelectionFallback(command,editor);} editor.focus(); }
function applyFormatToSelectionFallback(command,editor){ const selection=window.getSelection(); if(selection.rangeCount===0||selection.isCollapsed) return; const range=selection.getRangeAt(0); if(isSelectionFormatted(range,command)) removeFormatFromSelection(range,command); else applyFormatToRange(range,command); }
function isSelectionFormatted(range,command){ const tag=getTargetTagName(command); if(!tag) return false; let node=range.startContainer; while(node){ if(node.nodeType===1 && node.nodeName===tag) return true; node=node.parentNode; } return false; }
function getTargetTagName(command){ switch(command){ case 'bold': return 'B'; case 'italic': return 'I'; case 'underline': return 'U'; default: return null; } }
function applyFormatToRange(range,command){ const w=document.createElement(getTargetTagName(command)); w.appendChild(range.extractContents()); range.insertNode(w); }
function removeFormatFromSelection(range,command){ const tag=getTargetTagName(command); const container=range.commonAncestorContainer; const elements=(container.nodeType===1?container:container.parentNode).querySelectorAll(tag); elements.forEach(el=>{ if(range.intersectsNode(el)) unwrapElement(el); }); }
function unwrapElement(el){ const parent=el.parentNode; while(el.firstChild) parent.insertBefore(el.firstChild,el); parent.removeChild(el); }
function applyFontSizeToSelection(size,editor){ const selection=window.getSelection(); if(!selection.rangeCount) return; const range=selection.getRangeAt(0); if(selection.isCollapsed){ editor.style.fontSize=size; return; } const span=document.createElement('span'); span.style.fontSize=size; span.appendChild(range.extractContents()); range.insertNode(span); selection.removeAllRanges(); const nr=document.createRange(); nr.selectNodeContents(span); selection.addRange(nr); }
function clearAllContent(editor){ if(confirm('このカスタムラベルの内容と書式をすべてクリアしますか？')){ editor.innerHTML=''; editor.style.fontSize='12pt'; editor.style.lineHeight='1.2'; editor.style.textAlign='center'; editor.focus(); saveCustomLabels(); } }
function setupTextOnlyEditor(editor){ if(!editor) return; editor.addEventListener('paste', e=>{ e.preventDefault(); const text=(e.clipboardData||window.clipboardData).getData('text/plain'); const clean=text.replace(/<[^>]*>/g,''); const selection=window.getSelection(); if(selection.rangeCount>0){ const range=selection.getRangeAt(0); range.deleteContents(); const lines=clean.split('\n'); lines.forEach((line,i)=>{ if(i>0){ const br=document.createElement('br'); range.insertNode(br); range.setStartAfter(br);} if(line.length>0){ const tn=document.createTextNode(line); range.insertNode(tn); range.setStartAfter(tn);} }); range.collapse(true); selection.removeAllRanges(); selection.addRange(range);} }); editor.addEventListener('drop', e=>{ e.preventDefault(); if(e.dataTransfer.files.length>0) return false; const text=e.dataTransfer.getData('text/plain'); if(text){ const clean=text.replace(/<[^>]*>/g,''); const selection=window.getSelection(); if(selection.rangeCount>0){ const range=selection.getRangeAt(0); range.deleteContents(); const lines=clean.split('\n'); lines.forEach((line,i)=>{ if(i>0){ const br=document.createElement('br'); range.insertNode(br); range.setStartAfter(br);} if(line.length>0){ const tn=document.createTextNode(line); range.insertNode(tn); range.setStartAfter(tn);} }); range.collapse(true); selection.removeAllRanges(); selection.addRange(range);} } }); }

function scheduleDelayedPreviewUpdate(delay=500){ if(pendingUpdateTimer) clearTimeout(pendingUpdateTimer); pendingUpdateTimer=setTimeout(async()=>{ if(typeof updateCustomLabelsPreview==='function'){ await updateCustomLabelsPreview(); } pendingUpdateTimer=null; },delay); }

window.CustomLabels = { initializeCustomLabels, addCustomLabelItem, removeCustomLabelItem, getCustomLabelsFromUI, saveCustomLabels, updateCustomLabelsSummary, adjustCustomLabelsForTotal, setupCustomLabelEvents, clearAllCustomLabels, hasCustomLabelsWithContent, hasEmptyEnabledCustomLabels, removeEmptyCustomLabels, highlightEmptyCustomLabels, clearCustomLabelHighlights, reindexCustomLabelItems, toggleCustomLabelRow, scheduleDelayedPreviewUpdate, CustomLabelCalculator, isEditing: () => isEditingCustomLabel };
