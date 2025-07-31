// アプリケーション定数
const CONSTANTS = {
  LABEL: {
    TOTAL_LABELS_PER_SHEET: 44,
    LABELS_PER_ROW: 4
  },
  IMAGE: {
    SUPPORTED_FORMATS: /^image\/(jpeg|png|svg\+xml)$/,
    ACCEPTED_TYPES: 'image/jpeg, image/png, image/svg+xml'
  },
  QR: {
    EXPECTED_PARTS: 3
  },
  CSV: {
    PRODUCT_COLUMN: "商品ID / 数量 / 商品名",
    ORDER_NUMBER_COLUMN: "注文番号",
    PAYMENT_DATE_COLUMN: "支払い日時"
  }
};

// デバッグフラグの取得
const DEBUG_MODE = (() => {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get('debug') === '1';
})();

// デバッグログ用関数
function debugLog(...args) {
  if (DEBUG_MODE) {
    console.log(...args);
  }
}

// localStorage操作を管理するクラス
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
    CUSTOM_LABELS: 'customLabels', // 複数のカスタムラベルを保存
    ORDER_IMAGE_ENABLE: 'orderImageEnable' // 注文画像表示の有効/無効
  };

  // 設定値の取得
  static getSettings() {
    return {
      labelyn: this.get(this.KEYS.LABEL_SETTING, 'true') !== 'false',
      labelskip: parseInt(this.get(this.KEYS.LABEL_SKIP, '0'), 10),
      sortByPaymentDate: this.get(this.KEYS.SORT_BY_PAYMENT) === 'true',
      customLabelEnable: this.get(this.KEYS.CUSTOM_LABEL_ENABLE) === 'true',
      customLabelText: this.get(this.KEYS.CUSTOM_LABEL_TEXT, ''),
      customLabelCount: parseInt(this.get(this.KEYS.CUSTOM_LABEL_COUNT, '1'), 10),
      customLabels: this.getCustomLabels(),
      orderImageEnable: this.get(this.KEYS.ORDER_IMAGE_ENABLE) === 'true'
    };
  }

  // 複数のカスタムラベルを取得
  static getCustomLabels() {
    const labelsData = this.get(this.KEYS.CUSTOM_LABELS);
    if (!labelsData) return [];
    
    try {
      return JSON.parse(labelsData);
    } catch (e) {
      console.warn('Custom labels data parsing failed:', e);
      return [];
    }
  }

  // 複数のカスタムラベルを保存
  static setCustomLabels(labels) {
    this.set(this.KEYS.CUSTOM_LABELS, JSON.stringify(labels));
  }

  // 設定値の保存
  static saveSettings(settings) {
    Object.entries(settings).forEach(([key, value]) => {
      if (this.KEYS[key.toUpperCase()]) {
        this.set(this.KEYS[key.toUpperCase()], String(value));
      }
    });
  }

  // 注文画像の取得
  static getOrderImage(orderNumber = null) {
    const key = orderNumber ? 
      `${this.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : 
      this.KEYS.GLOBAL_ORDER_IMAGE;
    return this.get(key);
  }

  // 注文画像の保存
  static setOrderImage(imageData, orderNumber = null) {
    const key = orderNumber ? 
      `${this.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : 
      this.KEYS.GLOBAL_ORDER_IMAGE;
    this.set(key, imageData);
  }

  // 注文画像の削除
  static removeOrderImage(orderNumber = null) {
    const key = orderNumber ? 
      `${this.KEYS.ORDER_IMAGE_PREFIX}${orderNumber}` : 
      this.KEYS.GLOBAL_ORDER_IMAGE;
    this.remove(key);
  }

  // QR画像の一括削除
  static clearQRImages() {
    Object.keys(localStorage).forEach(key => {
      const value = localStorage.getItem(key);
      if (value?.includes('qrimage')) {
        this.remove(key);
      }
    });
  }

  // 注文画像の一括削除
  static clearOrderImages() {
    Object.keys(localStorage).forEach(key => {
      if (key === this.KEYS.GLOBAL_ORDER_IMAGE || 
          key.startsWith(this.KEYS.ORDER_IMAGE_PREFIX)) {
        this.remove(key);
      }
    });
  }

  // QRコードデータの取得
  static getQRData(orderNumber) {
    const data = this.get(orderNumber);
    if (!data) return null;
    
    try {
      return JSON.parse(data);
    } catch (e) {
      console.warn('QR data parsing failed:', e);
      return null;
    }
  }

  // QRコードデータの保存
  static setQRData(orderNumber, qrData) {
    this.set(orderNumber, JSON.stringify(qrData));
  }

  // 基本的なlocalStorage操作
  static get(key, defaultValue = null) {
    const value = localStorage.getItem(key);
    return value !== null ? value : defaultValue;
  }

  static set(key, value) {
    localStorage.setItem(key, value);
  }

  static remove(key) {
    localStorage.removeItem(key);
  }
}

window.addEventListener("load", function(){
  const settings = StorageManager.getSettings();
  
  debugLog(settings.labelyn, settings.labelskip);
  
  document.getElementById("labelyn").checked = settings.labelyn;
  document.getElementById("labelskipnum").value = settings.labelskip;
  document.getElementById("sortByPaymentDate").checked = settings.sortByPaymentDate;
  document.getElementById("customLabelEnable").checked = settings.customLabelEnable;
  document.getElementById("orderImageEnable").checked = settings.orderImageEnable;

  // カスタムラベル行の表示/非表示
  toggleCustomLabelRow(settings.customLabelEnable);

  // 注文画像行の表示/非表示
  toggleOrderImageRow(settings.orderImageEnable);

  // 複数のカスタムラベルを初期化
  initializeCustomLabels(settings.customLabels);

   // 画像ドロップゾーンの初期化
  const imageDropZoneElement = document.getElementById('imageDropZone');
  const imageDropZone = createOrderImageDropZone();
  imageDropZoneElement.appendChild(imageDropZone.element);
  window.orderImageDropZone = imageDropZone;

  // 全ての画像をクリアするボタンのイベントリスナーを追加
  const clearAllButton = document.getElementById('clearAllButton');
  clearAllButton.onclick = () => {
    if (confirm('本当に全てのQR画像をクリアしますか？')) {
      StorageManager.clearQRImages();
      alert('全てのQR画像をクリアしました');
      location.reload();
    }
  };

  // 全ての注文画像をクリアするボタンのイベントリスナーを追加
  const clearAllOrderImagesButton = document.getElementById('clearAllOrderImagesButton');
  clearAllOrderImagesButton.onclick = () => {
    if (confirm('本当に全ての注文画像（グローバル画像と個別画像）をクリアしますか？')) {
      StorageManager.clearOrderImages();
      alert('全ての注文画像をクリアしました');
      location.reload();
    }
  };

   // チェックボックスの状態が変更されたときにlocalStorageに保存
   document.getElementById("sortByPaymentDate").addEventListener("change", function() {
     StorageManager.set(StorageManager.KEYS.SORT_BY_PAYMENT, this.checked);
   });

   // 注文画像表示機能のイベントリスナー
   document.getElementById("orderImageEnable").addEventListener("change", function() {
     StorageManager.set(StorageManager.KEYS.ORDER_IMAGE_ENABLE, this.checked);
     toggleOrderImageRow(this.checked);
     
     // 画像表示をリアルタイムで更新
     updateAllOrderImagesVisibility(this.checked);
   });

  // カスタムラベル機能のイベントリスナー（遅延実行）
  setTimeout(function() {
    setupCustomLabelEvents();
  }, 100);

  // ボタンの初期状態を設定
  updateButtonStates();

  // スキップ数変更時の処理を追加
  document.getElementById("labelskipnum").addEventListener("input", function() {
    updateButtonStates();
  });

}, false);

function clickstart() {
  clearPreviousResults();
  const config = getConfigFromUI();
  
  Papa.parse(config.file, {
    header: true,
    skipEmptyLines: true,
    complete: function(results) {
      processCSVResults(results, config);
    }
  });
}

function executeCustomLabelsOnly() {
  clearPreviousResults();
  const config = getConfigFromUI();
  
  // カスタムラベルが有効でない場合は警告
  if (!config.customLabelEnable) {
    alert('カスタムラベル機能を有効にしてください。');
    return;
  }
  
  // カスタムラベルの内容が空の場合は警告
  if (!config.customLabels || config.customLabels.length === 0) {
    alert('印刷する文字列を入力してください。');
    return;
  }
  
  // 有効なカスタムラベルがあるかチェック
  const validLabels = config.customLabels.filter(label => label.text.trim() !== '');
  if (validLabels.length === 0) {
    alert('印刷する文字列を入力してください。');
    return;
  }
  
  // カスタムラベルのみを処理
  processCustomLabelsOnly(config);
}

function clearPreviousResults() {
  for (let sheet of document.querySelectorAll('section')) {
    sheet.parentNode.removeChild(sheet);
  }
}

function getConfigFromUI() {
  const file = document.getElementById("file").files[0];
  const labelyn = document.getElementById("labelyn").checked;
  const labelskip = document.getElementById("labelskipnum").value;
  const sortByPaymentDate = document.getElementById("sortByPaymentDate").checked;
  const customLabelEnable = document.getElementById("customLabelEnable").checked;
  
  // 複数のカスタムラベルを取得（有効なもののみ）
  const allCustomLabels = getCustomLabelsFromUI();
  const customLabels = customLabelEnable ? allCustomLabels.filter(label => label.enabled) : [];
  
  StorageManager.set(StorageManager.KEYS.LABEL_SETTING, labelyn);
  StorageManager.set(StorageManager.KEYS.LABEL_SKIP, labelskip);
  StorageManager.set(StorageManager.KEYS.CUSTOM_LABEL_ENABLE, customLabelEnable);
  StorageManager.setCustomLabels(allCustomLabels); // 全てのラベルを保存（有効/無効問わず）
  
  const labelarr = [];
  const labelskipNum = parseInt(labelskip, 10) || 0;
  if (labelskipNum > 0) {
    for (let i = 0; i < labelskipNum; i++) {
      labelarr.push("");
    }
  }
  
  return { 
    file, 
    labelyn, 
    labelskip, 
    sortByPaymentDate, 
    labelarr, 
    customLabelEnable, 
    customLabels
  };
}

function processCSVResults(results, config) {
  // CSV行数を取得
  const csvRowCount = results.data.length;
  
  // 複数カスタムラベルの総枚数を計算
  const totalCustomLabelCount = config.customLabels.reduce((sum, label) => sum + label.count, 0);
  
  // カスタムラベル枚数の上限を再計算・調整
  const skipCount = parseInt(config.labelskip, 10) || 0;
  const maxCustomLabels = Math.max(0, CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET - skipCount - csvRowCount);
  
  if (config.customLabelEnable && totalCustomLabelCount > maxCustomLabels) {
    adjustCustomLabelsForCSV(config.customLabels, maxCustomLabels);
    
    if (maxCustomLabels === 0) {
      alert(`CSVデータが${csvRowCount}行、スキップが${skipCount}枚で44枚シートが満杯のため、カスタムラベルは印刷できません。`);
    } else {
      alert(`CSVデータとスキップ分を考慮して、カスタムラベル総枚数を${maxCustomLabels}枚に調整しました。`);
    }
  }

  // データの並び替え
  if (config.sortByPaymentDate) {
    results.data.sort((a, b) => {
      const timeA = a[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      const timeB = b[CONSTANTS.CSV.PAYMENT_DATE_COLUMN] || "";
      return timeA.localeCompare(timeB);
    });
  }

  // 注文明細の生成
  generateOrderDetails(results.data, config.labelarr);
  
  // ラベル生成（注文分＋カスタムラベル）
  if (config.labelyn) {
    let totalLabelArray = [...config.labelarr];
    
    // カスタムラベルが有効な場合は追加
    if (config.customLabelEnable && config.customLabels.length > 0) {
      for (const customLabel of config.customLabels) {
        for (let i = 0; i < customLabel.count; i++) {
          totalLabelArray.push({ 
            type: 'custom', 
            content: customLabel.html || customLabel.text,
            fontSize: customLabel.fontSize || '10pt'
          });
        }
      }
    }
    
    if (totalLabelArray.length > 0) {
      generateLabels(totalLabelArray);
    }
  }
  
  // 印刷枚数の表示
  showPrintSummary();
  
  // ボタンの状態を更新
  updateButtonStates();
}

function processCustomLabelsOnly(config) {
  // スキップラベルの配列を作成
  const labelarr = [];
  const labelskipNum = parseInt(config.labelskip, 10) || 0;
  const maxAvailableLabels = CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET - labelskipNum;
  
  // 複数カスタムラベルの総枚数を計算
  const totalCustomLabelCount = config.customLabels.reduce((sum, label) => sum + label.count, 0);
  
  // カスタムラベル枚数が上限を超えている場合は調整
  if (totalCustomLabelCount > maxAvailableLabels) {
    adjustCustomLabelsForTotal(config.customLabels, maxAvailableLabels);
    
    if (maxAvailableLabels === 0) {
      alert(`スキップが${labelskipNum}枚で44枚シートが満杯のため、カスタムラベルは印刷できません。`);
      return;
    } else {
      alert(`スキップ分を考慮して、カスタムラベル総枚数を${maxAvailableLabels}枚に調整しました。`);
    }
  }
  
  if (labelskipNum > 0) {
    for (let i = 0; i < labelskipNum; i++) {
      labelarr.push("");
    }
  }
  
  // カスタムラベルを追加
  if (config.customLabelEnable && config.customLabels.length > 0) {
    for (const customLabel of config.customLabels) {
      for (let i = 0; i < customLabel.count; i++) {
        labelarr.push({ 
          type: 'custom', 
          content: customLabel.html || customLabel.text,
          fontSize: customLabel.fontSize || '10pt'
        });
      }
    }
  }
  
  // ラベル生成
  if (labelarr.length > 0) {
    generateLabels(labelarr);
  }
  
  // 印刷枚数の表示（カスタムラベル専用）
  const adjustedTotalCount = config.customLabels.reduce((sum, label) => sum + label.count, 0);
  showCustomLabelPrintSummary(adjustedTotalCount, labelskipNum);
  
  // ボタンの状態を更新
  updateButtonStates();
}

function generateOrderDetails(data, labelarr) {
  const tOrder = document.querySelector('#注文明細');
  
  for (let row of data) {
    const cOrder = document.importNode(tOrder.content, true);
    let orderNumber = '';
    
    // 注文情報の設定
    orderNumber = setOrderInfo(cOrder, row, labelarr);
    
    // 個別画像ドロップゾーンの作成
    createIndividualImageDropZone(cOrder, orderNumber);
    
    // 商品項目の処理
    processProductItems(cOrder, row);
    
    // 画像表示の処理
    displayOrderImage(cOrder, orderNumber);
    
    document.body.appendChild(cOrder);
  }
}

function setOrderInfo(cOrder, row, labelarr) {
  let orderNumber = '';
  
  for (let c of Object.keys(row).filter(key => key != CONSTANTS.CSV.PRODUCT_COLUMN)) {
    const divc = cOrder.querySelector("." + c);
    if (divc) {
      if (c == CONSTANTS.CSV.ORDER_NUMBER_COLUMN) {
        orderNumber = row[c];
        divc.textContent = "注文番号 : " + row[c];
        labelarr.push(row[c]);
      } else if (row[c]) {
        divc.textContent = row[c];
      }
    }
  }
  
  return orderNumber;
}

function createIndividualImageDropZone(cOrder, orderNumber) {
  debugLog(`個別画像ドロップゾーン作成開始 - 注文番号: "${orderNumber}"`);
  
  const individualDropZoneContainer = cOrder.querySelector('.individual-image-dropzone');
  const individualZone = cOrder.querySelector('.individual-order-image-zone');
  
  debugLog(`ドロップゾーンコンテナ発見: ${!!individualDropZoneContainer}`);
  debugLog(`個別ゾーン発見: ${!!individualZone}`);
  
  // 注文画像表示機能が無効の場合は個別画像ゾーン全体を非表示
  const settings = StorageManager.getSettings();
  debugLog(`注文画像表示設定: ${settings.orderImageEnable}`);
  
  if (!settings.orderImageEnable) {
    if (individualZone) {
      individualZone.style.display = 'none';
      debugLog('注文画像表示が無効のため個別ゾーンを非表示');
    }
    return;
  }

  // 有効な場合は表示
  if (individualZone) {
    individualZone.style.display = 'block';
    debugLog('注文画像表示が有効のため個別ゾーンを表示');
  }

  if (individualDropZoneContainer && orderNumber) {
    // 注文番号を正規化（"注文番号 : 66463556" → "66463556"）
    const normalizedOrderNumber = orderNumber.replace(/^.*?:\s*/, '').trim();
    debugLog(`個別画像ドロップゾーン作成 - 元の注文番号: "${orderNumber}" → 正規化後: "${normalizedOrderNumber}"`);
    
    try {
      const individualImageDropZone = createIndividualOrderImageDropZone(normalizedOrderNumber);
      if (individualImageDropZone && individualImageDropZone.element) {
        individualDropZoneContainer.appendChild(individualImageDropZone.element);
        debugLog(`個別画像ドロップゾーン作成成功: ${normalizedOrderNumber}`);
      } else {
        debugLog(`個別画像ドロップゾーン作成失敗: ${normalizedOrderNumber}`);
      }
    } catch (error) {
      debugLog(`個別画像ドロップゾーン作成エラー: ${error.message}`);
      console.error('個別画像ドロップゾーン作成エラー:', error);
    }
  } else {
    debugLog(`個別画像ドロップゾーン作成スキップ - コンテナ: ${!!individualDropZoneContainer}, 注文番号: "${orderNumber}"`);
  }
}

function processProductItems(cOrder, row) {
  const tItems = cOrder.querySelector('#商品');
  const trSpace = cOrder.querySelector('.spacerow');
  
  if (row[CONSTANTS.CSV.PRODUCT_COLUMN]) {
    for (let itemrow of row[CONSTANTS.CSV.PRODUCT_COLUMN].split('\n')) {
      const cItem = document.importNode(tItems.content, true);
      const productInfo = parseProductItemData(itemrow);
      
      if (productInfo.itemId && productInfo.quantity) {
        setProductItemElements(cItem, productInfo);
        trSpace.parentNode.parentNode.insertBefore(cItem, trSpace.parentNode);
      }
    }
  }
}

function parseProductItemData(itemrow) {
  const firstSplit = itemrow.split(' / ');
  const itemIdSplit = firstSplit[0].split(':');
  const itemId = itemIdSplit.length > 1 ? itemIdSplit[1].trim() : '';
  const quantitySplit = firstSplit[1] ? firstSplit[1].split(':') : [];
  const quantity = quantitySplit.length > 1 ? quantitySplit[1].trim() : '';
  const productName = firstSplit.slice(2).join(' / ');
  
  return { itemId, quantity, productName };
}

function setProductItemElements(cItem, productInfo) {
  const tdId = cItem.querySelector(".商品ID");
  if (tdId) {
    tdId.textContent = productInfo.itemId;
  }
  const tdQuantity = cItem.querySelector(".数量");
  if (tdQuantity) {
    tdQuantity.textContent = productInfo.quantity;
  }
  const tdName = cItem.querySelector(".商品名");
  if (tdName) {
    tdName.textContent = productInfo.productName;
  }
}

function displayOrderImage(cOrder, orderNumber) {
  // 注文画像表示機能が無効の場合は何もしない
  const settings = StorageManager.getSettings();
  if (!settings.orderImageEnable) {
    return;
  }

  let imageToShow = null;
  if (orderNumber) {
    // 注文番号を正規化
    const normalizedOrderNumber = orderNumber.replace(/^.*?:\s*/, '').trim();
    debugLog(`displayOrderImage - 元の注文番号: "${orderNumber}" → 正規化後: "${normalizedOrderNumber}"`);
    
    // 個別画像があるかチェック
    const individualImage = StorageManager.getOrderImage(normalizedOrderNumber);
    if (individualImage) {
      imageToShow = individualImage;
    } else {
      // 個別画像がない場合はグローバル画像を使用
      const globalImage = window.orderImageDropZone?.getImage();
      if (globalImage) {
        imageToShow = globalImage;
      }
    }
  }

  if (imageToShow) {
    const imageDiv = document.createElement('div');
    imageDiv.classList.add('order-image');

    const img = document.createElement('img');
    img.src = imageToShow;

    imageDiv.appendChild(img);
    const container = cOrder.querySelector('.order-image-container');
    if (container) {
      container.appendChild(imageDiv);
    }
  }
}

function generateLabels(labelarr) {
  if (labelarr.length % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET) {
    for (let i = 0; i < labelarr.length % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET; i++) {
      labelarr.push("");
    }
  }
  
  const tL44 = document.querySelector('#L44');
  let cL44 = document.importNode(tL44.content, true);
  let tableL44 = cL44.querySelector("table");
  let tr = document.createElement("tr");
  let i = 0;
  
  for (let label of labelarr) {
    if (i > 0 && i % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET == 0) {
      tableL44.appendChild(tr);
      tr = document.createElement("tr");
      document.body.insertBefore(cL44, tL44);
      cL44 = document.importNode(tL44.content, true);
      tableL44 = cL44.querySelector("table");
      tr = document.createElement("tr");
    } else if (i > 0 && i % CONSTANTS.LABEL.LABELS_PER_ROW == 0) {
      tableL44.appendChild(tr);
      tr = document.createElement("tr");
    }
    tr.appendChild(createLabel(label));
    i++;
  }
  tableL44.appendChild(tr);
  document.body.insertBefore(cL44, tL44);
}

function showPrintSummary() {
  const labelTable = document.querySelectorAll(".label44");
  const pageDiv = document.querySelectorAll(".page");
  
  if (labelTable.length > 0) {
    alert("印刷枚数\nA4 44面ラベルシール:" + labelTable.length + "枚\nA4普通紙:" + pageDiv.length + "枚");
  } else if (pageDiv.length > 0) {
    alert("印刷枚数\nA4普通紙:" + pageDiv.length + "枚");
  }
}

function showCustomLabelPrintSummary(customLabelCount, skipCount) {
  const labelTable = document.querySelectorAll(".label44");
  const totalLabels = skipCount + customLabelCount;
  
  if (labelTable.length > 0) {
    alert(`カスタムラベル印刷枚数\nA4 44面ラベルシール: ${labelTable.length}枚\nスキップ: ${skipCount}枚\nカスタムラベル: ${customLabelCount}枚\n合計ラベル数: ${totalLabels}枚`);
  }
}
function addP(div, text){
  const p = document.createElement("p");
  p.innerText = text;
  div.appendChild(p);
}
function createDiv(classname="", text=""){
  const div = document.createElement('div');
  if(classname){
    div.classList.add(classname);
  }
  if(text){
    addP(div,text);
  }
  return div;
}
function createDropzone(div){
  const divDrop = createDiv('dropzone', 'Paste QR image here!');
  divDrop.setAttribute("contenteditable", "true");
  divDrop.setAttribute("effectAllowed", "move");
  divDrop.addEventListener('dragover', function(event) {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
    showDropping(this);
  });
  divDrop.addEventListener('dragleave', function(event) {
    hideDropping(this);
  });
  divDrop.addEventListener('drop', function (event) {
    event.preventDefault();
    hideDropping(this);
    const elImage = document.createElement('img');
    if(event.dataTransfer.types.includes("text/uri-list")){
      const url = event.dataTransfer.getData('text/uri-list');
      //elImage.src = "{% url 'mp:boothreceipt' %}?url=" + url;
      elImage.src = url;
      this.parentNode.appendChild(elImage);
      readQR(elImage);
    }else{
      const file = event.dataTransfer.files[0];
      if(file.type.indexOf('image/') === 0){
        this.parentNode.appendChild(elImage);
        attachImage(file, elImage);
      }
    }
    this.classList.remove('dropzone');
    this.style.zIndex = -1;
    elImage.style.zIndex = 9;
    addEventQrReset(elImage);
  });
  divDrop.addEventListener("paste", function(event){
    event.preventDefault();
    if (!event.clipboardData 
            || !event.clipboardData.types
            || (!event.clipboardData.types.includes("Files"))) {
            return true;
    }
    for(let item of event.clipboardData.items){
      //console.log(item);
      if(item["kind"] == "file"){
        const imageFile = item.getAsFile();
        // FileReaderで読み込む
        const fr = new FileReader();
        fr.onload = function(e) {
          const elImage = document.createElement('img');
          elImage.src = e.target.result;
          elImage.onload = function() {
            divDrop.parentNode.appendChild(elImage);
            readQR(elImage);
            divDrop.classList.remove('dropzone');
            divDrop.style.zIndex = -1;
            elImage.style.zIndex = 9;
            addEventQrReset(elImage);
          };
        };
        fr.readAsDataURL(imageFile);
      }
    }
    // 画像以外がペーストされたときのために、元に戻しておく
    this.textContent = null;
    this.innerHTML = "<p>Paste QR image here!</p>";
  });
  div.appendChild(divDrop);
}

function createLabel(labelData=""){
  const divQr = createDiv('qr');
  const divOrdernum = createDiv('ordernum');
  const divYamato = createDiv('yamato');

  // ラベルデータが文字列の場合（既存の注文番号）
  if (typeof labelData === 'string') {
    if (labelData) {
      addP(divOrdernum, labelData);
      createDropzone(divQr);
      const qr = StorageManager.getQRData(labelData);
      if(qr){
        const elImage = document.createElement('img');
        elImage.src = qr['qrimage'];
        divQr.insertBefore(elImage, divQr.firstChild)
        addP(divYamato, qr['receiptnum']);
        addP(divYamato, qr['receiptpassword']);
        addEventQrReset(elImage);
      }
    }
  } 
  // ラベルデータがオブジェクトの場合（カスタムラベル）
  else if (typeof labelData === 'object' && labelData.type === 'custom') {
    divOrdernum.classList.add('custom-label');
    const contentDiv = document.createElement('div');
    contentDiv.classList.add('custom-label-text');
    contentDiv.innerHTML = labelData.content;
    
    // 文字サイズを適用
    if (labelData.fontSize) {
      contentDiv.style.fontSize = labelData.fontSize;
    }
    
    divOrdernum.appendChild(contentDiv);
    
    // カスタムラベルの場合はQRコードエリアとヤマトエリアを非表示
    divQr.style.display = 'none';
    divYamato.style.display = 'none';
    
    // カスタムラベルが全体を覆うようにスタイル調整
    divOrdernum.style.position = 'absolute';
    divOrdernum.style.width = '100%';
    divOrdernum.style.height = '100%';
    divOrdernum.style.top = '0';
    divOrdernum.style.left = '0';
    divOrdernum.style.display = 'flex';
    divOrdernum.style.alignItems = 'center';
    divOrdernum.style.justifyContent = 'center';
    divOrdernum.style.padding = '2px';
    divOrdernum.style.boxSizing = 'border-box';
  }

  const tdLabel = document.createElement('td');
  tdLabel.classList.add('qrlabel');
  tdLabel.appendChild(divQr);
  tdLabel.appendChild(divOrdernum);
  tdLabel.appendChild(divYamato);

  return tdLabel;
}

function addEventQrReset(elImage){
    elImage.addEventListener('click', function(event) {
      event.preventDefault();
      elDrop = elImage.parentNode.querySelector("div");
      elDrop.classList.add('dropzone');
      elDrop.style.zIndex = 99;
      elImage.parentNode.removeChild(elImage);
    });
}

function showDropping(elDrop) {
        elDrop.classList.add('dropover');
}

function hideDropping(elDrop) {
        elDrop.classList.remove('dropover');
}

function readQR(elImage){
  try {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.src = elImage.src;
    
    img.onload = function() {
      try {
        const canv = document.createElement("canvas");
        const context = canv.getContext("2d");
        canv.width = img.width;
        canv.height = img.height;
        context.drawImage(img, 0, 0, canv.width, canv.height);
        const b64data = canv.toDataURL("image/png");
        
        const imageData = context.getImageData(0, 0, canv.width, canv.height);
        const barcode = jsQR(imageData.data, imageData.width, imageData.height);
        
        if(barcode){
          debugLog("[" + barcode.data + "]");
          const b = String(barcode.data).replace(/^\s+|\s+$/g,'').replace(/ +/g,' ').split(" ");
          
          if(b.length === CONSTANTS.QR.EXPECTED_PARTS){
            const ordernum = elImage.closest("td").querySelector(".ordernum p").innerHTML;
            const d = elImage.closest("td").querySelector(".yamato");
            d.innerHTML = "";
            
            for(let i = 1; i < CONSTANTS.QR.EXPECTED_PARTS; i++){
              const p = document.createElement("p");
              p.innerText = b[i];
              d.appendChild(p);
            }
            
            const qrData = { 
              "receiptnum": b[1], 
              "receiptpassword": b[2], 
              "qrimage": b64data 
            };
            
            StorageManager.setQRData(ordernum, qrData);
          } else {
            console.warn('QRコードの形式が正しくありません');
          }
        }
      } catch (error) {
        console.error('QRコード処理エラー:', error);
      }
    };
    
    img.onerror = function() {
      console.error('画像の読み込みに失敗しました');
    };
  } catch (error) {
    console.error('QR読み取り関数エラー:', error);
  }
}

function attachImage(file, elImage) {
  if (!file || !elImage) {
    console.error('ファイルまたは画像要素が無効です');
    return;
  }

  const reader = new FileReader();
  reader.onload = function(event) {
    try {
      const src = event.target.result;
      elImage.src = src;
      elImage.setAttribute('title', file.name);
      elImage.onload = function() {
        readQR(elImage);
      };
    } catch (error) {
      console.error('画像の読み込みエラー:', error);
    }
  };
  
  reader.onerror = function() {
    console.error('ファイル読み込みエラー');
  };
  
  reader.readAsDataURL(file);
}

// 設定管理
const CONFIG = {
  SUPPORTED_IMAGE_TYPES: ['image/jpeg', 'image/png', 'image/svg+xml'],
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
  
  isImageFile(file) {
    return file && this.SUPPORTED_IMAGE_TYPES.includes(file.type);
  },
  
  isValidFileSize(file) {
    return file && file.size <= this.MAX_FILE_SIZE;
  },
  
  validateImageFile(file) {
    if (!file) {
      throw new Error('ファイルが選択されていません');
    }
    
    if (!this.isImageFile(file)) {
      throw new Error('サポートされていないファイル形式です（JPEG、PNG、SVGのみ）');
    }
    
    if (!this.isValidFileSize(file)) {
      throw new Error('ファイルサイズが大きすぎます（10MB以下）');
    }
    
    return true;
  }
};

// 共通のドラッグ&ドロップ機能を提供するベース関数
function createBaseImageDropZone(options = {}) {
  const {
    storageKey = 'orderImage',
    isIndividual = false,
    orderNumber = null,
    containerClass = 'order-image-drop',
    defaultMessage = '画像をドロップ or クリックで選択'
  } = options;

  debugLog(`ベース画像ドロップゾーン作成: ${storageKey}, 個別: ${isIndividual}, 注文番号: ${orderNumber}`);

  const dropZone = document.createElement('div');
  dropZone.classList.add(containerClass);
  
  if (isIndividual) {
    dropZone.style.cssText = 'min-height: 80px; border: 1px dashed #999; padding: 5px; background: #f9f9f9; cursor: pointer;';
  }

  let droppedImage = null;

  // localStorageから保存された画像を読み込む
  const savedImage = localStorage.getItem(storageKey);
  if (savedImage) {
    debugLog(`保存された画像を復元: ${storageKey}`);
    updatePreview(savedImage);
  } else {
    if (isIndividual) {
      dropZone.innerHTML = `<p style="margin: 5px; font-size: 12px; color: #666;">${defaultMessage}</p>`;
    } else {
      const defaultContent = document.getElementById('dropZoneDefaultContent').innerHTML;
      dropZone.innerHTML = defaultContent;
    }
    debugLog(`初期メッセージを設定: ${isIndividual ? defaultMessage : 'デフォルトコンテンツ'}`);
  }

  // 全ての注文明細の画像を更新する関数
  function updateAllOrderImages() {
    // 注文画像表示機能が無効の場合は何もしない
    const settings = StorageManager.getSettings();
    if (!settings.orderImageEnable) {
      return;
    }

    const allOrderSections = document.querySelectorAll('section');
    allOrderSections.forEach(orderSection => {
      const imageContainer = orderSection.querySelector('.order-image-container');
      if (!imageContainer) return;

      // 注文番号を複数の方法で取得を試行
      let orderNumber = null;
      
      // 方法1: .注文番号クラスから取得
      const orderNumberElement = orderSection.querySelector('.注文番号');
      if (orderNumberElement) {
        const rawOrderNumber = orderNumberElement.textContent.trim();
        // 注文番号を正規化（"注文番号 : 66463556" → "66463556"）
        orderNumber = rawOrderNumber.replace(/^.*?:\s*/, '').trim();
        debugLog(`生の注文番号: "${rawOrderNumber}" → 正規化後: "${orderNumber}"`);
      }
      
      // 方法2: .ordernum pから取得（ラベル用）
      if (!orderNumber) {
        const ordernumElement = orderSection.querySelector('.ordernum p');
        if (ordernumElement) {
          const rawOrderNumber = ordernumElement.textContent.trim();
          orderNumber = rawOrderNumber.replace(/^.*?:\s*/, '').trim();
          debugLog(`ラベル用注文番号: "${rawOrderNumber}" → 正規化後: "${orderNumber}"`);
        }
      }

      // 個別画像があるかチェック（個別画像を最優先）
      let imageToShow = null;
      if (orderNumber) {
        const individualImage = StorageManager.getOrderImage(orderNumber);
        const globalImage = StorageManager.getOrderImage(); // グローバル画像を取得
        
        debugLog(`注文番号: ${orderNumber}`);
        debugLog(`個別画像: ${individualImage ? 'あり' : 'なし'}`);
        debugLog(`グローバル画像: ${globalImage ? 'あり' : 'なし'}`);
        
        if (individualImage) {
          debugLog(`個別画像を優先使用: ${orderNumber}`);
          imageToShow = individualImage;
        } else {
          // 個別画像がない場合のみグローバル画像を使用
          debugLog(`個別画像なし、グローバル画像を使用: ${orderNumber}`);
          imageToShow = globalImage;
        }
      } else {
        // 注文番号がない場合はグローバル画像を使用
        const globalImage = StorageManager.getOrderImage(); // グローバル画像を取得
        debugLog('注文番号なし、グローバル画像を使用', globalImage ? 'あり' : 'なし');
        imageToShow = globalImage;
      }

      // 画像コンテナを更新
      imageContainer.innerHTML = '';
      if (imageToShow) {
        const imageDiv = document.createElement('div');
        imageDiv.classList.add('order-image');
        const img = document.createElement('img');
        img.src = imageToShow;
        imageDiv.appendChild(img);
        imageContainer.appendChild(imageDiv);
      }
    });
  }

  function updatePreview(imageUrl) {
    droppedImage = imageUrl;
    dropZone.innerHTML = '';
    const preview = document.createElement('img');
    preview.src = imageUrl;
    
    if (isIndividual) {
      preview.style.cssText = 'max-width: 100%; max-height: 60px; cursor: pointer;';
    } else {
      preview.classList.add('preview-image');
    }
    
    preview.title = 'クリックでリセット';
    dropZone.appendChild(preview);
    localStorage.setItem(storageKey, imageUrl);

    // 個別画像の場合は即座に表示を更新
    if (isIndividual && orderNumber) {
      updateOrderImageDisplay(imageUrl);
    } else if (!isIndividual) {
      // グローバル画像の場合は全ての注文明細の画像を更新
      updateAllOrderImages();
    }

    // 画像クリックでリセット
    preview.addEventListener('click', (e) => {
      e.stopPropagation();
      localStorage.removeItem(storageKey);
      droppedImage = null;
      
      if (isIndividual) {
        dropZone.innerHTML = `<p style="margin: 5px; font-size: 12px; color: #666;">${defaultMessage}</p>`;
        updateOrderImageDisplay(null);
      } else {
        const defaultContent = document.getElementById('dropZoneDefaultContent').innerHTML;
        dropZone.innerHTML = defaultContent;
        // グローバル画像がクリアされた場合も全ての注文明細を更新
        updateAllOrderImages();
      }
    });
  }

  // 個別画像用の表示更新関数
  function updateOrderImageDisplay(imageUrl) {
    // 注文画像表示機能が無効の場合は何もしない
    const settings = StorageManager.getSettings();
    if (!settings.orderImageEnable) {
      return;
    }

    const orderSection = dropZone.closest('section');
    if (!orderSection) return;

    const imageContainer = orderSection.querySelector('.order-image-container');
    if (!imageContainer) return;

    imageContainer.innerHTML = '';

    if (imageUrl) {
      const imageDiv = document.createElement('div');
      imageDiv.classList.add('order-image');
      const img = document.createElement('img');
      img.src = imageUrl;
      imageDiv.appendChild(img);
      imageContainer.appendChild(imageDiv);
    } else {
      // 個別画像がない場合はグローバル画像を表示
      const globalImage = window.orderImageDropZone?.getImage();
      if (globalImage) {
        const imageDiv = document.createElement('div');
        imageDiv.classList.add('order-image');
        const img = document.createElement('img');
        img.src = globalImage;
        imageDiv.appendChild(img);
        imageContainer.appendChild(imageDiv);
      }
    }
  }

  // 共通のイベントリスナー設定
  setupDragAndDropEvents(dropZone, updatePreview, isIndividual);
  setupClickEvent(dropZone, updatePreview, () => droppedImage);

  return {
    element: dropZone,
    getImage: () => droppedImage
  };
}

// ドラッグ&ドロップイベントの共通設定
function setupDragAndDropEvents(dropZone, updatePreview, isIndividual) {
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    if (isIndividual) {
      dropZone.style.backgroundColor = '#e6f3ff';
    } else {
      dropZone.classList.add('dragover');
    }
  });

  dropZone.addEventListener('dragleave', () => {
    if (isIndividual) {
      dropZone.style.backgroundColor = '#f9f9f9';
    } else {
      dropZone.classList.remove('dragover');
    }
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    if (isIndividual) {
      dropZone.style.backgroundColor = '#f9f9f9';
    } else {
      dropZone.classList.remove('dragover');
    }

    const file = e.dataTransfer.files[0];
    if (file && file.type.match(/^image\/(jpeg|png|svg\+xml)$/)) {
      const reader = new FileReader();
      reader.onload = (e) => {
        updatePreview(e.target.result);
      };
      reader.readAsDataURL(file);
    }
  });
}

// クリックイベントの共通設定
function setupClickEvent(dropZone, updatePreview, getDroppedImage) {
  dropZone.addEventListener('click', (e) => {
    // 画像が表示されている場合はクリックイベントを無視（画像のクリックリセット用）
    if (e.target.tagName === 'IMG') {
      return;
    }
    
    const currentImage = getDroppedImage();
    debugLog(`ドロップゾーンクリック - 現在の画像: ${currentImage ? 'あり' : 'なし'}`);
    
    if (!currentImage) {
      const fileInput = document.createElement('input');
      fileInput.type = 'file';
      fileInput.accept = 'image/jpeg, image/png, image/svg+xml';
      fileInput.style.display = 'none';
      document.body.appendChild(fileInput);
      
      debugLog('ファイル選択ダイアログを表示');
      fileInput.click();

      fileInput.addEventListener('change', () => {
        const file = fileInput.files[0];
        debugLog(`ファイル選択: ${file ? file.name : 'なし'}`);
        
        if (file && file.type.match(/^image\/(jpeg|png|svg\+xml)$/)) {
          const reader = new FileReader();
          reader.onload = (e) => {
            debugLog(`画像読み込み完了 - サイズ: ${e.target.result.length} bytes`);
            updatePreview(e.target.result);
          };
          reader.readAsDataURL(file);
        } else if (file) {
          alert('JPEG、PNG、SVGファイルのみサポートしています。');
        }
        document.body.removeChild(fileInput);
      });
    }
  });
}

function createOrderImageDropZone() {
  return createBaseImageDropZone({
    storageKey: 'orderImage',
    isIndividual: false,
    containerClass: 'order-image-drop'
  });
}

// 個別注文用の画像ドロップゾーンを作成する関数（リファクタリング済み）
function createIndividualOrderImageDropZone(orderNumber) {
  return createBaseImageDropZone({
    storageKey: `orderImage_${orderNumber}`,
    isIndividual: true,
    orderNumber: orderNumber,
    containerClass: 'individual-order-image-drop',
    defaultMessage: '画像をドロップ or クリックで選択'
  });
}

document.getElementById("file").addEventListener("change", function() {
  updateButtonStates();
});

// ページロード時に説明を表示し、ユーザーにファイル選択を促す
window.addEventListener("load", function() {
  // 初期状態でボタンを設定
  updateButtonStates();

  // ファイル選択を促すメッセージを表示
  alert("CSVファイルを選択するか、カスタムラベル機能を使用してください。");
});

// グローバルエラーハンドリング
window.addEventListener('error', function(event) {
  console.error('予期しないエラーが発生しました:', event.error);
});

window.addEventListener('unhandledrejection', function(event) {
  console.error('未処理のPromise拒否:', event.reason);
});

// パフォーマンス監視（開発時のみ）
if (window.performance && window.performance.mark) {
  window.performance.mark('app-start');
}

// カスタムラベル機能の関数群
function toggleCustomLabelRow(enabled) {
  const customLabelRow = document.getElementById('customLabelRow');
  customLabelRow.style.display = enabled ? 'table-row' : 'none';
}

// 注文画像表示機能の関数群
function toggleOrderImageRow(enabled) {
  const orderImageRow = document.getElementById('orderImageRow');
  orderImageRow.style.display = enabled ? 'table-row' : 'none';
}

// 全ての注文明細の画像表示可視性を更新
function updateAllOrderImagesVisibility(enabled) {
  debugLog(`画像表示機能が${enabled ? '有効' : '無効'}に変更されました`);
  
  const allOrderSections = document.querySelectorAll('section');
  allOrderSections.forEach(orderSection => {
    const imageContainer = orderSection.querySelector('.order-image-container');
    const individualZone = orderSection.querySelector('.individual-order-image-zone');
    
    if (enabled) {
      // 有効な場合：画像を表示し、個別画像ゾーンも表示
      if (individualZone) {
        individualZone.style.display = 'block';
      }
      
      if (imageContainer) {
        // 注文番号を取得
        let orderNumber = null;
        const orderNumberElement = orderSection.querySelector('.注文番号');
        if (orderNumberElement) {
          const rawOrderNumber = orderNumberElement.textContent.trim();
          orderNumber = rawOrderNumber.replace(/^.*?:\s*/, '').trim();
        }
        
        // 画像を表示
        let imageToShow = null;
        if (orderNumber) {
          const individualImage = StorageManager.getOrderImage(orderNumber);
          if (individualImage) {
            imageToShow = individualImage;
            debugLog(`個別画像を表示: ${orderNumber}`);
          } else {
            const globalImage = StorageManager.getOrderImage();
            if (globalImage) {
              imageToShow = globalImage;
              debugLog(`グローバル画像を表示: ${orderNumber}`);
            }
          }
        }
        
        imageContainer.innerHTML = '';
        if (imageToShow) {
          const imageDiv = document.createElement('div');
          imageDiv.classList.add('order-image');
          const img = document.createElement('img');
          img.src = imageToShow;
          imageDiv.appendChild(img);
          imageContainer.appendChild(imageDiv);
        }
      }
    } else {
      // 無効な場合：画像を非表示にし、個別画像ゾーンも非表示
      if (imageContainer) {
        imageContainer.innerHTML = '';
      }
      if (individualZone) {
        individualZone.style.display = 'none';
      }
    }
  });
}

// 複数カスタムラベルの初期化
function initializeCustomLabels(customLabels) {
  const container = document.getElementById('customLabelsContainer');
  container.innerHTML = '';
  
  // 説明文を一番上に追加
  const instructionDiv = document.createElement('div');
  instructionDiv.className = 'custom-labels-instructions';
  instructionDiv.innerHTML = `
    <div class="instructions-content">
      <strong>カスタムラベル設定について：</strong><br>
      • 実際のラベルサイズ: 48.3mm × 25.3mm<br>
      • テキストのみ入力可能<br>
      • 改行: Enterキー<br>
      • 書式設定: 右クリックでコンテキストメニューを表示（太字、斜体、フォントサイズ変更）
    </div>
  `;
  container.appendChild(instructionDiv);
  
  if (customLabels && customLabels.length > 0) {
    customLabels.forEach((label, index) => {
      addCustomLabelItem(label.html || label.text, label.count, index, label.enabled !== false);
      
      // 文字サイズを復元
      if (label.fontSize) {
        const item = container.children[index];
        const editor = item.querySelector('.rich-text-editor');
        
        if (editor) {
          // フォントサイズが既にpt単位なら そのまま使用、数値のみなら pt を追加
          const fontSize = label.fontSize.toString().includes('pt') ? label.fontSize : label.fontSize + 'pt';
          editor.style.fontSize = fontSize;
        }
      }
    });
  } else {
    // デフォルトで1つ追加
    addCustomLabelItem('', 1, 0, true);
  }
  
  updateCustomLabelsSummary();
}

// カスタムラベル項目を追加
function addCustomLabelItem(text = '', count = 1, index = null, enabled = true) {
  debugLog('addCustomLabelItem関数が呼び出されました'); // デバッグ用
  debugLog('引数:', { text, count, index }); // デバッグ用
  
  const container = document.getElementById('customLabelsContainer');
  debugLog('container要素:', container); // デバッグ用
  if (!container) {
    console.error('customLabelsContainer要素が見つかりません');
    return;
  }
  
  const itemIndex = index !== null ? index : container.children.length;
  debugLog('itemIndex:', itemIndex); // デバッグ用
  
  const item = document.createElement('div');
  item.classList.add('custom-label-item');
  item.dataset.index = itemIndex;
  
  item.innerHTML = `
    <div class="custom-label-item-header">
      <div class="custom-label-item-header-left">
        <input type="checkbox" class="custom-label-enabled" 
               id="customLabel_${itemIndex}_enabled" 
               data-index="${itemIndex}" 
               checked onchange="saveCustomLabels(); updateCustomLabelsSummary();">
        <label for="customLabel_${itemIndex}_enabled" class="custom-label-item-title">印刷する</label>
      </div>
      <button type="button" class="btn-base btn-danger btn-small" onclick="removeCustomLabelItem(${itemIndex})">削除</button>
    </div>
    <div class="custom-label-input-group">
      <div class="custom-label-editor-row">
        <div class="rich-text-editor" contenteditable="true" 
             placeholder="印刷したい文字列を入力してください..."
             data-index="${itemIndex}"
             style="font-size: 12pt;"></div>
        <div class="custom-label-count-group">
          <label>印刷枚数：</label>
          <input type="number" min="1" value="${count}" 
                 data-index="${itemIndex}" 
                 onchange="updateCustomLabelsSummary()">
        </div>
      </div>
    </div>
  `;
  
  container.appendChild(item);
  debugLog('item要素がコンテナに追加されました'); // デバッグ用
  
  // リッチテキストエディタのイベントリスナーを設定
  const editor = item.querySelector('.rich-text-editor');
  debugLog('editor要素:', editor); // デバッグ用
  if (editor) {
    // テキスト内容を設定（HTMLとして）
    if (text && text.trim() !== '') {
      editor.innerHTML = text;
    }
    
    setupRichTextFormatting(editor);
    setupTextOnlyEditor(editor);
    
    editor.addEventListener('input', function() {
      saveCustomLabels();
      updateButtonStates();
    });
  } else {
    console.error('editor要素が見つかりません');
  }
  
  // チェックボックスの状態を設定
  const enabledCheckbox = item.querySelector('.custom-label-enabled');
  if (enabledCheckbox) {
    enabledCheckbox.checked = enabled;
  }
  
  // 枚数入力のイベントリスナーを設定
  const countInput = item.querySelector('input[type="number"]');
  debugLog('countInput要素:', countInput); // デバッグ用
  if (countInput) {
    countInput.addEventListener('input', function() {
      saveCustomLabels();
      updateButtonStates();
    });
  } else {
    console.error('countInput要素が見つかりません');
  }
  
  updateCustomLabelsSummary();
}

// カスタムラベル項目を削除
function removeCustomLabelItem(index) {
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  
  if (items.length <= 1) {
    alert('最低1つのカスタムラベルは必要です。');
    return;
  }
  
  // 指定されたインデックスの項目を削除
  const itemToRemove = container.querySelector(`[data-index="${index}"]`);
  if (itemToRemove) {
    itemToRemove.remove();
  }
  
  // インデックスを再設定
  reindexCustomLabelItems();
  saveCustomLabels();
  updateCustomLabelsSummary();
  updateButtonStates();
}

// カスタムラベル項目のインデックスを再設定
function reindexCustomLabelItems() {
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  
  items.forEach((item, newIndex) => {
    item.dataset.index = newIndex;
    
    // タイトルを更新
    const title = item.querySelector('.custom-label-item-title');
    title.textContent = `ラベル ${newIndex + 1}`;
    
    // 削除ボタンのonclickを更新
    const removeBtn = item.querySelector('.btn-remove');
    removeBtn.onclick = () => removeCustomLabelItem(newIndex);
    
    // エディタのdata-indexを更新
    const editor = item.querySelector('.rich-text-editor');
    editor.dataset.index = newIndex;
    
    // 枚数入力のdata-indexを更新
    const countInput = item.querySelector('input[type="number"]');
    countInput.dataset.index = newIndex;
  });
}

// UIからカスタムラベルデータを取得
function getCustomLabelsFromUI() {
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  const labels = [];
  
  items.forEach(item => {
    const editor = item.querySelector('.rich-text-editor');
    const countInput = item.querySelector('input[type="number"]');
    const enabledCheckbox = item.querySelector('.custom-label-enabled');
    const text = editor.innerHTML.trim();
    const count = parseInt(countInput.value, 10) || 1;
    const enabled = enabledCheckbox ? enabledCheckbox.checked : true;
    
    // フォントサイズをエディタのスタイルから取得
    const computedStyle = window.getComputedStyle(editor);
    const fontSize = computedStyle.fontSize || '12pt';
    
    // テキストがある場合、または設定を保持するため常に保存
    labels.push({ 
      text, 
      count, 
      fontSize,
      html: text, // HTMLフォーマットも保存
      enabled // 印刷有効フラグを追加
    });
  });
  
  return labels;
}

// カスタムラベルデータを保存
function saveCustomLabels() {
  const labels = getCustomLabelsFromUI();
  StorageManager.setCustomLabels(labels);
}

// カスタムラベルの総計を更新
function updateCustomLabelsSummary() {
  const labels = getCustomLabelsFromUI();
  const enabledLabels = labels.filter(label => label.enabled);
  const totalCount = enabledLabels.reduce((sum, label) => sum + label.count, 0);
  const skipCount = parseInt(document.getElementById("labelskipnum").value, 10) || 0;
  const fileInput = document.getElementById("file");
  
  let csvRowCount = 0;
  if (fileInput.files.length > 0) {
    csvRowCount = 0; // 実際の行数は処理時に決定
  }
  
  const maxLabels = CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET - skipCount - csvRowCount;
  const summary = document.getElementById('customLabelsSummary');
  
  if (csvRowCount > 0) {
    summary.textContent = `合計 ${totalCount} 枚のカスタムラベル（有効なもののみ）。44枚シート中、スキップ${skipCount}枚 + CSV${csvRowCount}枚使用済み。残り${maxLabels}枚まで設定可能。`;
  } else {
    summary.textContent = `合計 ${totalCount} 枚のカスタムラベル（有効なもののみ）。44枚シート中、スキップ${skipCount}枚使用済み。残り${maxLabels}枚まで設定可能。`;
  }
  
  // 上限を超えている場合は警告色にする
  if (totalCount > maxLabels) {
    summary.style.color = '#dc3545';
    summary.style.fontWeight = 'bold';
  } else {
    summary.style.color = '#666';
    summary.style.fontWeight = 'normal';
  }
}

// CSV用の調整関数
function adjustCustomLabelsForCSV(customLabels, maxCustomLabels) {
  let remaining = maxCustomLabels;
  
  for (let i = 0; i < customLabels.length; i++) {
    if (remaining <= 0) {
      customLabels[i].count = 0;
    } else if (customLabels[i].count > remaining) {
      customLabels[i].count = remaining;
    }
    remaining -= customLabels[i].count;
  }
  
  // UIを更新
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  items.forEach((item, index) => {
    const countInput = item.querySelector('input[type="number"]');
    if (customLabels[index]) {
      countInput.value = customLabels[index].count;
    }
  });
  
  saveCustomLabels();
  updateCustomLabelsSummary();
}

// 総枚数用の調整関数
function adjustCustomLabelsForTotal(customLabels, maxTotalLabels) {
  let remaining = maxTotalLabels;
  
  for (let i = 0; i < customLabels.length; i++) {
    if (remaining <= 0) {
      customLabels[i].count = 0;
    } else if (customLabels[i].count > remaining) {
      customLabels[i].count = remaining;
    }
    remaining -= customLabels[i].count;
  }
  
  // UIを更新
  const container = document.getElementById('customLabelsContainer');
  const items = container.querySelectorAll('.custom-label-item');
  items.forEach((item, index) => {
    const countInput = item.querySelector('input[type="number"]');
    if (customLabels[index]) {
      countInput.value = customLabels[index].count;
    }
  });
  
  saveCustomLabels();
  updateCustomLabelsSummary();
}

function setupCustomLabelEvents() {
  // カスタムラベル有効化チェックボックス
  const customLabelEnable = document.getElementById('customLabelEnable');
  if (customLabelEnable) {
    customLabelEnable.addEventListener('change', function() {
      toggleCustomLabelRow(this.checked);
      StorageManager.set(StorageManager.KEYS.CUSTOM_LABEL_ENABLE, this.checked);
      updateButtonStates();
    });
  }

  // カスタムラベル追加ボタン
  const addButton = document.getElementById('addCustomLabelBtn');
  if (addButton) {
    addButton.addEventListener('click', function() {
      debugLog('ラベル追加ボタンがクリックされました'); // デバッグ用
      addCustomLabelItem('', 1, null, true);
      saveCustomLabels();
      updateButtonStates();
    });
  } else {
    console.error('addCustomLabelBtn要素が見つかりません');
  }

  // カスタムラベル全削除ボタン
  const clearButton = document.getElementById('clearCustomLabelsBtn');
  if (clearButton) {
    clearButton.addEventListener('click', function() {
      debugLog('ラベル全削除ボタンがクリックされました'); // デバッグ用
      if (confirm('本当に全てのカスタムラベルを削除しますか？')) {
        clearAllCustomLabels();
      }
    });
  } else {
    console.error('clearCustomLabelsBtn要素が見つかりません');
  }
}

// 全てのカスタムラベルを削除
function clearAllCustomLabels() {
  const container = document.getElementById('customLabelsContainer');
  
  // 説明文以外の全ての項目を削除
  const items = container.querySelectorAll('.custom-label-item');
  items.forEach(item => item.remove());
  
  // デフォルトで1つ追加
  addCustomLabelItem('', 1, 0, true);
  
  // 保存とUI更新
  saveCustomLabels();
  updateCustomLabelsSummary();
  updateButtonStates();
  
  debugLog('全てのカスタムラベルが削除されました');
}

function updateCustomLabelCountDescription(skipCount, csvRowCount, maxCustomLabels) {
  const customLabelCountGroup = document.querySelector('.custom-label-count-group');
  let descriptionElement = customLabelCountGroup.querySelector('.label-count-description');
  
  if (!descriptionElement) {
    descriptionElement = document.createElement('small');
    descriptionElement.classList.add('label-count-description');
    descriptionElement.style.cssText = 'display: block; color: #666; margin-top: 5px; font-size: 12px;';
    customLabelCountGroup.appendChild(descriptionElement);
  }
  
  const totalLabels = CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET;
  const usedLabels = skipCount + csvRowCount;
  const remainingLabels = totalLabels - usedLabels;
  
  if (csvRowCount > 0) {
    descriptionElement.textContent = `44枚シート中、スキップ${skipCount}枚 + CSV${csvRowCount}枚 = ${usedLabels}枚使用済み。残り${remainingLabels}枚まで設定可能。`;
  } else {
    descriptionElement.textContent = `44枚シート中、スキップ${skipCount}枚使用済み。残り${remainingLabels}枚まで設定可能。`;
  }
}

function updateButtonStates() {
  const fileInput = document.getElementById("file");
  const executeButton = document.getElementById("executeButton");
  const customLabelOnlyButton = document.getElementById("customLabelOnlyButton");
  const printButton = document.getElementById("printButton");
  const customLabelEnable = document.getElementById("customLabelEnable");

  // CSV処理実行ボタンの状態
  executeButton.disabled = fileInput.files.length === 0;

  // カスタムラベル専用実行ボタンの状態
  const hasValidCustomLabels = customLabelEnable.checked && hasCustomLabelsWithContent();
  customLabelOnlyButton.disabled = !hasValidCustomLabels;

  // 印刷ボタンの状態（何らかのコンテンツが生成されている場合に有効）
  const hasSheets = document.querySelectorAll('.sheet').length > 0;
  const hasLabels = document.querySelectorAll('.label44').length > 0;
  const hasContent = hasSheets || hasLabels;
  printButton.disabled = !hasContent;

  // カスタムラベル枚数の上限を更新
  updateCustomLabelsSummary();
}

// カスタムラベルに内容があるかチェック
function hasCustomLabelsWithContent() {
  const labels = getCustomLabelsFromUI();
  return labels.length > 0 && labels.some(label => label.text.trim() !== '');
}

function setupRichTextFormatting(editor) {
  debugLog('setupRichTextFormatting関数が呼び出されました, editor:', editor); // デバッグ用
  if (!editor) {
    console.error('setupRichTextFormatting: editor要素がnullです');
    return;
  }
  
  // Enterキーでの改行処理を改善
  editor.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      // Enterキーの処理を改善 - 1回で改行
      e.preventDefault();
      
      const selection = window.getSelection();
      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        
        // 選択範囲がある場合は削除
        range.deleteContents();
        
        // 改行要素を作成
        const br = document.createElement('br');
        range.insertNode(br);
        
        // さらにもう一つのbrを挿入してカーソル位置を確保
        const br2 = document.createElement('br');
        range.setStartAfter(br);
        range.insertNode(br2);
        
        // カーソルを2番目のbrの前に配置
        range.setStartBefore(br2);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
      }
    }
  });

  // コンテキストメニューの追加
  editor.addEventListener('contextmenu', function(e) {
    e.preventDefault();
    
    // 選択範囲の有無を確認
    const selection = window.getSelection();
    const hasSelection = selection.toString().length > 0;
    
    const menu = createFontSizeMenu(e.pageX, e.pageY, editor, hasSelection);
    document.body.appendChild(menu);
    
    // クリック外でメニューを閉じる
    setTimeout(() => {
      document.addEventListener('click', function closeMenu() {
        if (menu.parentNode) {
          menu.parentNode.removeChild(menu);
        }
        document.removeEventListener('click', closeMenu);
      });
    }, 100);
  });
}

// フォントサイズ選択メニューを作成
function createFontSizeMenu(x, y, editor, hasSelection = true) {
  const menu = document.createElement('div');
  menu.style.cssText = `
    position: fixed;
    left: ${x}px;
    top: ${y}px;
    background: white;
    border: 1px solid #ccc;
    border-radius: 4px;
    padding: 5px 0;
    box-shadow: 0 2px 10px rgba(0,0,0,0.2);
    z-index: 10000;
    font-family: sans-serif;
    min-width: 140px;
  `;
  
  // フォーマットオプション
  const formatOptions = [
    { label: '太字', command: 'bold', style: 'font-weight: bold;' },
    { label: '斜体', command: 'italic', style: 'font-style: italic;' },
    { label: '下線', command: 'underline', style: 'text-decoration: underline;' },
    { label: 'すべてクリア', command: 'clear', style: 'color: #dc3545; font-weight: bold;' }
  ];
  
  // 選択範囲がない場合はクリア機能のみ表示
  const availableOptions = hasSelection ? formatOptions : [formatOptions[3]]; // クリア機能のみ
  
  // フォーマットボタンを追加
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
    
    item.addEventListener('mouseenter', function() {
      this.style.backgroundColor = '#f0f0f0';
    });
    
    item.addEventListener('mouseleave', function() {
      this.style.backgroundColor = 'transparent';
    });
    
    item.addEventListener('mousedown', function(e) {
      e.preventDefault(); // 選択範囲がクリアされるのを防ぐ
      e.stopPropagation();
    });
    
    item.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      
      // 選択範囲を保持してフォーマットを適用
      setTimeout(() => {
        applyFormatToSelection(option.command, editor);
        
        // メニューを閉じる
        if (menu.parentNode) {
          menu.parentNode.removeChild(menu);
        }
        
        saveCustomLabels();
      }, 10);
    });
    
    menu.appendChild(item);
  });
  
  // 選択範囲がある場合のみフォントサイズオプションを追加
  if (hasSelection) {
    // 区切り線
    const separator = document.createElement('div');
    separator.style.cssText = `
      height: 1px;
      background-color: #ddd;
      margin: 5px 0;
    `;
    menu.appendChild(separator);
    
    // フォントサイズオプション
    const fontSizeLabel = document.createElement('div');
    fontSizeLabel.textContent = 'フォントサイズ';
    fontSizeLabel.style.cssText = `
      padding: 5px 15px;
      font-size: 11px;
      color: #666;
      font-weight: bold;
    `;
    menu.appendChild(fontSizeLabel);
    
    const fontSizes = [6, 8, 10, 12, 14, 16, 18, 20, 24, 28];
    
    fontSizes.forEach(size => {
      const item = document.createElement('div');
      item.textContent = `${size}pt`;
      item.style.cssText = `
        padding: 6px 20px;
        cursor: pointer;
        font-size: 11px;
        transition: background-color 0.2s;
      `;
      
      item.addEventListener('mouseenter', function() {
        this.style.backgroundColor = '#f0f0f0';
      });
      
      item.addEventListener('mouseleave', function() {
        this.style.backgroundColor = 'transparent';
      });
      
      item.addEventListener('mousedown', function(e) {
        e.preventDefault(); // 選択範囲がクリアされるのを防ぐ
        e.stopPropagation();
      });
      
      item.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        // 選択範囲を保持してフォントサイズを適用
        setTimeout(() => {
          applyFontSizeToSelection(size, editor);
          
          // メニューを閉じる
          if (menu.parentNode) {
            menu.parentNode.removeChild(menu);
          }
          
          saveCustomLabels();
        }, 10);
      });
      
      menu.appendChild(item);
    });
  }
  
  return menu;
}

// 選択範囲にフォーマットを適用
function applyFormatToSelection(command, editor) {
  const selection = window.getSelection();
  
  if (command === 'clear') {
    // すべてクリア：エディタ全体をクリア
    clearAllContent(editor);
    return;
  }
  
  if (selection.rangeCount > 0 && !selection.isCollapsed) {
    const range = selection.getRangeAt(0);
    const selectedContent = range.extractContents();
    
    let wrapper;
    switch (command) {
      case 'bold':
        wrapper = document.createElement('strong');
        break;
      case 'italic':
        wrapper = document.createElement('em');
        break;
      case 'underline':
        wrapper = document.createElement('u');
        break;
      default:
        wrapper = document.createElement('span');
    }
    
    wrapper.appendChild(selectedContent);
    range.insertNode(wrapper);
    
    // 新しい選択範囲を設定
    range.selectNodeContents(wrapper);
    selection.removeAllRanges();
    selection.addRange(range);
    
    // エディタにフォーカスを戻す
    editor.focus();
  }
}

// エディタの全内容をクリア（書式も含めて）
function clearAllContent(editor) {
  // 確認ダイアログを表示
  if (confirm('このカスタムラベルの内容と書式をすべてクリアしますか？')) {
    // エディタの内容を完全にクリア
    editor.innerHTML = '';
    
    // デフォルトのスタイルを再設定
    editor.style.fontSize = '12pt';
    editor.style.lineHeight = '1.2';
    editor.style.textAlign = 'center';
    
    // フォーカスを設定
    editor.focus();
    
    // カスタムラベルを保存
    saveCustomLabels();
  }
}

// 選択範囲にフォントサイズを適用
function applyFontSizeToSelection(fontSize, editor) {
  const selection = window.getSelection();
  if (selection.rangeCount > 0 && !selection.isCollapsed) {
    const range = selection.getRangeAt(0);
    const selectedContent = range.extractContents();
    
    // spanで囲んでフォントサイズを適用
    const span = document.createElement('span');
    span.style.fontSize = fontSize + 'pt';
    span.appendChild(selectedContent);
    
    range.insertNode(span);
    
    // 選択範囲を新しいspanの内容に設定
    range.selectNodeContents(span);
    selection.removeAllRanges();
    selection.addRange(range);
    
    // エディタにフォーカスを戻す
    editor.focus();
  }
}

// テキストのみ入力可能にする設定
function setupTextOnlyEditor(editor) {
  debugLog('setupTextOnlyEditor関数が呼び出されました, editor:', editor); // デバッグ用
  if (!editor) {
    console.error('setupTextOnlyEditor: editor要素がnullです');
    return;
  }
  
  // 画像やその他のメディアの貼り付けを防ぐ
  editor.addEventListener('paste', function(e) {
    e.preventDefault();
    
    // プレーンテキストのみを取得
    const text = (e.clipboardData || window.clipboardData).getData('text/plain');
    
    // HTMLタグを除去し、改行を<br>に変換
    const cleanText = text.replace(/<[^>]*>/g, '');
    
    // 現在の選択範囲にテキストを挿入
    const selection = window.getSelection();
    if (selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      range.deleteContents();
      
      // 改行を含むテキストを適切に処理
      const lines = cleanText.split('\n');
      lines.forEach((line, index) => {
        if (index > 0) {
          // 改行を挿入
          const br = document.createElement('br');
          range.insertNode(br);
          range.setStartAfter(br);
        }
        
        if (line.length > 0) {
          // テキストを挿入
          const textNode = document.createTextNode(line);
          range.insertNode(textNode);
          range.setStartAfter(textNode);
        }
      });
      
      // カーソルを最後に移動
      range.collapse(true);
      selection.removeAllRanges();
      selection.addRange(range);
    }
  });
  
  // ドラッグ&ドロップでの画像挿入を防ぐ
  editor.addEventListener('drop', function(e) {
    e.preventDefault();
    
    // ドロップされたファイルがある場合は何もしない
    if (e.dataTransfer.files.length > 0) {
      return false;
    }
    
    // テキストのみ許可
    const text = e.dataTransfer.getData('text/plain');
    if (text) {
      const cleanText = text.replace(/<[^>]*>/g, '');
      
      // 改行を含むテキストを適切に処理
      const selection = window.getSelection();
      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        range.deleteContents();
        
        const lines = cleanText.split('\n');
        lines.forEach((line, index) => {
          if (index > 0) {
            const br = document.createElement('br');
            range.insertNode(br);
            range.setStartAfter(br);
          }
          
          if (line.length > 0) {
            const textNode = document.createTextNode(line);
            range.insertNode(textNode);
            range.setStartAfter(textNode);
          }
        });
        
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
      }
    }
  });
  
  // ドラッグオーバー時の処理
  editor.addEventListener('dragover', function(e) {
    // ファイルのドラッグオーバーの場合は拒否
    if (e.dataTransfer.types.includes('Files')) {
      e.dataTransfer.dropEffect = 'none';
      return false;
    }
    e.preventDefault();
  });
  
  // 直接的なHTML挿入を監視してテキストのみに変換
  const observer = new MutationObserver(function(mutations) {
    mutations.forEach(function(mutation) {
      if (mutation.type === 'childList') {
        mutation.addedNodes.forEach(function(node) {
          if (node.nodeType === Node.ELEMENT_NODE) {
            // 画像やその他の要素が挿入された場合は削除
            if (node.tagName === 'IMG' || node.tagName === 'VIDEO' || node.tagName === 'AUDIO') {
              node.remove();
            }
          }
        });
      }
    });
  });
  
  observer.observe(editor, {
    childList: true,
    subtree: true
  });
}

function createFormatMenu(x, y) {
  const menu = document.createElement('div');
  menu.style.cssText = `
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    background: white;
    border: 1px solid #ccc;
    border-radius: 4px;
    padding: 5px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    z-index: 1000;
    font-size: 12px;
  `;

  const buttons = [
    { text: '太字', command: 'bold' },
    { text: '斜体', command: 'italic' },
    { text: '下線', command: 'underline' }
  ];

  buttons.forEach(btn => {
    const button = document.createElement('button');
    button.textContent = btn.text;
    button.style.cssText = `
      display: block;
      width: 100%;
      padding: 3px 8px;
      margin: 2px 0;
      border: none;
      background: #f5f5f5;
      cursor: pointer;
      border-radius: 2px;
    `;
    
    button.addEventListener('click', () => {
      document.execCommand(btn.command, false, null);
      menu.parentNode.removeChild(menu);
    });
    
    button.addEventListener('mouseenter', () => {
      button.style.backgroundColor = '#e0e0e0';
    });
    
    button.addEventListener('mouseleave', () => {
      button.style.backgroundColor = '#f5f5f5';
    });
    
    menu.appendChild(button);
  });

  return menu;
}

// 印刷ボタンのイベントリスナーを変更
document.addEventListener('DOMContentLoaded', function() {
  const printButton = document.getElementById('printButton');
  
  // 既存のonClickを保存
  const originalOnClick = printButton.onclick;
  
  // 新しい処理を設定
  printButton.onclick = function() {
    // まず印刷を実行
    window.print();
    
    // 印刷ダイアログが閉じた後に実行される
    setTimeout(() => {
      // 印刷完了の確認
      if (confirm('印刷が完了しましたか？完了した場合、次回のスキップ枚数を更新します。')) {
        updateSkipCount();
      }
    }, 100);
  };
});

// スキップ枚数を更新する関数
function updateSkipCount() {
  // 現在のスキップ枚数を取得
  const currentSkip = parseInt(document.getElementById("labelskipnum").value, 10) || 0;
  
  // 使用したラベル枚数を計算
  let usedLabels = 0;
  
  // CSV処理による注文ラベル数を取得
  const orderPages = document.querySelectorAll(".page");
  if (orderPages.length > 0) {
    // 注文ページの数 = CSV行数
    usedLabels = orderPages.length;
  }
  
  // カスタムラベルが有効な場合、その枚数を追加
  if (document.getElementById("customLabelEnable").checked) {
    const customLabels = getCustomLabelsFromUI();
    const totalCustomCount = customLabels.reduce((sum, label) => sum + label.count, 0);
    usedLabels += totalCustomCount;
  }
  
  // 合計使用枚数を計算
  const totalUsed = currentSkip + usedLabels;
  
  // 44枚のシートサイズに合わせて余りを計算
  const newSkipValue = totalUsed % CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET;
  
  // 新しいスキップ枚数を設定
  document.getElementById("labelskipnum").value = newSkipValue;
  StorageManager.set(StorageManager.KEYS.LABEL_SKIP, newSkipValue);
  
  // 更新完了メッセージ
  alert(`次回のスキップ枚数を ${newSkipValue} 枚に更新しました。`);
  
  // カスタムラベルの上限も更新
  updateCustomLabelsSummary();
}
