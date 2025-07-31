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

// localStorage操作を管理するクラス
class StorageManager {
  static KEYS = {
    ORDER_IMAGE_PREFIX: 'orderImage_',
    GLOBAL_ORDER_IMAGE: 'orderImage',
    LABEL_SETTING: 'labelyn',
    LABEL_SKIP: 'labelskip',
    SORT_BY_PAYMENT: 'sortByPaymentDate'
  };

  // 設定値の取得
  static getSettings() {
    return {
      labelyn: this.get(this.KEYS.LABEL_SETTING, 'true') !== 'false',
      labelskip: parseInt(this.get(this.KEYS.LABEL_SKIP, '0'), 10),
      sortByPaymentDate: this.get(this.KEYS.SORT_BY_PAYMENT) === 'true'
    };
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

  // QRコードの重複チェック
  static checkQRDuplicate(qrContent, currentOrderNumber) {
    const qrHash = this.generateQRHash(qrContent);
    const duplicates = [];
    
    Object.keys(localStorage).forEach(key => {
      if (key !== currentOrderNumber) {
        const data = this.get(key);
        if (data) {
          try {
            const parsedData = JSON.parse(data);
            if (parsedData.qrhash === qrHash) {
              duplicates.push(key);
            }
          } catch (e) {
            // 無視
          }
        }
      }
    });
    
    return duplicates;
  }

  // QRコンテンツのハッシュ化
  static generateQRHash(qrContent) {
    // シンプルなハッシュ関数（本格的な場合はCrypto APIを使用）
    let hash = 0;
    if (qrContent.length === 0) return hash;
    for (let i = 0; i < qrContent.length; i++) {
      const char = qrContent.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // 32bit整数に変換
    }
    return hash.toString();
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
  
  console.log(settings.labelyn, settings.labelskip);
  
  document.getElementById("labelyn").checked = settings.labelyn;
  document.getElementById("labelskipnum").value = settings.labelskip;
  document.getElementById("sortByPaymentDate").checked = settings.sortByPaymentDate;

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
  
  StorageManager.set(StorageManager.KEYS.LABEL_SETTING, labelyn);
  StorageManager.set(StorageManager.KEYS.LABEL_SKIP, labelskip);
  
  const labelarr = [];
  const labelskipNum = parseInt(labelskip, 10) || 0;
  if (labelskipNum > 0) {
    for (let i = 0; i < labelskipNum; i++) {
      labelarr.push("");
    }
  }
  
  return { file, labelyn, labelskip, sortByPaymentDate, labelarr };
}

function processCSVResults(results, config) {
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
  
  // ラベル生成
  if (config.labelyn && config.labelarr.length > 0) {
    generateLabels(config.labelarr);
  }
  
  // 印刷枚数の表示
  showPrintSummary();
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
  const individualDropZoneContainer = cOrder.querySelector('.individual-image-dropzone');
  if (individualDropZoneContainer && orderNumber) {
    const individualImageDropZone = createIndividualOrderImageDropZone(orderNumber);
    individualDropZoneContainer.appendChild(individualImageDropZone.element);
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
  let imageToShow = null;
  if (orderNumber) {
    // 個別画像があるかチェック
    const individualImage = StorageManager.getOrderImage(orderNumber);
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
function setupDropzoneEvents(dropzone) {
  dropzone.addEventListener('dragover', function(event) {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
    showDropping(this);
  });

  dropzone.addEventListener('dragleave', function(event) {
    hideDropping(this);
  });

  dropzone.addEventListener('drop', function (event) {
    event.preventDefault();
    hideDropping(this);
    const elImage = document.createElement('img');
    
    if(event.dataTransfer.types.includes("text/uri-list")){
      const url = event.dataTransfer.getData('text/uri-list');
      elImage.src = url;
      this.parentNode.appendChild(elImage);
      readQR(elImage);
    } else {
      const file = event.dataTransfer.files[0];
      if(file && file.type.indexOf('image/') === 0){
        this.parentNode.appendChild(elImage);
        attachImage(file, elImage);
      }
    }
    
    this.classList.remove('dropzone');
    this.style.zIndex = -1;
    elImage.style.zIndex = 9;
    addEventQrReset(elImage);
  });

  dropzone.addEventListener("paste", function(event){
    event.preventDefault();
    if (!event.clipboardData 
            || !event.clipboardData.types
            || (!event.clipboardData.types.includes("Files"))) {
            return true;
    }
    
    for(let item of event.clipboardData.items){
      if(item["kind"] == "file"){
        const imageFile = item.getAsFile();
        const fr = new FileReader();
        fr.onload = function(e) {
          const elImage = document.createElement('img');
          elImage.src = e.target.result;
          elImage.onload = function() {
            dropzone.parentNode.appendChild(elImage);
            readQR(elImage);
            dropzone.classList.remove('dropzone');
            dropzone.style.zIndex = -1;
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
}

function createDropzone(div){
  const divDrop = createDiv('dropzone', 'Paste QR image here!');
  divDrop.setAttribute("contenteditable", "true");
  divDrop.setAttribute("effectAllowed", "move");
  
  // 共通のイベントリスナーを設定
  setupDropzoneEvents(divDrop);
  
  div.appendChild(divDrop);
}

function createLabel(ordernum=""){
  const divQr = createDiv('qr');
  const divOrdernum = createDiv('ordernum', ordernum);
  const divYamato = createDiv('yamato');

  if(ordernum){
    createDropzone(divQr);
    const qr = StorageManager.getQRData(ordernum);
    if(qr){
      const elImage = document.createElement('img');
      elImage.src = qr['qrimage'];
      divQr.insertBefore(elImage, divQr.firstChild)
      addP(divYamato, qr['receiptnum']);
      addP(divYamato, qr['receiptpassword']);
      addEventQrReset(elImage);
    }
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
      const elDrop = elImage.parentNode.querySelector("div");
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
          console.log("[" + barcode.data + "]");
          const b = String(barcode.data).replace(/^\s+|\s+$/g,'').replace(/ +/g,' ').split(" ");
          
          if(b.length === CONSTANTS.QR.EXPECTED_PARTS){
            const ordernum = elImage.closest("td").querySelector(".ordernum p").innerHTML;
            
            // 重複チェック
            const duplicates = StorageManager.checkQRDuplicate(barcode.data, ordernum);
            if (duplicates.length > 0) {
              const duplicateList = duplicates.join(', ');
              const confirmMessage = `警告: このQRコードは既に以下の注文で使用されています:\n${duplicateList}\n\n同じQRコードを使用すると配送ミスの原因となる可能性があります。\n続行しますか？`;
              
              if (!confirm(confirmMessage)) {
                console.log('QRコード登録がキャンセルされました');
                // 画像を削除して元の状態に戻す
                const parentQr = elImage.parentNode;
                const dropzone = parentQr.querySelector('.dropzone, div[class="dropzone"]');
                
                if (dropzone) {
                  // 既存のドロップゾーンを復元
                  dropzone.classList.add('dropzone');
                  dropzone.style.zIndex = '99';
                  dropzone.innerHTML = '<p>Paste QR image here!</p>';
                  dropzone.style.display = 'block';
                } else {
                  // ドロップゾーンが見つからない場合は新しく作成
                  const newDropzone = document.createElement('div');
                  newDropzone.className = 'dropzone';
                  newDropzone.contentEditable = 'true';
                  newDropzone.setAttribute('effectallowed', 'move');
                  newDropzone.style.zIndex = '99';
                  newDropzone.innerHTML = '<p>Paste QR image here!</p>';
                  parentQr.appendChild(newDropzone);
                  
                  // ドロップゾーンのイベントリスナーを再設定
                  setupDropzoneEvents(newDropzone);
                }
                
                // 画像を削除
                elImage.parentNode.removeChild(elImage);
                return;
              }
            }
            
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
              "qrimage": b64data,
              "qrhash": StorageManager.generateQRHash(barcode.data)
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

  const dropZone = document.createElement('div');
  dropZone.classList.add(containerClass);
  
  if (isIndividual) {
    dropZone.style.cssText = 'min-height: 80px; border: 1px dashed #999; padding: 5px; background: #f9f9f9;';
  }

  let droppedImage = null;

  // localStorageから保存された画像を読み込む
  const savedImage = localStorage.getItem(storageKey);
  if (savedImage) {
    updatePreview(savedImage);
  } else {
    if (isIndividual) {
      dropZone.innerHTML = `<p style="margin: 5px; font-size: 12px; color: #666;">${defaultMessage}</p>`;
    } else {
      const defaultContentElement = document.getElementById('dropZoneDefaultContent');
      const defaultContent = defaultContentElement ? defaultContentElement.innerHTML : defaultMessage;
      dropZone.innerHTML = defaultContent;
    }
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
        const defaultContentElement = document.getElementById('dropZoneDefaultContent');
        const defaultContent = defaultContentElement ? defaultContentElement.innerHTML : defaultMessage;
        dropZone.innerHTML = defaultContent;
      }
    });
  }

  // 個別画像用の表示更新関数
  function updateOrderImageDisplay(imageUrl) {
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
  setupClickEvent(dropZone, updatePreview, droppedImage);

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
function setupClickEvent(dropZone, updatePreview, droppedImage) {
  dropZone.addEventListener('click', () => {
    if (!droppedImage) {
      const fileInput = document.createElement('input');
      fileInput.type = 'file';
      fileInput.accept = 'image/jpeg, image/png, image/svg+xml';
      fileInput.style.display = 'none';
      document.body.appendChild(fileInput);
      fileInput.click();

      fileInput.addEventListener('change', () => {
        const file = fileInput.files[0];
        if (file && file.type.match(/^image\/(jpeg|png|svg\+xml)$/)) {
          const reader = new FileReader();
          reader.onload = (e) => {
            updatePreview(e.target.result);
          };
          reader.readAsDataURL(file);
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
  const fileInput = document.getElementById("file");
  const executeButton = document.getElementById("executeButton");
  const printButton = document.getElementById("printButton");

  if (fileInput.files.length > 0) {
    executeButton.disabled = false;
    printButton.disabled = false;
  } else {
    executeButton.disabled = true;
    printButton.disabled = true;
  }
});

// ページロード時に説明を表示し、ユーザーにファイル選択を促す
window.addEventListener("load", function() {
  const fileInput = document.getElementById("file");
  const executeButton = document.getElementById("executeButton");
  const printButton = document.getElementById("printButton");

  // 初期状態でボタンを無効化
  executeButton.disabled = true;
  printButton.disabled = true;

  // ファイル選択を促すメッセージを表示
  alert("CSVファイルを選択してください。");
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
