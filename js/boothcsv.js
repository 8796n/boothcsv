window.addEventListener("load", function(){
  const labelyn = localStorage.getItem("labelyn");
  const labelskip = localStorage.getItem("labelskip");
  console.log(labelyn, labelskip);
  if(labelyn == null || labelyn == "true"){
    document.getElementById("labelyn").checked = true;
  }else{
    document.getElementById("labelyn").checked = false;
  }
  document.getElementById("labelskipnum").value = labelskip != null ? labelskip : 0;

   // 画像ドロップゾーンの初期化
  const imageDropZoneElement = document.getElementById('imageDropZone');
  const imageDropZone = createOrderImageDropZone();
  imageDropZoneElement.appendChild(imageDropZone.element);
  window.orderImageDropZone = imageDropZone;

  // 全ての画像をクリアするボタンのイベントリスナーを追加
  const clearAllButton = document.getElementById('clearAllButton');
  clearAllButton.onclick = () => {
    if (confirm('本当に全てのQR画像をクリアしますか？')) {
      // localStorageからQR画像を削除
      Object.keys(localStorage).forEach(key => {
        if (localStorage.getItem(key)?.includes('qrimage')) {
          localStorage.removeItem(key);
        }
      });
      alert('全てのQR画像をクリアしました');
      location.reload(); // ページをリロードして反映
    }
  };

  // 全ての注文画像をクリアするボタンのイベントリスナーを追加
  const clearAllOrderImagesButton = document.getElementById('clearAllOrderImagesButton');
  clearAllOrderImagesButton.onclick = () => {
    if (confirm('本当に全ての注文画像（グローバル画像と個別画像）をクリアしますか？')) {
      // localStorageから注文画像を削除
      Object.keys(localStorage).forEach(key => {
        if (key === 'orderImage' || key.startsWith('orderImage_')) {
          localStorage.removeItem(key);
        }
      });
      alert('全ての注文画像をクリアしました');
      location.reload(); // ページをリロードして反映
    }
  };

   // ページロード時にチェックボックスの状態を反映
   const sortByPaymentDate = localStorage.getItem("sortByPaymentDate") === "true";
   document.getElementById("sortByPaymentDate").checked = sortByPaymentDate;

   // チェックボックスの状態が変更されたときにlocalStorageに保存
   document.getElementById("sortByPaymentDate").addEventListener("change", function() {
     localStorage.setItem("sortByPaymentDate", this.checked);
   });

}, false);

function clickstart() {
  for (let sheet of document.querySelectorAll('section')) {
    sheet.parentNode.removeChild(sheet);
  }
  const file = document.getElementById("file").files[0];
  const labelyn = document.getElementById("labelyn").checked;
  localStorage.setItem("labelyn", labelyn);
  const labelskip = document.getElementById("labelskipnum").value;
  localStorage.setItem("labelskip", labelskip);
  const labelarr = [];
  const labelskipNum = parseInt(labelskip, 10) || 0;
  if (labelskipNum > 0) {
    for (let i = 0; i < labelskipNum; i++) {
      labelarr.push("");
    }
  }
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: function(results) {
      // チェックボックスの状態に応じて並び替え
      const sortByPaymentDate = document.getElementById("sortByPaymentDate").checked;
      if (sortByPaymentDate) {
        results.data.sort((a, b) => {
          const timeA = a["支払い日時"] || "";
          const timeB = b["支払い日時"] || "";
          return timeA.localeCompare(timeB);
        });
      }

      const tOrder = document.querySelector('#注文明細');
      for (let row of results.data) {
        const cOrder = document.importNode(tOrder.content, true);
        let orderNumber = '';
        for (let c of Object.keys(row).filter(key => key != "商品ID / 数量 / 商品名")) {
          const divc = cOrder.querySelector("." + c);
          if (divc) {
            if (c == "注文番号") {
              orderNumber = row[c];
              divc.textContent = "注文番号 : " + row[c];
              labelarr.push(row[c]);
            } else if (row[c]) {
              divc.textContent = row[c];
            }
          }
        }

        // 個別注文画像ドロップゾーンを作成
        const individualDropZoneContainer = cOrder.querySelector('.individual-image-dropzone');
        if (individualDropZoneContainer && orderNumber) {
          const individualImageDropZone = createIndividualOrderImageDropZone(orderNumber);
          individualDropZoneContainer.appendChild(individualImageDropZone.element);
        }
        const tItems = cOrder.querySelector('#商品');
        const trSpace = cOrder.querySelector('.spacerow');
        if (row["商品ID / 数量 / 商品名"]) {
          for(let itemrow of row["商品ID / 数量 / 商品名"].split('\n')){
            const cItem = document.importNode(tItems.content, true);
            // 最初の2つの ' / ' でのみ分割する
            const firstSplit = itemrow.split(' / ');
          // 商品IDと数量をコロンで分割して2番目の要素を取得し、trimで空白を除去
          const itemIdSplit = firstSplit[0].split(':');
          const itemId = itemIdSplit.length > 1 ? itemIdSplit[1].trim() : '';
          const quantitySplit = firstSplit[1] ? firstSplit[1].split(':') : [];
          const quantity = quantitySplit.length > 1 ? quantitySplit[1].trim() : '';
          const productName = firstSplit.slice(2).join(' / '); // 残りの部分を商品名として結合

            if (itemId && quantity) {
              const tdId = cItem.querySelector(".商品ID");
              if (tdId) {
                tdId.textContent = itemId;
              }
              const tdQuantity = cItem.querySelector(".数量");
              if (tdQuantity) {
                tdQuantity.textContent = quantity;
              }
              const tdName = cItem.querySelector(".商品名");
              if (tdName) {
                tdName.textContent = productName;
              }
            }
            trSpace.parentNode.parentNode.insertBefore(cItem, trSpace.parentNode);
          }
        }

        // 画像表示の優先度: 個別画像 > グローバル画像
        let imageToShow = null;
        if (orderNumber) {
          // 個別画像があるかチェック
          const individualImage = localStorage.getItem(`orderImage_${orderNumber}`);
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
        document.body.appendChild(cOrder);
      }
      if(labelyn && labelarr.length > 0){
        if(labelarr.length % 44){
          for(let i = 0; i < labelarr.length % 44; i++){
            labelarr.push("");
          }
        }
        const tL44 = document.querySelector('#L44');
        let cL44 = document.importNode(tL44.content, true);
        let tableL44 = cL44.querySelector("table");
        let tr = document.createElement("tr");
        let i = 0;
        for(let label of labelarr){
          if(i > 0 && i % 44 == 0){
            tableL44.appendChild(tr);
            tr = document.createElement("tr");
            document.body.insertBefore(cL44, tL44);
            cL44 = document.importNode(tL44.content, true);
            tableL44 = cL44.querySelector("table");
            tr = document.createElement("tr");
          }else if(i > 0 && i % 4 == 0){
            tableL44.appendChild(tr);
            tr = document.createElement("tr");
          }
            tr.appendChild(createLabel(label));
          i++;
        }
        tableL44.appendChild(tr);

        document.body.insertBefore(cL44, tL44);
      }
      const labelTable = document.querySelectorAll(".label44");
      const pageDiv = document.querySelectorAll(".page");
      if(labelTable.length > 0){
        alert("印刷枚数\nA4 44面ラベルシール:" + labelTable.length + "枚\nA4普通紙:" + pageDiv.length + "枚");
      }else if(pageDiv.length > 0){
        alert("印刷枚数\nA4普通紙:" + pageDiv.length + "枚");
      }
    }
  });
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

function createLabel(ordernum=""){
  const divQr = createDiv('qr');
  const divOrdernum = createDiv('ordernum', ordernum);
  const divYamato = createDiv('yamato');

  if(ordernum){
    createDropzone(divQr);
    const qr = JSON.parse(localStorage.getItem(ordernum));
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
  const img = new Image();
  img.crossOrigin = 'anonymous';
  img.src = elImage.src;
  img.onload = function() {
    const canv = document.createElement("canvas");
    const context = canv.getContext("2d");
    canv.width = img.width;
    canv.height = img.height;
    context.drawImage(img, 0, 0, canv.width, canv.height);
    const b64data = canv.toDataURL("image/png");
    //console.log(b64data);
    const imageData = context.getImageData(0, 0, canv.width, canv.height);
    const barcode = jsQR(imageData.data, imageData.width, imageData.height);
    if(barcode){
      console.log("[" + barcode.data + "]");
      const b = String(barcode.data).replace(/^\s+|\s+$/g,'').replace(/ +/g,' ').split(" ");
      if(b.length == 3){
        const ordernum = elImage.closest("td").querySelector(".ordernum p").innerHTML;
        const d = elImage.closest("td").querySelector(".yamato");
        d.innerHTML = "";
        for(let i=1; i < 3; i++){
          const p = document.createElement("p");
          p.innerText = b[i];
          d.appendChild(p);
        }
        const json = { "receiptnum": b[1], "receiptpassword": b[2], "qrimage": b64data };
        //console.log(JSON.stringify(json));
       
        localStorage.setItem(ordernum, JSON.stringify(json));
      }
    }
  };
}

function attachImage(file, elImage) {
  const reader = new FileReader();
  reader.onload = function(event) {
      const src = event.target.result;
      elImage.src = src;
      elImage.setAttribute('title', file.name);
      elImage.onload = function() {
        readQR(elImage);
      }
  };
  reader.readAsDataURL(file);
}

function createOrderImageDropZone() {
  const dropZone = document.createElement('div');
  dropZone.classList.add('order-image-drop');

  let droppedImage = null;

  // localStorageから保存された画像を読み込む
  const savedImage = localStorage.getItem('orderImage');
  if (savedImage) {
    updatePreview(savedImage);
  } else {
    const defaultContent = document.getElementById('dropZoneDefaultContent').innerHTML;
    dropZone.innerHTML = defaultContent;
  }

  function updatePreview(imageUrl) {
    droppedImage = imageUrl;
    dropZone.innerHTML = ''; // クリア
    const preview = document.createElement('img');
    preview.src = imageUrl;
    preview.classList.add('preview-image');
    preview.title = 'クリックでリセット'; // Tooltipを追加
    dropZone.appendChild(preview);
    // localStorageに保存
    localStorage.setItem('orderImage', imageUrl);

    // 画像クリックでリセット
    preview.addEventListener('click', (e) => {
      e.stopPropagation(); // イベントの伝播を停止
      localStorage.removeItem('orderImage');
      droppedImage = null;
      const defaultContent = document.getElementById('dropZoneDefaultContent').innerHTML;
      dropZone.innerHTML = defaultContent;
    });
  }

  // ドラッグ&ドロップイベントの設定
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');

    const file = e.dataTransfer.files[0];
    if (file && file.type.match(/^image\/(jpeg|png|svg\+xml)$/)) {
      const reader = new FileReader();
      reader.onload = (e) => {
        updatePreview(e.target.result);
      };
      reader.readAsDataURL(file);
    }
  });

  // クリックイベント：画像がない場合はファイル選択、ある場合は何もしない（画像のクリックでリセット）
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

  return {
    element: dropZone,
    getImage: () => droppedImage
  };
}

// 個別注文用の画像ドロップゾーンを作成する関数
function createIndividualOrderImageDropZone(orderNumber) {
  const dropZone = document.createElement('div');
  dropZone.classList.add('individual-order-image-drop');
  dropZone.style.cssText = 'min-height: 80px; border: 1px dashed #999; padding: 5px; background: #f9f9f9;';

  let droppedImage = null;
  const storageKey = `orderImage_${orderNumber}`;

  // localStorageから保存された画像を読み込む
  const savedImage = localStorage.getItem(storageKey);
  if (savedImage) {
    updatePreview(savedImage);
  } else {
    dropZone.innerHTML = '<p style="margin: 5px; font-size: 12px; color: #666;">画像をドロップ or クリックで選択</p>';
  }

  // 注文明細の画像表示を更新する関数
  function updateOrderImageDisplay(imageUrl) {
    // 現在の注文明細のコンテナを取得
    const orderSection = dropZone.closest('section');
    if (!orderSection) return;

    const imageContainer = orderSection.querySelector('.order-image-container');
    if (!imageContainer) return;

    // 既存の画像をクリア
    imageContainer.innerHTML = '';

    if (imageUrl) {
      // 新しい画像を表示
      const imageDiv = document.createElement('div');
      imageDiv.classList.add('order-image');

      const img = document.createElement('img');
      img.src = imageUrl;

      imageDiv.appendChild(img);
      imageContainer.appendChild(imageDiv);
    } else {
      // 画像がない場合はグローバル画像があれば表示
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

  function updatePreview(imageUrl) {
    droppedImage = imageUrl;
    dropZone.innerHTML = ''; // クリア
    const preview = document.createElement('img');
    preview.src = imageUrl;
    preview.style.cssText = 'max-width: 100%; max-height: 60px; cursor: pointer;';
    preview.title = 'クリックでリセット';
    dropZone.appendChild(preview);
    // localStorageに保存
    localStorage.setItem(storageKey, imageUrl);

    // 注文明細の画像表示を即座に更新
    updateOrderImageDisplay(imageUrl);

    // 画像クリックでリセット
    preview.addEventListener('click', (e) => {
      e.stopPropagation();
      localStorage.removeItem(storageKey);
      droppedImage = null;
      dropZone.innerHTML = '<p style="margin: 5px; font-size: 12px; color: #666;">画像をドロップ or クリックで選択</p>';
      
      // 画像をリセットした時も表示を更新
      updateOrderImageDisplay(null);
    });
  }

  // ドラッグ&ドロップイベントの設定
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.style.backgroundColor = '#e6f3ff';
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.style.backgroundColor = '#f9f9f9';
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.style.backgroundColor = '#f9f9f9';

    const file = e.dataTransfer.files[0];
    if (file && file.type.match(/^image\/(jpeg|png|svg\+xml)$/)) {
      const reader = new FileReader();
      reader.onload = (e) => {
        updatePreview(e.target.result);
      };
      reader.readAsDataURL(file);
    }
  });

  // クリックイベント
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

  return {
    element: dropZone,
    getImage: () => droppedImage
  };
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
