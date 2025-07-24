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
}, false);

function clickstart() {
  for(let sheet of document.querySelectorAll('section')){
        sheet.parentNode.removeChild(sheet);
  }
  const file = document.getElementById("file").files[0];
  const labelyn = document.getElementById("labelyn").checked;
  localStorage.setItem("labelyn", labelyn);
  const labelskip = document.getElementById("labelskipnum").value;
  localStorage.setItem("labelskip", labelskip);
  const labelarr = [];
  if(labelskip > 0){
    for(let i = 0; i < labelskip; i++){
      labelarr.push("");
    }
  }
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: function(results) {
      const tOrder = document.querySelector('#注文明細');
      for(let row of results.data){
        const cOrder = document.importNode(tOrder.content, true);
        for(let c of Object.keys(row).filter(key => key != "商品ID / 数量 / 商品名")){
          const divc = cOrder.querySelector("." + c);
          if(divc){
            if(c=="注文番号"){
              divc.textContent = "注文番号 : " + row[c]
              labelarr.push(row[c]);
            }else if(row[c]){
              divc.textContent = row[c];
            }
          }
        }
        const tItems = cOrder.querySelector('#商品');
        const trSpace = cOrder.querySelector('.spacerow');
        const trItem = cOrder.querySelector('.items');
        for(let itemrow of row["商品ID / 数量 / 商品名"].split('\n')){
          const cItem = document.importNode(tItems.content, true);
          // 最初の2つの ' / ' でのみ分割する
          const firstSplit = itemrow.split(' / ');
          // 商品IDと数量から余計な文字列を取り除く
          const itemId = firstSplit[0].replace('商品ID : ', '');
          const quantity = firstSplit[1].replace('数量 : ', '');
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
        document.body.appendChild(cOrder);
      }
      if(labelyn && labelarr.length > 0){
        if(labelarr.length % 44){
          for(let i = 0; i < labelarr.length % 44; i++){
            labelarr.push("");
          }
        }
        const tL44 = document.querySelector('#L44');
        const cL44 = document.importNode(tL44.content, true);
        const tableL44 = cL44.querySelector("table");
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