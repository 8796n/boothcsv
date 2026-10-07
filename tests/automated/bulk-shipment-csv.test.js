const test = require('node:test');
const assert = require('node:assert/strict');
const Papa = require('../../js/papaparse.min.js');
const { buildShipmentCsv, buildShipmentCsvFromRecords } = require('../../js/bulk-shipment-csv.js');
const { parseOrderCsv, anonymousAddressLabel, SHIPMENT_HEADERS } = require('../../js/order-csv.js');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

test('発送通知CSVは選択注文だけを残し、列順とコメント以外の値を保持する', function() {
  const header = ['フォーマット番号', '注文番号', '配送番号', 'ユーザー識別コード', 'お支払い方法', '注文日時', '支払い日時', 'BOOST合計', '合計金額', '発送サイズ', 'ゆうゆうBOOTHパック/ポスト発送確認符号', '郵便番号', '都道府県', '市区町村・丁目・番地', 'マンション・建物名・部屋番号', '氏名', '電話番号', '商品ID / 数量 / 商品名', '発送用伝票番号', '発送通知コメント'];
  const row = ['1', '000123', '000456', 'test-user', 'クレジットカード', '2026-10-06 15:41:09', '2026-10-06 15:41:09', '0', '1200', 'あんしんBOOTHパック ネコポス', '', '', '', '', '', '', '', '商品ID : 1 / 数量 : 1 / 商品,"A"\r\n商品ID : 2 / 数量 : 2 / 商品B', '-', '元のコメント'];
  const other = row.map((value, index) => index === 1 ? '789' : value);
  const source = '\uFEFF' + Papa.unparse([header, row, other], { quotes: true });
  const comment = '発送しました。\n備考: "取扱注意", よろしくお願いします。';
  const result = buildShipmentCsv(source, [' 000123 ', '000123', '999'], comment);
  const output = Papa.parse(result.csv.replace(/^\uFEFF/, '')).data;
  assert.deepEqual(output, [header, [...row.slice(0, -1), comment]]);
  assert.deepEqual(result.orderNumbers, ['000123']);
  assert.deepEqual(result.missingOrderNumbers, ['999']);
  assert.equal(result.rowCount, 1);
  assert.equal(result.csv.charCodeAt(0), 0xfeff);
  assert.equal(buildShipmentCsv(result.csv, ['000123'], comment).csv, result.csv);
  assert.equal(Papa.parse(buildShipmentCsv(source, ['000123'], '').csv).data[1][19], '');

  const empty = buildShipmentCsv(Papa.unparse([header]), ['000123'], comment);
  assert.equal(empty.rowCount, 0);
  assert.deepEqual(empty.orderNumbers, []);
  assert.deepEqual(empty.missingOrderNumbers, ['000123']);
  assert.equal(buildShipmentCsv(source, [], comment).rowCount, 0);
  assert.equal(buildShipmentCsv(source, ['123'], comment).rowCount, 0, '注文番号を数値化しない');
  assert.throws(() => buildShipmentCsv('注文番号,商品名\n123,商品', ['123'], comment), /列構成/);
  assert.throws(() => buildShipmentCsv(Papa.unparse([header, ['1', '123']]), ['123'], comment), /列数/);
  assert.throws(() => buildShipmentCsv(Papa.unparse([header, row]) + '\r\n"閉じない引用符', ['123'], comment), /読み取れません/);
  assert.throws(() => buildShipmentCsv(Papa.unparse([[...header, '注文番号']]), ['123'], comment), /列構成/);
  assert.throws(() => buildShipmentCsv(Papa.unparse([header, row, row]), ['123'], comment), /重複/);
});

function shipmentFixture(overrides = {}) {
  const values = SHIPMENT_HEADERS.map(name => ({
    'フォーマット番号': '1', '注文番号': '000123', '配送番号': '000456',
    'お支払い方法': 'atone翌月払い（コンビニ/口座振替）',
    '発送サイズ': 'あんしんBOOTHパック ネコポス', '発送用伝票番号': '-',
    '商品ID / 数量 / 商品名': '商品ID : 1 / 数量 : 2 / 商品,"A"\r\n商品ID : 2 / 数量 : 1 / 商品B',
    ...overrides
  })[name] || '');
  return '\uFEFF' + Papa.unparse([SHIPMENT_HEADERS, values], { quotes: true });
}

test('新旧を判別し、空欄と元の20列を保ち、未知形式を保存前に拒否できる', function() {
  const oldText = fs.readFileSync(path.join(__dirname, '../../sample/booth_orders_sample.csv'), 'utf8');
  const legacy = parseOrderCsv(oldText);
  assert.equal(legacy.kind, 'legacy');
  assert.equal(legacy.shipmentCsvByOrder.size, 0);
  const parsed = parseOrderCsv(shipmentFixture(), { sourceName: 'new.csv' });
  const row = parsed.data[0];
  assert.equal(parsed.kind, 'shipment');
  assert.equal(Object.keys(row).length, 14);
  assert.equal(row['お支払方法'], 'atone翌月払い（コンビニ/口座振替）');
  assert.equal(row['氏名'], '');
  assert.equal(Object.hasOwn(row, '注文状況'), false);
  assert.equal(Object.hasOwn(row, '発送日時'), false);
  const source = parsed.shipmentCsvByOrder.get('000123');
  assert.equal(source.formatVersion, '1');
  assert.equal(source.sourceName, 'new.csv');
  assert.equal(source.values[2], '000456');
  assert.equal(source.values[19], '');
  assert.equal(anonymousAddressLabel(row), '【あんしんBOOTHパック】');
  assert.equal(anonymousAddressLabel({'発送サイズ':'ゆうゆうBOOTHパック 仮のサイズ'}), '【ゆうゆうBOOTHパック】');
  assert.equal(anonymousAddressLabel({'発送サイズ':'宅急便'}), '');
  assert.equal(anonymousAddressLabel({}), '');
  assert.throws(() => parseOrderCsv(shipmentFixture({'フォーマット番号':'2'})), /フォーマット番号/);
  assert.throws(() => parseOrderCsv(shipmentFixture({'フォーマット番号':''})), /フォーマット番号/);
  const mixed = Papa.parse(shipmentFixture().replace(/^\uFEFF/, '')).data;
  mixed.push(mixed[1].map((value, i) => i === 0 ? '2' : value));
  assert.throws(() => parseOrderCsv(Papa.unparse(mixed)), /フォーマット番号/);
  assert.throws(() => parseOrderCsv(shipmentFixture({'注文番号':' '})), /注文番号/);
  assert.equal(parseOrderCsv(Papa.unparse([SHIPMENT_HEADERS])).data.length, 0);
});

test('新CSV未取り込み・発送済みは出力せず、保存元データを変更しない', function() {
  const parsed = parseOrderCsv(shipmentFixture());
  const source = parsed.shipmentCsvByOrder.get('000123');
  const records = [
    {orderNumber:'000123', shipmentCsv:source, shippedAt:null},
    {orderNumber:'legacy-only', row:{'注文番号':'legacy-only'}, shippedAt:null},
    {orderNumber:'shipped', shippedAt:'2026-10-01T00:00:00Z'}
  ];
  const comment = '発送しました。\n"コメント", テスト';
  const before = JSON.stringify(source);
  const result = buildShipmentCsvFromRecords(records, comment);
  assert.deepEqual(result.orderNumbers, ['000123']);
  assert.deepEqual(result.missingOrderNumbers, ['legacy-only']);
  assert.equal(result.shippedComment, comment);
  const output = Papa.parse(result.csv.replace(/^\uFEFF/, '')).data;
  assert.deepEqual(output[0], source.headers);
  assert.deepEqual(output[1].slice(0, -1), source.values.slice(0, -1));
  assert.equal(output[1][19], comment);
  assert.equal(JSON.stringify(source), before);
  assert.equal(buildShipmentCsvFromRecords([records[1]], '').rowCount, 0);
  assert.equal(buildShipmentCsvFromRecords([records[2]], '').rowCount, 0);
  assert.throws(() => buildShipmentCsvFromRecords([{orderNumber:'other',shipmentCsv:source}], ''), /注文番号/);
  assert.throws(() => buildShipmentCsvFromRecords([{orderNumber:'000123',shipmentCsv:{...source,formatVersion:'2'}}], ''), /フォーマット番号/);
});

test('交互取り込みと再読み込みで実績・QR・元CSVを保持し、コメントは完了記録時だけ更新する', async function() {
  const saved = new Map();
  const db = {
    getAllOrders: async () => structuredClone([...saved.values()]),
    saveOrder: async record => saved.set(record.orderNumber, structuredClone(record))
  };
  const context = {window:{}, DEBUG_MODE:false, CONSTANTS:{CSV:{ORDER_NUMBER_COLUMN:'注文番号'}}};
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../../js/order-repository.js'), 'utf8'), context);
  const repo = new context.window.OrderRepository(db);
  const parsed = parseOrderCsv(shipmentFixture());
  const oldRow = {...parsed.data[0], '注文状況':'支払済み', '発送日時':''};
  delete oldRow['発送サイズ'];
  await repo.bulkUpsert([oldRow]);
  const record = repo.get('000123');
  assert.equal(record.shippedComment, null);
  await repo.setOrderQRData('000123', {qrimage:new ArrayBuffer(3), receiptnum:'test', receiptpassword:'test'});
  await repo.setOrderImage('000123', {data:new ArrayBuffer(2),mimeType:'image/png'});
  await repo.markPrinted('000123','printed');
  await repo.markShipped('000123','shipped','出力時の文面');
  const createdAt = record.createdAt;
  await repo.bulkUpsert(parsed.data, parsed.shipmentCsvByOrder);
  await repo.bulkUpsert([{...oldRow,'氏名':''}]);
  assert.equal(repo.getAll().length, 1);
  assert.equal(record.row['発送サイズ'], 'あんしんBOOTHパック ネコポス');
  assert.equal(record.row['注文状況'], '支払済み');
  assert.equal(record.row['氏名'], '');
  assert.equal(record.shipmentCsv.values[19], '');
  assert.equal(record.shippedComment, '出力時の文面');
  assert.equal(record.printedAt, 'printed');
  assert.equal(record.shippedAt, 'shipped');
  assert.equal(record.createdAt, createdAt);
  assert.equal(record.qr.qrimage.byteLength, 3);
  assert.equal(record.image.data.byteLength, 2);
  await repo.markShipped('000123','resynced');
  assert.equal(record.shippedComment, '出力時の文面');
  await repo.markShipped('000123','resynced','');
  assert.equal(record.shippedComment, '');
  const reopened = new context.window.OrderRepository(db);
  await reopened.init();
  assert.equal(reopened.get('000123').shipmentCsv.formatVersion, '1');
  assert.equal(reopened.get('000123').shippedComment, '');
  await reopened.clearShipped('000123');
  assert.equal(reopened.get('000123').shippedComment, null);
});

test('通知画面は保存済みCSVを再利用し、出力した注文とコメントだけを完了時に記録する', async function() {
  const app = fs.readFileSync(path.join(__dirname, '../../js/boothcsv.js'), 'utf8');
  const parsed = parseOrderCsv(shipmentFixture());
  const records = new Map([
    ['000123',{orderNumber:'000123',row:parsed.data[0],shipmentCsv:parsed.shipmentCsvByOrder.get('000123')}],
    ['old',{orderNumber:'old',row:{'注文番号':'old'}}]
  ]);
  const marked = [];
  const repo = {get:key=>records.get(key), markShipped:async (...args)=>marked.push(args)};
  const elements = new Map();
  function element(id) {
    if (!elements.has(id)) elements.set(id, {
      textContent:'',value:'',files:[],disabled:false,hidden:false,open:false,
      showModal(){this.open=true;},close(){this.open=false;},click(){},remove(){}
    });
    return elements.get(id);
  }
  const blobs = [];
  let saved = 0;
  const context = {
    ensureOrderRepository:async()=>repo,
    document:{getElementById:element,createElement:()=>element('downloadLink'),body:{appendChild(){}}},
    window:{BoothCSVOrders:require('../../js/order-csv.js'),BoothCSVBulkShipment:require('../../js/bulk-shipment-csv.js'),open(){}},
    URL:{createObjectURL:blob=>{blobs.push(blob);return 'blob:test';}},Blob,
    setTimeout(){},revokeBlobUrl(){},refreshProcessedOrdersPanel:async()=>{},
    formatShippingCommentPreview:text=>text,
    persistCsvToRepository:async result=>{
      saved++;
      result.data.forEach(row=>records.set(row['注文番号'],{orderNumber:row['注文番号'],row,shipmentCsv:result.shipmentCsvByOrder.get(row['注文番号'])}));
    }
  };
  vm.runInNewContext(app.slice(app.indexOf('async function openShipmentCsvDialog('),app.indexOf('function buildShipmentConfirmationMessage(')),context);
  await context.openShipmentCsvDialog(['000123'],'出力時のコメント');
  assert.equal(element('shipmentCsvSourceSteps').hidden,true);
  assert.equal(element('shipmentCsvDownload').disabled,false);
  assert.equal(element('shipmentCsvRecord').disabled,true);
  await element('shipmentCsvRecord').onclick();
  assert.equal(marked.length,0,'ダウンロード前は実績を記録しない');
  await context.openShipmentCsvDialog(['000123','old'],'出力時のコメント');
  assert.match(element('shipmentCsvStatus').textContent,/出力対象外.*old/);
  assert.equal(element('shipmentCsvDownload').disabled,false,'取り込み済みの注文だけで続行できる');
  element('shipmentCsvDownload').onclick();
  assert.equal(marked.length,0,'ダウンロードだけでは実績を記録しない');
  assert.equal(Papa.parse((await blobs[0].text()).replace(/^\uFEFF/, '')).data[1][19],'出力時のコメント');
  await element('shipmentCsvRecord').onclick();
  assert.equal(marked.length,1);
  assert.equal(marked[0][0],'000123');
  assert.equal(marked[0][2],'出力時のコメント');
  await context.openShipmentCsvDialog(['old'],'');
  assert.equal(element('shipmentCsvDownload').disabled,true);
  assert.equal(element('shipmentCsvSourceSteps').hidden,false);
  element('shipmentCsvFile').files=[{name:'unknown.csv',text:async()=>shipmentFixture({'フォーマット番号':'2'})}];
  await element('shipmentCsvFile').onchange();
  assert.match(element('shipmentCsvFileError').textContent,/フォーマット番号/);
  assert.equal(saved,0,'未知形式では保存処理を呼ばない');
  element('shipmentCsvFile').files=[{name:'new.csv',text:async()=>shipmentFixture({'注文番号':'old'})}];
  await element('shipmentCsvFile').onchange();
  assert.equal(saved,1);
  assert.equal(element('shipmentCsvDownload').disabled,false);
  assert.equal(element('shipmentCsvSourceSteps').hidden,true);
});

test('拡張から通知した文面だけを保存し、発送日時の再取得ではコメントを更新しない', async function() {
  const app = fs.readFileSync(path.join(__dirname, '../../js/boothcsv.js'), 'utf8');
  const marked = [];
  const context = {
    normalizeOrderSelection:numbers=>numbers,canUseExtensionBridge:()=>true,
    settingsCache:{shippingMessageTemplate:'定型文'},
    ensureOrderRepository:async()=>({markShipped:async(...args)=>marked.push(args)}),
    window:{BoothCSVExtensionBridge:{
      collectOrderShipmentStatus:async()=>({ok:true}),
      notifyOrderShipment:async()=>({ok:true,submitted:true,shippedAt:'new-time',shippedComment:'フォームに入った文面'})
    }},
    collectShipmentConfirmationStatus:async()=>({responses:new Map([['done',{ok:true,shippedAt:'old-time'}]])}),
    buildShipmentConfirmationMessage:()=>'',confirm:()=>true,
    refreshProcessedOrdersPanel:async()=>{},alert(){}
  };
  vm.runInNewContext(app.slice(app.indexOf('async function notifySelectedOrdersShipment('),app.indexOf('async function persistCsvToRepository(')),context);
  await context.notifySelectedOrdersShipment(['done','new']);
  assert.equal(marked[0].length,2);
  assert.equal(marked[1][2],'フォームに入った文面');
});
