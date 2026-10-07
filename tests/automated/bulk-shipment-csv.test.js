const test = require('node:test');
const assert = require('node:assert/strict');
const Papa = require('../../js/papaparse.min.js');
const { buildShipmentCsv } = require('../../js/bulk-shipment-csv.js');

test('発送通知CSVは選択注文の全配送行だけを残し、列順とコメント以外の値を保持する', function() {
  const header = ['フォーマット番号', '注文番号', '配送番号', 'ユーザー識別コード', 'お支払い方法', '注文日時', '支払い日時', 'BOOST合計', '合計金額', '発送サイズ', 'ゆうゆうBOOTHパック/ポスト発送確認符号', '郵便番号', '都道府県', '市区町村・丁目・番地', 'マンション・建物名・部屋番号', '氏名', '電話番号', '商品ID / 数量 / 商品名', '発送用伝票番号', '発送通知コメント'];
  const row = ['1', '000123', '000456', 'test-user', 'クレジットカード', '2026-10-06 15:41:09', '2026-10-06 15:41:09', '0', '1200', 'あんしんBOOTHパック ネコポス', '', '', '', '', '', '', '', '商品ID : 1 / 数量 : 1 / 商品,"A"\r\n商品ID : 2 / 数量 : 2 / 商品B', '-', '元のコメント'];
  const other = row.map((value, index) => index === 1 ? '789' : value);
  const secondDelivery = row.map((value, index) => index === 2 ? '000457' : value);
  const source = '\uFEFF' + Papa.unparse([header, row, other, secondDelivery], { quotes: true });
  const comment = '発送しました。\n備考: "取扱注意", よろしくお願いします。';
  const result = buildShipmentCsv(source, [' 000123 ', '000123', '999'], comment);
  const output = Papa.parse(result.csv.replace(/^\uFEFF/, '')).data;
  assert.deepEqual(output, [header, [...row.slice(0, -1), comment], [...secondDelivery.slice(0, -1), comment]]);
  assert.deepEqual(result.orderNumbers, ['000123']);
  assert.deepEqual(result.missingOrderNumbers, ['999']);
  assert.equal(result.rowCount, 2);
  assert.equal(result.csv.charCodeAt(0), 0xfeff);
  assert.equal(buildShipmentCsv(result.csv, ['000123'], comment).csv, result.csv);
  assert.equal(Papa.parse(buildShipmentCsv(source, ['000123'], '').csv).data[1][19], '');

  const empty = buildShipmentCsv(Papa.unparse([header]), ['000123'], comment);
  assert.equal(empty.rowCount, 0);
  assert.deepEqual(empty.orderNumbers, []);
  assert.deepEqual(empty.missingOrderNumbers, ['000123']);
  assert.equal(buildShipmentCsv(source, [], comment).rowCount, 0);
  assert.equal(buildShipmentCsv(source, ['123'], comment).rowCount, 0, '注文番号を数値化しない');
  assert.throws(() => buildShipmentCsv('注文番号,商品名\n123,商品', ['123'], comment), /発送通知用CSV/);
  assert.throws(() => buildShipmentCsv(Papa.unparse([header, ['1', '123']]), ['123'], comment), /列数/);
  assert.throws(() => buildShipmentCsv(Papa.unparse([header, row]) + '\r\n"閉じない引用符', ['123'], comment), /読み取れません/);
  assert.throws(() => buildShipmentCsv(Papa.unparse([[...header, '注文番号']]), ['123'], comment), /発送通知用CSV/);
});
