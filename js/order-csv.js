(function(root, factory) {
  const api = factory(typeof module === 'object' && module.exports ? require('./papaparse.min.js') : root.Papa);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.BoothCSVOrders = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(Papa) {
  'use strict';

  const LEGACY_HEADERS = ['注文番号', 'ユーザー識別コード', 'お支払方法', '注文状況', '注文日時', '支払い日時', '発送日時', '合計金額', '郵便番号', '都道府県', '市区町村・丁目・番地', 'マンション・建物名・部屋番号', '氏名', '電話番号', '商品ID / 数量 / 商品名'];
  const SHIPMENT_HEADERS = ['フォーマット番号', '注文番号', '配送番号', 'ユーザー識別コード', 'お支払い方法', '注文日時', '支払い日時', 'BOOST合計', '合計金額', '発送サイズ', 'ゆうゆうBOOTHパック/ポスト発送確認符号', '郵便番号', '都道府県', '市区町村・丁目・番地', 'マンション・建物名・部屋番号', '氏名', '電話番号', '商品ID / 数量 / 商品名', '発送用伝票番号', '発送通知コメント'];

  function validateRows(headers, rows, kind) {
    if (kind === 'shipment' && rows.some(row => !Array.isArray(row) || row[0] !== '1')) {
      throw new Error('未対応のフォーマット番号です。対応している番号は「1」です。');
    }
    const expected = kind === 'shipment' ? SHIPMENT_HEADERS : LEGACY_HEADERS;
    if (!Array.isArray(headers) || headers.length !== expected.length || headers.some((name, i) => name !== expected[i])) {
      throw new Error('CSVの列構成が対応形式と一致しません。BOOTHから取得したCSVを確認してください。');
    }
    const seen = new Set();
    for (const row of rows) {
      if (!Array.isArray(row) || row.length !== headers.length || row.some(value => typeof value !== 'string')) {
        throw new Error('CSVの列数または値の形式が一致しません。');
      }
      const number = row[kind === 'shipment' ? 1 : 0].trim();
      if (!number) throw new Error('注文番号が空の行があります。');
      if (seen.has(number)) throw new Error('同じ注文番号の行が重複しています。');
      seen.add(number);
    }
  }

  function parseOrderCsv(csvText, options = {}) {
    const parsed = Papa.parse(String(csvText).replace(/^\uFEFF/, ''), { delimiter: ',', skipEmptyLines: 'greedy' });
    if (parsed.errors.length) throw new Error('CSVを読み取れませんでした。BOOTHから取得したCSVを確認してください。');
    const [headers, ...rows] = parsed.data;
    const kind = headers && headers[0] === 'フォーマット番号' ? 'shipment'
      : headers && headers[0] === '注文番号' ? 'legacy' : null;
    if (!kind) throw new Error('対応していないCSVです。未発送注文のCSVまたは宛名印刷用CSVを選択してください。');
    validateRows(headers, rows, kind);
    const shipmentCsvByOrder = new Map();
    const importedAt = new Date().toISOString();
    const data = rows.map(values => {
      const raw = Object.fromEntries(headers.map((header, i) => [header, values[i]]));
      if (kind === 'legacy') return raw;
      const row = {};
      for (const field of LEGACY_HEADERS) {
        const source = field === 'お支払方法' ? 'お支払い方法' : field;
        if (Object.hasOwn(raw, source)) row[field] = raw[source];
      }
      row['発送サイズ'] = raw['発送サイズ'];
      shipmentCsvByOrder.set(row['注文番号'].trim(), {
        formatVersion: raw['フォーマット番号'], headers: [...headers], values: [...values],
        sourceName: options.sourceName || '', importedAt
      });
      return row;
    });
    return { kind, headers, rows, data, shipmentCsvByOrder };
  }

  function anonymousAddressLabel(row) {
    const size = String(row && row['発送サイズ'] || '');
    const service = ['あんしんBOOTHパック', 'ゆうゆうBOOTHパック'].find(name => size.includes(name));
    return service ? `【${service}】` : '';
  }

  return { parseOrderCsv, validateRows, anonymousAddressLabel, SHIPMENT_HEADERS };
});
