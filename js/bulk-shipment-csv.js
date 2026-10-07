(function(root, factory) {
  const api = factory(typeof module === 'object' && module.exports ? require('./papaparse.min.js') : root.Papa);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.BoothCSVBulkShipment = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(Papa) {
  'use strict';

  function buildShipmentCsv(csvText, orderNumbers, messageTemplate) {
    // 配列として扱い、BOOTHの列順・文字列・商品欄の改行を保つ。
    const parsed = Papa.parse(csvText.replace(/^\uFEFF/, ''), { delimiter: ',', skipEmptyLines: 'greedy' });
    if (parsed.errors.length) throw new Error('CSVを読み取れませんでした。BOOTHからダウンロードしたCSVを選択してください。');
    const [header, ...rows] = parsed.data;
    const required = ['フォーマット番号', '注文番号', '配送番号', '発送用伝票番号', '発送通知コメント'];
    if (!header || required.some(name => !header.includes(name)) || new Set(header).size !== header.length) {
      throw new Error('発送通知用CSVではありません。「未発送注文のCSV」を選択してください。');
    }
    if (rows.some(row => row.length !== header.length)) throw new Error('CSVの列数が一致しません。BOOTHから再ダウンロードしてください。');

    const orderIndex = header.indexOf('注文番号');
    const commentIndex = header.indexOf('発送通知コメント');
    const selected = new Set(orderNumbers.map(String).map(value => value.trim()).filter(Boolean));
    const matched = new Set();
    const filtered = rows.filter(row => selected.has(row[orderIndex].trim())).map(row => {
      matched.add(row[orderIndex].trim());
      row[commentIndex] = messageTemplate;
      return row;
    });
    return {
      csv: '\uFEFF' + Papa.unparse([header, ...filtered], { quotes: true, newline: '\r\n' }),
      rowCount: filtered.length,
      orderNumbers: Array.from(matched),
      missingOrderNumbers: Array.from(selected).filter(number => !matched.has(number))
    };
  }

  return { buildShipmentCsv };
});
