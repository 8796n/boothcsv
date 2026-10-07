(function(root, factory) {
  const commonJS = typeof module === 'object' && module.exports;
  const api = factory(commonJS ? require('./papaparse.min.js') : root.Papa,
    commonJS ? require('./order-csv.js') : root.BoothCSVOrders);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.BoothCSVBulkShipment = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(Papa, Orders) {
  'use strict';

  function buildShipmentCsv(csvText, orderNumbers, messageTemplate) {
    const parsed = Orders.parseOrderCsv(csvText);
    if (parsed.kind !== 'shipment') {
      throw new Error('発送通知用CSVではありません。「未発送注文のCSV」を選択してください。');
    }
    return buildResult(parsed.headers, parsed.rows, orderNumbers, messageTemplate);
  }

  function buildShipmentCsvFromRecords(records, messageTemplate) {
    const pending = records.filter(record => !record.shippedAt);
    const rows = [];
    for (const record of pending) {
      const source = record.shipmentCsv;
      if (!source) continue;
      if (source.formatVersion !== '1') throw new Error('保存されたCSVのフォーマット番号に対応していません。');
      Orders.validateRows(source.headers, [source.values], 'shipment');
      if (source.values[1].trim() !== record.orderNumber) throw new Error('保存されたCSVの注文番号が一致しません。');
      rows.push(source.values);
    }
    return buildResult(Orders.SHIPMENT_HEADERS, rows, pending.map(record => record.orderNumber), messageTemplate);
  }

  function buildResult(header, rows, orderNumbers, messageTemplate) {
    const orderIndex = header.indexOf('注文番号');
    const commentIndex = header.indexOf('発送通知コメント');
    const comment = typeof messageTemplate === 'string' ? messageTemplate : '';
    const selected = new Set(orderNumbers.map(String).map(value => value.trim()).filter(Boolean));
    const matched = new Set();
    const filtered = rows.filter(row => selected.has(row[orderIndex].trim())).map(source => {
      const row = [...source];
      matched.add(row[orderIndex].trim());
      row[commentIndex] = comment;
      return row;
    });
    return {
      csv: '\uFEFF' + Papa.unparse([header, ...filtered], { quotes: true, newline: '\r\n' }),
      rowCount: filtered.length,
      shippedComment: comment,
      orderNumbers: Array.from(matched),
      missingOrderNumbers: Array.from(selected).filter(number => !matched.has(number))
    };
  }

  return { buildShipmentCsv, buildShipmentCsvFromRecords };
});
