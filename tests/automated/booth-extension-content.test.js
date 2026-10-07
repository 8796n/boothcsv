const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../../js/booth-extension-content.js'), 'utf8');
const announcement = '「自宅から発送」の注文を一括で発送完了にできるようになりました';

async function request(type, initialText, loadedText) {
  let listener;
  let elapsed = 0;
  const mount = {
    innerText: initialText,
    childElementCount: 1,
    querySelectorAll: () => []
  };
  const document = {
    body: { get innerText() { return announcement + '\n' + (mount.innerText || ''); } },
    getElementById: () => mount.innerText === null ? null : mount,
    querySelectorAll: () => [],
    querySelector: () => null
  };
  vm.runInNewContext(source, {
    document,
    window: { location: { pathname: '/orders/12345678' } },
    chrome: { runtime: { onMessage: { addListener(callback) { listener = callback; } } } },
    Date: class extends Date { static now() { return elapsed; } },
    setTimeout(callback, delay) {
      elapsed += delay;
      if (loadedText !== undefined) mount.innerText = loadedText;
      queueMicrotask(callback);
    }
  });
  const response = await new Promise(resolve => listener({ type }, {}, resolve));
  return { response, elapsed };
}

test('お知らせの発送完了を無視して注文詳細の描画を待ち、未描画なら取得を中断する', async function() {
  const pendingText = '注文詳細 - 注文番号 : 12345678\n商品を発送してください\n未発送';
  // 空のマウント、ローディング要素、マウント自体が未配置のいずれも待機する。
  for (const initialText of ['', '読み込み中', null]) {
    const { response, elapsed } = await request('boothcsv:collect-order-shipment-status', initialText, pendingText);
    assert.equal(response.ok, true);
    assert.equal(response.shipped, false);
    assert.ok(elapsed > 0, 'お知らせやローディング要素だけで待機を終了しない');
  }

  const shipped = await request('boothcsv:collect-order-shipment-status', '注文番号 : 12345678\n発送完了');
  assert.equal(shipped.response.shipped, true, '注文自体の発送完了は検出する');

  const qr = await request('boothcsv:collect-order-qr', '', pendingText);
  assert.match(qr.response.error, /発送コードを発行するボタンが見つかりません/);
  assert.ok(qr.elapsed > 0, 'QR取得でも未発送の詳細が描画されてから判定する');

  for (const type of ['boothcsv:collect-order-qr', 'boothcsv:collect-order-shipment-status', 'boothcsv:submit-order-shipment-notification']) {
    const { response, elapsed } = await request(type, '読み込み中');
    assert.equal(response.ok, false);
    assert.match(response.error, /注文詳細.*タイムアウト/);
    assert.ok(elapsed >= 12000);
  }
});
