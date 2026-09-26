const test = require('node:test');
const assert = require('node:assert/strict');

const {
  mergeRuns,
  updateRange,
  toggleRange
} = require('../../js/label-rich-text.js');

test('書式の適用は範囲内の文字だけに効き、同じ書式の隣接ランは結合される', function() {
  const runs = [{ text: 'ABCDE' }];
  const serif = updateRange(runs, 1, 3, function(format) { format.fontFamily = 'serif'; return format; });
  assert.deepEqual(serif, [
    { text: 'A' },
    { text: 'BC', fontFamily: 'serif' },
    { text: 'DE' }
  ]);
  const back = updateRange(serif, 1, 3, function(format) { format.fontFamily = ''; return format; });
  assert.deepEqual(back, [{ text: 'ABCDE' }]);
});

test('太字の一部を解除しても、その部分のフォントや斜体は残る', function() {
  const runs = [
    { text: 'A' },
    { text: 'B', bold: true },
    { text: 'C', bold: true, italic: true, fontFamily: 'serif' },
    { text: 'D', bold: true },
    { text: 'E' }
  ];
  assert.deepEqual(toggleRange(runs, 2, 3, 'bold'), [
    { text: 'A' },
    { text: 'B', bold: true },
    { text: 'C', italic: true, fontFamily: 'serif' },
    { text: 'D', bold: true },
    { text: 'E' }
  ]);
});

test('一部だけ太字の範囲をトグルすると全体が太字になり、空白と改行は判定に使わない', function() {
  const runs = [{ text: 'AB', bold: true }, { text: ' \n​' }, { text: 'C' }];
  assert.deepEqual(toggleRange(runs, 0, 6, 'bold'), [{ text: 'AB \n​C', bold: true }]);
  // 見える文字が AB（太字）だけの範囲は「全体が太字」とみなして解除する
  assert.deepEqual(toggleRange(runs, 0, 5, 'bold'), [{ text: 'AB \n​C' }]);
});

test('書式の違う隣接ランは結合せず、改行を含む同じ書式のランは結合する', function() {
  assert.deepEqual(mergeRuns([
    { text: '明朝', bold: true, fontFamily: 'serif' },
    { text: 'ゴシ', bold: true, fontFamily: 'sans-serif' },
    { text: '' },
    { text: 'A\n', italic: true },
    { text: 'B', italic: true }
  ]), [
    { text: '明朝', bold: true, fontFamily: 'serif' },
    { text: 'ゴシ', bold: true, fontFamily: 'sans-serif' },
    { text: 'A\nB', italic: true }
  ]);
});
