const test = require('node:test');
const assert = require('node:assert/strict');

const {
  buildSheetLayouts,
  calculateMultiSheetDistribution
} = require('../../js/label-sheet-layout.js');

test('最初のシートだけにスキップ面が付く', function() {
  const labels = Array.from({ length: 50 }, function(_, index) {
    return 'ORDER-' + String(index + 1);
  });

  const sheets = buildSheetLayouts(labels, {
    skipOnFirstSheet: 3,
    totalLabelsPerSheet: 44
  });

  assert.equal(sheets.length, 2);
  assert.deepEqual(
    sheets[0].items.filter(function(item) { return item.isSkipFace; }).map(function(item) { return item.skipFaceNumber; }),
    [1, 2, 3]
  );
  assert.deepEqual(
    sheets[1].items.filter(function(item) { return item.isSkipFace; }),
    []
  );
});

test('スキップ面数を含む複数シート分配は2ページ目でリセットされる', function() {
  const sheets = calculateMultiSheetDistribution(50, 3, 44);

  assert.deepEqual(sheets, [
    { sheetNumber: 1, skipCount: 3, labelCount: 41, remainingCount: 0, totalInSheet: 44 },
    { sheetNumber: 2, skipCount: 0, labelCount: 9, remainingCount: 35, totalInSheet: 9 }
  ]);
});

test('1ページ目がほぼ埋まるケースでも2ページ目にはスキップ面が付かない', function() {
  const sheets = calculateMultiSheetDistribution(2, 43, 44);

  assert.deepEqual(sheets, [
    { sheetNumber: 1, skipCount: 43, labelCount: 1, remainingCount: 0, totalInSheet: 44 },
    { sheetNumber: 2, skipCount: 0, labelCount: 1, remainingCount: 43, totalInSheet: 1 }
  ]);
});

test('意味のあるラベルがない場合はシートを生成しない', function() {
  const sheets = buildSheetLayouts(['', '   ', null], {
    skipOnFirstSheet: 5,
    totalLabelsPerSheet: 44
  });

  assert.deepEqual(sheets, []);
});
