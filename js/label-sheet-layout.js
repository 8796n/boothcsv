(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) {
    module.exports = api;
  }
  root.BoothCSVLabelLayout = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  'use strict';

  function normalizeSheetSize(totalLabelsPerSheet) {
    const parsed = parseInt(totalLabelsPerSheet, 10);
    return parsed > 0 ? parsed : 44;
  }

  function normalizeSkipCount(skipCount, totalLabelsPerSheet) {
    const parsed = parseInt(skipCount, 10);
    if (!Number.isFinite(parsed) || parsed <= 0) return 0;
    return Math.min(parsed, totalLabelsPerSheet - 1);
  }

  function hasMeaningfulLabel(label) {
    if (!label) return false;
    if (typeof label === 'string') return label.trim() !== '';
    return true;
  }

  function padLabelsToFullSheets(labelarr, totalLabelsPerSheet) {
    const padded = labelarr.slice();
    const remainder = padded.length % totalLabelsPerSheet;
    if (remainder === 0) return padded;
    const toFill = totalLabelsPerSheet - remainder;
    for (let i = 0; i < toFill; i++) {
      padded.push('');
    }
    return padded;
  }

  function calculateMultiSheetDistribution(totalLabels, skipCount, totalLabelsPerSheet) {
    const sheetSize = normalizeSheetSize(totalLabelsPerSheet);
    const normalizedTotalLabels = Math.max(0, parseInt(totalLabels, 10) || 0);
    let remainingLabels = normalizedTotalLabels;
    let currentSkip = normalizeSkipCount(skipCount, sheetSize);
    let sheetNumber = 1;
    const sheetsInfo = [];

    while (remainingLabels > 0) {
      const availableInSheet = Math.max(0, sheetSize - currentSkip);
      const labelsInThisSheet = Math.min(remainingLabels, availableInSheet);
      const remainingInSheet = availableInSheet - labelsInThisSheet;

      sheetsInfo.push({
        sheetNumber,
        skipCount: currentSkip,
        labelCount: labelsInThisSheet,
        remainingCount: remainingInSheet,
        totalInSheet: currentSkip + labelsInThisSheet
      });

      remainingLabels -= labelsInThisSheet;
      currentSkip = 0;
      sheetNumber++;
    }

    if (sheetsInfo.length > 0) {
      return sheetsInfo;
    }

    return [{
      sheetNumber: 1,
      skipCount: currentSkip,
      labelCount: 0,
      remainingCount: sheetSize - currentSkip,
      totalInSheet: currentSkip
    }];
  }

  function buildSheetLayouts(labelarr, options) {
    const opts = options || {};
    const totalLabelsPerSheet = normalizeSheetSize(opts.totalLabelsPerSheet);
    const skipOnFirstSheet = normalizeSkipCount(opts.skipOnFirstSheet, totalLabelsPerSheet);

    if (!Array.isArray(labelarr) || labelarr.length === 0) {
      return [];
    }

    if (!labelarr.some(hasMeaningfulLabel)) {
      return [];
    }

    const paddedLabels = padLabelsToFullSheets(labelarr, totalLabelsPerSheet);
    const sheets = [];

    paddedLabels.forEach(function(label, index) {
      const sheetIndex = Math.floor(index / totalLabelsPerSheet);
      const posInSheet = index % totalLabelsPerSheet;

      if (!sheets[sheetIndex]) {
        sheets[sheetIndex] = {
          sheetIndex,
          items: []
        };
      }

      sheets[sheetIndex].items.push({
        label,
        sheetIndex,
        posInSheet,
        isSkipFace: sheetIndex === 0 && posInSheet < skipOnFirstSheet,
        skipFaceNumber: sheetIndex === 0 && posInSheet < skipOnFirstSheet ? posInSheet + 1 : null
      });
    });

    return sheets;
  }

  return {
    calculateMultiSheetDistribution,
    buildSheetLayouts
  };
});
