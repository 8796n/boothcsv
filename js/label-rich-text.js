// label-rich-text.js
// カスタムラベルの書式モデル。
// エディタの内容を「ラン」（同じ書式が続く文字列）の平らな配列として扱い、
// 書式操作のたびに決まった形の HTML へ組み直す。DOM を直接切り貼りしないので、
// 入れ子の不整合や、無関係な箇所の書式が巻き込まれる問題が起きない。
//   ラン = { text, bold, italic, underline, fontFamily, fontSize }（text 内の '\n' は <br>）
(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) {
    module.exports = api;
  }
  root.BoothCSVLabelRichText = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  'use strict';

  const FORMAT_KEYS = ['bold', 'italic', 'underline', 'fontFamily', 'fontSize'];
  const TAG_FORMATS = { STRONG: 'bold', B: 'bold', EM: 'italic', I: 'italic', U: 'underline' };
  const BLOCK_TAGS = new Set(['DIV', 'P']);

  function pickFormat(run) {
    const format = {};
    FORMAT_KEYS.forEach(function(key) {
      if (run[key]) format[key] = run[key];
    });
    return format;
  }

  function sameFormat(a, b) {
    return FORMAT_KEYS.every(function(key) { return (a[key] || '') === (b[key] || ''); });
  }

  // 空のランを捨て、書式が完全に同じ隣接ランだけを結合する
  function mergeRuns(runs) {
    const merged = [];
    runs.forEach(function(run) {
      if (!run.text) return;
      const last = merged[merged.length - 1];
      if (last && sameFormat(last, run)) last.text += run.text;
      else merged.push(Object.assign(pickFormat(run), { text: run.text }));
    });
    return merged;
  }

  // [start, end) の文字の書式を update(書式) の戻り値に置き換える
  function updateRange(runs, start, end, update) {
    const result = [];
    let pos = 0;
    runs.forEach(function(run) {
      const runEnd = pos + run.text.length;
      const clamp = function(value) { return Math.min(Math.max(value, pos), runEnd); };
      const cuts = [pos, clamp(start), clamp(end), runEnd];
      for (let i = 0; i < 3; i++) {
        const text = run.text.slice(cuts[i] - pos, cuts[i + 1] - pos);
        if (!text) continue;
        const format = i === 1 ? pickFormat(update(pickFormat(run))) : pickFormat(run);
        result.push(Object.assign(format, { text }));
      }
      pos = runEnd;
    });
    return mergeRuns(result);
  }

  // [start, end) の見える文字（空白・改行・ZWSP 以外）がすべて key を持つか
  function rangeHasFormat(runs, start, end, key) {
    let pos = 0;
    let found = false;
    for (const run of runs) {
      const text = run.text.slice(Math.max(start - pos, 0), Math.max(end - pos, 0));
      if (/[^\s​]/.test(text)) {
        if (!run[key]) return false;
        found = true;
      }
      pos += run.text.length;
    }
    return found;
  }

  // 範囲全体が key を持っていれば外し、そうでなければ全体に付ける
  function toggleRange(runs, start, end, key) {
    const value = !rangeHasFormat(runs, start, end, key);
    return updateRange(runs, start, end, function(format) {
      format[key] = value;
      return format;
    });
  }

  function readFormat(el, parent) {
    const format = Object.assign({}, parent);
    if (TAG_FORMATS[el.tagName]) format[TAG_FORMATS[el.tagName]] = true;
    const style = el.style;
    if (!style) return format;
    // 太字化時に固定されていた line-height などの旧データは読み捨てる
    if (style.fontFamily) format.fontFamily = style.fontFamily;
    if (style.fontSize) format.fontSize = style.fontSize;
    if (style.fontWeight) format.bold = style.fontWeight === 'bold' || style.fontWeight === 'bolder' || Number(style.fontWeight) >= 600;
    if (style.fontStyle) format.italic = style.fontStyle !== 'normal';
    if ((style.textDecorationLine || style.textDecoration || '').includes('underline')) format.underline = true;
    return format;
  }

  // root の子孫をランに変換する。positions には各ノードの文字オフセット範囲を記録する
  function parseNodes(root) {
    const runs = [];
    const positions = new Map();
    let length = 0;
    let lastChar = '';
    let pendingBreak = false;
    const append = function(text, format) {
      runs.push(Object.assign({}, format, { text }));
      length += text.length;
      lastChar = text.slice(-1);
    };
    const breakLine = function(format) {
      if (length > 0 && lastChar !== '\n') append('\n', format);
      pendingBreak = false;
    };
    const walk = function(node, format) {
      if (node.nodeType === 3) {
        if (pendingBreak) breakLine(format);
        positions.set(node, { start: length, end: length + node.nodeValue.length });
        append(node.nodeValue, format);
        return;
      }
      if (node.nodeType !== 1) return;
      if (node.tagName === 'BR') {
        if (pendingBreak) breakLine(format);
        positions.set(node, { start: length, end: length + 1 });
        append('\n', format);
        return;
      }
      const isBlock = BLOCK_TAGS.has(node.tagName);
      if (isBlock) breakLine(format);
      const start = length;
      const childFormat = readFormat(node, format);
      node.childNodes.forEach(function(child) { walk(child, childFormat); });
      positions.set(node, { start, end: length });
      if (isBlock) pendingBreak = true;
    };
    root.childNodes.forEach(function(child) { walk(child, {}); });
    positions.set(root, { start: 0, end: length });
    return { runs: mergeRuns(runs), positions };
  }

  // DOM 上の位置 (container, offset) を文字オフセットに変換する
  function offsetOf(positions, container, offset) {
    const own = positions.get(container);
    if (container.nodeType === 3) return own ? own.start + Math.min(offset, own.end - own.start) : 0;
    for (let i = offset; i < container.childNodes.length; i++) {
      const child = positions.get(container.childNodes[i]);
      if (child) return child.start;
    }
    return own ? own.end : 0;
  }

  // ランを <span style>（フォント・サイズ）> <strong> > <em> > <u> の固定形で描画する。
  // フォント・サイズが同じランが続く間は 1 つの span を共有する
  function renderRuns(doc, runs) {
    const fragment = doc.createDocumentFragment();
    const segments = [];
    let pos = 0;
    let span = null;
    let spanKey = '';
    mergeRuns(runs).forEach(function(run) {
      let node = doc.createDocumentFragment();
      run.text.split('\n').forEach(function(part, i) {
        if (i > 0) {
          const br = doc.createElement('br');
          node.appendChild(br);
          segments.push({ node: br, start: pos, length: 1 });
          pos += 1;
        }
        if (part) {
          const text = doc.createTextNode(part);
          node.appendChild(text);
          segments.push({ node: text, start: pos, length: part.length });
          pos += part.length;
        }
      });
      [['underline', 'u'], ['italic', 'em'], ['bold', 'strong']].forEach(function(pair) {
        if (!run[pair[0]]) return;
        const el = doc.createElement(pair[1]);
        el.appendChild(node);
        node = el;
      });
      const key = run.fontFamily || run.fontSize ? (run.fontFamily || '') + '|' + (run.fontSize || '') : '';
      if (!key) {
        span = null;
        spanKey = '';
        fragment.appendChild(node);
        return;
      }
      if (key !== spanKey) {
        span = doc.createElement('span');
        if (run.fontFamily) span.style.fontFamily = run.fontFamily;
        if (run.fontSize) span.style.fontSize = run.fontSize;
        fragment.appendChild(span);
        spanKey = key;
      }
      span.appendChild(node);
    });
    return { fragment, segments };
  }

  // 文字オフセットを描画後の DOM 上の位置に戻す
  function pointAt(segments, root, offset) {
    for (const segment of segments) {
      if (offset > segment.start + segment.length) continue;
      if (segment.node.nodeType === 3) return [segment.node, offset - segment.start];
      const parent = segment.node.parentNode;
      const index = Array.prototype.indexOf.call(parent.childNodes, segment.node);
      return [parent, offset <= segment.start ? index : index + 1];
    }
    return [root, root.childNodes.length];
  }

  // range の範囲に update(runs, start, end) を適用してエディタを組み直し、選択を復元する
  function editRange(editor, range, update) {
    const parsed = parseNodes(editor);
    const start = offsetOf(parsed.positions, range.startContainer, range.startOffset);
    const end = offsetOf(parsed.positions, range.endContainer, range.endOffset);
    if (end <= start) return false;
    const doc = editor.ownerDocument;
    const rendered = renderRuns(doc, update(parsed.runs, start, end));
    editor.replaceChildren(rendered.fragment);
    const next = doc.createRange();
    next.setStart.apply(next, pointAt(rendered.segments, editor, start));
    next.setEnd.apply(next, pointAt(rendered.segments, editor, end));
    const selection = doc.defaultView.getSelection();
    selection.removeAllRanges();
    selection.addRange(next);
    return true;
  }

  return {
    mergeRuns,
    updateRange,
    rangeHasFormat,
    toggleRange,
    parseNodes,
    renderRuns,
    editRange
  };
});
