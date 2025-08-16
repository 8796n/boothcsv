#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');
const argv = require('minimist')(process.argv.slice(2));
let Jimp; // 遅延読み込み（Jimp は cut 生成時のみ使用）

(async ()=>{
  const outDir = path.resolve(argv.out || 'docs/images');
  const html = path.resolve(argv.html || 'boothcsv.html');
  const chromePath = argv.chrome || 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe' || undefined;

  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  if (!fs.existsSync(html)) {
    console.error('HTML not found:', html);
    process.exit(1);
  }

  const launchOpts = { headless: true, args: ['--no-sandbox','--disable-setuid-sandbox'] };
  if (chromePath) launchOpts.executablePath = chromePath;

  let browser;
  try {
    browser = await puppeteer.launch(launchOpts);
  } catch (e) {
    console.error('Failed to launch browser:', e.message);
    console.error('If you don\'t have Chrome installed, install Chrome or pass --chrome <path>');
    process.exit(1);
  }

  const page = await browser.newPage();

  // ブラウザのコンソール出力をNode.jsのコンソールに転送
  page.on('console', msg => {
    for (let i = 0; i < msg.args().length; ++i) {
      console.log(`[Browser] ${i}: ${msg.args()[i]}`);
    }
  });

  await page.goto('http://localhost:8080/' + path.basename(html) + '?debug=1', { waitUntil: 'networkidle2' });
  await page.setViewport({ width: 1300, height: 700 });

  // UIの状態を設定する共通関数
  async function setupUI(config) {
    await page.evaluate((cfg) => {
      // 差分適用ユーティリティ
      const log = (...args) => { if (window.DEBUG_SETUP_UI) console.debug('[setupUI]', ...args); };

      // サイドバー：ピン留めのON/OFFで制御
      if (cfg.sidebar !== undefined) {
        const sidebarPin = document.getElementById('sidebarPin');
        const sidebarToggle = document.getElementById('sidebarToggle');
        const sidebar = document.getElementById('appSidebar');

        if (cfg.sidebar === true) { // true: ピン留めして表示
          if (sidebar && sidebarPin && !document.body.classList.contains('sidebar-docked')) {
            // サイドバーが開いていなければ開く
            if (!sidebar.classList.contains('open') && sidebarToggle) {
              sidebarToggle.click();
              log('sidebarToggle clicked to open sidebar');
            }
            // ピン留めする
            sidebarPin.click();
            log('sidebarPin clicked to dock sidebar');
          }
        } else { // false: ピン留めを解除して非表示
          if (sidebarPin && document.body.classList.contains('sidebar-docked')) {
            sidebarPin.click();
            log('sidebarPin clicked to undock sidebar');
          }
        }
      }

      // チェックボックス（変更時のみ）
      const checkboxes = {
        labelyn: cfg.labelyn,
        customLabelEnable: cfg.customLabel,
        orderImageEnable: cfg.orderImage
      };
      for (const [id, value] of Object.entries(checkboxes)) {
        if (value !== undefined) {
          const elem = document.getElementById(id);
          if (elem && elem.checked !== value) {
            elem.checked = value;
            elem.dispatchEvent(new Event('change'));
            log('checkbox changed', id, '->', value);
          }
        }
      }

      // labelskipnum （数値文字列比較）
      const skipVal = (
        cfg.labelSkip !== undefined ? cfg.labelSkip :
        (cfg.labelskip !== undefined ? cfg.labelskip : cfg.labelskipnum)
      );
      if (skipVal !== undefined) {
        const skipInput = document.getElementById('labelskipnum');
        if (skipInput) {
          const newValStr = String(skipVal);
          if (skipInput.value !== newValStr) {
            skipInput.value = newValStr;
            skipInput.dispatchEvent(new Event('input', { bubbles: true }));
            skipInput.dispatchEvent(new Event('change', { bubbles: true }));
            log('labelskipnum changed ->', newValStr);
          }
        }
      }

      // CSVファイルのクリア（差分不要。常に空にする指示のときのみ）
      if (cfg.clearFile) {
        const fileInput = document.getElementById('file');
        if (fileInput && fileInput.value !== '') {
          fileInput.value = '';
          log('file input cleared');
        }
      }
    }, config);

    // labelskipnum 適用後のラベル再生成完了を簡易的に待つ（既に CSV が読み込まれているケース）
    if (config && (config.labelSkip !== undefined || config.labelskip !== undefined || config.labelskipnum !== undefined)) {
      try {
        // 直後は autoProcessCSV が走るのでラベルシート数の DOM 変化を短時間監視
        await page.waitForFunction(() => {
          // settingsCache の反映と最低1フレーム経過を期待
          return !!window.settingsCache && typeof window.settingsCache.labelskip === 'number';
        }, { timeout: 500 });
      } catch(_) {}
    }
  }

  // スクリーンショット撮影関数
  async function shot(name, snapName, config) {
    console.log('👉 Capturing:', name);
    
    // UI状態を設定
    if (config.ui) {
      await setupUI(config.ui);
      await new Promise(r => setTimeout(r, 500)); // UI更新待ち
    }

    // 追加のセットアップ処理
    if (config.setup) {
      try {
        await page.evaluate(config.setup);
        await new Promise(r => setTimeout(r, config.setupDelay || 500));
      } catch(e) {
        console.warn('Setup failed:', e.message);
      }
    }

    // ビューポートサイズの調整
    if (config.viewport) {
      await page.setViewport({
        width: config.viewport.width || 1570,
        height: config.viewport.height || 1200
      });
    }

    // snap-anchorベースの撮影
    const rect = await page.evaluate(async (snapName, options = {}) => {
      let target;
      const targetName = options.targetSelector || snapName;

      // ターゲット要素が見つかるまで待機
      for (let i = 0; i < 30; i++) {
        if (options.targetSelector) {
          target = document.querySelector(options.targetSelector);
        } else if (snapName) {
          // 複数のアンカーがある場合は、ビューポート内で最初に見つかるものを使用
          const anchors = document.querySelectorAll(`.snap-anchor[data-snap="${snapName}"]`);
          let visibleAnchor = null;
          for (const anchor of anchors) {
            const rect = anchor.getBoundingClientRect();
            // ビューポート内にあるかチェック
            if (rect.top >= 0 && rect.top < window.innerHeight) {
              visibleAnchor = anchor;
              break;
            }
          }
          // ビューポート内のアンカーが見つからない場合は最初のアンカーを使用
          const anchor = visibleAnchor || anchors[0];
          if (anchor) {
            target = anchor.closest('section.sheet') || anchor.parentElement;
          }
        }
        if (target) break;
        await new Promise(r => setTimeout(r, 100));
      }
      
      if (!target) {
        console.warn(`Target not found for: ${targetName}`);
        return null;
      }
      
      // templateタグをスキップ
      while (target && target.tagName === 'TEMPLATE') {
        target = target.parentElement;
      }
      
      if (!target) {
        console.warn(`Target not found for: ${targetName} after template skip`);
        return null;
      }

      // ビューポートに収める
      target.scrollIntoView({ behavior: 'instant', block: 'center' });
      await new Promise(r => setTimeout(r, 200)); // スクロール待ち
      
      const rect = target.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) {
        console.warn(`Invalid dimensions for: ${targetName}`);
        return null;
      }

      // 拡張セレクタ（コンテキストメニューなど）を含める
      let x = rect.left, y = rect.top, w = rect.width, h = rect.height;
      
      if (options.expandSelectors) {
        options.expandSelectors.forEach(sel => {
          const el = document.querySelector(sel);
          if (!el) return;
          const r = el.getBoundingClientRect();
          if (r.width === 0 || r.height === 0) return;
          x = Math.min(x, r.left);
          y = Math.min(y, r.top);
          const x2 = Math.max(x + w, r.right);
          const y2 = Math.max(y + h, r.bottom);
          w = x2 - x;
          h = y2 - y;
        });
      }

      // パディング
      const padding = options.padding || 10;
      x -= padding;
      y -= padding;
      w += padding * 2;
      h += padding * 2;

      return {
        x: Math.max(0, Math.round(x)),
        y: Math.max(0, Math.round(y)),
        width: Math.round(w),
        height: Math.round(h)
      };
    }, snapName, config.snapOptions);

    // 最終出力のみ: rect があればフルスクリーンをバッファ取得→crop→ name で保存。無ければそのまま name で保存。
    const outPath = path.join(outDir, name);
    const ext = path.extname(name).toLowerCase();

    if (rect) {
      const vp = page.viewport();
      rect.width = Math.min(rect.width, vp.width - rect.x);
      rect.height = Math.min(rect.height, vp.height - rect.y);
      try {
        // フルビューポートをバッファ取得
        const fullBuffer = await page.screenshot({ type: ext === '.jpg' || ext === '.jpeg' ? 'jpeg' : undefined });
        if (!Jimp) Jimp = require('jimp');
        const img = await Jimp.read(fullBuffer);
        const safeW = Math.min(rect.width, img.bitmap.width - rect.x);
        const safeH = Math.min(rect.height, img.bitmap.height - rect.y);
        if (safeW > 0 && safeH > 0) {
          img.crop(rect.x, rect.y, safeW, safeH);
          await img.writeAsync(outPath);
          console.log(`✔ Saved cropped (${safeW}x${safeH}) ->`, outPath);
        } else {
          // 幅高さ不正ならそのまま保存
            await fs.promises.writeFile(outPath, fullBuffer);
            console.warn('Crop skipped (invalid dims). Saved full viewport instead.');
        }
      } catch (e) {
        console.error('Final screenshot (crop) failed:', e.message);
      }
    } else {
      try {
        await page.screenshot({ path: outPath, type: ext === '.jpg' || ext === '.jpeg' ? 'jpeg' : undefined });
        console.log('✔ Saved (no rect fallback):', outPath);
      } catch (e) {
        console.error('Final screenshot (no rect) failed:', e.message);
      }
    }
  }

  // スクリーンショット定義
  const shots = [
    {
      name: 'main-interface.png',
      ui: {
        sidebar: true,
        labelyn: true,
        customLabel: false,
        orderImage: false,
        clearFile: true
      }
    },
    {
      name: 'custom-labels.png',
      snap: 'custom-labels',
      ui: {
        sidebar: true,
        labelyn: true,
        customLabel: true,
        orderImage: false,
        labelSkip: 1  // ラベルスキップ数を0に設定
      },
      setup: () => {
        // リッチテキストエディタにサンプルテキストを挿入
        const editor = document.querySelector('.rich-text-editor');
        if (editor) {
          editor.innerHTML = '<div><b>【サンプル商品】</b></div><div>アクリルキーホルダー</div><div style="color: #666;">商品コード: AKH-001</div>';
          editor.dispatchEvent(new Event('input', {bubbles: true}));
          
          // コンテキストメニューを表示
          setTimeout(() => {
            const bold = editor.querySelector('b');
            if (bold) {
              const range = document.createRange();
              range.selectNodeContents(bold);
              const sel = window.getSelection();
              sel.removeAllRanges();
              sel.addRange(range);
              
              const rect = bold.getBoundingClientRect();
              const evt = new MouseEvent('contextmenu', {
                bubbles: true,
                cancelable: true,
                view: window,
                button: 2,
                clientX: rect.left + 10,
                clientY: rect.top + 10
              });
              bold.dispatchEvent(evt);
            }
          }, 300);
        }
      },
      setupDelay: 1000,
      snapOptions: {
        expandSelectors: ['.custom-label-context-menu'],
        padding: 6
      }
    },
    {
      name: 'custom-labels-sheet.png',
      snap: 'label-sheet',
      ui: {
        sidebar: false,  // サイドバーを閉じる
        labelyn: true,
        customLabel: true,
        orderImage: false
      },
      setup: () => {
        // コンテキストメニューを閉じる
        document.querySelectorAll('.custom-label-context-menu').forEach(n => n.remove());
        const editor = document.querySelector('.rich-text-editor');
        if (editor) editor.blur();
        window.getSelection().removeAllRanges();
      },
      viewport: { height: 1400 }
    },
    {
      name: 'custom-fonts.png',
      snap: 'custom-fonts',
      ui: {
        sidebar: true,
        labelyn: true,
        customLabel: true,
        orderImage: false
      },
      setup: () => {
        // フォントセクションを開く
        const content = document.getElementById('fontSectionContent');
        if (content) {
          content.style.maxHeight = 'none';
          content.style.overflow = 'visible';
          content.hidden = false;
        }
      }
    },
    {
      name: 'image-function.png',
      snap: 'image',
      ui: {
        sidebar: true,
        labelyn: true,
        customLabel: false,
        orderImage: true
      }
    },
    {
      name: 'image-dropzone.png',
      snap: 'image',
      ui: {
        sidebar: true,
        labelyn: true,
        customLabel: false,
        orderImage: true
      },
      setup: async () => {
        // サンプル画像をロード
        try {
          const res = await fetch('/sample/footersample.png');
          const blob = await res.blob();
          const url = URL.createObjectURL(blob);
          if (window.orderImageDropZone && window.orderImageDropZone.setImage) {
            window.orderImageDropZone.setImage(url, await blob.arrayBuffer());
          }
        } catch(e) {
          console.warn('Image load failed:', e);
        }
      }
    },
    {
      name: 'csv-labels.png',
      snap: 'label-sheet',
      ui: {
        sidebar: false,
        labelyn: true,
        customLabel: false,
        orderImage: true
      },
      setup: async () => {
        // 全てのシートをクリア
        document.querySelectorAll('section.sheet').forEach(s => s.remove());
        
        // CSVファイルをロード
        try {
          const resCsv = await fetch('/sample/booth_orders_sample.csv');
          const blobCsv = await resCsv.blob();
          const fileCsv = new File([blobCsv], 'booth_orders_sample.csv');
          const input = document.getElementById('file');
          const dt = new DataTransfer();
          dt.items.add(fileCsv);
          Object.defineProperty(input, 'files', {value: dt.files});
          input.dispatchEvent(new Event('change'));
          console.log('CSV file loaded for csv-labels');
        } catch(e) {
          console.warn('CSV load failed:', e);
        }

        // QRコード画像をロードして最初のドロップゾーンに設定
        try {
          await new Promise(r => setTimeout(r, 1500)); // CSV処理を少し待つ
          const dropzone = document.querySelector('.dropzone');
          if (!dropzone) {
            console.warn('QR dropzone not found');
            return;
          }
          const resQr = await fetch('/sample/qrcodedsample.png');
          const blobQr = await resQr.blob();
          const urlQr = URL.createObjectURL(blobQr);
          const img = document.createElement('img');
          img.src = urlQr;
          img.style.maxWidth = '100%';
          img.style.height = 'auto';
          await new Promise(resolve => {
            img.onload = () => {
              dropzone.innerHTML = '';
              dropzone.appendChild(img);
              if (typeof window.readQR === 'function') {
                window.readQR(img); // QRコードを読み取らせる
              }
              console.log('QR code image loaded into first dropzone');
              resolve();
            };
            img.onerror = () => {
              console.warn('QR image could not be loaded into img tag');
              resolve();
            }
          });
        } catch(e) {
          console.warn('QR Code image load failed:', e);
        }
      },
      setupDelay: 2500
    },
    {
      name: 'csv-orders.png',
      snap: 'order-page',
      ui: {
        sidebar: false,
        labelyn: true,
        customLabel: false,
        orderImage: true
      }
    },
    {
      name: 'usage-guide.png',
      ui: {
        sidebar: true,
        labelyn: true,
        customLabel: true,
        orderImage: true
      }
    }
  ];

  // 順番に実行
  for (const shotConfig of shots) {
    // targetSelectorが指定されている場合は、snapNameをundefinedとして渡す
    const snapName = shotConfig.snapOptions?.targetSelector ? undefined : shotConfig.snap;
    await shot(shotConfig.name, snapName, shotConfig);
    await new Promise(r => setTimeout(r, 1000));
  }

  await browser.close();
  console.log('All done. Images saved to', outDir);
})();