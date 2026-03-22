(function() {
  'use strict';

  const TRIGGER_PATTERNS = [
    /発送コードを発行する/,
    /発送コード/,
    /QRコード/,
    /再発行/
  ];

  const DEFAULT_ISSUE_SETTINGS = {
    packageSizeId: '16',
    codeType: '0',
    description: '',
    includeOrderNumber: true,
    handlingCodes: []
  };

  const QR_IMAGE_PATTERNS = [
    /qr/i,
    /yamato_invoice/i,
    /qr_code/i
  ];

  function isElementVisible(element) {
    if (!element) return false;
    const rect = element.getBoundingClientRect();
    const style = window.getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
  }

  function findTriggerButton() {
    const candidates = Array.from(document.querySelectorAll('button, a, [role="button"]'));
    return candidates.find(function(element) {
      const text = (element.innerText || element.textContent || '').trim();
      return isElementVisible(element) && TRIGGER_PATTERNS.some(function(pattern) {
        return pattern.test(text);
      });
    }) || null;
  }

  function findReissueToggle() {
    const candidates = Array.from(document.querySelectorAll('button, a, summary, [role="button"]'));
    return candidates.find(function(element) {
      const text = (element.innerText || element.textContent || '').trim();
      if (!text) return false;
      return /サイズを変更.*再発行|発送コードを再発行|再発行/.test(text);
    }) || null;
  }

  function findShipmentNotifyTrigger() {
    const candidates = Array.from(document.querySelectorAll('button, a, [role="button"], input[type="submit"], input[type="button"]'));
    return candidates.find(function(element) {
      const text = normalizeInlineText(element.innerText || element.textContent || element.value || '');
      if (!text) return false;
      return isElementVisible(element) && /発送完了を通知|発送完了する|完了を通知/.test(text);
    }) || null;
  }

  function getOrderDetailText() {
    const mount = getOrderDetailMount();
    const mountText = normalizeInlineText(mount ? (mount.innerText || mount.textContent || '') : '');
    if (mountText) return mountText;
    return normalizeInlineText(getOrderPageText());
  }

  function getShippingStateSignals() {
    const detailText = getOrderDetailText();
    const statusHead = detailText.slice(0, 400);
    const hasPendingStatus = /商品を発送してください|未発送/.test(statusHead);
    const hasShippedStatus = /発送完了|発送済み/.test(statusHead);
    const hasTrackingNumber = /商品追跡用\s*伝票番号/.test(detailText);
    const hasShippedAt = /発送日時/.test(detailText);

    return {
      detailText,
      statusHead,
      hasPendingStatus,
      hasShippedStatus,
      hasTrackingNumber,
      hasShippedAt
    };
  }

  function getIssueForm() {
    let form = document.querySelector('form[action*="/register_yamato_invoice"]');
    if (form) return form;

    const toggle = findReissueToggle();
    if (toggle) {
      toggle.click();
      form = document.querySelector('form[action*="/register_yamato_invoice"]');
    }

    return form;
  }

  function isShippedOrderPage() {
    const signals = getShippingStateSignals();
    if (!signals.detailText) return false;
    if (signals.hasShippedAt || signals.hasTrackingNumber) return true;
    if (signals.hasPendingStatus) return false;
    return signals.hasShippedStatus;
  }

  function getOrderPageText() {
    return document.body ? (document.body.innerText || document.body.textContent || '') : '';
  }

  function getOrderDetailMount() {
    return document.getElementById('js-mount-point-manage-order-show');
  }

  function hasOrderDetailContent() {
    const mount = getOrderDetailMount();
    if (!mount) return false;
    if (mount.childElementCount > 0) return true;

    const mountText = normalizeInlineText(mount.innerText || mount.textContent || '');
    return /注文番号|発送完了|発送済み|発送日時|商品追跡用\s*伝票番号|発送コード/.test(mountText);
  }

  async function waitForOrderDetailContent(timeoutMs) {
    const deadline = Date.now() + (timeoutMs || 12000);
    while (Date.now() < deadline) {
      if (hasOrderDetailContent()) return true;

      const pageText = getOrderPageText();
      if (/注文番号|発送完了|発送済み|発送日時|商品追跡用\s*伝票番号|発送コード/.test(pageText)) {
        return true;
      }

      await new Promise(function(resolve) {
        setTimeout(resolve, 250);
      });
    }

    return hasOrderDetailContent();
  }

  function normalizeInlineText(value) {
    return typeof value === 'string' ? value.replace(/\s+/g, ' ').trim() : '';
  }

  function escapeRegExp(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function normalizeShippedAt(rawValue) {
    const raw = normalizeInlineText(rawValue);
    if (!raw) return { raw: '', iso: '' };

    const structuredMatch = raw.match(/(\d{4})\s*[\/-年]\s*(\d{1,2})\s*[\/-月]\s*(\d{1,2})(?:日)?\s*(午前|午後|AM|PM)?\s*(\d{1,2})(?:\s*[:時]\s*|時)(\d{1,2})(?:\s*[:分]\s*|分)?(\d{1,2})?(?:秒)?/i);
    if (structuredMatch) {
      const year = Number(structuredMatch[1]);
      const month = Number(structuredMatch[2]);
      const day = Number(structuredMatch[3]);
      const meridiem = (structuredMatch[4] || '').toUpperCase();
      let hours = Number(structuredMatch[5]);
      const minutes = Number(structuredMatch[6]);
      const seconds = Number(structuredMatch[7] || 0);

      if (meridiem === '午後' || meridiem === 'PM') {
        if (hours < 12) hours += 12;
      } else if (meridiem === '午前' || meridiem === 'AM') {
        if (hours === 12) hours = 0;
      }

      const date = new Date(year, month - 1, day, hours, minutes, seconds);
      if (!Number.isNaN(date.getTime())) {
        return {
          raw,
          iso: date.toISOString()
        };
      }
    }

    const candidate = raw
      .replace(/年/g, '/')
      .replace(/月/g, '/')
      .replace(/日/g, '')
      .replace(/時/g, ':')
      .replace(/分/g, ':')
      .replace(/秒/g, '')
      .replace(/午前/g, ' AM ')
      .replace(/午後/g, ' PM ')
      .replace(/\s+/g, ' ')
      .trim();

    const parsed = Date.parse(candidate.replace(/-/g, '/'));
    return {
      raw,
      iso: Number.isNaN(parsed) ? '' : new Date(parsed).toISOString()
    };
  }

  function extractDateTimeValueFromText(text, labelPattern) {
    const normalizedText = normalizeInlineText(text);
    if (!normalizedText) return '';

    const labelSource = escapeRegExp(labelPattern.source || String(labelPattern));
    const dateTimePattern = new RegExp(
      `${labelSource}\\s*(?:[:：]\\s*)?((?:\\d{4}\\s*[\\/-年]\\s*\\d{1,2}\\s*[\\/-月]\\s*\\d{1,2}(?:日)?)\\s*(?:午前|午後|AM|PM)?\\s*\\d{1,2}(?:\\s*[:時]\\s*|時)\\d{1,2}(?:(?:\\s*[:分]\\s*|分)\\d{1,2})?(?:秒)?)`,
      labelPattern.flags || ''
    );
    const match = normalizedText.match(dateTimePattern);
    return match && match[1] ? normalizeInlineText(match[1]) : '';
  }

  function extractValueFromStructuredRow(labelPattern) {
    const containers = Array.from(document.querySelectorAll('div, li, dt, dd, tr'));
    for (const container of containers) {
      const children = Array.from(container.children || []);
      if (children.length < 2) continue;

      for (let index = 0; index < children.length - 1; index++) {
        const labelText = normalizeInlineText(children[index].innerText || children[index].textContent || '');
        if (!labelText || !labelPattern.test(labelText)) continue;

        const exactLabel = labelText === labelPattern.source || new RegExp(`^${labelPattern.source}$`, labelPattern.flags || '').test(labelText);
        const looksLikeLabelValuePair = exactLabel || /^発送日時\b/.test(labelText);
        if (!looksLikeLabelValuePair) continue;

        const valueText = normalizeInlineText(children[index + 1].innerText || children[index + 1].textContent || '');
        if (!valueText) continue;

        const directDateTime = extractDateTimeValueFromText(`発送日時 ${valueText}`, labelPattern);
        if (directDateTime) return directDateTime;
        return valueText;
      }
    }
    return '';
  }

  function extractFieldValueByLabel(labelPattern) {
    const structuredRowValue = extractValueFromStructuredRow(labelPattern);
    if (structuredRowValue) return structuredRowValue;

    const selectors = ['th', 'dt', 'label', '.definition-label', '.item-label', 'span', 'div', 'p'];
    for (const selector of selectors) {
      const elements = Array.from(document.querySelectorAll(selector));
      for (const element of elements) {
        const text = normalizeInlineText(element.innerText || element.textContent || '');
        if (!text || !labelPattern.test(text)) continue;

        const directDateTime = extractDateTimeValueFromText(text, labelPattern);
        if (directDateTime) return directDateTime;

        const sameLineMatch = text.match(/[:：]\s*(.+)$/);
        if (sameLineMatch && sameLineMatch[1]) {
          return sameLineMatch[1].trim();
        }

        const next = element.nextElementSibling;
        if (next) {
          const nextText = normalizeInlineText(next.innerText || next.textContent || '');
          if (nextText) return nextText;
        }

        const parent = element.parentElement;
        if (parent) {
          const children = Array.from(parent.children);
          const index = children.indexOf(element);
          if (index >= 0 && index < children.length - 1) {
            const siblingText = normalizeInlineText(children[index + 1].innerText || children[index + 1].textContent || '');
            if (siblingText) return siblingText;
          }

          const parentText = normalizeInlineText(parent.innerText || parent.textContent || '');
          const parentDateTime = extractDateTimeValueFromText(parentText, labelPattern);
          if (parentDateTime) return parentDateTime;

          const parentMatch = parentText.match(new RegExp(`${labelPattern.source}\\s*[:：]?\\s*(.+)$`));
          if (parentMatch && parentMatch[1]) {
            const candidate = parentMatch[1].trim();
            if (candidate && candidate !== text) return candidate;
          }
        }
      }
    }

    const pageText = getOrderPageText();
    const pageDateTime = extractDateTimeValueFromText(pageText, labelPattern);
    if (pageDateTime) return pageDateTime;

    const directMatch = pageText.match(/発送日時\s*[:：]\s*([^\n\r]+)/);
    if (directMatch && directMatch[1]) return directMatch[1].trim();

    const inlineMatch = pageText.match(/発送日時\s+([^\n\r]+)/);
    if (inlineMatch && inlineMatch[1]) return inlineMatch[1].trim();

    const multilineMatch = pageText.match(/発送日時\s*[\n\r]+\s*([^\n\r]+)/);
    return multilineMatch && multilineMatch[1] ? multilineMatch[1].trim() : '';
  }

  async function collectOrderShipmentStatus() {
    await waitForOrderDetailContent(12000);

    const shippedAtText = extractFieldValueByLabel(/発送日時/);
    const shippingSignals = getShippingStateSignals();
    const shipped = !!shippedAtText || isShippedOrderPage();
    const normalized = normalizeShippedAt(shippedAtText);
    const orderNumber = getOrderNumberFromPage(null) || '';
    const mount = getOrderDetailMount();
    const mountText = mount ? normalizeInlineText(mount.innerText || mount.textContent || '').slice(0, 80) : '';

    return {
      ok: true,
      orderNumber,
      shipped,
      shippedAt: normalized.iso || '',
      shippedAtRaw: normalized.raw || shippedAtText || '',
      diagnosticsSummary: `shipped=${shipped ? 'yes' : 'no'}, shippedAtRaw=${normalized.raw || shippedAtText || ''}, pending=${shippingSignals.hasPendingStatus ? 'yes' : 'no'}, shippedWord=${shippingSignals.hasShippedStatus ? 'yes' : 'no'}, trackingWord=${shippingSignals.hasTrackingNumber ? 'yes' : 'no'}, ${mount ? `mountChildren=${mount.childElementCount}` : 'mountMissing'}${mountText ? `, mountText=${mountText}` : ''}${shippingSignals.statusHead ? `, statusHead=${shippingSignals.statusHead.slice(0, 120)}` : ''}`
    };
  }

  function findIssuedQrImage() {
    const issuedContainer = findIssuedQrContainer();
    const containerCandidates = issuedContainer ? Array.from(issuedContainer.querySelectorAll('img')) : [];
    const globalCandidates = Array.from(document.querySelectorAll('img[alt="QR"], img[src*="yamato"], img[src*="qr" i], img[data-src*="yamato"], img[data-src*="qr" i]'));
    const candidates = Array.from(new Set(containerCandidates.concat(globalCandidates)));

    const scoredCandidates = candidates
      .map(function(image) {
        const width = image.naturalWidth || image.width;
        const height = image.naturalHeight || image.height;
        const alt = (image.getAttribute('alt') || '').trim();
        const src = image.currentSrc || image.src || image.getAttribute('src') || '';
        const inIssuedContainer = !!issuedContainer && issuedContainer.contains(image);
        const squareDelta = Math.abs(width - height);
        const visibleBonus = isElementVisible(image) ? 200 : 0;
        const containerBonus = inIssuedContainer ? 150 : 0;
        const exactQrAltBonus = alt === 'QR' ? 1000 : 0;
        const qrUrlBonus = /yamato_invoices|qr_code/i.test(src) ? 500 : 0;
        const genericQrBonus = (/qr/i.test(alt) || /qr/i.test(src)) ? 100 : 0;
        const penalty = /base_resized|pximg|booth\.pximg\.net/i.test(src) ? 400 : 0;
        return {
          image,
          width,
          height,
          alt,
          src,
          score: exactQrAltBonus + qrUrlBonus + genericQrBonus + visibleBonus + containerBonus + Math.min(width, height) - squareDelta - penalty
        };
      })
      .filter(function(candidate) {
        if (candidate.width < 80 || candidate.height < 80) return false;
        if (candidate.alt && candidate.alt !== 'QR' && !/qr/i.test(candidate.alt) && /base_resized|pximg|booth\.pximg\.net/i.test(candidate.src)) {
          return false;
        }
        return true;
      })
      .sort(function(left, right) {
        return right.score - left.score;
      });

    return scoredCandidates.find(function(candidate) {
      const image = candidate.image;
      const width = image.naturalWidth || image.width;
      const height = image.naturalHeight || image.height;
      const inIssuedContainer = !!issuedContainer && issuedContainer.contains(image);
      return (inIssuedContainer && width >= 60 && height >= 60) || (width >= 100 && height >= 100);
    })?.image || null;
  }

  function findIssuedQrContainer() {
    const candidates = Array.from(document.querySelectorAll('section, article, div, li, dl'));
    return candidates.find(function(element) {
      const text = (element.innerText || element.textContent || '').trim();
      if (!text) return false;
      return /受付番号/.test(text) && /パスワード/.test(text);
    }) || null;
  }

  function isLikelyQrImage(image) {
    if (!image) return false;
    const alt = image.getAttribute('alt') || '';
    const src = image.currentSrc || image.src || '';
    return alt === 'QR' || QR_IMAGE_PATTERNS.some(function(pattern) {
      return pattern.test(src) || pattern.test(alt);
    });
  }

  async function blobToDataUrl(blob) {
    return await new Promise(function(resolve, reject) {
      const reader = new FileReader();
      reader.onload = function() { resolve(reader.result); };
      reader.onerror = function() { reject(reader.error || new Error('Blobの読み込みに失敗しました')); };
      reader.readAsDataURL(blob);
    });
  }

  async function ensureImageReady(image, timeoutMs) {
    if (!image) return;
    if (image.complete && (image.naturalWidth || image.width) > 0) return;

    await new Promise(function(resolve) {
      let settled = false;
      const done = function() {
        if (settled) return;
        settled = true;
        image.removeEventListener('load', done);
        image.removeEventListener('error', done);
        resolve();
      };

      const timerId = setTimeout(done, timeoutMs || 3000);
      image.addEventListener('load', function onLoad() {
        clearTimeout(timerId);
        done();
      }, { once: true });
      image.addEventListener('error', function onError() {
        clearTimeout(timerId);
        done();
      }, { once: true });
    });

    if (typeof image.decode === 'function') {
      try {
        await image.decode();
      } catch (_error) {
      }
    }
  }

  function getImageUrl(image) {
    if (!image) return '';
    return image.getAttribute('src') || image.currentSrc || image.src || image.getAttribute('data-src') || '';
  }

  function resolveImageUrl(url) {
    if (!url) return '';
    try {
      return new URL(url, window.location.href).href;
    } catch (_error) {
      return url;
    }
  }

  function isHttpUrl(url) {
    return /^https?:/i.test(url || '');
  }

  function isInlineUrl(url) {
    return /^(data:|blob:)/i.test(url || '');
  }

  function isBlobUrl(url) {
    return /^blob:/i.test(url || '');
  }

  function isDataUrl(url) {
    return /^data:/i.test(url || '');
  }

  function findQrImageUrlFromMarkup(scopeElement) {
    const scopes = [scopeElement, document.documentElement].filter(Boolean);
    const patterns = [
      /https:\/\/s2\.booth\.pm\/yamato_invoices\/[^"'\s>]+\.png/i,
      /https:\/\/[^"'\s>]*booth\.pm\/yamato_invoices\/[^"'\s>]+\.png/i,
      /\/yamato_invoices\/[^"'\s>]+\.png/i
    ];

    for (const scope of scopes) {
      const html = scope.innerHTML || '';
      for (const pattern of patterns) {
        const match = html.match(pattern);
        if (match && match[0]) {
          return resolveImageUrl(match[0]);
        }
      }
    }

    return '';
  }

  function findQrImageUrlFromPerformance() {
    if (!window.performance || typeof window.performance.getEntriesByType !== 'function') {
      return '';
    }

    const entries = window.performance.getEntriesByType('resource');
    if (!Array.isArray(entries) || entries.length === 0) {
      return '';
    }

    const candidates = entries
      .map(function(entry) {
        return entry && typeof entry.name === 'string' ? entry.name : '';
      })
      .filter(function(url) {
        return /\/yamato_invoices\/.*\/qr_code\/.*\.png(?:$|[?#])/i.test(url);
      });

    return candidates.length > 0 ? candidates[candidates.length - 1] : '';
  }

  function findQrImageUrlFromImages() {
    const images = Array.from(document.querySelectorAll('img[alt="QR"], img'));
    const urls = images
      .map(function(image) {
        return [
          image.getAttribute('src'),
          image.getAttribute('data-src'),
          image.currentSrc,
          image.src
        ].filter(Boolean);
      })
      .flat()
      .map(resolveImageUrl)
      .filter(function(url) {
        return /\/yamato_invoices\/.*\/qr_code\/.*\.png(?:$|[?#])/i.test(url);
      });

    return urls.length > 0 ? urls[0] : '';
  }

  function getDirectQrImageUrl(image, scopeElement) {
    const candidates = [
      image && image.getAttribute('src'),
      image && image.getAttribute('data-src'),
      image && image.currentSrc,
      image && image.src,
      findQrImageUrlFromMarkup(scopeElement),
      findQrImageUrlFromImages(),
      findQrImageUrlFromPerformance()
    ].filter(Boolean).map(resolveImageUrl);

    return candidates.find(isHttpUrl) || candidates[0] || '';
  }

  function getQrCandidateDiagnostics(image, scopeElement) {
    const markupUrl = findQrImageUrlFromMarkup(scopeElement);
    const imageListUrl = findQrImageUrlFromImages();
    const performanceUrl = findQrImageUrlFromPerformance();
    const diagnostics = {
      hasImageElement: !!image,
      imageAlt: image ? (image.getAttribute('alt') || '') : '',
      imageSrcAttr: image ? (image.getAttribute('src') || '') : '',
      imageDataSrcAttr: image ? (image.getAttribute('data-src') || '') : '',
      imageCurrentSrc: image ? (image.currentSrc || '') : '',
      imageSrc: image ? (image.src || '') : '',
      imageWidth: image ? (image.naturalWidth || image.width || 0) : 0,
      imageHeight: image ? (image.naturalHeight || image.height || 0) : 0,
      markupUrl,
      imageListUrl,
      performanceUrl,
      directUrl: getDirectQrImageUrl(image, scopeElement)
    };
    return diagnostics;
  }

  function summarizeDiagnostics(diagnostics) {
    if (!diagnostics) return '';
    const parts = [];
    parts.push(`img=${diagnostics.hasImageElement ? 'yes' : 'no'}`);
    if (diagnostics.imageAlt) parts.push(`alt=${diagnostics.imageAlt}`);
    if (diagnostics.imageWidth || diagnostics.imageHeight) parts.push(`size=${diagnostics.imageWidth}x${diagnostics.imageHeight}`);
    if (diagnostics.imageSrcAttr) parts.push(`srcAttr=${diagnostics.imageSrcAttr.slice(0, 120)}`);
    if (diagnostics.imageCurrentSrc) parts.push(`currentSrc=${diagnostics.imageCurrentSrc.slice(0, 120)}`);
    if (diagnostics.markupUrl) parts.push(`markup=${diagnostics.markupUrl.slice(0, 120)}`);
    if (diagnostics.imageListUrl) parts.push(`images=${diagnostics.imageListUrl.slice(0, 120)}`);
    if (diagnostics.performanceUrl) parts.push(`perf=${diagnostics.performanceUrl.slice(0, 120)}`);
    if (diagnostics.directUrl) parts.push(`direct=${diagnostics.directUrl.slice(0, 120)}`);
    return parts.join(', ');
  }

  function getCurrentQrDiagnostics() {
    const scopeElement = findIssuedQrContainer();
    const image = findIssuedQrImage();
    return getQrCandidateDiagnostics(image, scopeElement);
  }

  async function loadImageFromUrl(imageUrl, timeoutMs) {
    return await new Promise(function(resolve, reject) {
      const image = new Image();
      let settled = false;
      const finish = function(callback, value) {
        if (settled) return;
        settled = true;
        clearTimeout(timerId);
        callback(value);
      };
      const timerId = setTimeout(function() {
        finish(reject, new Error('画像の読み込みがタイムアウトしました'));
      }, timeoutMs || 5000);

      image.onload = function() {
        finish(resolve, image);
      };
      image.onerror = function() {
        finish(reject, new Error('画像の読み込みに失敗しました'));
      };
      image.src = imageUrl;
    });
  }

  async function fetchUrlAsDataUrl(imageUrl, requestOptions) {
    const response = await fetch(imageUrl, requestOptions || {});
    if (!response.ok) {
      throw new Error('QR画像の取得に失敗しました');
    }
    return await blobToDataUrl(await response.blob());
  }

  function drawImageToDataUrl(image) {
    const canvas = document.createElement('canvas');
    canvas.width = image.naturalWidth || image.width;
    canvas.height = image.naturalHeight || image.height;
    const context = canvas.getContext('2d');
    if (!context || !canvas.width || !canvas.height) return null;
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL('image/png');
  }

  async function imageElementToDataUrl(image) {
    const imageUrl = getImageUrl(image);
    if (!image || !imageUrl) return null;

    await ensureImageReady(image, 3000);

    if (isDataUrl(imageUrl)) {
      return imageUrl;
    }

    if (isBlobUrl(imageUrl)) {
      try {
        return await fetchUrlAsDataUrl(imageUrl, { cache: 'no-store' });
      } catch (_blobFetchError) {
        try {
          const loadedImage = await loadImageFromUrl(imageUrl, 5000);
          return drawImageToDataUrl(loadedImage);
        } catch (_blobReloadError) {
        }
      }
    }

    try {
      return drawImageToDataUrl(image);
    } catch (_canvasFirstError) {
    }

    if (isHttpUrl(imageUrl) || isInlineUrl(imageUrl)) {
      try {
        return await fetchUrlAsDataUrl(imageUrl, {
          credentials: 'include',
          cache: 'no-store'
        });
      } catch (_fetchError) {
        try {
          const loadedImage = await loadImageFromUrl(imageUrl, 5000);
          return drawImageToDataUrl(loadedImage);
        } catch (_reloadError) {
        }
      }
    }

    return null;
  }

  function extractIssuedQrMetadata() {
    const text = document.body ? document.body.innerText : '';
    const receiptMatch = text.match(/受付番号[:：]\s*(\d+)/);
    const passwordMatch = text.match(/パスワード[:：]\s*(\d+)/);
    const orderMatch = text.match(/注文番号[:：]\s*(\d+)/);
    const expiresMatch = text.match(/有効期限[:：]\s*([^\n]+)/);
    const sizeMatch = text.match(/サイズ[:：]\s*([^\n]+)/);
    return {
      receiptnum: receiptMatch ? receiptMatch[1] : null,
      receiptpassword: passwordMatch ? passwordMatch[1] : null,
      orderNumber: orderMatch ? orderMatch[1] : null,
      expiresAt: expiresMatch ? expiresMatch[1].trim() : null,
      packageSizeLabel: sizeMatch ? sizeMatch[1].trim() : null
    };
  }

  function hasIssuedQrMetadata() {
    const metadata = extractIssuedQrMetadata();
    return !!(metadata.receiptnum && metadata.receiptpassword);
  }

  async function getStructuredQrCandidate() {
    const scopeElement = findIssuedQrContainer();
    const image = findIssuedQrImage();
    if (!image) return null;
    const imageUrl = getDirectQrImageUrl(image, scopeElement);
    const dataUrl = await imageElementToDataUrl(image);
    if (!dataUrl && !imageUrl) return null;
    const metadata = extractIssuedQrMetadata();
    const diagnostics = getQrCandidateDiagnostics(image, scopeElement);
    return {
      type: 'img',
      dataUrl: dataUrl,
      imageUrl: imageUrl,
      diagnostics,
      receiptnum: metadata.receiptnum,
      receiptpassword: metadata.receiptpassword,
      orderNumber: metadata.orderNumber,
      expiresAt: metadata.expiresAt,
      packageSizeLabel: metadata.packageSizeLabel
    };
  }

  function getNormalizedIssueSettings(rawSettings) {
    const settings = Object.assign({}, DEFAULT_ISSUE_SETTINGS, rawSettings || {});
    if (!Array.isArray(settings.handlingCodes)) settings.handlingCodes = [];
    settings.handlingCodes = Array.from(new Set(settings.handlingCodes.filter(Boolean)));
    settings.packageSizeId = String(settings.packageSizeId || DEFAULT_ISSUE_SETTINGS.packageSizeId);
    settings.codeType = String(settings.codeType || DEFAULT_ISSUE_SETTINGS.codeType);
    settings.description = typeof settings.description === 'string' ? settings.description.trim() : DEFAULT_ISSUE_SETTINGS.description;
    settings.includeOrderNumber = settings.includeOrderNumber !== false;
    return settings;
  }

  function getOrderNumberFromPage(form) {
    const text = document.body ? document.body.innerText : '';
    const textMatch = text.match(/注文番号\s*[:：]\s*(\d+)/);
    if (textMatch) return textMatch[1];

    const headingMatch = text.match(/注文詳細\s*-\s*注文番号\s*[:：]\s*(\d+)/);
    if (headingMatch) return headingMatch[1];

    const action = form && form.getAttribute('action');
    if (action) {
      const actionMatch = action.match(/\/orders\/(\d+)\/register_yamato_invoice/);
      if (actionMatch) return actionMatch[1];
    }

    const pathMatch = window.location && window.location.pathname
      ? window.location.pathname.match(/\/orders\/(\d+)/)
      : null;
    if (pathMatch) return pathMatch[1];

    return '';
  }

  function buildDescriptionValue(form, settings) {
    const baseDescription = settings.description || DEFAULT_ISSUE_SETTINGS.description;
    if (!settings.includeOrderNumber) {
      return baseDescription;
    }

    const orderNumber = getOrderNumberFromPage(form);
    if (!orderNumber) {
      return baseDescription;
    }

    return `${baseDescription}:${orderNumber}`;
  }

  function setValueIfPresent(form, selector, value) {
    const input = form.querySelector(selector);
    if (!input) return;
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function setFieldValue(element, value) {
    if (!element) return;
    const prototype = element.tagName === 'TEXTAREA'
      ? window.HTMLTextAreaElement && window.HTMLTextAreaElement.prototype
      : window.HTMLInputElement && window.HTMLInputElement.prototype;
    const descriptor = prototype ? Object.getOwnPropertyDescriptor(prototype, 'value') : null;
    element.focus();
    if (descriptor && typeof descriptor.set === 'function') {
      descriptor.set.call(element, value);
    } else {
      element.value = value;
    }
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
    element.blur();
  }

  function findShipmentCommentField(scope) {
    const root = scope || document;
    const directMatch = root.querySelector('textarea[name="shipment_request[shipment_comment]"]');
    if (directMatch && isElementVisible(directMatch)) {
      return directMatch;
    }

    const candidates = Array.from(root.querySelectorAll('textarea, input[type="text"]'));
    return candidates.find(function(element) {
      if (!isElementVisible(element)) return false;
      const name = element.getAttribute('name') || '';
      const placeholder = element.getAttribute('placeholder') || '';
      const ariaLabel = element.getAttribute('aria-label') || '';
      const containerText = normalizeInlineText((element.closest('form, section, div, dialog') || root).innerText || '');
      return name === 'shipment_request[shipment_comment]'
        || /comment|message|note/i.test(name)
        || /発送通知コメント|コメント/.test(placeholder)
        || /発送通知コメント|コメント/.test(ariaLabel)
        || /発送通知コメント/.test(containerText);
    }) || null;
  }

  function findShipmentSubmitButton(scope, excludedElement) {
    const root = scope || document;
    const directCandidates = Array.from(root.querySelectorAll('button[type="submit"][name="commit"], input[type="submit"][name="commit"]'));
    const directMatch = directCandidates.find(function(element) {
      if (element === excludedElement) return false;
      const text = normalizeInlineText(element.innerText || element.textContent || element.value || '');
      return isElementVisible(element) && /発送完了通知を送信/.test(text);
    });
    if (directMatch) return directMatch;

    const candidates = Array.from(root.querySelectorAll('button, [role="button"], input[type="submit"], input[type="button"]'));
    return candidates.find(function(element) {
      if (element === excludedElement) return false;
      const text = normalizeInlineText(element.innerText || element.textContent || element.value || '');
      if (!text) return false;
      return isElementVisible(element) && /発送完了通知を送信|発送完了を通知|通知する|送信する|完了を通知/.test(text);
    }) || null;
  }

  async function waitForShipmentNotificationUI(trigger, timeoutMs) {
    const deadline = Date.now() + (timeoutMs || 6000);
    while (Date.now() < deadline) {
      if (isShippedOrderPage()) {
        return {
          commentField: null,
          submitButton: null,
          form: null,
          alreadyShipped: true
        };
      }

      const commentField = findShipmentCommentField();
      const submitButton = findShipmentSubmitButton(document, trigger);
      const form = (commentField && commentField.closest('form')) || (submitButton && submitButton.closest('form')) || null;
      if (commentField || submitButton) {
        return {
          commentField,
          submitButton,
          form,
          alreadyShipped: false
        };
      }

      await new Promise(function(resolve) {
        setTimeout(resolve, 250);
      });
    }

    return {
      commentField: findShipmentCommentField(),
      submitButton: findShipmentSubmitButton(document, trigger),
      form: null,
      alreadyShipped: isShippedOrderPage()
    };
  }

  function syncHandlingCodes(form, handlingCodes) {
    const checkboxes = Array.from(form.querySelectorAll('input[name="yamato_invoice_form[handling_codes][]"]'));
    checkboxes.forEach(function(input) {
      input.checked = handlingCodes.includes(input.value);
      input.dispatchEvent(new Event('change', { bubbles: true }));
    });
  }

  function syncIncludeOrderNumber(form, includeOrderNumber) {
    const checkbox = form.querySelector('input[type="checkbox"][name="yamato_invoice_form[display_order_id]"]')
      || form.querySelector('input[type="checkbox"][name*="display_order"]')
      || Array.from(form.querySelectorAll('input[type="checkbox"]')).find(function(input) {
        const label = input.closest('label');
        const text = label ? (label.innerText || label.textContent || '') : '';
        return /品名に注文番号を記載する/.test(text);
      });
    if (!checkbox) return;
    checkbox.checked = !!includeOrderNumber;
    checkbox.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function applyIssueSettings(form, rawSettings) {
    const settings = getNormalizedIssueSettings(rawSettings);
    const descriptionValue = buildDescriptionValue(form, settings);
    setValueIfPresent(form, '[name="yamato_invoice_form[package_size_id]"]', settings.packageSizeId);
    setValueIfPresent(form, '[name="yamato_invoice_form[code_type]"]', settings.codeType);
    setValueIfPresent(form, '[name="yamato_invoice_form[description]"]', descriptionValue);
    syncIncludeOrderNumber(form, settings.includeOrderNumber);
    syncHandlingCodes(form, settings.handlingCodes);
  }

  function submitIssueForm(form) {
    const submitButton = form.querySelector('button[type="submit"], input[type="submit"]') || findTriggerButton();
    if (submitButton) {
      submitButton.click();
      return;
    }
    if (typeof form.requestSubmit === 'function') {
      form.requestSubmit();
      return;
    }
    form.submit();
  }

  function scoreCandidate(width, height, visible) {
    const squareDelta = Math.abs(width - height);
    const sizeScore = Math.min(width, height);
    return (visible ? 200 : 0) + sizeScore - squareDelta;
  }

  function getImageCandidates() {
    const candidates = [];

    Array.from(document.images).forEach(function(image) {
      if (!isLikelyQrImage(image)) return;
      const width = image.naturalWidth || image.width;
      const height = image.naturalHeight || image.height;
      if (width < 80 || height < 80) return;
      candidates.push({
        type: 'img',
        score: scoreCandidate(width, height, isElementVisible(image)),
        dataUrl: image.src
      });
    });

    Array.from(document.querySelectorAll('canvas')).forEach(function(canvas) {
      const width = canvas.width || canvas.clientWidth;
      const height = canvas.height || canvas.clientHeight;
      if (width < 80 || height < 80) return;
      candidates.push({
        type: 'canvas',
        score: scoreCandidate(width, height, isElementVisible(canvas)),
        dataUrl: canvas.toDataURL('image/png')
      });
    });

    candidates.sort(function(left, right) {
      return right.score - left.score;
    });
    return candidates;
  }

  async function waitForQrCandidate(timeoutMs) {
    const startedAt = Date.now();
    while (Date.now() - startedAt < timeoutMs) {
      const structuredCandidate = await getStructuredQrCandidate();
      if (structuredCandidate) {
        return structuredCandidate;
      }
      const candidates = getImageCandidates();
      if (candidates.length > 0) {
        return candidates[0];
      }
      await new Promise(function(resolve) {
        setTimeout(resolve, 400);
      });
    }
    return null;
  }

  async function waitForExistingQr(timeoutMs) {
    const startedAt = Date.now();
    while (Date.now() - startedAt < timeoutMs) {
      const structuredCandidate = await getStructuredQrCandidate();
      if (structuredCandidate) {
        return structuredCandidate;
      }
      await new Promise(function(resolve) {
        setTimeout(resolve, 300);
      });
    }
    return null;
  }

  async function collectOrderQr(rawSettings) {
    const existingQr = await getStructuredQrCandidate();
    if (existingQr) {
      return {
        ok: true,
        dataUrl: existingQr.dataUrl,
        imageUrl: existingQr.imageUrl,
        diagnostics: existingQr.diagnostics,
        diagnosticsSummary: summarizeDiagnostics(existingQr.diagnostics),
        sourceType: existingQr.type,
        receiptnum: existingQr.receiptnum,
        receiptpassword: existingQr.receiptpassword,
        orderNumber: existingQr.orderNumber,
        expiresAt: existingQr.expiresAt,
        packageSizeLabel: existingQr.packageSizeLabel,
        alreadyIssued: true
      };
    }

    if (isShippedOrderPage()) {
      return {
        ok: false,
        error: 'この注文は発送済みのため、発送コードの取得や再発行はできません'
      };
    }

    if (hasIssuedQrMetadata()) {
      const delayedExistingQr = await waitForExistingQr(3000);
      if (delayedExistingQr) {
        return {
          ok: true,
          dataUrl: delayedExistingQr.dataUrl,
          imageUrl: delayedExistingQr.imageUrl,
          diagnostics: delayedExistingQr.diagnostics,
          diagnosticsSummary: summarizeDiagnostics(delayedExistingQr.diagnostics),
          sourceType: delayedExistingQr.type,
          receiptnum: delayedExistingQr.receiptnum,
          receiptpassword: delayedExistingQr.receiptpassword,
          orderNumber: delayedExistingQr.orderNumber,
          expiresAt: delayedExistingQr.expiresAt,
          packageSizeLabel: delayedExistingQr.packageSizeLabel,
          alreadyIssued: true
        };
      }

      const diagnostics = getCurrentQrDiagnostics();
      return {
        ok: false,
        error: `発送コードは既に発行済みですが、QR画像を取得できませんでした${summarizeDiagnostics(diagnostics) ? ` | ${summarizeDiagnostics(diagnostics)}` : ''}`,
        diagnostics,
        diagnosticsSummary: summarizeDiagnostics(diagnostics)
      };
    }

    const issueForm = getIssueForm();
    if (issueForm) {
      applyIssueSettings(issueForm, rawSettings);
      submitIssueForm(issueForm);
    } else {
      const trigger = findTriggerButton();
      if (!trigger) {
        return { ok: false, error: '発送コードを発行するボタンが見つかりません' };
      }

      trigger.click();
    }

    const candidate = await waitForQrCandidate(15000);
    if (!candidate) {
      return { ok: false, error: 'QRコード画像またはcanvasを検出できませんでした' };
    }

    return {
      ok: true,
      dataUrl: candidate.dataUrl,
      sourceType: candidate.type,
      receiptnum: candidate.receiptnum || null,
      receiptpassword: candidate.receiptpassword || null,
      orderNumber: candidate.orderNumber || null,
      expiresAt: candidate.expiresAt || null,
      packageSizeLabel: candidate.packageSizeLabel || null,
      alreadyIssued: false
    };
  }

  async function submitOrderShipmentNotification(messageTemplate) {
    await waitForOrderDetailContent(12000);

    const currentStatus = await collectOrderShipmentStatus();
    if (currentStatus.shipped) {
      return Object.assign({}, currentStatus, {
        alreadyShipped: true,
        submitted: false
      });
    }

    const trigger = findShipmentNotifyTrigger();
    if (!trigger) {
      return {
        ok: false,
        error: '発送完了を通知ボタンが見つかりません',
        diagnosticsSummary: `mountChildren=${getOrderDetailMount() ? getOrderDetailMount().childElementCount : 'missing'}`
      };
    }

    trigger.click();
    const uiState = await waitForShipmentNotificationUI(trigger, 6000);
    if (uiState.alreadyShipped) {
      return Object.assign({}, await collectOrderShipmentStatus(), {
        alreadyShipped: true,
        submitted: false
      });
    }

    if (uiState.commentField && typeof messageTemplate === 'string') {
      setFieldValue(uiState.commentField, messageTemplate);
    }

    const submitButton = uiState.submitButton || findShipmentSubmitButton(uiState.form || document, trigger);
    if (submitButton) {
      submitButton.click();
      return {
        ok: true,
        submitted: true,
        orderNumber: getOrderNumberFromPage(uiState.form),
        diagnosticsSummary: `shipmentSubmitted=yes, usedComment=${uiState.commentField ? 'yes' : 'no'}`
      };
    }

    if (uiState.form && typeof uiState.form.requestSubmit === 'function') {
      uiState.form.requestSubmit();
      return {
        ok: true,
        submitted: true,
        orderNumber: getOrderNumberFromPage(uiState.form),
        diagnosticsSummary: `shipmentSubmitted=yes, usedComment=${uiState.commentField ? 'yes' : 'no'}, submitMode=form` 
      };
    }

    return {
      ok: false,
      error: '発送通知の送信ボタンが見つかりません',
      diagnosticsSummary: `shipmentSubmitted=no, usedComment=${uiState.commentField ? 'yes' : 'no'}`
    };
  }

  chrome.runtime.onMessage.addListener(function(message, _sender, sendResponse) {
    if (!message || !message.type) return undefined;

    if (message.type === 'boothcsv:collect-order-qr') {
      collectOrderQr(message.yamatoIssueSettings)
        .then(sendResponse)
        .catch(function(error) {
          sendResponse({
            ok: false,
            error: error && error.message ? error.message : 'QR取得に失敗しました'
          });
        });
      return true;
    }

    if (message.type === 'boothcsv:collect-order-shipment-status') {
      Promise.resolve(collectOrderShipmentStatus())
        .then(sendResponse)
        .catch(function(error) {
          sendResponse({
            ok: false,
            error: error && error.message ? error.message : '発送状態の取得に失敗しました'
          });
        });
      return true;
    }

    if (message.type === 'boothcsv:submit-order-shipment-notification') {
      Promise.resolve(submitOrderShipmentNotification(message.messageTemplate || ''))
        .then(sendResponse)
        .catch(function(error) {
          sendResponse({
            ok: false,
            error: error && error.message ? error.message : '発送通知の送信に失敗しました'
          });
        });
      return true;
    }

    return undefined;
  });
})();