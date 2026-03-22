'use strict';

const APP_URL = chrome.runtime.getURL('boothcsv.html');
const CSV_URL = 'https://manage.booth.pm/orders/csv?state=paid';
const ORDER_URL_PREFIX = 'https://manage.booth.pm/orders/';

chrome.action.onClicked.addListener(async () => {
  await chrome.tabs.create({ url: APP_URL });
});

async function fetchBoothCsv() {
  const response = await fetch(CSV_URL, {
    credentials: 'include',
    cache: 'no-store',
    redirect: 'follow'
  });

  const csvText = await response.text();
  if (!response.ok) {
    throw new Error(`CSV取得に失敗しました: HTTP ${response.status}`);
  }
  if (!csvText.includes('注文番号')) {
    throw new Error('CSV内容を確認できませんでした。BOOTHにログイン済みか確認してください。');
  }

  return {
    csvText,
    fileName: `booth-orders-${new Date().toISOString().slice(0, 10)}.csv`
  };
}

async function waitForTabComplete(tabId, timeoutMs = 20000) {
  return await new Promise((resolve, reject) => {
    const timeoutId = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      reject(new Error('注文詳細ページの読み込みがタイムアウトしました'));
    }, timeoutMs);

    function listener(updatedTabId, changeInfo) {
      if (updatedTabId !== tabId || changeInfo.status !== 'complete') return;
      clearTimeout(timeoutId);
      chrome.tabs.onUpdated.removeListener(listener);
      resolve();
    }

    chrome.tabs.onUpdated.addListener(listener);
  });
}

async function sendMessageWithRetry(tabId, message, retries = 15) {
  for (let attempt = 0; attempt < retries; attempt++) {
    try {
      const response = await chrome.tabs.sendMessage(tabId, message);
      if (response) return response;
    } catch (error) {
      if (attempt === retries - 1) throw error;
    }
    await new Promise(resolve => setTimeout(resolve, 500));
  }
  throw new Error('content scriptとの通信に失敗しました');
}

async function blobToDataUrl(blob) {
  const buffer = await blob.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }
  return `data:${blob.type || 'image/png'};base64,${btoa(binary)}`;
}

function isHttpUrl(url) {
  return /^https?:/i.test(url || '');
}

function isDataUrl(url) {
  return /^data:/i.test(url || '');
}

function resolveAbsoluteUrl(url, baseUrl) {
  if (!url) return '';
  try {
    return new URL(url, baseUrl || ORDER_URL_PREFIX).href;
  } catch (_error) {
    return url;
  }
}

async function fetchImageAsDataUrl(imageUrl, referrerUrl) {
  const requestOptions = {
    cache: 'no-store',
    credentials: 'include'
  };

  if (referrerUrl) {
    requestOptions.referrer = referrerUrl;
  }

  try {
    const response = await fetch(imageUrl, requestOptions);
    if (!response.ok) {
      throw new Error(`QR画像の取得に失敗しました: HTTP ${response.status}`);
    }
    return await blobToDataUrl(await response.blob());
  } catch (_firstError) {
    const response = await fetch(imageUrl, {
      cache: 'no-store',
      credentials: 'omit'
    });
    if (!response.ok) {
      throw new Error(`QR画像の取得に失敗しました: HTTP ${response.status}`);
    }
    return await blobToDataUrl(await response.blob());
  }
}

async function normalizeQrResponse(response, pageUrl) {
  if (!response || !response.ok) return response;

  if (isDataUrl(response.dataUrl)) {
    return response;
  }

  const candidateUrl = response.imageUrl || response.dataUrl;
  const imageUrl = resolveAbsoluteUrl(candidateUrl, pageUrl);
  if (!isHttpUrl(imageUrl)) return response;

  try {
    return {
      ...response,
      dataUrl: await fetchImageAsDataUrl(imageUrl, pageUrl)
    };
  } catch (error) {
    const originalMessage = error && error.message ? error.message : 'QR画像の取得に失敗しました';
    const diagnosticSuffix = response.diagnosticsSummary ? ` | ${response.diagnosticsSummary}` : '';
    throw new Error(`${originalMessage} (${imageUrl})${diagnosticSuffix}`);
  }
}

async function collectOrderQr(orderNumber, yamatoIssueSettings) {
  const tab = await chrome.tabs.create({
    url: `${ORDER_URL_PREFIX}${encodeURIComponent(orderNumber)}`,
    active: false
  });

  try {
    await waitForTabComplete(tab.id);
    const response = await sendMessageWithRetry(tab.id, {
      type: 'boothcsv:collect-order-qr',
      orderNumber,
      yamatoIssueSettings
    });
    if (response && response.diagnostics && !response.diagnosticsSummary) {
      response.diagnosticsSummary = [
        `img=${response.diagnostics.hasImageElement ? 'yes' : 'no'}`,
        response.diagnostics.imageAlt ? `alt=${response.diagnostics.imageAlt}` : '',
        response.diagnostics.imageWidth || response.diagnostics.imageHeight ? `size=${response.diagnostics.imageWidth}x${response.diagnostics.imageHeight}` : '',
        response.diagnostics.imageSrcAttr ? `srcAttr=${response.diagnostics.imageSrcAttr}` : '',
        response.diagnostics.imageCurrentSrc ? `currentSrc=${response.diagnostics.imageCurrentSrc}` : '',
        response.diagnostics.markupUrl ? `markup=${response.diagnostics.markupUrl}` : '',
        response.diagnostics.imageListUrl ? `images=${response.diagnostics.imageListUrl}` : '',
        response.diagnostics.performanceUrl ? `perf=${response.diagnostics.performanceUrl}` : '',
        response.diagnostics.directUrl ? `direct=${response.diagnostics.directUrl}` : ''
      ].filter(Boolean).join(', ');
    }
    const normalizedResponse = await normalizeQrResponse(response, tab.url);
    if (normalizedResponse && normalizedResponse.ok && !normalizedResponse.dataUrl) {
      return {
        ok: false,
        error: `QR画像のデータ化に失敗しました${normalizedResponse.diagnosticsSummary ? ` | ${normalizedResponse.diagnosticsSummary}` : ''}`
      };
    }
    return normalizedResponse;
  } finally {
    if (tab.id) {
      await chrome.tabs.remove(tab.id).catch(() => {});
    }
  }
}

async function collectOrderShipmentStatus(orderNumber) {
  const tab = await chrome.tabs.create({
    url: `${ORDER_URL_PREFIX}${encodeURIComponent(orderNumber)}`,
    active: false
  });

  try {
    await waitForTabComplete(tab.id);
    return await sendMessageWithRetry(tab.id, {
      type: 'boothcsv:collect-order-shipment-status',
      orderNumber
    });
  } finally {
    if (tab.id) {
      await chrome.tabs.remove(tab.id).catch(() => {});
    }
  }
}

async function notifyOrderShipment(orderNumber, messageTemplate) {
  const tab = await chrome.tabs.create({
    url: `${ORDER_URL_PREFIX}${encodeURIComponent(orderNumber)}`,
    active: false
  });

  try {
    await waitForTabComplete(tab.id);

    const initialStatus = await sendMessageWithRetry(tab.id, {
      type: 'boothcsv:collect-order-shipment-status',
      orderNumber
    });
    if (initialStatus && initialStatus.ok && initialStatus.shippedAt) {
      return {
        ...initialStatus,
        alreadyShipped: true,
        submitted: false
      };
    }

    const submitResponse = await sendMessageWithRetry(tab.id, {
      type: 'boothcsv:submit-order-shipment-notification',
      orderNumber,
      messageTemplate: typeof messageTemplate === 'string' ? messageTemplate : ''
    });

    if (!submitResponse || !submitResponse.ok) {
      return submitResponse || { ok: false, error: '発送通知の送信に失敗しました' };
    }

    if (submitResponse.alreadyShipped) {
      return submitResponse;
    }

    await waitForTabComplete(tab.id, 20000).catch(() => {});

    const startedAt = Date.now();
    let latestStatus = null;
    while (Date.now() - startedAt < 20000) {
      try {
        latestStatus = await sendMessageWithRetry(tab.id, {
          type: 'boothcsv:collect-order-shipment-status',
          orderNumber
        }, 5);
      } catch (_error) {
        latestStatus = null;
      }

      if (latestStatus && latestStatus.ok && (latestStatus.shippedAt || latestStatus.shippedAtRaw)) {
        return {
          ...latestStatus,
          submitted: true,
          diagnosticsSummary: [submitResponse.diagnosticsSummary, latestStatus.diagnosticsSummary].filter(Boolean).join(' | ')
        };
      }

      await new Promise(resolve => setTimeout(resolve, 500));
    }

    return {
      ok: false,
      error: '発送通知後の発送日時確認に失敗しました',
      diagnosticsSummary: [submitResponse.diagnosticsSummary, latestStatus && latestStatus.diagnosticsSummary].filter(Boolean).join(' | ')
    };
  } finally {
    if (tab.id) {
      await chrome.tabs.remove(tab.id).catch(() => {});
    }
  }
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (!message || !message.type) return undefined;

  if (message.type === 'boothcsv:fetch-csv') {
    fetchBoothCsv()
      .then(result => sendResponse({ ok: true, ...result }))
      .catch(error => sendResponse({ ok: false, error: error && error.message ? error.message : 'CSV取得失敗' }));
    return true;
  }

  if (message.type === 'boothcsv:collect-order-qr') {
    collectOrderQr(message.orderNumber, message.yamatoIssueSettings)
      .then(response => sendResponse(response))
      .catch(error => sendResponse({ ok: false, error: error && error.message ? error.message : 'QR取得失敗' }));
    return true;
  }

  if (message.type === 'boothcsv:collect-order-shipment-status') {
    collectOrderShipmentStatus(message.orderNumber)
      .then(response => sendResponse(response))
      .catch(error => sendResponse({ ok: false, error: error && error.message ? error.message : '発送状態取得失敗' }));
    return true;
  }

  if (message.type === 'boothcsv:notify-order-shipment') {
    notifyOrderShipment(message.orderNumber, message.messageTemplate)
      .then(response => sendResponse(response))
      .catch(error => sendResponse({ ok: false, error: error && error.message ? error.message : '発送通知失敗' }));
    return true;
  }

  return undefined;
});