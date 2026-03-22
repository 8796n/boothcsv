(function() {
  'use strict';

  function canUseExtensionBridge() {
    return !!window.BoothCSVAppRuntime?.canUseExtensionBridge?.();
  }

  function setStatus(elementId, message, state) {
    const statusElement = document.getElementById(elementId);
    if (!statusElement) return;
    statusElement.textContent = message || '';
    statusElement.classList.toggle('is-busy', state === 'busy');
    statusElement.classList.toggle('is-error', state === 'error');
  }

  async function fetchCsvFromExtension() {
    setStatus('extensionFetchCsvStatus', 'CSV取得中...', 'busy');
    const result = await chrome.runtime.sendMessage({ type: 'boothcsv:fetch-csv' });
    if (!result || !result.ok) {
      throw new Error(result && result.error ? result.error : 'CSV取得に失敗しました');
    }
    const imported = await window.BoothCSVExtensionBridge.importCSVText(result.csvText, {
      sourceName: result.fileName || 'BOOTH CSV'
    });
    setStatus('extensionFetchCsvStatus', `${imported.rowCount}件を読込`, 'idle');
  }

  async function collectOrderQRCode(orderNumber) {
    const bridge = window.BoothCSVExtensionBridge;
    if (!bridge) {
      throw new Error('アプリの初期化が完了していません');
    }

    const normalizedOrderNumber = orderNumber ? String(orderNumber).trim() : '';
    if (!normalizedOrderNumber) {
      throw new Error('注文番号が不正です');
    }

    const yamatoIssueSettings = typeof bridge.prepareYamatoIssueSettings === 'function'
      ? await bridge.prepareYamatoIssueSettings()
      : (typeof bridge.getValidatedYamatoIssueSettings === 'function'
      ? bridge.getValidatedYamatoIssueSettings()
      : (typeof bridge.getYamatoIssueSettings === 'function'
      ? bridge.getYamatoIssueSettings()
      : null));

    const response = await chrome.runtime.sendMessage({
      type: 'boothcsv:collect-order-qr',
      orderNumber: normalizedOrderNumber,
      yamatoIssueSettings
    });

    if (!response || !response.ok) {
      throw new Error(response && response.error ? response.error : 'QR取得に失敗しました');
    }

    if (!response.dataUrl) {
      throw new Error('QR画像データが取得できませんでした');
    }

    await bridge.importOrderQRCodeImage(normalizedOrderNumber, response.dataUrl, {
      allowDuplicate: false,
      receiptnum: response.receiptnum || null,
      receiptpassword: response.receiptpassword || null,
      barcodeData: response.barcodeData || null,
      alreadyIssued: !!response.alreadyIssued
    });

    return response;
  }

  async function collectOrderShipmentStatus(orderNumber) {
    const normalizedOrderNumber = orderNumber ? String(orderNumber).trim() : '';
    if (!normalizedOrderNumber) {
      throw new Error('注文番号が不正です');
    }

    const response = await chrome.runtime.sendMessage({
      type: 'boothcsv:collect-order-shipment-status',
      orderNumber: normalizedOrderNumber
    });

    if (!response || !response.ok) {
      throw new Error(response && response.error ? response.error : '発送状態の取得に失敗しました');
    }

    return response;
  }

  async function notifyOrderShipment(orderNumber, messageTemplate) {
    const normalizedOrderNumber = orderNumber ? String(orderNumber).trim() : '';
    if (!normalizedOrderNumber) {
      throw new Error('注文番号が不正です');
    }

    const response = await chrome.runtime.sendMessage({
      type: 'boothcsv:notify-order-shipment',
      orderNumber: normalizedOrderNumber,
      messageTemplate: typeof messageTemplate === 'string' ? messageTemplate : ''
    });

    if (!response || !response.ok) {
      throw new Error(response && response.error ? response.error : '発送通知の送信に失敗しました');
    }

    return response;
  }

  async function syncQrCodes() {
    const bridge = window.BoothCSVExtensionBridge;
    if (!bridge) {
      throw new Error('アプリの初期化が完了していません');
    }

    const orderNumbers = bridge.getPendingQrOrderNumbers();
    if (!orderNumbers.length) {
      setStatus('extensionQrStatus', '未取得QRなし', 'idle');
      return;
    }

    const failures = [];
    for (let index = 0; index < orderNumbers.length; index++) {
      const orderNumber = orderNumbers[index];
      setStatus('extensionQrStatus', `QR一括取得 ${index + 1}/${orderNumbers.length}: ${orderNumber}`, 'busy');

      try {
        await collectOrderQRCode(orderNumber);
      } catch (error) {
        failures.push({ orderNumber, error: error && error.message ? error.message : '保存失敗' });
      }
    }

    if (failures.length > 0) {
      const summary = failures.slice(0, 3).map(item => `${item.orderNumber}: ${item.error}`).join(' / ');
      setStatus('extensionQrStatus', `${failures.length}件失敗`, 'error');
      alert(`QR一括取得で失敗した注文があります。\n${summary}`);
      return;
    }

    setStatus('extensionQrStatus', `${orderNumbers.length}件のQR一括取得完了`, 'idle');
  }

  window.BoothCSVExtensionBridge = Object.assign(window.BoothCSVExtensionBridge || {}, {
    collectOrderQRCode,
    collectOrderShipmentStatus,
    notifyOrderShipment
  });

  document.addEventListener('DOMContentLoaded', function() {
    if (!canUseExtensionBridge()) return;

    const controls = document.getElementById('extensionHeaderCsvControls');
    const csvButton = document.getElementById('extensionFetchCsvButton');
    const qrButton = document.getElementById('extensionSyncQrButton');

    if (csvButton) {
      csvButton.addEventListener('click', async function() {
        csvButton.disabled = true;
        try {
          await fetchCsvFromExtension();
        } catch (error) {
          setStatus('extensionFetchCsvStatus', 'CSV取得失敗', 'error');
          alert(error && error.message ? error.message : 'CSV取得に失敗しました');
        } finally {
          csvButton.disabled = false;
        }
      });
    }

    if (qrButton) {
      qrButton.addEventListener('click', async function() {
        qrButton.disabled = true;
        try {
          await syncQrCodes();
        } catch (error) {
          setStatus('extensionQrStatus', 'QR一括取得失敗', 'error');
          alert(error && error.message ? error.message : 'QR一括取得に失敗しました');
        } finally {
          qrButton.disabled = false;
        }
      });
    }
  });
})();