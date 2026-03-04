(() => {
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  // UI elements
  const btnExport = $('#btn-export');
  const btnCancel = $('#btn-cancel');
  const progressEl = $('#progress');
  const progressStatus = $('#progress-status');
  const progressCount = $('#progress-count');
  const progressFill = $('#progress-fill');
  const errorBanner = $('#error-banner');
  const customDates = $('#custom-dates');
  const dateFrom = $('#date-from');
  const dateTo = $('#date-to');

  // Set default date range for custom inputs
  const today = new Date().toISOString().split('T')[0];
  const weekAgo = new Date(Date.now() - 7 * 86400000).toISOString().split('T')[0];
  dateTo.value = today;
  dateFrom.value = weekAgo;

  // Show/hide custom date inputs
  $$('input[name="period"]').forEach((radio) => {
    radio.addEventListener('change', () => {
      customDates.classList.toggle('hidden', radio.value !== 'custom' || !radio.checked);
    });
  });

  let activeTabId = null;

  function showError(msg) {
    errorBanner.textContent = msg;
    errorBanner.classList.remove('hidden');
  }

  function hideError() {
    errorBanner.classList.add('hidden');
  }

  function getSelectedValue(name) {
    const checked = $(`input[name="${name}"]:checked`);
    return checked ? checked.value : null;
  }

  function calcTargetDate(period) {
    const now = new Date();
    switch (period) {
      case '24h': return new Date(now.getTime() - 24 * 60 * 60 * 1000);
      case '7d':  return new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      case '30d': return new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
      case 'custom': {
        const from = dateFrom.value;
        if (!from) return null;
        return new Date(from + 'T00:00:00');
      }
      default: return null;
    }
  }

  function calcEndDate(period) {
    if (period === 'custom') {
      const to = dateTo.value;
      if (!to) return null;
      return new Date(to + 'T23:59:59.999');
    }
    return new Date();
  }

  function setExporting(active) {
    btnExport.disabled = active;
    btnCancel.classList.toggle('hidden', !active);
    progressEl.classList.toggle('hidden', !active);
    if (active) {
      progressStatus.textContent = '収集中...';
      progressCount.textContent = '0 件';
      progressFill.classList.remove('determinate');
      progressFill.style.width = '';
    }
  }

  // Export button
  btnExport.addEventListener('click', async () => {
    hideError();

    const period = getSelectedValue('period');
    const format = getSelectedValue('format');
    const includeReactions = $('#include-reactions').checked;
    const includeSystem = $('#include-system').checked;

    // Validate custom dates
    if (period === 'custom') {
      if (!dateFrom.value) {
        showError('開始日を入力してください。');
        return;
      }
      if (!dateTo.value) {
        showError('終了日を入力してください。');
        return;
      }
      if (dateFrom.value > dateTo.value) {
        showError('開始日は終了日より前にしてください。');
        return;
      }
    }

    const targetDate = calcTargetDate(period);
    const endDate = calcEndDate(period);
    if (!targetDate) {
      showError('日付が正しくありません。');
      return;
    }

    // Get active tab (URL check is skipped — activeTab permission may not expose tab.url)
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab) {
        showError('アクティブなタブが見つかりません。');
        return;
      }
      activeTabId = tab.id;
    } catch (e) {
      showError('タブ情報の取得に失敗しました。');
      return;
    }

    setExporting(true);

    try {
      const response = await chrome.tabs.sendMessage(activeTabId, {
        type: 'startExport',
        payload: {
          targetDate: targetDate.toISOString(),
          endDate: endDate.toISOString(),
          format,
          includeReactions,
          includeSystem,
        },
      });

      if (response && response.error) {
        showError(response.error);
        setExporting(false);
      }
    } catch (e) {
      // Content script が注入されていない = Teams ページではない
      showError('Teams のページを開いた状態で実行してください。ページを再読み込みすると解決する場合があります。');
      setExporting(false);
    }
  });

  // Cancel button
  btnCancel.addEventListener('click', async () => {
    if (activeTabId) {
      try {
        await chrome.tabs.sendMessage(activeTabId, { type: 'cancelExport' });
      } catch (_) {
        // ignore
      }
    }
    setExporting(false);
  });

  // Listen for progress and completion messages from content script
  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg.type === 'exportProgress') {
      progressCount.textContent = `${msg.count} 件`;
      progressStatus.textContent = msg.status || '収集中...';
    } else if (msg.type === 'exportComplete') {
      progressStatus.textContent = '完了！';
      progressFill.classList.add('determinate');
      progressFill.style.width = '100%';
      progressCount.textContent = `${msg.count} 件`;

      if (msg.count === 0) {
        showError('指定期間内のメッセージが見つかりませんでした。');
        setExporting(false);
        return;
      }

      // Trigger download
      const blob = new Blob([msg.data], { type: msg.mimeType });
      const url = URL.createObjectURL(blob);
      chrome.downloads.download({
        url,
        filename: msg.filename,
        saveAs: true,
      }, () => {
        URL.revokeObjectURL(url);
        setTimeout(() => setExporting(false), 1500);
      });
    } else if (msg.type === 'exportError') {
      showError(msg.error);
      setExporting(false);
    }
    sendResponse({ received: true });
  });
})();
