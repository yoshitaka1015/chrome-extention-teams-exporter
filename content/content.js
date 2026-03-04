/**
 * メインオーケストレーター
 * popup からのメッセージを受信し、スクロール・収集・フォーマット・通知を制御する
 */
window.TeamsExporter = window.TeamsExporter || {};

(() => {
  const Scroller = () => window.TeamsExporter.Scroller;
  const Formatter = () => window.TeamsExporter.Formatter;

  let abortController = null;

  /**
   * popup へメッセージを送信する (popup が閉じていてもエラーにしない)
   */
  function sendToPopup(msg) {
    try {
      chrome.runtime.sendMessage(msg);
    } catch (_) {
      // popup が閉じられた場合は無視
    }
  }

  /**
   * ファイル名を生成する
   */
  function generateFilename(extension) {
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);
    const timeStr = now.toTimeString().slice(0, 5).replace(':', '');
    return `teams_export_${dateStr}_${timeStr}.${extension}`;
  }

  /**
   * メッセージを期間でフィルタリングする
   */
  function filterByDateRange(messages, targetDate, endDate) {
    const target = new Date(targetDate);
    const end = new Date(endDate);

    return messages.filter((msg) => {
      if (!msg.timestampDate) return true; // タイムスタンプ不明は含める
      return msg.timestampDate >= target && msg.timestampDate <= end;
    });
  }

  /**
   * エクスポートを実行する
   */
  async function runExport(payload) {
    const {
      targetDate,
      endDate,
      format,
      includeReactions,
      includeSystem,
    } = payload;

    // 既に実行中ならキャンセル
    if (abortController) {
      abortController.abort();
    }
    abortController = new AbortController();
    const { signal } = abortController;

    try {
      // メッセージ収集
      const messagesMap = await Scroller().collectMessages({
        targetDate: new Date(targetDate),
        signal,
        onProgress: (count) => {
          sendToPopup({
            type: 'exportProgress',
            count,
            status: '収集中...',
          });
        },
      });

      if (signal.aborted) return;

      // Map → 配列
      let messages = Array.from(messagesMap.values());

      // 期間フィルタリング
      sendToPopup({ type: 'exportProgress', count: messages.length, status: 'フィルタリング中...' });
      messages = filterByDateRange(messages, targetDate, endDate);

      // フォーマット
      sendToPopup({ type: 'exportProgress', count: messages.length, status: 'フォーマット中...' });
      const result = Formatter().format(messages, format, {
        includeReactions,
        includeSystem,
      });

      const filename = generateFilename(result.extension);

      // 完了通知
      sendToPopup({
        type: 'exportComplete',
        count: messages.length,
        data: result.data,
        mimeType: result.mimeType,
        filename,
      });
    } catch (err) {
      if (signal.aborted) return;
      sendToPopup({
        type: 'exportError',
        error: err.message || 'エクスポート中にエラーが発生しました。',
      });
    } finally {
      abortController = null;
    }
  }

  // popup からのメッセージを受信
  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg.type === 'startExport') {
      // Teams ページかチェック
      if (!document.querySelector(window.TeamsExporter.SELECTORS.SCROLL_CONTAINER)) {
        sendResponse({ error: 'メッセージペインが見つかりません。チャットまたはチャネルを開いてください。' });
        return true;
      }

      sendResponse({ ok: true });
      runExport(msg.payload);
      return true;
    }

    if (msg.type === 'cancelExport') {
      if (abortController) {
        abortController.abort();
        abortController = null;
      }
      sendResponse({ ok: true });
      return true;
    }

    return false;
  });
})();
