/**
 * 自動スクロール制御
 * 仮想スクロールされた Teams チャットを上方向にスクロールしながらメッセージを収集する
 */
window.TeamsExporter = window.TeamsExporter || {};

window.TeamsExporter.Scroller = (() => {
  const S = () => window.TeamsExporter.SELECTORS;
  const Parser = () => window.TeamsExporter.Parser;

  const SCROLL_STEP = 800;        // 1回あたりのスクロール量 (px)
  const SCROLL_WAIT = 600;        // スクロール後の基本待機時間 (ms)
  const LOADER_WAIT = 4500;       // ローダー検出時の追加待機上限 (ms)
  const LOADER_CHECK_INTERVAL = 300;
  const STALL_THRESHOLD = 5;      // 新メッセージなし許容回数
  const MAX_MESSAGES = 50000;     // メモリ保護用上限

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * ローダーが消えるまで待機する
   */
  async function waitForLoader(signal) {
    const start = Date.now();
    while (Date.now() - start < LOADER_WAIT) {
      if (signal && signal.aborted) return;
      const loader = document.querySelector(S().VIRTUAL_LIST_LOADER);
      const placeholders = document.querySelector(S().PLACEHOLDERS);
      if (!loader && !placeholders) return;
      await sleep(LOADER_CHECK_INTERVAL);
    }
  }

  /**
   * メッセージを収集しながら上方向にスクロールする
   * @param {object} options
   * @param {Date} options.targetDate - ここまで遡る目標日付
   * @param {AbortSignal} options.signal - キャンセル用シグナル
   * @param {function} options.onProgress - 進捗コールバック (count)
   * @returns {Map<string, object>} mid → メッセージオブジェクトの Map
   */
  async function collectMessages({ targetDate, signal, onProgress }) {
    const container = document.querySelector(S().SCROLL_CONTAINER);
    if (!container) {
      throw new Error('メッセージペインが見つかりません。チャットまたはチャネルを開いてください。');
    }

    const messages = new Map();
    let stallCount = 0;

    // まず現在表示されているメッセージを収集
    collectVisible(messages);
    if (onProgress) onProgress(messages.size);

    while (true) {
      if (signal && signal.aborted) break;

      // 上限チェック
      if (messages.size >= MAX_MESSAGES) break;

      const prevSize = messages.size;
      const prevScrollTop = container.scrollTop;

      // 上方向にスクロール
      container.scrollTop = Math.max(0, container.scrollTop - SCROLL_STEP);

      // 基本待機
      await sleep(SCROLL_WAIT);
      if (signal && signal.aborted) break;

      // ローダーが表示されていれば追加待機
      await waitForLoader(signal);
      if (signal && signal.aborted) break;

      // 現在表示されているメッセージを収集
      collectVisible(messages);

      if (onProgress) onProgress(messages.size);

      // 目標日付チェック
      if (hasReachedTarget(messages, targetDate)) break;

      // スクロール不能チェック (最上部到達)
      if (container.scrollTop === 0 && prevScrollTop === 0) {
        // ローダーがなければ完全に最上部
        const loader = document.querySelector(S().VIRTUAL_LIST_LOADER);
        if (!loader) break;
      }

      // 新メッセージなしカウント
      if (messages.size === prevSize) {
        stallCount++;
        if (stallCount >= STALL_THRESHOLD) break;
      } else {
        stallCount = 0;
      }
    }

    return messages;
  }

  /**
   * 現在表示されているメッセージを Map に追加する
   */
  function collectVisible(messages) {
    const parsed = Parser().parseVisibleMessages();
    for (const msg of parsed) {
      if (!messages.has(msg.mid)) {
        messages.set(msg.mid, msg);
      }
    }
  }

  /**
   * 目標日付より古いメッセージがあるか判定する
   */
  function hasReachedTarget(messages, targetDate) {
    for (const msg of messages.values()) {
      if (msg.timestampDate && msg.timestampDate < targetDate) {
        return true;
      }
    }
    return false;
  }

  return {
    collectMessages,
  };
})();
