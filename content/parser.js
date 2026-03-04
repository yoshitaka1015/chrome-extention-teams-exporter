/**
 * メッセージパーサー
 * DOM 要素からプレーンオブジェクトとしてメッセージ情報を抽出する
 */
window.TeamsExporter = window.TeamsExporter || {};

window.TeamsExporter.Parser = (() => {
  const S = () => window.TeamsExporter.SELECTORS;

  /**
   * メッセージ本文のテキストを取得する
   * エモーティコンは alt 属性からテキスト化する
   */
  function extractMessageText(contentEl) {
    if (!contentEl) return '';

    const clone = contentEl.cloneNode(true);

    // エモーティコンの img を alt テキストに置換
    clone.querySelectorAll(S().EMOTICON).forEach((img) => {
      const alt = img.getAttribute('alt') || '';
      img.replaceWith(alt);
    });

    return clone.textContent.trim();
  }

  /**
   * 添付ファイル情報を取得する
   */
  function extractAttachments(messageEl) {
    const grid = messageEl.querySelector(S().FILE_ATTACHMENT);
    if (!grid) return [];

    const attachments = [];
    // ファイル名はリンクテキストやaria-labelから取得
    grid.querySelectorAll('a, [aria-label]').forEach((el) => {
      const name = el.textContent.trim() || el.getAttribute('aria-label') || '';
      if (name) {
        attachments.push({
          name,
          url: el.href || '',
        });
      }
    });

    return attachments;
  }

  /**
   * リアクション情報を取得する (推測ベース)
   */
  function extractReactions(messageEl) {
    const reactions = [];

    // リアクションコンテナを探索
    const containers = messageEl.querySelectorAll(S().REACTION_CONTAINER);
    containers.forEach((container) => {
      const btn = container.querySelector('button');
      if (!btn) return;

      const label = btn.getAttribute('aria-label') || btn.textContent.trim();
      const countEl = container.querySelector('span');
      const count = countEl ? parseInt(countEl.textContent, 10) || 1 : 1;

      if (label) {
        reactions.push({ label, count });
      }
    });

    return reactions;
  }

  /**
   * 単一のメッセージ要素をパースする
   * @param {Element} messageEl - [data-tid="chat-pane-message"] 要素
   * @returns {object|null} パース結果のプレーンオブジェクト
   */
  function parseMessage(messageEl) {
    const mid = messageEl.getAttribute(S().MESSAGE_ID_ATTR);
    if (!mid) return null;

    // タイムスタンプ
    const timeEl = messageEl.querySelector(S().TIMESTAMP);
    const datetime = timeEl ? timeEl.getAttribute('datetime') : null;
    const timestamp = datetime ? new Date(datetime) : null;

    // 投稿者
    const authorEl = messageEl.querySelector(S().AUTHOR_NAME);
    const author = authorEl ? authorEl.textContent.trim() : '';

    // 自分のメッセージか判定
    const isMyMessage = !!messageEl.querySelector(S().MY_MESSAGE_BODY);

    // 本文
    const contentEl = messageEl.querySelector(S().MESSAGE_CONTENT);
    const text = extractMessageText(contentEl);

    // 添付ファイル
    const attachments = extractAttachments(messageEl);

    // リアクション
    const reactions = extractReactions(messageEl);

    return {
      mid,
      timestamp: timestamp ? timestamp.toISOString() : null,
      timestampDate: timestamp,
      author: author || (isMyMessage ? '自分' : '不明'),
      text,
      isMyMessage,
      attachments,
      reactions,
      type: 'message',
    };
  }

  /**
   * システムメッセージをパースする
   * @param {Element} el - [role="heading"][aria-level="4"] 要素
   * @returns {object|null}
   */
  function parseSystemMessage(el) {
    const text = el.textContent.trim();
    if (!text) return null;

    // 前後のメッセージからタイムスタンプを推測
    const parent = el.closest(S().CHAT_ITEM);
    let timestamp = null;
    if (parent) {
      const timeEl = parent.querySelector(S().TIMESTAMP);
      if (timeEl) {
        const dt = timeEl.getAttribute('datetime');
        if (dt) timestamp = new Date(dt);
      }
    }

    return {
      mid: 'sys_' + text.substring(0, 40).replace(/\s+/g, '_') + '_' + (timestamp ? timestamp.getTime() : Date.now()),
      timestamp: timestamp ? timestamp.toISOString() : null,
      timestampDate: timestamp,
      author: 'システム',
      text,
      isMyMessage: false,
      attachments: [],
      reactions: [],
      type: 'system',
    };
  }

  /**
   * 現在表示されている全メッセージをパースして返す
   * @returns {object[]} メッセージオブジェクトの配列
   */
  function parseVisibleMessages() {
    const results = [];

    // 通常メッセージ
    document.querySelectorAll(S().MESSAGE).forEach((el) => {
      const msg = parseMessage(el);
      if (msg) results.push(msg);
    });

    // システムメッセージ
    document.querySelectorAll(S().SYSTEM_MESSAGE).forEach((el) => {
      const msg = parseSystemMessage(el);
      if (msg) results.push(msg);
    });

    return results;
  }

  return {
    parseMessage,
    parseSystemMessage,
    parseVisibleMessages,
  };
})();
