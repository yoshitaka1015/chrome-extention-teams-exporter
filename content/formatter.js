/**
 * 出力フォーマッター
 * 収集したメッセージを Markdown / CSV / JSON 形式に変換する
 */
window.TeamsExporter = window.TeamsExporter || {};

window.TeamsExporter.Formatter = (() => {

  /**
   * メッセージを時系列順にソートする
   */
  function sortMessages(messages) {
    return [...messages].sort((a, b) => {
      const ta = a.timestampDate ? a.timestampDate.getTime() : 0;
      const tb = b.timestampDate ? b.timestampDate.getTime() : 0;
      return ta - tb;
    });
  }

  /**
   * 日付文字列 (YYYY-MM-DD) を取得する
   */
  function toDateStr(date) {
    if (!date) return '不明';
    return date.toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' });
  }

  /**
   * 時刻文字列 (HH:MM) を取得する
   */
  function toTimeStr(date) {
    if (!date) return '';
    return date.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' });
  }

  // --- Markdown ---

  function toMarkdown(messages, options) {
    const sorted = sortMessages(messages);
    const lines = [];
    lines.push('# Teams チャットエクスポート');
    lines.push(`エクスポート日時: ${new Date().toLocaleString('ja-JP')}`);
    lines.push(`メッセージ数: ${sorted.length}`);
    lines.push('');

    let currentDate = '';

    for (const msg of sorted) {
      const dateStr = msg.timestampDate ? toDateStr(msg.timestampDate) : '不明';

      if (dateStr !== currentDate) {
        currentDate = dateStr;
        lines.push(`## ${currentDate}`);
        lines.push('');
      }

      if (msg.type === 'system' && !options.includeSystem) continue;

      const time = msg.timestampDate ? toTimeStr(msg.timestampDate) : '';

      if (msg.type === 'system') {
        lines.push(`*--- ${msg.text} ---*`);
        lines.push('');
        continue;
      }

      lines.push(`**${msg.author}** (${time})`);
      if (msg.text) {
        lines.push(msg.text);
      }

      // 添付ファイル
      if (msg.attachments.length > 0) {
        for (const att of msg.attachments) {
          lines.push(`📎 [${att.name}](${att.url || '#'})`);
        }
      }

      // リアクション
      if (options.includeReactions && msg.reactions.length > 0) {
        const reactionText = msg.reactions.map((r) => `${r.label} (${r.count})`).join(', ');
        lines.push(`> リアクション: ${reactionText}`);
      }

      lines.push('');
    }

    return lines.join('\n');
  }

  // --- CSV ---

  function escapeCsvField(value) {
    const str = String(value || '');
    if (str.includes('"') || str.includes(',') || str.includes('\n') || str.includes('\r')) {
      return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  }

  function toCsv(messages, options) {
    const sorted = sortMessages(messages);
    const BOM = '\uFEFF';
    const headers = ['日時', '投稿者', 'メッセージ', '種別', 'リアクション', '添付ファイル'];
    const rows = [headers.map(escapeCsvField).join(',')];

    for (const msg of sorted) {
      if (msg.type === 'system' && !options.includeSystem) continue;

      const datetime = msg.timestampDate
        ? msg.timestampDate.toLocaleString('ja-JP')
        : '';

      const reactions = options.includeReactions
        ? msg.reactions.map((r) => `${r.label}(${r.count})`).join('; ')
        : '';

      const attachments = msg.attachments.map((a) => a.name).join('; ');

      const row = [
        datetime,
        msg.author,
        msg.text,
        msg.type === 'system' ? 'システム' : 'メッセージ',
        reactions,
        attachments,
      ];

      rows.push(row.map(escapeCsvField).join(','));
    }

    return BOM + rows.join('\r\n');
  }

  // --- JSON ---

  function toJson(messages, options) {
    const sorted = sortMessages(messages);

    const filtered = sorted.filter((msg) => {
      if (msg.type === 'system' && !options.includeSystem) return false;
      return true;
    });

    const output = {
      exportedAt: new Date().toISOString(),
      messageCount: filtered.length,
      messages: filtered.map((msg) => {
        const entry = {
          id: msg.mid,
          timestamp: msg.timestamp,
          author: msg.author,
          text: msg.text,
          type: msg.type,
          isMyMessage: msg.isMyMessage,
        };

        if (msg.attachments.length > 0) {
          entry.attachments = msg.attachments;
        }

        if (options.includeReactions && msg.reactions.length > 0) {
          entry.reactions = msg.reactions;
        }

        return entry;
      }),
    };

    return JSON.stringify(output, null, 2);
  }

  // --- Public ---

  /**
   * メッセージをフォーマットする
   * @param {object[]} messages - メッセージ配列
   * @param {string} format - "markdown" | "csv" | "json"
   * @param {object} options - { includeReactions, includeSystem }
   * @returns {{ data: string, mimeType: string, extension: string }}
   */
  function format(messages, formatType, options) {
    switch (formatType) {
      case 'markdown':
        return {
          data: toMarkdown(messages, options),
          mimeType: 'text/markdown; charset=utf-8',
          extension: 'md',
        };
      case 'csv':
        return {
          data: toCsv(messages, options),
          mimeType: 'text/csv; charset=utf-8',
          extension: 'csv',
        };
      case 'json':
        return {
          data: toJson(messages, options),
          mimeType: 'application/json; charset=utf-8',
          extension: 'json',
        };
      default:
        throw new Error(`未対応の出力形式: ${formatType}`);
    }
  }

  return { format };
})();
