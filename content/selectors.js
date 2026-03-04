/**
 * Teams DOM セレクタ定数
 * Teams の DOM 構造が変更された場合、このファイルのみ修正すればよい
 */
window.TeamsExporter = window.TeamsExporter || {};

window.TeamsExporter.SELECTORS = {
  // スクロール・リスト
  SCROLL_CONTAINER: '[data-tid="message-pane-list-viewport"]',
  MESSAGE_LIST: '#chat-pane-list',
  CHAT_ITEM: '[data-tid="chat-pane-item"]',

  // メッセージ
  MESSAGE: '[data-tid="chat-pane-message"]',
  MESSAGE_ID_ATTR: 'data-mid',
  MY_MESSAGE_BODY: '.fui-ChatMyMessage__body',
  OTHER_MESSAGE_BODY: '.fui-ChatMessage__body',

  // メッセージ内容
  AUTHOR_NAME: '[data-tid="message-author-name"]',
  TIMESTAMP: 'time[datetime]',
  MESSAGE_CONTENT: '[id^="content-"][data-message-content]',
  EMOTICON: '[data-tid="emoticon-renderer"] img',
  FILE_ATTACHMENT: '[data-tid="file-attachment-grid"]',

  // システムメッセージ
  SYSTEM_MESSAGE: '[role="heading"][aria-level="4"]',

  // 仮想スクロール
  VIRTUAL_LIST_LOADER: '[data-testid="virtual-list-loader"]',
  PLACEHOLDERS: '[data-testid="vl-placeholders"]',

  // リアクション (推測ベース、動作確認時に調整)
  REACTION_BUTTON: 'button[aria-label*="リアクション"], button[aria-label*="reaction"]',
  REACTION_CONTAINER: '[class*="reaction"]',
};
