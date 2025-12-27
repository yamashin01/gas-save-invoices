/**
 * 請求書自動管理システム - Gmail操作
 */

/**
 * 検索クエリを構築
 * @param {Object} sender - 送信元設定
 * @param {Date} startDate - 検索開始日
 * @param {Date} endDate - 検索終了日
 * @returns {string} Gmail検索クエリ
 */
function buildSearchQuery(sender, startDate, endDate) {
  const parts = [
    `from:${sender.email}`,
    `subject:${sender.keyword}`,
    `after:${formatDateForGmail(startDate)}`,
    `before:${formatDateForGmail(new Date(endDate.getTime() + 86400000))}`, // 終了日の翌日
    'has:attachment'
  ];
  return parts.join(' ');
}

/**
 * 指定条件でメールを検索
 * @param {Object} sender - 送信元設定
 * @param {Date} startDate - 検索開始日
 * @param {Date} endDate - 検索終了日
 * @returns {Array<GmailMessage>} メッセージの配列
 */
function searchInvoiceEmails(sender, startDate, endDate) {
  const query = buildSearchQuery(sender, startDate, endDate);
  console.log(`検索クエリ: ${query}`);

  const threads = GmailApp.search(query);
  console.log(`検索結果: ${threads.length}スレッド`);

  // 全スレッドからメッセージを取得
  const messages = threads.flatMap(thread => thread.getMessages());
  console.log(`メッセージ数: ${messages.length}件`);

  return messages;
}

/**
 * メッセージからPDF添付ファイルを取得
 * @param {GmailMessage} message - Gmailメッセージ
 * @returns {Array<GmailAttachment>} PDF添付ファイルの配列
 */
function getPdfAttachments(message) {
  const attachments = message.getAttachments();
  const pdfAttachments = attachments.filter(att => 
    att.getContentType() === 'application/pdf' || 
    att.getName().toLowerCase().endsWith('.pdf')
  );

  console.log(`添付ファイル: ${attachments.length}件, うちPDF: ${pdfAttachments.length}件`);
  return pdfAttachments;
}

/**
 * メッセージ情報を取得
 * @param {GmailMessage} message - Gmailメッセージ
 * @returns {Object} メッセージ情報
 */
function getMessageInfo(message) {
  return {
    id: message.getId(),
    subject: message.getSubject(),
    from: message.getFrom(),
    date: message.getDate()
  };
}