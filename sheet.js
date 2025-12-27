/**
 * 請求書自動管理システム - スプレッドシート操作
 */

/**
 * 設定シートから有効な送信元一覧を取得
 * @returns {Array<{email: string, company: string, keyword: string}>} 送信元設定の配列
 */
function getEnabledSenders() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    throw new Error(`シート「${SHEET_NAMES.SETTINGS}」が見つかりません`);
  }

  const data = sheet.getDataRange().getValues();
  // ヘッダー行をスキップ
  const senders = data.slice(1)
    .filter(row => row[SETTINGS_COLUMNS.ENABLED] === true || row[SETTINGS_COLUMNS.ENABLED] === 'TRUE')
    .map(row => ({
      email: row[SETTINGS_COLUMNS.EMAIL],
      company: row[SETTINGS_COLUMNS.COMPANY],
      keyword: row[SETTINGS_COLUMNS.KEYWORD]
    }))
    .filter(sender => sender.email && sender.company);

  console.log(`有効な送信元: ${senders.length}件`);
  return senders;
}

/**
 * 処理済みのMessage ID一覧を取得
 * @returns {Set<string>} Message IDのセット
 */
function getProcessedMessageIds() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.INVOICE_LIST);
  if (!sheet) {
    return new Set();
  }

  const data = sheet.getDataRange().getValues();
  const messageIds = new Set(
    data.slice(1).map(row => row[INVOICE_COLUMNS.MESSAGE_ID]).filter(id => id)
  );

  console.log(`処理済みMessage ID: ${messageIds.size}件`);
  return messageIds;
}

/**
 * 請求書一覧に記録を追加
 * @param {Object} record - 追加するレコード
 */
function addInvoiceRecord(record) {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.INVOICE_LIST);
  if (!sheet) {
    throw new Error(`シート「${SHEET_NAMES.INVOICE_LIST}」が見つかりません`);
  }

  const row = [
    record.messageId,
    record.receivedDate,
    record.company,
    record.amountAuto || '',
    record.amountFinal || '',
    record.status,
    record.pdfLink,
    record.processedAt
  ];

  sheet.appendRow(row);
  console.log(`請求書一覧に追加: ${record.company} (${record.receivedDate})`);
}

/**
 * エラーログに記録を追加
 * @param {Object} error - エラー情報
 */
function addErrorLog(error) {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.ERROR_LOG);
  if (!sheet) {
    console.warn(`シート「${SHEET_NAMES.ERROR_LOG}」が見つかりません。エラーログをスキップします。`);
    return;
  }

  const row = [
    formatDateTime(new Date()),
    error.target,
    error.message,
    '未対応'
  ];

  sheet.appendRow(row);
  console.error(`エラーログに追加: ${error.target} - ${error.message}`);
}

/**
 * スプレッドシートの初期設定（シートが存在しない場合に作成）
 */
function initializeSheets() {
  const ss = getSpreadsheet();

  // 設定シート
  let settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    settingsSheet.appendRow(['送信元メールアドレス', '会社名', '検索キーワード', '有効']);
    settingsSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    console.log(`シート「${SHEET_NAMES.SETTINGS}」を作成しました`);
  }

  // 請求書一覧シート
  let invoiceSheet = ss.getSheetByName(SHEET_NAMES.INVOICE_LIST);
  if (!invoiceSheet) {
    invoiceSheet = ss.insertSheet(SHEET_NAMES.INVOICE_LIST);
    invoiceSheet.appendRow([
      'Message ID', '受信日', '請求元', '金額（自動）', '金額（確定）', 
      'ステータス', 'PDFリンク', '処理日時'
    ]);
    invoiceSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    // Message ID列を非表示にする（管理用）
    invoiceSheet.hideColumns(1);
    console.log(`シート「${SHEET_NAMES.INVOICE_LIST}」を作成しました`);
  }

  // エラーログシート
  let errorSheet = ss.getSheetByName(SHEET_NAMES.ERROR_LOG);
  if (!errorSheet) {
    errorSheet = ss.insertSheet(SHEET_NAMES.ERROR_LOG);
    errorSheet.appendRow(['発生日時', '処理対象', 'エラー内容', '対応状況']);
    errorSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    console.log(`シート「${SHEET_NAMES.ERROR_LOG}」を作成しました`);
  }

  console.log('シートの初期化が完了しました');
}