/**
 * 請求書自動管理システム - デバッグ・テスト用
 */

/**
 * 設定確認用: スクリプトプロパティを表示
 */
function debugShowProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  console.log('=== スクリプトプロパティ ===');
  for (const [key, value] of Object.entries(props)) {
    // 値は一部マスク
    const maskedValue = value.length > 10 
      ? `${value.substring(0, 5)}...${value.substring(value.length - 5)}`
      : value;
    console.log(`${key}: ${maskedValue}`);
  }
}

/**
 * 設定確認用: 送信元一覧を表示
 */
function debugShowSenders() {
  try {
    const senders = getEnabledSenders();
    console.log('=== 有効な送信元一覧 ===');
    senders.forEach((sender, index) => {
      console.log(`${index + 1}. ${sender.company}`);
      console.log(`   Email: ${sender.email}`);
      console.log(`   Keyword: ${sender.keyword}`);
    });
  } catch (error) {
    console.error(`エラー: ${error.message}`);
  }
}

/**
 * テスト用: 特定期間のメールを検索（処理はしない）
 * @param {string} email - 送信元メールアドレス
 * @param {string} keyword - 検索キーワード
 * @param {number} year - 年
 * @param {number} month - 月（1-12）
 */
function debugSearchEmails(email, keyword, year, month) {
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0);

  const sender = { email, keyword, company: 'テスト' };
  const query = buildSearchQuery(sender, startDate, endDate);

  console.log('=== メール検索テスト ===');
  console.log(`期間: ${formatDateDisplay(startDate)} 〜 ${formatDateDisplay(endDate)}`);
  console.log(`クエリ: ${query}`);

  const messages = searchInvoiceEmails(sender, startDate, endDate);

  console.log(`結果: ${messages.length}件`);
  messages.forEach((msg, index) => {
    const info = getMessageInfo(msg);
    const attachments = msg.getAttachments();
    const pdfCount = attachments.filter(a => 
      a.getContentType() === 'application/pdf' || 
      a.getName().toLowerCase().endsWith('.pdf')
    ).length;

    console.log(`${index + 1}. ${info.subject}`);
    console.log(`   日付: ${formatDateDisplay(info.date)}`);
    console.log(`   添付: ${attachments.length}件 (PDF: ${pdfCount}件)`);
  });
}

/**
 * テスト用: 前月のメールを検索（設定シートの全送信元）
 */
function debugSearchLastMonth() {
  const { start, end } = getLastMonthRange();
  console.log('=== 前月メール検索テスト ===');
  console.log(`期間: ${formatDateDisplay(start)} 〜 ${formatDateDisplay(end)}`);

  const senders = getEnabledSenders();
  let totalCount = 0;

  for (const sender of senders) {
    console.log(`--- ${sender.company} ---`);
    const messages = searchInvoiceEmails(sender, start, end);
    totalCount += messages.length;

    messages.forEach((msg, index) => {
      const info = getMessageInfo(msg);
      const pdfAttachments = getPdfAttachments(msg);
      console.log(`  ${index + 1}. ${info.subject} (PDF: ${pdfAttachments.length}件)`);
    });
  }

  console.log(`合計: ${totalCount}件`);
}

/**
 * テスト用: 単一メッセージの処理テスト（実際に保存する）
 * @param {string} messageId - Gmail Message ID
 * @param {string} company - 会社名
 */
function debugProcessSingleMessage(messageId, company) {
  console.log('=== 単一メッセージ処理テスト ===');
  
  const message = GmailApp.getMessageById(messageId);
  if (!message) {
    console.error('メッセージが見つかりません');
    return;
  }

  const messageInfo = getMessageInfo(message);
  console.log(`件名: ${messageInfo.subject}`);
  console.log(`日付: ${formatDateDisplay(messageInfo.date)}`);

  const sender = { company };

  try {
    processMessage(message, sender, messageInfo);
    console.log('処理完了');
  } catch (error) {
    console.error(`エラー: ${error.message}`);
  }
}

/**
 * テスト用: ドライフォルダの確認
 */
function debugCheckDriveFolder() {
  try {
    const folder = getDriveFolder();
    console.log('=== 保存先フォルダ ===');
    console.log(`名前: ${folder.getName()}`);
    console.log(`URL: ${folder.getUrl()}`);
    
    const files = folder.getFiles();
    let count = 0;
    console.log('--- 既存ファイル（最新10件）---');
    while (files.hasNext() && count < 10) {
      const file = files.next();
      console.log(`  ${file.getName()}`);
      count++;
    }
  } catch (error) {
    console.error(`エラー: ${error.message}`);
  }
}

/**
 * スクリプトプロパティを設定するヘルパー
 * （GUIで設定する代わりにコードから設定する場合）
 */
function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  
  // 以下の値を実際のIDに置き換えてから実行
  // props.setProperty('DRIVE_FOLDER_ID', 'your-folder-id-here');
  // props.setProperty('NOTIFICATION_EMAIL', 'your-email@example.com');
  // props.setProperty('GEMINI_API_KEY', 'your-gemini-api-key-here');

  console.log('スクリプトプロパティを設定しました');
  debugShowProperties();
}

/**
 * Gemini API接続テスト
 */
function debugTestGeminiConnection() {
  console.log('=== Gemini API 接続テスト ===');
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    console.error('GEMINI_API_KEY が設定されていません');
    return;
  }
  
  console.log(`APIキー: ${apiKey.substring(0, 5)}...${apiKey.substring(apiKey.length - 5)}`);
  
  // シンプルなテストリクエスト
  const testUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${apiKey}`;
  
  const requestBody = {
    contents: [
      {
        parts: [
          { text: '1+1は？数字だけで答えて。' }
        ]
      }
    ]
  };
  
  try {
    const response = UrlFetchApp.fetch(testUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const text = response.getContentText();
    
    console.log(`レスポンスコード: ${code}`);
    
    if (code === 200) {
      const result = JSON.parse(text);
      const answer = result.candidates?.[0]?.content?.parts?.[0]?.text;
      console.log(`応答: ${answer}`);
      console.log('接続成功！');
    } else {
      console.error(`エラー: ${text}`);
    }
  } catch (error) {
    console.error(`接続エラー: ${error.message}`);
  }
}

/**
 * 指定したDriveファイルで金額抽出をテスト
 * GASエディタから実行する場合は引数なしで実行し、
 * ファイルIDをコード内で指定してください
 */
function debugTestExtractAmountFromDrive() {
  // テストしたいPDFファイルのIDを指定
  const fileId = '1-nuKdFZENXcX3stqWWucnferJSe6WUCL';

  if (fileId === 'YOUR_PDF_FILE_ID_HERE') {
    console.error('fileId を実際のPDFファイルIDに置き換えてください');
    return;
  }
  
  debugExtractAmountFromFile(fileId);
}