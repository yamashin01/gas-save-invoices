/**
 * 請求書自動管理システム - メイン処理
 * 
 * 【初期設定】
 * 1. スクリプトプロパティに以下を設定:
 *    - DRIVE_FOLDER_ID: PDF保存先フォルダのID
 *    - NOTIFICATION_EMAIL: 通知先メールアドレス（任意）
 * 
 * 2. initializeSheets() を実行してシートを初期化
 * 
 * 3. 「設定」シートに送信元情報を入力
 * 
 * 4. processLastMonthInvoices() を実行（または月次トリガーを設定）
 */

/**
 * 前月の請求書を処理（メインエントリーポイント）
 */
function processLastMonthInvoices() {
  console.log('=== 請求書自動処理 開始 ===');

  const { start, end } = getLastMonthRange();
  console.log(`対象期間: ${formatDateDisplay(start)} 〜 ${formatDateDisplay(end)}`);

  try {
    const result = processInvoices(start, end);
    sendResultNotification(result);
    console.log('=== 請求書自動処理 完了 ===');
  } catch (error) {
    console.error(`致命的エラー: ${error.message}`);
    sendErrorNotification(error);
    throw error;
  }
}

/**
 * 指定期間の請求書を処理
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @returns {Object} 処理結果
 */
function processInvoices(startDate, endDate) {
  const senders = getEnabledSenders();
  const processedIds = getProcessedMessageIds();

  const result = {
    startDate,
    endDate,
    successCount: 0,
    skippedCount: 0,
    errorCount: 0,
    errors: [],
    hasErrors: false
  };

  for (const sender of senders) {
    console.log(`--- 送信元: ${sender.company} (${sender.email}) ---`);

    try {
      const senderResult = processSenderInvoices(sender, startDate, endDate, processedIds);
      result.successCount += senderResult.successCount;
      result.skippedCount += senderResult.skippedCount;
      result.errorCount += senderResult.errorCount;
      result.errors.push(...senderResult.errors);
    } catch (error) {
      console.error(`送信元処理エラー: ${error.message}`);
      const errorInfo = {
        target: `${sender.company} (${sender.email})`,
        message: error.message
      };
      addErrorLog(errorInfo);
      result.errors.push(errorInfo);
      result.errorCount++;
    }
  }

  result.hasErrors = result.errorCount > 0;
  console.log(`処理結果 - 成功: ${result.successCount}, スキップ: ${result.skippedCount}, エラー: ${result.errorCount}`);

  return result;
}

/**
 * 特定の送信元の請求書を処理
 * @param {Object} sender - 送信元設定
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @param {Set<string>} processedIds - 処理済みID
 * @returns {Object} 処理結果
 */
function processSenderInvoices(sender, startDate, endDate, processedIds) {
  const result = {
    successCount: 0,
    skippedCount: 0,
    errorCount: 0,
    errors: []
  };

  const messages = searchInvoiceEmails(sender, startDate, endDate);

  for (const message of messages) {
    const messageInfo = getMessageInfo(message);

    // 重複チェック
    if (processedIds.has(messageInfo.id)) {
      console.log(`スキップ（処理済み）: ${messageInfo.subject}`);
      result.skippedCount++;
      continue;
    }

    try {
      processMessage(message, sender, messageInfo);
      processedIds.add(messageInfo.id); // 処理済みに追加
      result.successCount++;
    } catch (error) {
      console.error(`メッセージ処理エラー: ${error.message}`);
      const errorInfo = {
        target: `${sender.email} / ${formatDateDisplay(messageInfo.date)} / ${messageInfo.subject}`,
        message: error.message
      };
      addErrorLog(errorInfo);
      result.errors.push(errorInfo);
      result.errorCount++;
    }
  }

  return result;
}

/**
 * 個別メッセージを処理
 * @param {GmailMessage} message - メッセージ
 * @param {Object} sender - 送信元設定
 * @param {Object} messageInfo - メッセージ情報
 */
function processMessage(message, sender, messageInfo) {
  // PDF添付ファイルを取得
  const pdfAttachments = getPdfAttachments(message);

  if (pdfAttachments.length === 0) {
    throw new Error('PDF添付ファイルがありません');
  }

  // 最初のPDFを処理（複数ある場合は最初の1つ）
  const pdf = pdfAttachments[0];
  if (pdfAttachments.length > 1) {
    console.warn(`複数のPDF添付（${pdfAttachments.length}件）がありますが、最初の1件のみ処理します`);
  }

  // ファイル名を生成
  const fileName = generateInvoiceFileName(messageInfo.date, sender.company);

  // ドライブに保存
  const savedFile = savePdfToDrive(pdf, fileName);
  const pdfLink = getFileLink(savedFile);

  // スプレッドシートに記録
  addInvoiceRecord({
    messageId: messageInfo.id,
    receivedDate: formatDateDisplay(messageInfo.date),
    company: sender.company,
    amountAuto: '', // Phase 1では空
    amountFinal: '',
    status: STATUS.NEEDS_CHECK,
    pdfLink,
    processedAt: formatDateTime(new Date())
  });

  console.log(`処理完了: ${fileName}`);
}

/**
 * 月次トリガーを設定
 * 毎月1日の午前9時に実行
 */
function createMonthlyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processLastMonthInvoices') {
      ScriptApp.deleteTrigger(trigger);
      console.log('既存のトリガーを削除しました');
    }
  }

  // 新しいトリガーを作成
  ScriptApp.newTrigger('processLastMonthInvoices')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  console.log('月次トリガーを作成しました（毎月1日 9:00）');
}

/**
 * トリガーを削除
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  console.log(`${triggers.length}件のトリガーを削除しました`);
}