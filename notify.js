/**
 * 請求書自動管理システム - 通知機能
 */

/**
 * 処理結果を通知
 * @param {Object} result - 処理結果
 */
function sendResultNotification(result) {
  if (!NOTIFICATION_EMAIL) {
    console.log('通知メールアドレスが設定されていないため、通知をスキップします');
    return;
  }

  const subject = result.hasErrors 
    ? `【要確認】請求書自動処理 - エラーあり (${result.errorCount}件)`
    : `【完了】請求書自動処理 - ${result.successCount}件処理`;

  const body = buildNotificationBody(result);

  GmailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
  console.log(`通知メールを送信しました: ${NOTIFICATION_EMAIL}`);
}

/**
 * 通知メール本文を構築
 * @param {Object} result - 処理結果
 * @returns {string} メール本文
 */
function buildNotificationBody(result) {
  const lines = [
    '請求書自動処理が完了しました。',
    '',
    `処理日時: ${formatDateTime(new Date())}`,
    `対象期間: ${formatDateDisplay(result.startDate)} 〜 ${formatDateDisplay(result.endDate)}`,
    '',
    `処理成功: ${result.successCount}件`,
    `スキップ（処理済み）: ${result.skippedCount}件`,
    `エラー: ${result.errorCount}件`,
    ''
  ];

  if (result.hasErrors) {
    lines.push('--- エラー詳細 ---');
    result.errors.forEach((error, index) => {
      lines.push(`${index + 1}. ${error.target}`);
      lines.push(`   ${error.message}`);
    });
    lines.push('');
    lines.push('エラーログシートを確認してください。');
  }

  if (result.successCount > 0) {
    lines.push('');
    lines.push('処理した請求書は「請求書一覧」シートで確認できます。');
    lines.push('ステータスが「要確認」の項目は金額を確認・入力してください。');
    lines.push('（金額抽出の信頼度が高い場合は「完了」になっています）');
  }

  return lines.join('\n');
}

/**
 * エラー通知を送信（即時通知用）
 * @param {Error} error - エラーオブジェクト
 */
function sendErrorNotification(error) {
  if (!NOTIFICATION_EMAIL) {
    console.error('通知メールアドレスが未設定。エラー:', error.message);
    return;
  }

  const subject = '【エラー】請求書自動処理で問題が発生しました';
  const body = [
    '請求書自動処理でエラーが発生しました。',
    '',
    `発生日時: ${formatDateTime(new Date())}`,
    `エラー内容: ${error.message}`,
    '',
    'スクリプトの設定を確認してください。'
  ].join('\n');

  GmailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
  console.log('エラー通知メールを送信しました');
}