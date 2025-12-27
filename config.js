/**
 * 請求書自動管理システム - 設定・定数
 */

// シート名
const SHEET_NAMES = {
  SETTINGS: '設定',
  INVOICE_LIST: '請求書一覧',
  ERROR_LOG: 'エラーログ'
};

// 保存先フォルダID（実際のIDに置き換えてください）
const DRIVE_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID') || '';

// 通知設定
const NOTIFICATION_EMAIL = PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL') || '';

/**
 * 設定シートのカラム定義
 */
const SETTINGS_COLUMNS = {
  EMAIL: 0,      // 送信元メールアドレス
  COMPANY: 1,    // 会社名
  KEYWORD: 2,    // 検索キーワード
  ENABLED: 3     // 有効フラグ
};

/**
 * 請求書一覧シートのカラム定義
 */
const INVOICE_COLUMNS = {
  MESSAGE_ID: 0,    // Message ID
  RECEIVED_DATE: 1, // 受信日
  COMPANY: 2,       // 請求元
  AMOUNT_AUTO: 3,   // 金額（自動）
  AMOUNT_FINAL: 4,  // 金額（確定）
  STATUS: 5,        // ステータス
  PDF_LINK: 6,      // PDFリンク
  PROCESSED_AT: 7   // 処理日時
};

/**
 * ステータス定義
 */
const STATUS = {
  COMPLETED: '完了',
  NEEDS_CHECK: '要確認',
  ERROR: 'エラー'
};

/**
 * スプレッドシートを取得（紐づいたスプレッドシートを使用）
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 保存先フォルダを取得
 */
function getDriveFolder() {
  if (!DRIVE_FOLDER_ID) {
    throw new Error('DRIVE_FOLDER_ID が設定されていません。スクリプトプロパティを確認してください。');
  }
  return DriveApp.getFolderById(DRIVE_FOLDER_ID);
}