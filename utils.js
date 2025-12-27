/**
 * 請求書自動管理システム - ユーティリティ
 */

/**
 * 日付をyyyymmdd形式にフォーマット
 * @param {Date} date - 日付オブジェクト
 * @returns {string} yyyymmdd形式の文字列
 */
function formatDateYYYYMMDD(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

/**
 * 日付をyyyy/MM/dd形式にフォーマット
 * @param {Date} date - 日付オブジェクト
 * @returns {string} yyyy/MM/dd形式の文字列
 */
function formatDateDisplay(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

/**
 * 日時をyyyy/MM/dd HH:mm形式にフォーマット
 * @param {Date} date - 日付オブジェクト
 * @returns {string} yyyy/MM/dd HH:mm形式の文字列
 */
function formatDateTime(date) {
  const dateStr = formatDateDisplay(date);
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${dateStr} ${hours}:${minutes}`;
}

/**
 * 前月の開始日と終了日を取得
 * @returns {{start: Date, end: Date}} 前月の開始日と終了日
 */
function getLastMonthRange() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const end = new Date(now.getFullYear(), now.getMonth(), 0); // 前月末日
  return { start, end };
}

/**
 * Gmail検索用の日付フォーマット（yyyy/MM/dd）
 * @param {Date} date - 日付オブジェクト
 * @returns {string} yyyy/MM/dd形式の文字列
 */
function formatDateForGmail(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

/**
 * ファイル名として安全な文字列に変換
 * @param {string} str - 元の文字列
 * @returns {string} ファイル名として安全な文字列
 */
function sanitizeFileName(str) {
  return str.replace(/[\\/:*?"<>|]/g, '_');
}