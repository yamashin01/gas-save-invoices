/**
 * 請求書自動管理システム - ドライブ操作
 */

/**
 * PDFファイルをドライブに保存
 * @param {Blob} pdfBlob - PDFのBlob
 * @param {string} fileName - 保存するファイル名
 * @returns {File} 保存されたファイル
 */
function savePdfToDrive(pdfBlob, fileName) {
  const folder = getDriveFolder();
  const file = folder.createFile(pdfBlob);
  file.setName(fileName);

  console.log(`ファイル保存: ${fileName}`);
  return file;
}

/**
 * ファイルへのリンクを取得
 * @param {File} file - ドライブファイル
 * @returns {string} ファイルへのURL
 */
function getFileLink(file) {
  return file.getUrl();
}

/**
 * 請求書ファイル名を生成（Phase 1: 金額なし）
 * @param {Date} date - 受信日
 * @param {string} company - 会社名
 * @returns {string} ファイル名
 */
function generateInvoiceFileName(date, company) {
  const dateStr = formatDateYYYYMMDD(date);
  const safeCompany = sanitizeFileName(company);
  return `${dateStr}_${safeCompany}.pdf`;
}

/**
 * 請求書ファイル名を生成（Phase 2: 金額あり）
 * @param {Date} date - 受信日
 * @param {string} company - 会社名
 * @param {number} amount - 金額
 * @returns {string} ファイル名
 */
function generateInvoiceFileNameWithAmount(date, company, amount) {
  const dateStr = formatDateYYYYMMDD(date);
  const safeCompany = sanitizeFileName(company);
  const amountStr = amount.toLocaleString('ja-JP');
  return `${dateStr}_${safeCompany}_${amountStr}円.pdf`;
}