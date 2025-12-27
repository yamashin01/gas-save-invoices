/**
 * 請求書自動管理システム - 為替レート・通貨換算
 */

/**
 * USD/JPY為替レートを取得（Google Finance経由）
 * @returns {number} 為替レート
 */
function getUsdJpyRate() {
  try {
    // Google Sheetsの GOOGLEFINANCE 関数を使用
    const ss = getSpreadsheet();
    let tempSheet = ss.getSheetByName('_temp_rate');
    
    // 一時シートを作成
    if (!tempSheet) {
      tempSheet = ss.insertSheet('_temp_rate');
    }
    
    // GOOGLEFINANCE関数で為替レートを取得
    tempSheet.getRange('A1').setFormula('=GOOGLEFINANCE("CURRENCY:USDJPY")');
    SpreadsheetApp.flush(); // 計算を強制実行
    
    // 少し待ってから値を取得
    Utilities.sleep(1000);
    const rate = tempSheet.getRange('A1').getValue();
    
    // 一時シートを削除
    ss.deleteSheet(tempSheet);
    
    if (typeof rate === 'number' && rate > 0) {
      console.log(`USD/JPY為替レート: ${rate}`);
      return rate;
    }
    
    throw new Error('為替レートの取得に失敗しました');
    
  } catch (error) {
    console.error(`為替レート取得エラー: ${error.message}`);
    // フォールバック: 固定レート（要確認フラグを立てる）
    const fallbackRate = 150;
    console.warn(`フォールバックレートを使用: ${fallbackRate}`);
    return fallbackRate;
  }
}

/**
 * 為替レートをキャッシュ付きで取得
 * 同一実行内では同じレートを使用
 */
let cachedExchangeRate = null;

function getCachedUsdJpyRate() {
  if (cachedExchangeRate === null) {
    cachedExchangeRate = getUsdJpyRate();
  }
  return cachedExchangeRate;
}

/**
 * キャッシュをクリア
 */
function clearExchangeRateCache() {
  cachedExchangeRate = null;
}

/**
 * 金額を円に換算
 * @param {number} amount - 金額
 * @param {string} currency - 通貨コード（JPY / USD）
 * @returns {{amountJpy: number, rate: number|null}} 円換算結果
 */
function convertToJpy(amount, currency) {
  if (amount === null || amount === undefined) {
    return { amountJpy: null, rate: null };
  }
  
  const upperCurrency = (currency || 'JPY').toUpperCase().trim();
  
  if (upperCurrency === 'JPY' || upperCurrency === '円' || upperCurrency === '') {
    return { amountJpy: amount, rate: null };
  }
  
  if (upperCurrency === 'USD' || upperCurrency === 'ドル' || upperCurrency === '$') {
    const rate = getCachedUsdJpyRate();
    const amountJpy = Math.round(amount * rate);
    console.log(`通貨換算: $${amount} × ${rate} = ¥${amountJpy}`);
    return { amountJpy, rate };
  }
  
  // 未対応の通貨
  console.warn(`未対応の通貨: ${currency}`);
  return { amountJpy: amount, rate: null };
}
  
