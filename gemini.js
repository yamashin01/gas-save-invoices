/**
 * 請求書自動管理システム - Gemini API（PDF金額抽出）
 */

const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
const GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent';

/**
 * PDFから請求金額を抽出
 * @param {Blob} pdfBlob - PDFのBlob
 * @returns {{amount: number|null, confidence: string, rawText: string}} 抽出結果
 */
function extractAmountFromPdf(pdfBlob) {
  if (!GEMINI_API_KEY) {
    console.warn('GEMINI_API_KEY が設定されていません。金額抽出をスキップします。');
    return { amount: null, confidence: 'スキップ', rawText: '' };
  }

  try {
    const base64Data = Utilities.base64Encode(pdfBlob.getBytes());
    const mimeType = 'application/pdf';

    const requestBody = {
      contents: [
        {
          parts: [
            {
              inline_data: {
                mime_type: mimeType,
                data: base64Data
              }
            },
            {
              text: `この請求書PDFから請求金額（税込合計）を抽出してください。

以下のJSON形式のみで回答してください。他の文章は不要です。
{
  "amount": 数値（円単位、カンマなし）,
  "confidence": "high" または "medium" または "low",
  "note": "抽出した金額の説明（例：税込合計、請求金額など）"
}

注意事項：
- 税込の合計金額を優先してください
- 金額が見つからない場合は amount を null にしてください
- 複数の金額がある場合は、請求合計・税込合計を選んでください`
            }
          ]
        }
      ],
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 256
      }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(`${GEMINI_API_URL}?key=${GEMINI_API_KEY}`, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      console.error(`Gemini API エラー: ${responseCode} - ${response.getContentText()}`);
      return { amount: null, confidence: 'エラー', rawText: response.getContentText() };
    }

    const result = JSON.parse(response.getContentText());
    const text = result.candidates?.[0]?.content?.parts?.[0]?.text || '';

    console.log(`Gemini応答: ${text}`);

    return parseGeminiResponse(text);

  } catch (error) {
    console.error(`金額抽出エラー: ${error.message}`);
    return { amount: null, confidence: 'エラー', rawText: error.message };
  }
}

/**
 * Geminiの応答をパース
 * @param {string} text - Geminiからの応答テキスト
 * @returns {{amount: number|null, confidence: string, rawText: string}} パース結果
 */
function parseGeminiResponse(text) {
  try {
    // JSON部分を抽出（余計な文字があっても対応）
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      console.warn('JSON形式の応答が見つかりません');
      return { amount: null, confidence: '解析失敗', rawText: text };
    }

    const parsed = JSON.parse(jsonMatch[0]);

    const amount = parsed.amount !== null && parsed.amount !== undefined 
      ? Number(parsed.amount) 
      : null;

    const confidenceMap = {
      'high': '高',
      'medium': '中',
      'low': '低'
    };
    const confidence = confidenceMap[parsed.confidence] || parsed.confidence || '不明';

    const note = parsed.note || '';

    console.log(`抽出結果 - 金額: ${amount}, 信頼度: ${confidence}, 備考: ${note}`);

    return { amount, confidence, rawText: note };

  } catch (error) {
    console.error(`応答パースエラー: ${error.message}`);
    return { amount: null, confidence: '解析失敗', rawText: text };
  }
}

/**
 * 金額抽出のテスト（デバッグ用）
 * @param {string} fileId - Google DriveのファイルID
 */
function debugExtractAmountFromFile(fileId) {
  console.log('=== 金額抽出テスト ===');

  try {
    const file = DriveApp.getFileById(fileId);
    console.log(`ファイル名: ${file.getName()}`);

    const blob = file.getBlob();
    const result = extractAmountFromPdf(blob);

    console.log(`金額: ${result.amount !== null ? result.amount.toLocaleString() + '円' : '抽出失敗'}`);
    console.log(`信頼度: ${result.confidence}`);
    console.log(`備考: ${result.rawText}`);

  } catch (error) {
    console.error(`エラー: ${error.message}`);
  }
}