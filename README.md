# 請求書自動管理システム - セットアップガイド

## 概要

毎月届く請求書メールを自動で収集し、PDFをGoogle Driveに保存、スプレッドシートで管理するシステムです。Gemini APIを使ってPDFから金額を自動抽出します。

## 前提条件

- Googleアカウント
- Google Apps Script（GAS）の基本的な操作知識
- Gemini APIキー（金額自動抽出を使う場合）

## セットアップ手順

### 1. スプレッドシートを作成

1. 新しいGoogleスプレッドシートを作成

### 2. PDF保存用フォルダを作成

1. Google Driveで新しいフォルダを作成（例: `請求書管理`）
2. フォルダのURLからIDをコピー
   - URL例: `https://drive.google.com/drive/folders/【この部分がID】`

### 3. GASプロジェクトを作成

1. スプレッドシートを開く
2. メニュー「拡張機能」→「Apps Script」をクリック
3. 各 `.gs` ファイルの内容をコピー:
   - `config.gs`
   - `utils.gs`
   - `sheet.gs`
   - `gmail.gs`
   - `drive.gs`
   - `gemini.gs`
   - `notify.gs`
   - `main.gs`
   - `debug.gs`

### 4. appsscript.json を設定

1. GASエディタで「プロジェクトの設定」（歯車アイコン）をクリック
2. 「「appsscript.json」マニフェスト ファイルをエディタで表示する」にチェック
3. `appsscript.json` を以下の内容に置き換え:

```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}
```

### 5. スクリプトプロパティを設定

1. GASエディタで「プロジェクトの設定」（歯車アイコン）をクリック
2. 「スクリプト プロパティ」セクションで以下を追加:

| プロパティ名 | 値 |
|------------|---|
| `DRIVE_FOLDER_ID` | 手順2で取得したフォルダID |
| `GEMINI_API_KEY` | Gemini APIのAPIキー（金額自動抽出用） |
| `NOTIFICATION_EMAIL` | 通知を受け取るメールアドレス（任意） |

#### Gemini APIキーの取得方法

1. https://aistudio.google.com/apikey にアクセス
2. 「APIキーを作成」をクリック
3. 発行されたAPIキーをコピーしてスクリプトプロパティに設定

### 6. シートを初期化

1. GASエディタで `initializeSheets` 関数を選択
2. 「実行」ボタンをクリック
3. 権限の承認を求められたら許可

### 7. 設定シートに送信元を登録

スプレッドシートの「設定」シートに、請求書が届く送信元を登録:

| 送信元メールアドレス | 会社名 | 検索キーワード | 有効 |
|:--|:--|:--|:--|
| invoice@example.co.jp | Example社 | 請求書 | TRUE |
| billing@service.com | サービス株式会社 | ご請求のお知らせ | TRUE |

### 8. 動作確認

1. `debugShowSenders` を実行して設定が正しく読み込まれるか確認
2. `debugSearchLastMonth` を実行してメールが検索できるか確認
3. `debugCheckDriveFolder` を実行してフォルダにアクセスできるか確認

### 9. 本番実行

1. `processLastMonthInvoices` を実行

### 10. 月次トリガーを設定（任意）

毎月自動で実行したい場合:

1. `createMonthlyTrigger` を実行
2. 毎月1日の午前9時に自動実行されるようになります

---

## ファイル構成

```
├── config.gs    ... 設定・定数
├── utils.gs     ... ユーティリティ関数
├── sheet.gs     ... スプレッドシート操作
├── gmail.gs     ... Gmail操作
├── drive.gs     ... ドライブ操作
├── gemini.gs    ... Gemini API（金額抽出）
├── notify.gs    ... 通知機能
├── main.gs      ... メイン処理・トリガー
└── debug.gs     ... デバッグ・テスト用
```

---

## 主な関数

### メイン処理

| 関数名 | 説明 |
|-------|------|
| `processLastMonthInvoices` | 前月の請求書を処理（メインエントリーポイント） |
| `initializeSheets` | シートを初期化 |
| `createMonthlyTrigger` | 月次トリガーを設定 |
| `deleteAllTriggers` | トリガーをすべて削除 |

### デバッグ用

| 関数名 | 説明 |
|-------|------|
| `debugShowProperties` | スクリプトプロパティを表示 |
| `debugShowSenders` | 設定シートの送信元一覧を表示 |
| `debugSearchLastMonth` | 前月のメールを検索（処理はしない） |
| `debugCheckDriveFolder` | 保存先フォルダを確認 |
| `debugProcessSingleMessage` | 単一メッセージを処理（テスト用） |
| `debugExtractAmountFromFile` | 指定PDFから金額抽出テスト |

---

## 処理の流れ

1. 設定シートから有効な送信元一覧を取得
2. 各送信元について前月のメールを検索
3. 各メールについて:
   - 重複チェック（Message IDで判定）
   - PDF添付ファイルを取得
   - Gemini APIで金額を抽出
   - ファイル名を生成（日付_会社名_金額.pdf）
   - ドライブに保存
   - スプレッドシートに記録
4. 処理結果を通知メールで送信

---

## ステータスについて

| ステータス | 説明 |
|----------|------|
| 完了 | 金額抽出に成功し、信頼度が高い |
| 要確認 | 金額抽出に失敗、または信頼度が低い |
| エラー | 処理中にエラーが発生 |

「要確認」の場合は、スプレッドシートで金額を手動確認・入力してください。

---

## トラブルシューティング

### 「DRIVE_FOLDER_ID が設定されていません」エラー

→ スクリプトプロパティが正しく設定されているか確認

### メールが検索されない

1. 検索キーワードが件名と一致しているか確認
2. 送信元メールアドレスが正しいか確認
3. `debugSearchEmails` でテスト検索を実行

### PDFが保存されない

1. 保存先フォルダIDが正しいか確認
2. フォルダへの書き込み権限があるか確認
3. `debugCheckDriveFolder` でフォルダにアクセスできるか確認

### 金額が抽出されない

1. `GEMINI_API_KEY` が設定されているか確認
2. `debugExtractAmountFromFile` で個別PDFをテスト
3. PDFが画像ベース（スキャン）の場合は精度が落ちることがあります

### 権限エラーが発生する

→ `appsscript.json` のスコープ設定を確認し、再度権限を承認

---

## ライセンス

MIT License