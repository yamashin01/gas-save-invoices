# 請求書自動管理システム - セットアップガイド

## 概要

毎月届く請求書メールを自動で収集し、PDFをGoogle Driveに保存、スプレッドシートで管理するシステムです。

## 前提条件

- Googleアカウント
- Google Apps Script（GAS）の基本的な操作知識

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
   - `notify.gs`
   - `main.gs`
   - `debug.gs`

### 4. スクリプトプロパティを設定

1. GASエディタで「プロジェクトの設定」（歯車アイコン）をクリック
2. 「スクリプト プロパティ」セクションで以下を追加:

| プロパティ名 | 値 |
|------------|---|
| `DRIVE_FOLDER_ID` | 手順2で取得したフォルダID |
| `NOTIFICATION_EMAIL` | 通知を受け取るメールアドレス（任意） |

### 5. シートを初期化

1. GASエディタで `initializeSheets` 関数を選択
2. 「実行」ボタンをクリック
3. 権限の承認を求められたら許可

### 6. 設定シートに送信元を登録

スプレッドシートの「設定」シートに、請求書が届く送信元を登録:

| 送信元メールアドレス | 会社名 | 検索キーワード | 有効 |
|:--|:--|:--|:--|
| invoice@example.co.jp | Example社 | 請求書 | TRUE |
| billing@service.com | サービス株式会社 | ご請求のお知らせ | TRUE |

### 7. 動作確認

1. `debugShowSenders` を実行して設定が正しく読み込まれるか確認
2. `debugSearchLastMonth` を実行してメールが検索できるか確認
3. `debugCheckDriveFolder` を実行してフォルダにアクセスできるか確認

### 8. 本番実行

1. `processLastMonthInvoices` を実行

### 9. 月次トリガーを設定（任意）

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

### 重複して処理される

→ 通常は Message ID で重複チェックされますが、「請求書一覧」シートのMessage ID列が空になっていないか確認

---

## Phase 2 への拡張（予定）

- PDF解析による金額自動抽出
- ファイル名への金額追加
- ステータス自動判定

---

## ライセンス

MIT License# gas-save-invoices
