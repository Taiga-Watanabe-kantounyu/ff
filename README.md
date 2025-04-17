# Gmail Watcher

このプロジェクトは、Gmail APIを使用してメールボックスの変更を監視し、添付ファイル（Excelファイル）を処理して、Googleスプレッドシートに転記し、発注用ファイルを生成するシステムです。

## 機密情報の取り扱いについて

このリポジトリは、機密情報（認証情報、お届け先情報など）を含まずに公開されています。実際に使用する際は、以下のファイルを適切に設定する必要があります：

- **認証情報**：
  - `creds.json` - Gmail API用の認証情報
  - `service-account-key.json` - Google Sheets API用の認証情報
  - `token.json` - アクセストークン（初回実行時に自動生成）

- **設定ファイル**：
  - `config.js` - スプレッドシートID、プリンタ設定などの設定情報

- **データファイル**：
  - `data/ffmasta.csv` - お届け先マスタCSV
  - `data/freightMaster.json` - お届け先マスタJSON
  - `data/processed_sheets.json` - 処理済みシート記録
  - `data/last_mail_check.json` - 最終メールチェック時間

- **テンプレートファイル**：
  - `templates/saito_irai.xlsx` - 発注用テンプレートファイル

これらのファイルは`.gitignore`に登録されており、リポジトリには含まれていません。代わりに、各ファイルのサンプル（`.example`拡張子付き）が提供されています。

## セットアップ手順

1. **リポジトリのクローン**
   ```
   git clone https://github.com/あなたのユーザー名/ff_2.git
   cd ff_2
   ```

2. **依存パッケージのインストール**
   ```
   npm install
   ```

3. **サンプルファイルから実際のファイルを作成**
   ```
   cp config/config.js.example config/config.js
   cp data/ffmasta.csv.example data/ffmasta.csv
   cp data/freightMaster.json.example data/freightMaster.json
   cp data/processed_sheets.json.example data/processed_sheets.json
   cp data/last_mail_check.json.example data/last_mail_check.json
   ```

4. **認証情報の設定**
   - Google Cloud Consoleから認証情報を取得し、以下のファイルを配置：
     - `config/creds.json` - Gmail API用の認証情報
     - `config/service-account-key.json` - Google Sheets API用の認証情報
   - 初回実行時に`token.json`は自動生成されます

5. **必要なディレクトリの作成**
   ```
   mkdir -p data orders invoices
   ```

6. **設定ファイルの確認と更新**
   - `config/config.js`のスプレッドシートIDやプリンタ設定を環境に合わせて更新

## プロジェクト構造

```
ff_2/
├── src/                  # ソースコード
│   ├── core/             # コア機能
│   │   ├── gmailWatcher.js       # Gmailの監視機能
│   │   ├── excelProcessor.js     # Excelファイルの処理機能
│   │   └── orderFileGenerator.js # 発注用ファイル生成機能
│   ├── utils/            # ユーティリティ
│   │   ├── checkExcelStructure.js # Excelファイルの構造確認
│   │   └── csvToJson.js          # CSVからJSONへの変換
│   └── models/           # データモデル
│       └── deliveryMasterManager.js # お届け先マスタ管理
├── config/               # 設定ファイル
│   ├── config.js                 # アプリケーション設定
│   ├── creds.json                # Gmail API認証情報
│   ├── service-account-key.json  # Google Sheets API認証情報
│   └── token.json                # アクセストークン
├── data/                 # データファイル
│   ├── deliveryMaster.json       # お届け先マスタ
│   ├── processed_sheets.json     # 処理済みシート記録
│   └── ffmasta.csv               # お届け先マスタCSV
├── templates/            # テンプレートファイル
│   └── saito_irai.xlsx           # 発注用テンプレート
├── docs/                 # ドキュメント
│   ├── excelフォーマット仕様.md   # Excelフォーマット仕様
│   ├── processAttachments.md     # 添付ファイル処理仕様
│   └── 発注仕様.md               # 発注仕様
├── samples/              # サンプルファイル
│   ├── 依頼書フォーマット_テスト_発注.xlsx
│   ├── 依頼書フォーマット_テスト.xlsx
│   ├── 依頼書フォーマット_テスト2.xlsx
│   ├── 依頼書フォーマット_テスト3.xlsx
│   └── 依頼書フォーマット_テスト4.xlsx
├── tests/                # テストファイル（将来的に追加）
├── memory-bank/          # メモリーバンク
│   ├── projectbrief.md
│   ├── productContext.md
│   ├── systemPatterns.md
│   ├── techContext.md
│   ├── activeContext.md
│   └── progress.md
├── package.json          # NPM設定
└── README.md             # プロジェクト説明
```

## 機能概要

1. **Gmail監視機能** (`src/core/gmailWatcher.js`)
   - Gmail APIを使用してメールボックスの変更を監視
   - 特定の条件に一致するメールの添付ファイルを取得

2. **Excel処理機能** (`src/core/excelProcessor.js`)
   - 受信したExcelファイルを処理
   - データをGoogleスプレッドシートに転記

3. **発注ファイル生成機能** (`src/core/orderFileGenerator.js`)
   - 受注データから発注用ファイルを生成
   - テンプレートを使用して正しいフォーマットで出力

4. **お届け先マスタ管理** (`src/models/deliveryMasterManager.js`)
   - お届け先情報（名前、電話番号、住所）の管理
   - 追加/更新/削除/検索機能

## 使用技術

- Node.js
- Google Cloud Pub/Sub
- Gmail API
- Google Sheets API
- xlsx / exceljs（Excelファイル処理）

## セットアップ

1. 必要なパッケージのインストール
   ```
   npm install
   ```

2. 認証情報の設定
   - `config/creds.json` - Gmail API用の認証情報
   - `config/service-account-key.json` - Google Sheets API用の認証情報

3. 設定ファイルの確認
   - `config/config.js` - スプレッドシートIDなどの設定

## 使用方法

### Gmail監視の開始

```
node src/core/gmailWatcher.js
```

### Excel処理の単独実行

```
node src/core/excelProcessor.js <Excelファイルパス> [メッセージID] [内部日付]
```

### 発注ファイル生成の単独実行

```
node src/core/orderFileGenerator.js <受注ファイルパス>
```

### お届け先マスタ管理

```
node src/models/deliveryMasterManager.js                     # インタラクティブモード
node src/models/deliveryMasterManager.js list                # 一覧表示
node src/models/deliveryMasterManager.js add <名前> <電話> <住所> # 追加/更新
node src/models/deliveryMasterManager.js delete <名前>       # 削除
node src/models/deliveryMasterManager.js search <キーワード> # 検索
```

### CSVからお届け先マスタの生成

```
node src/utils/csvToJson.js
```

## 注意事項

- スプレッドシートのシート名は現在「Sheet1」に固定されています。
- お届け先マスタに登録されていないお届け先の場合、電話番号と住所が空欄になります。
