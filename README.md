# Gmail Watcher

このプロジェクトは、Gmail APIを使用してメールボックスの変更を監視し、添付ファイル（Excelファイル）を処理して、Googleスプレッドシートに転記し、発注用ファイルを生成するシステムです。

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
│   ├── ff01-455323-24aa6cec6617.json # Google Sheets API認証情報
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
   - `config/ff01-455323-24aa6cec6617.json` - Google Sheets API用の認証情報

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
