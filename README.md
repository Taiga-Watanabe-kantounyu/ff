# FF_2 自動化システム

このプロジェクトは、Gmail API と Google Cloud Pub/Sub を利用し、メール受信から発注ファイル作成・請求書出力までの一連の業務フローを自動化します。

## 主な機能

- Gmail のリアルタイム監視・新着メールの添付 Excel ファイルを取得  
- 添付 Excel → Google スプレッドシートへの自動転記  
- スプレッドシートのデータを元に発注用 Excel（saito_irai.xlsx）をテンプレートから生成  
- 休日・祝日を考慮したリードタイム計算による配送日自動算出  
- 請求書の HTML テンプレート生成＆Puppeteer による PDF 出力・自動保存・自動印刷  
- 外部設定（認証情報／スプレッドシートID／プリンタ設定など）による柔軟な管理  

## 必要ファイル（サンプル）

| ファイル                                | 役割                                      |
| --------------------------------------- | ----------------------------------------- |
| `config/config.js`                      | アプリ設定（スプレッドシートIDなど）     |
| `config/creds.json`                     | Gmail API 認証情報                        |
| `config/service-account-key.json`       | Google Sheets API 認証情報               |
| `data/ffmasta.csv`                      | お届け先マスタ CSV                        |
| `data/freightMaster.json`               | お届け先マスタ JSON                       |
| `data/processed_sheets.json`            | 処理済みシート記録                        |
| `data/last_mail_check.json`             | 最終メールチェック時刻                    |
| `templates/saito_irai.xlsx`             | 発注用テンプレート                        |

> ※ サンプルファイルは `*.example` 拡張子で同梱。実運用前にコピーして利用してください。

## セットアップ

```bash
git clone https://github.com/あなたのユーザー名/ff_2.git
cd ff_2

npm install

# サンプルファイルの作成
cp config/config.js.example config/config.js
cp data/ffmasta.csv.example data/ffmasta.csv
cp data/freightMaster.json.example data/freightMaster.json
cp data/processed_sheets.json.example data/processed_sheets.json
cp data/last_mail_check.json.example data/last_mail_check.json
cp templates/saito_irai.xlsx.example.txt templates/saito_irai.xlsx

# 必要ディレクトリの作成
mkdir -p data orders invoices
```

## 実行コマンド例

### Gmail 監視

```bash
node src/core/gmailWatcher.js
```

### 添付 Excel の単独処理

```bash
node src/core/excelProcessor.js <Excelファイルパス> [messageId] [内部日付]
```

### 発注用ファイル生成

```bash
node src/core/orderFileGenerator.js <受注ファイルパス>
```

### 請求書 PDF 出力

```bash
node src/tools/generateInvoice.js <請求データJSON>
# または Puppeteer 経由
node src/utils/puppeteerInvoiceGenerator.js
```

### お届け先マスタ管理

```bash
# インタラクティブモード
node src/models/deliveryMasterManager.js

# コマンド操作例
node src/models/deliveryMasterManager.js list
node src/models/deliveryMasterManager.js add <名前> <電話> <住所>
node src/models/deliveryMasterManager.js delete <名前>
node src/models/deliveryMasterManager.js search <キーワード>
```

### CSV → マスタ JSON 変換

```bash
node src/utils/csvToJson.js
```

## プロジェクト構成

```
ff_2/
├── src/                  # ソースコード
│   ├── core/             # コア処理
│   ├── utils/            # 共通ユーティリティ
│   └── models/           # データモデル
├── config/               # 設定・認証情報
├── data/                 # データファイル
├── templates/            # Excel / HTML テンプレート
├── orders/               # 発注書出力先
├── invoices/             # 請求書出力先
├── memory-bank/          # ドキュメント（メモリーバンク）
├── README.md             # プロジェクト概要（本ファイル）
└── SETUP.md              # 詳細セットアップ手順
```

## 技術スタック

- Node.js  
- Gmail API, Google Sheets API  
- Google Cloud Pub/Sub  
- exceljs / xlsx  
- Puppeteer  

## GitHub

リポジトリを更新し GitHub に反映する手順:

```bash
git add .
git commit -m "Update README"
git push origin main
```
