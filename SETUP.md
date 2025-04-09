# 他のPCでのセットアップ手順詳細

このドキュメントでは、このプロジェクトを別のPCで動かすための詳細な手順を説明します。

## 前提条件

- Node.js（バージョン22.13.0以上）がインストールされていること
- Gitがインストールされていること
- Google Cloud Platformのアカウントを持っていること

## 1. リポジトリのクローン

```bash
git clone https://github.com/あなたのユーザー名/ff_2.git
cd ff_2
```

## 2. 依存パッケージのインストール

```bash
npm install
```

これにより、package.jsonに記載されている以下のパッケージがインストールされます：
- @google-cloud/local-auth
- axios
- exceljs
- googleapis
- xlsx

## 3. 認証情報の設定

### Gmail API用の認証情報

1. [Google Cloud Console](https://console.cloud.google.com/)にアクセス
2. プロジェクトを作成または選択
3. 「APIとサービス」→「認証情報」を選択
4. 「認証情報を作成」→「OAuthクライアントID」を選択
5. アプリケーションの種類で「デスクトップアプリ」を選択
6. 名前を入力して「作成」をクリック
7. ダウンロードしたJSONファイルを`config/creds.json`として保存

### Google Sheets API用の認証情報

1. 同じGoogle Cloud Consoleプロジェクトで「APIとサービス」→「ライブラリ」を選択
2. Google Sheets APIを検索して有効化
3. 「認証情報」→「認証情報を作成」→「サービスアカウント」を選択
4. サービスアカウント名を入力して「作成」をクリック
5. 役割として「編集者」を選択
6. 「完了」をクリック
7. 作成したサービスアカウントを選択
8. 「鍵」タブ→「鍵を追加」→「新しい鍵を作成」→「JSON」を選択
9. ダウンロードしたJSONファイルを`config/ff01-455323-24aa6cec6617.json`として保存

## 4. 必要なディレクトリの作成

```bash
mkdir -p data orders invoices
```

## 5. 設定ファイルの更新

`config/config.js`を開き、以下の項目を環境に合わせて更新します：

- `spreadsheetId`: 使用するGoogleスプレッドシートのID
- `printerSettings`: プリンタ名や印刷設定

## 6. データファイルの準備

### お届け先マスタの準備

1. `data/ffmasta.csv`が存在することを確認
2. CSVからJSONマスタを生成：
   ```bash
   node src/utils/csvToJson.js
   ```
   これにより`data/deliveryMaster.json`が生成されます

## 7. 初回実行

Gmail監視機能を初めて実行すると、ブラウザが開いてGoogleアカウントへのアクセス許可を求められます。許可すると`config/token.json`が生成されます。

```bash
node src/core/gmailWatcher.js
```

## トラブルシューティング

### 認証エラー

- `token.json`を削除して再認証を試みる
- Google Cloud Consoleで適切なAPIが有効化されているか確認

### スプレッドシートへのアクセスエラー

- サービスアカウントのメールアドレスをスプレッドシートの共有設定に追加（編集権限付与）

### 日本語ファイル名の問題

- Windows環境では日本語ファイル名の処理に問題が生じる場合があります
- ファイル名は英数字を使用することを推奨

### プリンタ関連の問題

- `config/config.js`の`printerSettings.enabled`を`false`に設定して印刷機能を無効化できます
