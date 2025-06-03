# Progress

## 現在の状況
- Gmail APIとGoogle Cloud Pub/Subによるメールボックス変更のリアルタイム監視を実装。
- `gmailWatcher.js`で受信メールの添付Excelファイルをダウンロードし、`excelProcessor.js`でGoogleスプレッドシートに自動転記。
- `processed_sheets.json`への処理済みファイル記録機能を実装し、冪等性を確保。
- `orderFileGenerator.js`で受注データから発注用Excel（saito_irai.xlsx）をテンプレート保持しつつ生成。
- `deliveryMasterManager.js`と`deliveryMaster.json`による配送先マスタ管理機能を実装。
- `src/utils/addLeadTime.js`と`orderFileGenerator.js`で祝日API（holidays-jp）をキャッシュ利用し、日曜祝日をスキップする配送日計算を実装。
- `config.js`にMAIL_SEARCH_CONFIG、sheetName、プリンタ設定を追加し、動的に読み込む仕組みを実装。
- `src/utils/printFile.js`でPowerShellを利用したExcel自動印刷機能を実装。
- `config/invoice-config.js`で消費税率・会社情報・プリンタ設定などを一元管理。
- `htmlTemplateRenderer.js`と`formatUtils.js`でHTMLテンプレート（templates/inv_temp.html）ベースの請求書生成機能を実装。
- `puppeteerInvoiceGenerator.js`でPuppeteerサーバーサイドPDF生成機能を追加。
- `generateInvoice.js`ツールに`--puppeteer`オプションを追加し、デフォルトでPuppeteer出力を選択可能に。
- `billing/invoices`ディレクトリへの請求書HTML/PDF保存を実装。
- `sendErrorMail.js`をオブザーバーパターンでエラー通知に登録し、自動メール送信を実装。
- GmailWatcher→ExcelProcessor→OrderFileGenerator→PrintFileの処理フローをパイプライン（オーケストレーション）パターンで統括。

## 残りの作業
- メール変更監視およびスプレッドシート転記の追加テストとエッジケース検証。
- `config.js`からのシート名指定を厳密に検証し、任意のシート名対応を強化。
- エラーハンドリングとログ出力のさらなる改善、監視システムの安定化。
- ドキュメント整備とメモリバンク・READMEの同期。
- `deliveryMaster.json`の配送先データ拡充と管理UIまたはCLIの追加検討。

## 既知の問題
- スプレッドシートのシート名指定はまだ設定ファイル依存であり、シート存在チェックは限定的。
- 未登録配送先は電話番号・住所が空欄になる。
- フロー中の非同期処理タイミングによってログ出力が前後する場合がある。
