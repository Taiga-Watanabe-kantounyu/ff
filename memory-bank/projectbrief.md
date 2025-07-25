# Project Brief

このプロジェクトは、Gmail APIを使用してメールボックスの変更を監視し、受信した添付ファイルをGoogleスプレッドシートに自動で転記する機能を提供します。さらに、Excelデータを基に発注用ファイル（saito_irai.xlsx）を生成し、請求書をHTMLテンプレートおよびPuppeteerを用いてPDF形式で生成・保存・自動印刷するビジネスプロセスの自動化を実現します。Google Cloud Pub/Subによるリアルタイム通知、OAuth 2.0による認証、休日を考慮したリードタイム計算、設定ファイルベースの柔軟な管理機能を備えています。

また、FAX送信処理を自動化する機能を追加しました。`sendfax.ahk` を使用して、FAX番号の入力、送信ボタンのクリック、完了ボタンの操作を含むプロセスを効率化しています。
