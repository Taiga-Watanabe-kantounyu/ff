/**
 * 請求書生成に関する設定ファイル
 */

module.exports = {
  // 消費税率
  TAX_RATE: 0.10,
  // 請求書の保存先ディレクトリ
  INVOICE_OUTPUT_DIR: 'billing/invoices',
  
  // 会社情報
  COMPANY_NAME: '株式会社FF物流',
  COMPANY_ADDRESS: '〒123-4567 東京都千代田区FF町1-2-3',
  COMPANY_TEL: '03-1234-5678',
  COMPANY_FAX: '03-1234-5679',
  
  // 銀行情報
  BANK_INFO: '○○銀行 ○○支店 普通 1234567',
  
  // 支払期限（当月末）
  PAYMENT_TERM_DAYS: 30,
  
  // HTMLテンプレートのパス
  INVOICE_TEMPLATE_PATH: 'templates/inv_temp.html',
  
  // PDFファイル名のフォーマット
  PDF_FILENAME_FORMAT: '請求書_{{year}}年{{month}}月_{{date}}.pdf',
  
  // HTMLファイル名のフォーマット
  HTML_FILENAME_FORMAT: '請求書_{{year}}年{{month}}月_{{date}}.html',
  
  // スタイル設定
  STYLES: {
    // ここに必要に応じてスタイル設定を追加
  }
};
