/**
 * 請求書生成に関する設定ファイル
 */

module.exports = {
  // 消費税率
  TAX_RATE: 0.10,
  // 請求書の保存先ディレクトリ
  INVOICE_OUTPUT_DIR: 'billing/invoices',
  
  // HTMLテンプレートのパス
  INVOICE_TEMPLATE_PATH: 'templates/inv_temp.html',

  // PDFファイル名のフォーマット
  PDF_FILENAME_FORMAT: '請求書_{{year}}年{{month}}月_{{date}}.pdf',

  // HTMLファイル名のフォーマット
  HTML_FILENAME_FORMAT: '請求書_{{year}}年{{month}}月_{{date}}.html'
};
