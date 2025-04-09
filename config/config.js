// スプレッドシートのID
const SPREADSHEET_ID = '1bAMaFbZvfO3LDQ_fIdbUE4J6UcZvKpWXtHdwvurylYo';

// 認証情報ファイルのパス
const CREDENTIALS_PATH = './config/ff01-455323-24aa6cec6617.json';

// メール検索設定
const MAIL_SEARCH_CONFIG = {
  // 検索対象となる差出人のメールアドレス
  FROM_EMAIL: 'taiga-watanabe@kanto-unyu.co.jp'
};

// プリンタ設定
const PRINTER_CONFIG = {
  // 発注書の印刷に使用するプリンタ名
  // システムにインストールされているプリンタ名を指定
  // 例: 'Microsoft Print to PDF', '東京事務所 bizhub C360i'
  PRINTER_NAME: '東京事務所 bizhub C360i',
  
  // 印刷を有効にするかどうか
  ENABLE_PRINTING: true,
  
  // 印刷部数
  COPIES: 1
};

module.exports = {
  SPREADSHEET_ID,
  CREDENTIALS_PATH,
  PRINTER_CONFIG,
  MAIL_SEARCH_CONFIG
};
