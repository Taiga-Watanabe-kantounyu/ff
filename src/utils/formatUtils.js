/**
 * 数値をフォーマットするユーティリティ関数
 */

/**
 * 数値を通貨形式にフォーマットする
 * @param {number} value - フォーマットする数値
 * @param {string} locale - ロケール（デフォルト: 'ja-JP'）
 * @param {string} currency - 通貨コード（デフォルト: 'JPY'）
 * @returns {string} - フォーマットされた通貨文字列
 */
function formatCurrency(value, locale = 'ja-JP', currency = 'JPY') {
  return new Intl.NumberFormat(locale, {
    style: 'currency',
    currency: currency,
    minimumFractionDigits: 0
  }).format(value);
}

/**
 * 数値を3桁区切りの数値形式にフォーマットする
 * @param {number} value - フォーマットする数値
 * @param {string} locale - ロケール（デフォルト: 'ja-JP'）
 * @returns {string} - フォーマットされた数値文字列
 */
function formatNumber(value, locale = 'ja-JP') {
  return new Intl.NumberFormat(locale, {
    style: 'decimal',
    minimumFractionDigits: 0
  }).format(value);
}

/**
 * 日付をフォーマットする
 * @param {Date|string} date - フォーマットする日付オブジェクトまたは日付文字列
 * @param {string} format - 出力形式 ('yyyy/MM/dd', 'yyyy年MM月dd日', 'MM/dd' のいずれか)
 * @returns {string} - フォーマットされた日付文字列
 */
function formatDate(date, format = 'yyyy/MM/dd') {
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  switch (format) {
    case 'yyyy/MM/dd':
      return `${year}/${month}/${day}`;
    case 'yyyy年MM月dd日':
      return `${year}年${month}月${day}日`;
    case 'MM/dd':
      return `${month}/${day}`;
    default:
      return `${year}/${month}/${day}`;
  }
}

module.exports = {
  formatCurrency,
  formatNumber,
  formatDate
};
