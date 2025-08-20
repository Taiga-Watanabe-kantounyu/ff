/**
 * Puppeteerを使用した請求書PDF生成ユーティリティ
 * HTMLテンプレートを使用してサーバーサイドでPDFを生成する
 */

const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');
const { renderTemplate, saveRenderedHtml } = require('./htmlTemplateRenderer');
const invoiceConfig = require('../../config/invoice-config');

/**
 * HTMLファイルからPDFを生成する
 * @param {string} htmlPath - HTML文書のパス
 * @param {string} pdfPath - 出力PDFのパス
 * @param {Object} options - PDFオプション
 * @returns {Promise<string>} - 生成されたPDFのパス
 */
async function generatePdfFromHtml(htmlPath, pdfPath, options = {}) {
  if (!fs.existsSync(htmlPath)) {
    throw new Error(`HTMLファイルが存在しません: ${htmlPath}`);
  }
  
  // 出力ディレクトリを確認
  const outputDir = path.dirname(pdfPath);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  
  // デフォルトのPDFオプション
  const defaultOptions = {
    format: 'A4',
    printBackground: true,
    margin: {
      top: '1cm',
      right: '1cm',
      bottom: '1cm',
      left: '1cm'
    },
    preferCSSPageSize: true
  };
  
  // オプションをマージ
  const pdfOptions = { ...defaultOptions, ...options };
  
  let browser;
  try {
    // ブラウザを起動
    console.log('Puppeteerを起動中...');
    browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
    // 新しいページを開く
    const page = await browser.newPage();
    
    // HTMLファイルを読み込む
    const htmlContent = fs.readFileSync(htmlPath, 'utf8');
    
    // HTMLコンテンツを設定
    await page.setContent(htmlContent, {
      waitUntil: 'networkidle0', // すべてのリソースが読み込まれるまで待機
      timeout: 30000 // タイムアウト30秒
    });
    
    // HTMLからPDFを生成
    console.log('PDFを生成中...');
    await page.pdf({
      path: pdfPath,
      ...pdfOptions
    });
    
    console.log(`PDFが生成されました: ${pdfPath}`);
    return pdfPath;
  } catch (error) {
    console.error(`PDF生成中にエラーが発生しました: ${error.message}`);
    throw error;
  } finally {
    // ブラウザを閉じる
    if (browser) {
      await browser.close();
    }
  }
}

/**
 * HTMLテンプレートから直接PDFを生成する
 * @param {string} templatePath - テンプレートファイルのパス
 * @param {Object} data - テンプレートに適用するデータ
 * @param {string} pdfPath - 出力PDFのパス
 * @param {Object} options - PDFオプション
 * @returns {Promise<string>} - 生成されたPDFのパス
 */
async function generatePdfFromTemplate(templatePath, data, pdfPath, options = {}) {
  try {
    // 一時HTMLファイルのパスを生成
    const tempDir = path.join(process.cwd(), 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    const tempHtmlPath = path.join(tempDir, `invoice_${Date.now()}.html`);
    
    // テンプレートをレンダリング
    const renderedHtml = renderTemplate(templatePath, data);
    
    // 一時HTMLファイルに保存
    fs.writeFileSync(tempHtmlPath, renderedHtml, 'utf8');
    
    // HTMLからPDFを生成
    const result = await generatePdfFromHtml(tempHtmlPath, pdfPath, options);
    
    // 一時ファイルを削除
    fs.unlinkSync(tempHtmlPath);
    
    return result;
  } catch (error) {
    console.error(`PDF生成中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * 請求書データからPDFを生成する
 * @param {Object} invoiceData - 請求書データ
 * @param {string} year - 年 (YYYY)
 * @param {string} month - 月 (MM)
 * @param {Object} options - PDFオプション
 * @returns {Promise<string>} - 生成されたPDFのパス
 */
async function generateInvoicePdf(invoiceData, year, month, options = {}) {
  try {
    // 今日の日付
    const today = new Date();
    const todayStr = today.toISOString().split('T')[0].replace(/-/g, '');
    
    // ファイル名用に荷主名をサニタイズ
    let shipperName = '';
    if (invoiceData && invoiceData.billTo && invoiceData.billTo.shipperName) {
      shipperName = invoiceData.billTo.shipperName.replace(/[\\\/:*?"<>| \u3000]/g, '_');
    } else {
      shipperName = 'unknown';
    }

    // PDFファイル名の生成
    const pdfFileName = invoiceConfig.PDF_FILENAME_FORMAT
      .replace('{{shipper}}', shipperName)
      .replace('{{year}}', year)
      .replace('{{month}}', month)
      .replace('{{date}}', todayStr);
    
    // 出力先ディレクトリとファイルパスの設定
    const outputDir = path.join(process.cwd(), invoiceConfig.INVOICE_OUTPUT_DIR);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    const pdfPath = path.join(outputDir, pdfFileName);
    
    // テンプレートのパス
    const templatePath = path.join(process.cwd(), invoiceConfig.INVOICE_TEMPLATE_PATH);
    
    // テンプレートからPDFを生成
    return await generatePdfFromTemplate(templatePath, invoiceData, pdfPath, options);
  } catch (error) {
    console.error(`請求書PDF生成中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

module.exports = {
  generatePdfFromHtml,
  generatePdfFromTemplate,
  generateInvoicePdf
};
