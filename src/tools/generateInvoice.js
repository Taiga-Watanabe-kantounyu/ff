/**
 * 月次請求書を手動で生成するためのユーティリティスクリプト（Puppeteer版）
 * 
 * 使用方法：
 *   node src/tools/generateInvoice.js [年] [月]
 * 
 * 例：
 *   node src/tools/generateInvoice.js 2025 04         # 2025年4月の請求書を生成
 *   node src/tools/generateInvoice.js                 # 前月の請求書を生成
 */

const path = require('path');
const fs = require('fs');
const { collectInvoiceData } = require('../utils/invoiceGenerator');
const { generateInvoicePdf } = require('../utils/puppeteerInvoiceGenerator');
const config = require('../../config/config');
const invoiceConfig = require('../../config/invoice-config');

/**
 * PDFファイルをブラウザで開く
 * @param {string} filePath - ファイルパス
 */
function openPdfFile(filePath) {
    if (filePath && fs.existsSync(filePath)) {
        // OSによって適切なコマンドを実行
        const isWindows = process.platform === 'win32';
        
        const { exec } = require('child_process');
        if (isWindows) {
            // Windowsの場合、ファイルの関連付けを使ってデフォルトのアプリで開く
            exec(`cmd /c start "" "${filePath}"`, (error) => {
                if (error) {
                    console.error(`ファイルを開く際にエラーが発生しました: ${error.message}`);
                }
            });
        } else {
            // Mac/Linux
            const command = process.platform === 'darwin' ? 'open' : 'xdg-open';
            exec(`${command} "${filePath}"`, (error) => {
                if (error) {
                    console.error(`ファイルを開く際にエラーが発生しました: ${error.message}`);
                }
            });
        }
    }
}


async function main() {
    try {
        const args = process.argv.slice(2);
        
        // 引数に基づいて処理を分岐
        if (args.length === 0) {
            // 引数なしの場合は前月の請求書を生成
            console.log('前月の請求書をPDFで生成します...');
            const invoiceData = await generateMonthlyPuppeteerData();
            
            if (invoiceData) {
                // 年月情報を取得
                const now = new Date();
                let targetYear, targetMonth;
                
                if (now.getMonth() === 0) {
                    // 1月の場合は前年の12月
                    targetYear = (now.getFullYear() - 1).toString();
                    targetMonth = '12';
                } else {
                    // それ以外は前月
                    targetYear = now.getFullYear().toString();
                    targetMonth = String(now.getMonth()).padStart(2, '0');
                }
                
                const pdfPath = await generateInvoicePdf(invoiceData, targetYear, targetMonth);
                console.log(`請求書PDFを生成しました: ${pdfPath}`);
                console.log('PDFを開きます...');
                openPdfFile(pdfPath);
            } else {
                console.log('請求書生成中にエラーが発生したか、対象データがありませんでした。');
            }
        } else if (args.length === 2) {
            // 年と月が指定された場合
            const year = args[0];
            const month = args[1].padStart(2, '0');
            
            console.log(`${year}年${month}月の請求書をPDF生成します...`);
            const data = await collectInvoiceData(year, month);
            
            if (data.totalItems === 0) {
                console.log(`${year}年${month}月のデータが見つかりませんでした。請求書は生成されません。`);
                return;
            }
            
            // 請求書データの作成
            const invoiceData = prepareInvoiceData(data, year, month);
            
            // Puppeteerを使用してPDF生成
            const pdfPath = await generateInvoicePdf(invoiceData, year, month);
            
            if (pdfPath) {
                console.log(`請求書PDFを生成しました: ${pdfPath}`);
                console.log('PDFを開きます...');
                openPdfFile(pdfPath);
            } else {
                console.log('請求書生成中にエラーが発生したか、対象データがありませんでした。');
            }
        } else {
            console.log('使用方法:');
            console.log('  node src/tools/generateInvoice.js [年] [月]');
            console.log('例:');
            console.log('  node src/tools/generateInvoice.js 2025 04             # 2025年4月の請求書を生成');
            console.log('  node src/tools/generateInvoice.js                     # 前月の請求書を生成');
        }
    } catch (error) {
        console.error('エラーが発生しました:', error.message);
        process.exit(1);
    }
}

/**
 * Puppeteerを使った月次請求書生成のためのデータを準備
 * @returns {Promise<Object>} - 請求書データ
 */
async function generateMonthlyPuppeteerData() {
    try {
        // 対象年月を設定
        const now = new Date();
        let targetYear, targetMonth;
        
        // 前月を使用
        if (now.getMonth() === 0) {
            // 1月の場合は前年の12月
            targetYear = (now.getFullYear() - 1).toString();
            targetMonth = '12';
        } else {
            // それ以外は前月
            targetYear = now.getFullYear().toString();
            targetMonth = String(now.getMonth()).padStart(2, '0');
        }
        
        console.log(`${targetYear}年${targetMonth}月の請求書データを収集中...`);
        
        // データを収集
        const data = await collectInvoiceData(targetYear, targetMonth);
        
        if (data.totalItems === 0) {
            console.log(`${targetYear}年${targetMonth}月のデータが見つかりませんでした。請求書は生成されません。`);
            return null;
        }
        
        // 請求書データの作成
        return prepareInvoiceData(data, targetYear, targetMonth);
        
    } catch (error) {
        console.error(`月次請求書データ収集中にエラーが発生しました: ${error.message}`);
        throw error;
    }
}

/**
 * 請求書データを準備する
 * @param {Object} data - 収集したデータ
 * @param {string} year - 年 (YYYY)
 * @param {string} month - 月 (MM)
 * @returns {Object} - テンプレートに適用するデータ
 */
function prepareInvoiceData(data, year, month) {
    // 請求期間の設定
    const periodStart = `${year}/${month}/01`;
    const lastDay = new Date(parseInt(year), parseInt(month), 0).getDate();
    const periodEnd = `${year}/${month}/${lastDay}`;
    
    // 請求日（今日の日付）
    const today = new Date();
    const todayStr = `${today.getFullYear()}年${String(today.getMonth() + 1).padStart(2, '0')}月${String(today.getDate()).padStart(2, '0')}日`;
    
    // 支払期限の計算
    const dueDate = new Date(parseInt(year), parseInt(month), invoiceConfig.PAYMENT_TERM_DAYS);
    const dueDateStr = `${dueDate.getFullYear()}年${String(dueDate.getMonth() + 1).padStart(2, '0')}月${String(dueDate.getDate()).padStart(2, '0')}日`;
    
    // 請求書番号
    const invoiceNumber = `INV-${year}${month}-001`;
    
    // テーブル行のHTML生成
    let itemsHtml = '';
    let currentDelivery = '';
    
    // お届け先ごとのデータ生成
    Object.entries(data.deliveryTotals).forEach(([deliveryName, deliveryData]) => {
        // お届け先の見出し行
        itemsHtml += `
            <tr class="delivery-header">
                <td><strong>${deliveryName}</strong></td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td><strong>¥${formatNumberToCurrency(deliveryData.total)}</strong></td>
            </tr>
        `;
        
        // ランク別小計
        if (deliveryData.aRankTotal > 0) {
            itemsHtml += `
                <tr class="rank-total">
                    <td></td>
                    <td>Aランク小計</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td>¥${formatNumberToCurrency(deliveryData.aRankTotal)}</td>
                </tr>
            `;
        }
        
        if (deliveryData.bRankTotal > 0) {
            itemsHtml += `
                <tr class="rank-total">
                    <td></td>
                    <td>Bランク小計</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td>¥${formatNumberToCurrency(deliveryData.bRankTotal)}</td>
                </tr>
            `;
        }
        
        // 明細行
        deliveryData.items.forEach(item => {
            itemsHtml += `
                <tr>
                    <td>${item.pickupDate}</td>
                    <td>${item.rank || '-'}</td>
                    <td>${item.productName || '-'}</td>
                    <td>${item.quantity || '-'}</td>
                    <td>¥${formatNumberToCurrency(item.freightPerUnit)}</td>
                    <td>¥${formatNumberToCurrency(item.totalFreight)}</td>
                </tr>
            `;
        });
        
        // 区切り行
        itemsHtml += `
            <tr class="separator">
                <td colspan="6"></td>
            </tr>
        `;
    });
    
    // 消費税の計算
    const taxAmount = Math.floor(data.grandTotal * invoiceConfig.TAX_RATE);
    const totalWithTax = data.grandTotal + taxAmount;
    
    // 集計行
    itemsHtml += `
        <tr class="total-section">
            <td>小計</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td>¥${formatNumberToCurrency(data.grandTotal)}</td>
        </tr>
        <tr>
            <td>消費税（10%）</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td>¥${formatNumberToCurrency(taxAmount)}</td>
        </tr>
        <tr>
            <td><strong>合計金額</strong></td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td><strong>¥${formatNumberToCurrency(totalWithTax)}</strong></td>
        </tr>
        <tr>
            <td>Aランク小計</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td>¥${formatNumberToCurrency(data.rankTotals.A)}</td>
        </tr>
        <tr>
            <td>Bランク小計</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td>¥${formatNumberToCurrency(data.rankTotals.B)}</td>
        </tr>
    `;
    
    // テンプレートに適用するデータを準備
    return {
        customerName: config.MAIL_SEARCH_CONFIG && config.MAIL_SEARCH_CONFIG.FROM_EMAIL ? config.MAIL_SEARCH_CONFIG.FROM_EMAIL.split('@')[0] : invoiceConfig.COMPANY_NAME,
        invoiceDate: todayStr,
        periodStart: `${year}年${month}月1日`,
        periodEnd: `${year}年${month}月${lastDay}日`,
        invoiceNumber: invoiceNumber,
        dueDate: dueDateStr,
        bankInfo: invoiceConfig.BANK_INFO,
        grandTotal: formatNumberToCurrency(totalWithTax),
        itemsHtml: itemsHtml,
        // 会社情報
        companyName: invoiceConfig.COMPANY_NAME,
        companyAddress: invoiceConfig.COMPANY_ADDRESS,
        companyTel: invoiceConfig.COMPANY_TEL,
        companyFax: invoiceConfig.COMPANY_FAX
    };
}

/**
 * 数値を通貨形式にフォーマットする
 * @param {number} value - 値
 * @returns {string} - フォーマットされた文字列
 */
function formatNumberToCurrency(value) {
    return new Intl.NumberFormat('ja-JP', {
        style: 'decimal',
        minimumFractionDigits: 0
    }).format(value);
}

main();
