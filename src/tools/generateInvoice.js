/**
 * 月次請求明細を手動で生成するためのユーティリティスクリプト（Puppeteer版）
 * 
 * 使用方法：
 *   node src/tools/generateInvoice.js [年] [月]
 * 
 * 例：
 *   node src/tools/generateInvoice.js 2025 04         # 2025年4月の請求明細を生成
 *   node src/tools/generateInvoice.js                 # 前月の請求明細を生成
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
            // 引数なしの場合は前月の請求明細を生成
            console.log('前月の請求明細をPDFで生成します...');
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
                console.log(`請求明細PDFを生成しました: ${pdfPath}`);
                console.log('PDFを開きます...');
                openPdfFile(pdfPath);
            } else {
                console.log('請求明細生成中にエラーが発生したか、対象データがありませんでした。');
            }
        } else if (args.length === 2) {
            // 年と月が指定された場合
            const year = args[0];
            const month = args[1].padStart(2, '0');
            
            console.log(`${year}年${month}月の請求明細をPDF生成します...`);
            const data = await collectInvoiceData(year, month);
            
            if (data.totalItems === 0) {
                console.log(`${year}年${month}月のデータが見つかりませんでした。請求明細は生成されません。`);
                return;
            }
            
            // 請求明細データの作成
            const invoiceData = prepareInvoiceData(data, year, month);
            
            // Puppeteerを使用してPDF生成
            const pdfPath = await generateInvoicePdf(invoiceData, year, month);
            
            if (pdfPath) {
                console.log(`請求明細PDFを生成しました: ${pdfPath}`);
                console.log('PDFを開きます...');
                openPdfFile(pdfPath);
            } else {
                console.log('請求明細生成中にエラーが発生したか、対象データがありませんでした。');
            }
        } else {
            console.log('使用方法:');
            console.log('  node src/tools/generateInvoice.js [年] [月]');
            console.log('例:');
            console.log('  node src/tools/generateInvoice.js 2025 04             # 2025年4月の請求明細を生成');
            console.log('  node src/tools/generateInvoice.js                     # 前月の請求明細を生成');
        }
    } catch (error) {
        console.error('エラーが発生しました:', error.message);
        process.exit(1);
    }
}

/**
 * Puppeteerを使った月次請求明細生成のためのデータを準備
 * @returns {Promise<Object>} - 請求明細データ
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
        
        console.log(`${targetYear}年${targetMonth}月の請求明細データを収集中...`);
        
        // データを収集
        const data = await collectInvoiceData(targetYear, targetMonth);
        
        if (data.totalItems === 0) {
            console.log(`${targetYear}年${targetMonth}月のデータが見つかりませんでした。請求明細は生成されません。`);
            return null;
        }
        
        // 請求明細データの作成
        return prepareInvoiceData(data, targetYear, targetMonth);
        
    } catch (error) {
        console.error(`月次請求明細データ収集中にエラーが発生しました: ${error.message}`);
        throw error;
    }
}

/**
 * 請求明細データを準備する
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
    
    // 請求明細番号
    const invoiceNumber = `INV-${year}${month}-001`;
    
    // テーブル行のHTML生成
    let itemsHtml = '';
    let currentDelivery = '';
    
    // お届け先ごとのデータ生成（テーブルフッターなし）
    Object.entries(data.deliveryTotals).forEach(([deliveryName, deliveryData], index) => {
        // 2番目以降の配送先から改ページを入れる
        const pageBreakClass = index > 0 ? 'page-break-before' : '';
        
        // お届け先の見出し行
        itemsHtml += `
            <tr class="delivery-header ${pageBreakClass}">
                <td><strong>${deliveryName}</strong></td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td><strong>¥${formatNumberToCurrency(deliveryData.total)}</strong></td>
            </tr>
        `;
        
        // お届け先ごとのランク別小計は非表示
        
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
        
        // 最後のお届け先以外に区切り行を追加
        if (index < Object.entries(data.deliveryTotals).length - 1) {
            itemsHtml += `
                <tr class="separator">
                    <td colspan="6"></td>
                </tr>
            `;
        }
    });
    
    // 消費税の計算
    const taxAmount = Math.round(data.grandTotal * invoiceConfig.TAX_RATE);
    const totalWithTax = data.grandTotal + taxAmount;
    
    // 合計部分の前に空白行を追加して視覚的に分離
    itemsHtml += `
        <tr>
            <td colspan="6" style="height: 20px; border: none;"></td>
        </tr>
        <tr>
            <td colspan="6" style="border-top: 2px solid #4a90e2; border-left: none; border-right: none; border-bottom: none; padding: 0;"></td>
        </tr>
        <tr>
            <td colspan="6" style="text-align: center; font-weight: bold; font-size: 16px; border: none; padding: 15px 0;">【請求明細合計】</td>
        </tr>
        <tr class="total-section" style="background-color: #f0f4f8;">
            <td style="text-align: right; font-weight: bold;">小計</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td style="font-weight: bold;">¥${formatNumberToCurrency(data.grandTotal)}</td>
        </tr>
        <tr style="background-color: #f0f4f8;">
            <td style="text-align: right; font-weight: bold;">消費税（10%）</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td style="font-weight: bold;">¥${formatNumberToCurrency(taxAmount)}</td>
        </tr>
        <tr style="background-color: #e6eff8;">
            <td style="text-align: right; font-weight: bold; font-size: 16px;">合計金額</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td style="font-weight: bold; font-size: 16px;">¥${formatNumberToCurrency(totalWithTax)}</td>
        </tr>
    `;
    
    // テンプレートに適用するデータを準備
    return {
        customerName: config.MAIL_SEARCH_CONFIG && config.MAIL_SEARCH_CONFIG.FROM_EMAIL ? config.MAIL_SEARCH_CONFIG.FROM_EMAIL.split('@')[0] : invoiceConfig.COMPANY_NAME,
        invoiceDate: todayStr,
        periodStart: `${year}年${month}月1日`,
        periodEnd: `${year}年${month}月${lastDay}日`,
        invoiceNumber: invoiceNumber,
        grandTotal: formatNumberToCurrency(totalWithTax),
        itemsHtml: itemsHtml,
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
