const path = require('path');
const { google } = require('googleapis');
const { SPREADSHEET_ID } = require('../../config/config');

/**
 * 請求書データ収集機能
 * スプレッドシートのデータを基に請求書生成に必要なデータを収集する
 */

/**
 * 請求書生成に必要なデータを収集する
 * @param {string} year - 対象年 (YYYY)
 * @param {string} month - 対象月 (MM)
 * @returns {Promise<Object>} - 請求書生成に必要なデータ
 */
const { loadDeliveryMaster } = require('../models/deliveryMasterManager');

async function collectInvoiceData(year, month) {
    try {
        // Google Sheets APIの認証
        const auth = new google.auth.GoogleAuth({
            keyFile: path.join(process.cwd(), 'config', 'ff01-455323-24aa6cec6617.json'),
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });
        const sheets = google.sheets({ version: 'v4', auth });

        // スプレッドシートから対象月のデータを取得
        console.log(`${year}年${month}月のデータを取得中...`);
        
        // シート名（デフォルトは'Sheet1'）
        const TARGET_SHEET_NAME = 'Sheet1';
        
        // スプレッドシートのデータを全て取得
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${TARGET_SHEET_NAME}!A:I`, // A列からI列まで全て取得
        });
        
        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            console.log('データが見つかりませんでした。');
            return {
                totalItems: 0,
                deliveryTotals: {},
                rankTotals: { A: 0, B: 0 },
                grandTotal: 0,
                details: []
            };
        }
        
        // ヘッダー行を除外
        const dataRows = rows.slice(1);
        
        // 対象月のデータをフィルタリング
        const targetData = dataRows.filter(row => {
            if (!row[0]) return false; // 集荷日がない行はスキップ
            
            // 集荷日のフォーマットを確認（YYYY/MM/DD形式を想定）
            const dateParts = row[0].split('/');
            if (dateParts.length !== 3) return false;
            
            const rowYear = dateParts[0];
            const rowMonth = dateParts[1].padStart(2, '0');
            
            return rowYear === year && rowMonth === month;
        });
        
        console.log(`対象データ数: ${targetData.length}行`);
        
        // データ集計
        const deliveryTotals = {}; // お届け先ごとの合計
        const rankTotals = { A: 0, B: 0 }; // ランク別合計
        let grandTotal = 0; // 総合計
        
        // 配送マスタを一括取得
const deliveryMaster = loadDeliveryMaster();

 // 配送先ごとの合計ケース数を事前計算
// ※仕様：同一配送先・同一集荷日ごとに合計5ケース以上なら全商品に5ケース以上料金を適用する
const deliveryQuantities = {};
targetData.forEach(row => {
    const qty = parseInt(row[6]) || 1;
    const name = row[7] || '';
    const pickupDate = row[0] || '';
    if (name && pickupDate) {
        deliveryQuantities[name] = deliveryQuantities[name] || {};
        deliveryQuantities[name][pickupDate] = (deliveryQuantities[name][pickupDate] || 0) + qty;
    }
});

        const details = targetData.map(row => {
            const pickupDate = row[0] || '';      // 集荷日
            const temperatureZone = row[1] || ''; // 温度帯
            const productCode = row[2] || '';     // 品番
            const productName = row[3] || '';     // 品名
            const rank = row[4] || '';            // ランク
            const weight = row[5] || '';          // 重量(kg)
            const quantity = parseInt(row[6]) || 1;  // 数量（個口数）
            const deliveryName = row[7] || '';    // お届け先名

            // 配送マスタから該当先情報取得
            const master = deliveryMaster[deliveryName] || {};

            // ランク判定
            const rankStr = rank ? rank.toString().toUpperCase() : '';
            // 5ケース以下はマスタのLow料金を参照
            let freightPerUnit;
const dailyQty = ((deliveryQuantities[deliveryName] || {})[pickupDate] || 0);
if (dailyQty < 5) {
                if (rankStr === 'A') {
                    freightPerUnit = master.aRankFeeLow !== undefined ? master.aRankFeeLow : parseFloat(row[8]) || 0;
                } else if (rankStr === 'B') {
                    freightPerUnit = master.bRankFeeLow !== undefined ? master.bRankFeeLow : parseFloat(row[8]) || 0;
                } else if (rankStr === 'C') {
                    freightPerUnit = master.cRankFeeLow !== undefined ? master.cRankFeeLow : parseFloat(row[8]) || 0;
                } else if (rankStr === 'D') {
                    freightPerUnit = master.dRankFeeLow !== undefined ? master.dRankFeeLow : parseFloat(row[8]) || 0;
                } else {
                    freightPerUnit = parseFloat(row[8]) || 0;
                }
            } else {
                if (rankStr === 'A') {
                    freightPerUnit = master.aRankFee !== undefined ? master.aRankFee : parseFloat(row[8]) || 0;
                } else if (rankStr === 'B') {
                    freightPerUnit = master.bRankFee !== undefined ? master.bRankFee : parseFloat(row[8]) || 0;
                } else if (rankStr === 'C') {
                    freightPerUnit = master.cRankFee !== undefined ? master.cRankFee : parseFloat(row[8]) || 0;
                } else if (rankStr === 'D') {
                    freightPerUnit = master.dRankFee !== undefined ? master.dRankFee : parseFloat(row[8]) || 0;
                } else {
                    freightPerUnit = parseFloat(row[8]) || 0;
                }
            }

            const totalFreight = freightPerUnit * quantity; // 総運賃 = 個口当たりの運賃 × 個口数
            
            // お届け先ごとの合計を計算
            if (deliveryName) {
                if (!deliveryTotals[deliveryName]) {
                    deliveryTotals[deliveryName] = {
                        total: 0,
                        aRankTotal: 0,
                        bRankTotal: 0,
                        items: []
                    };
                }
                
                deliveryTotals[deliveryName].total += totalFreight;
                
                // ランク別の合計を計算
                if (rank && rank.toString().toUpperCase() === 'A') {
                    rankTotals.A += totalFreight;
                    deliveryTotals[deliveryName].aRankTotal += totalFreight;
                } else {
                    rankTotals.B += totalFreight;
                    deliveryTotals[deliveryName].bRankTotal += totalFreight;
                }
                
                // 明細を追加
                deliveryTotals[deliveryName].items.push({
                    pickupDate,
                    temperatureZone,
                    productCode,
                    productName,
                    rank,
                    weight,
                    quantity,
                    freightPerUnit,
                    totalFreight
                });
            }
            
            // 総合計に加算
            grandTotal += totalFreight;
            
            return {
                pickupDate,
                temperatureZone,
                productCode,
                productName,
                rank,
                weight,
                quantity,
                deliveryName,
                freightPerUnit,
                totalFreight
            };
        });
        
        return {
            totalItems: targetData.length,
            deliveryTotals,
            rankTotals,
            grandTotal,
            details
        };
        
    } catch (error) {
        console.error(`データ収集中にエラーが発生しました: ${error.message}`);
        throw error;
    }
}

module.exports = {
    collectInvoiceData
};
