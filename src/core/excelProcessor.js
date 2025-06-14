const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { google } = require('googleapis');
const { SPREADSHEET_ID, CREDENTIALS_PATH } = require('../../config/config');
const PROCESSED_SHEETS_FILE = path.join(process.cwd(), 'data', 'processed_sheets.json');
const FREIGHT_MASTER_FILE = path.join(process.cwd(), 'data', 'freightMaster.json');

// スプレッドシートのシート名（デフォルトは'Sheet1'）
const TARGET_SHEET_NAME = 'Sheet1';

// 運賃マスタを読み込む
function loadFreightMaster() {
    if (fs.existsSync(FREIGHT_MASTER_FILE)) {
        const data = fs.readFileSync(FREIGHT_MASTER_FILE, 'utf8');
        return JSON.parse(data || '{}');
    }
    return {};
}

function loadProcessedSheets() {
    if (fs.existsSync(PROCESSED_SHEETS_FILE)) {
        const data = fs.readFileSync(PROCESSED_SHEETS_FILE, 'utf8');
        return JSON.parse(data || '[]');
    }
    return [];
}

async function processExcelAndUpdateSheet(directoryPath, messageId, internalDate) {
    const processedFiles = loadProcessedSheets();
    // ファイル名を取得して一意のIDに含める
    const fileName = path.basename(directoryPath);
    const uniqueId = `${messageId}-${internalDate}-${fileName}`;
    if (processedFiles.includes(uniqueId)) {
        console.log(`ファイル: ${fileName} (メールID: ${messageId}) は既に処理済みです。スキップします。`);
        return false; // 処理をスキップした場合はfalseを返す
    }
    try {
        await processSingleExcelFile(directoryPath, uniqueId);
        return true; // 処理を実行した場合はtrueを返す
    } catch (error) {
        console.error(`ファイルの処理中にエラー発生: ${error.message}`);
        return false; // エラーが発生した場合はfalseを返す
    }
}

function saveProcessedSheet(sheetTitle) {
    const processedSheets = loadProcessedSheets();
    if (!processedSheets.includes(sheetTitle)) {
        processedSheets.push(sheetTitle);
        fs.writeFileSync(PROCESSED_SHEETS_FILE, JSON.stringify(processedSheets));
    }
}

module.exports = { processExcelAndUpdateSheet, processSingleExcelFile };

// コマンドラインから直接実行された場合の処理
if (require.main === module) {
    const args = process.argv.slice(2);
    if (args.length < 1) {
        console.log('使用方法: node excelProcessor.js <Excelファイルパス> [メッセージID] [内部日付]');
        process.exit(1);
    }
    
    const filePath = args[0];
    const messageId = args[1] || 'test-message-id';
    const internalDate = args[2] || Date.now().toString();
    const uniqueId = `${messageId}-${internalDate}`;
    
    console.log(`ファイル: ${filePath}`);
    console.log(`メッセージID: ${messageId}`);
    console.log(`内部日付: ${internalDate}`);
    console.log(`一意のID: ${uniqueId}`);
    
    processSingleExcelFile(filePath, uniqueId)
        .then(() => console.log('処理が完了しました。'))
        .catch(err => console.error(`エラーが発生しました: ${err.message}`));
}
async function processSingleExcelFile(filePath, uniqueId) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // C1セルから集荷日を取得
    const pickupDateCell = 'C1';
    let pickupDate = '';
    
    if (worksheet[pickupDateCell]) {
        const cell = worksheet[pickupDateCell];
        // セルの詳細情報をログ出力
        console.log('C1セルの詳細情報:');
        console.log(`  型(t): ${cell.t}`);
        console.log(`  値(v): ${cell.v}`);
        console.log(`  表示形式(w): ${cell.w}`);
        console.log(`  値の型: ${typeof cell.v}`);
        
        // Excelの日付セルかどうかを確認
        if (cell.t === 'd') {
            // 既にJavaScript Dateオブジェクトに変換されている場合
            const date = cell.v;
            pickupDate = `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(2, '0')}/${String(date.getDate()).padStart(2, '0')}`;
            console.log(`日付オブジェクトとして処理: ${pickupDate}`);
        } else if (cell.w && cell.w.includes('/')) {
            // 表示形式が日付の場合（例：2023/04/01）
            // 日本の日付形式（YYYY/MM/DD）に変換
            const parts = cell.w.split('/');
            if (parts.length === 3) {
                // 年が2桁の場合は2000年代として解釈
                let year = parseInt(parts[2], 10);
                if (year < 100) {
                    year = 2000 + year;
                }
                pickupDate = `${year}/${parts[0].padStart(2, '0')}/${parts[1].padStart(2, '0')}`;
                console.log(`表示形式から処理: 元の形式=${cell.w}, 変換後=${pickupDate}`);
            } else {
                pickupDate = cell.w;
                console.log(`表示形式をそのまま使用: ${pickupDate}`);
            }
        } else if (typeof cell.v === 'number') {
            // 数値の場合はExcelのシリアル値として扱い、日付に変換
            const date = xlsx.SSF.parse_date_code(cell.v);
            // 年が2桁の場合は2000年代として解釈し、4桁に変換
            const year = date.y < 100 ? 2000 + date.y : date.y;
            pickupDate = `${year}/${String(date.m).padStart(2, '0')}/${String(date.d).padStart(2, '0')}`;
            console.log(`日付変換: シリアル値=${cell.v}, 解析結果=${date.y}/${date.m}/${date.d}, 変換後=${pickupDate}`);
        } else {
            // その他の場合は文字列としてそのまま使用
            pickupDate = cell.v.toString();
            console.log(`文字列として処理: ${pickupDate}`);
        }
    }
    
    console.log(`集荷日: ${pickupDate}`);
    
    // ヘッダー行をスキップして2行目以降のデータを取得
    // headerオプションを'A'に設定して列名をA, B, C, ...として取得
    // range: 2を指定して2行目以降のデータを取得（ヘッダー行をスキップ）
    const rawData = xlsx.utils.sheet_to_json(worksheet, { range: 2, header: 'A' });
    console.log(`Processing file: ${filePath}`);
    console.log(rawData);
    
    // 運賃マスタを読み込む
    const freightMaster = loadFreightMaster();
    
    // 配送先・集荷日ごとの合計ケース数を集計（集荷日はC1セル値を使用）
    const caseCountMap = {};
    rawData.forEach(row => {
        const deliveryName = row['H'] || '';
        const quantity = parseInt(row['G']) || 0;
        if (!deliveryName || !pickupDate) return;
        if (!caseCountMap[deliveryName]) caseCountMap[deliveryName] = {};
        caseCountMap[deliveryName][pickupDate] = (caseCountMap[deliveryName][pickupDate] || 0) + quantity;
    });

    // フォーマット仕様に従ってデータを整形
    const formattedData = rawData.map(row => {
        const temperatureZone = row['B'] || '';
        const productCode = row['C'] || '';
        const productName = row['D'] || '';
        const rank = row['E'] || '';
        const weight = row['F'] || '';
        const quantity = row['G'] || '';
        const deliveryName = row['H'] || '';

        // ケース数合計取得（集荷日はC1セル値を使用）
        const totalCases = (caseCountMap[deliveryName] && caseCountMap[deliveryName][pickupDate]) ? caseCountMap[deliveryName][pickupDate] : 0;
        console.log(`totalCases: ${totalCases}`);
        // 運賃を計算
        let freight = 0;
        if (deliveryName && freightMaster[deliveryName]) {
            const rankStr = rank ? rank.toString().toUpperCase() : '';
            if (totalCases < 5) {
                // 5ケース未満はLow料金
                if (rankStr === 'A') {
                    freight = freightMaster[deliveryName].aRankFeeLow || 0;
                } else if (rankStr === 'B') {
                    freight = freightMaster[deliveryName].bRankFeeLow || 0;
                } else if (rankStr === 'C') {
                    freight = freightMaster[deliveryName].cRankFeeLow || 0;
                } else if (rankStr === 'D') {
                    freight = freightMaster[deliveryName].dRankFeeLow || 0;
                } else {
                    freight = freightMaster[deliveryName].bRankFeeLow || 0;
                }
            } else {
                // 5ケース以上は通常料金
                if (rankStr === 'A') {
                    freight = freightMaster[deliveryName].aRankFee || 0;
                } else if (rankStr === 'B') {
                    freight = freightMaster[deliveryName].bRankFee || 0;
                } else if (rankStr === 'C') {
                    freight = freightMaster[deliveryName].cRankFee || 0;
                } else if (rankStr === 'D') {
                    freight = freightMaster[deliveryName].dRankFee || 0;
                } else {
                    freight = freightMaster[deliveryName].bRankFee || 0;
                }
            }
        } else {
            console.log(`配送先 "${deliveryName}" の運賃情報が見つかりません`);
        }

        return [
            pickupDate,
            temperatureZone,
            productCode,
            productName,
            rank,
            weight,
            quantity,
            deliveryName,
            freight
        ];
    });
    
    console.log('整形後のデータ:');
    console.log(formattedData);
const auth = new google.auth.GoogleAuth({
    keyFile: path.join(process.cwd(), 'config', 'ff01-455323-24aa6cec6617.json'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({version: 'v4', auth});

// スプレッドシートの最終行を取得
const getLastRow = async () => {
    try {
        console.log(`使用するシート名: ${TARGET_SHEET_NAME}`);
        console.log(`使用するスプレッドシートID: ${SPREADSHEET_ID}`);
        
        // スプレッドシートの情報を取得して、シートの存在を確認
        const spreadsheet = await sheets.spreadsheets.get({
            spreadsheetId: SPREADSHEET_ID
        });
        
        // スプレッドシートの全シート名を表示
        console.log('スプレッドシートの全シート:');
        spreadsheet.data.sheets.forEach(sheet => {
            console.log(`- ${sheet.properties.title}`);
        });
        
        // 指定したシート名が存在するか確認
        const sheetExists = spreadsheet.data.sheets.some(sheet => 
            sheet.properties.title === TARGET_SHEET_NAME
        );
        
        if (!sheetExists) {
            console.error(`エラー: シート名 "${TARGET_SHEET_NAME}" はスプレッドシートに存在しません。`);
            console.error('利用可能なシート:');
            spreadsheet.data.sheets.forEach(sheet => {
                console.error(`- ${sheet.properties.title}`);
            });
            throw new Error(`シート名 "${TARGET_SHEET_NAME}" が見つかりません`);
        }
        
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${TARGET_SHEET_NAME}!A:A`,
        });
        return response.data.values ? response.data.values.length : 0;
    } catch (err) {
        console.error(`最終行の取得中にエラーが発生しました: ${err.message}`);
        if (err.response) {
            console.error('エラーレスポンス:', JSON.stringify(err.response.data, null, 2));
        }
        return 0;
    }
};

try {
    const lastRow = await getLastRow();
    
    const request = {
        spreadsheetId: SPREADSHEET_ID,
        range: `${TARGET_SHEET_NAME}!A${lastRow + 1}:I${lastRow + formattedData.length}`, // I列まで（運賃）
        valueInputOption: 'RAW',
        resource: {
            values: formattedData,
        },
    };
    
    console.log(`スプレッドシート更新リクエスト:`, JSON.stringify(request, null, 2));
    
    try {
        const response = await sheets.spreadsheets.values.update(request);
        console.log(`スプレッドシート更新成功: ${filePath}`);
        console.log(`更新された行数: ${response.data.updatedRows}`);
        
        // 処理が完全に成功した場合のみ、処理済みファイルを保存
        if (uniqueId) {
            saveProcessedSheet(uniqueId);
            const fileName = uniqueId.split('-')[2] || 'unknown';
            console.log(`ファイル: ${fileName} (メールID: ${uniqueId.split('-')[0]}) を処理済みとして記録しました。`);
        }
    } catch (updateErr) {
        console.error(`スプレッドシート更新中にエラーが発生しました: ${updateErr.message}`);
        if (updateErr.response) {
            console.error('エラーレスポンス:', JSON.stringify(updateErr.response.data, null, 2));
        }
        throw updateErr; // エラーを再スローして呼び出し元に伝播
    }
} catch (err) {
    console.error(`スプレッドシート処理中にエラーが発生しました: ${err.message}`);
    if (err.response) {
        console.error('エラーレスポンス:', JSON.stringify(err.response.data, null, 2));
    }
    throw err; // エラーを再スローして呼び出し元に伝播
}
}
