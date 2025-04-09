const fs = require('fs');
const path = require('path');

// CSVファイルのパス
const CSV_FILE_PATH = path.join(process.cwd(), 'data', 'ffmasta.csv');
// 出力するJSONファイルのパス
const OUTPUT_JSON_PATH = path.join(process.cwd(), 'data', 'freightMaster.json');

/**
 * CSVファイルをJSONに変換する
 * @param {string} csvFilePath - CSVファイルのパス
 * @returns {Object} - 変換されたJSONオブジェクト
 */
function convertCsvToJson(csvFilePath) {
  try {
    // CSVファイルを読み込む
    const csvData = fs.readFileSync(csvFilePath, 'utf8');
    
    // BOMを削除
    const cleanedCsvData = csvData.replace(/^\uFEFF/, '');
    
    // 行に分割
    const rows = cleanedCsvData.split('\n');
    
    // ヘッダー行を取得
    const headers = rows[0].split(',');
    
    // データ行を処理
    const result = {};
    
    for (let i = 1; i < rows.length; i++) {
      // 空行をスキップ
      if (!rows[i].trim()) continue;
      
      const values = rows[i].split(',');
      const deliveryName = values[0].trim();
      
      if (deliveryName) {
        // 各列のインデックスを取得
        const addressIndex = headers.findIndex(header => header.includes('住所'));
        const leadTimeIndex = headers.findIndex(header => header.includes('リードタイム'));
        const bRankFeeIndex = headers.findIndex(header => header.includes('Bランク料金'));
        const aRankFeeIndex = headers.findIndex(header => header.includes('Aランク料金'));
        const phoneIndex = headers.findIndex(header => header.includes('電話'));
        
        // 納品先名をキーとして、すべての情報を値として設定
        result[deliveryName] = {
          address: addressIndex >= 0 && addressIndex < values.length ? values[addressIndex].trim() : "",
          leadTime: leadTimeIndex >= 0 && leadTimeIndex < values.length ? values[leadTimeIndex].trim() : "",
          bRankFee: bRankFeeIndex >= 0 && bRankFeeIndex < values.length ? parseInt(values[bRankFeeIndex].trim(), 10) || 0 : 0,
          aRankFee: aRankFeeIndex >= 0 && aRankFeeIndex < values.length ? parseInt(values[aRankFeeIndex].trim(), 10) || 0 : 0,
          phone: phoneIndex >= 0 && phoneIndex < values.length ? values[phoneIndex].trim() : ""
        };
      }
    }
    
    return result;
  } catch (error) {
    console.error(`CSVの変換中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * メイン処理
 */
function main() {
  try {
    // CSVをJSONに変換
    const jsonData = convertCsvToJson(CSV_FILE_PATH);
    
    // JSONファイルに書き込む
    fs.writeFileSync(OUTPUT_JSON_PATH, JSON.stringify(jsonData, null, 2));
    
    console.log(`JSONファイルを生成しました: ${OUTPUT_JSON_PATH}`);
    console.log(`登録された納品先数: ${Object.keys(jsonData).length}`);
    
    // 変換結果の一部を表示
    const entries = Object.entries(jsonData);
    if (entries.length > 0) {
      console.log('\n変換結果のサンプル:');
      console.log(JSON.stringify(entries[0][1], null, 2));
    }
  } catch (error) {
    console.error(`処理中にエラーが発生しました: ${error.message}`);
    process.exit(1);
  }
}

// スクリプトを実行
main();
