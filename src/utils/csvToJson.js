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
    
    // 電話番号、正式名称、リードタイム、料金の列インデックスを取得
    const phoneIndex = headers.findIndex(header => header.includes('電話'));
    const formalNameIndex = headers.findIndex(header => header.includes('正式名称'));
    const leadTimeIndex = headers.findIndex(header => header.includes('リードタイム'));
    const bRankFeeIndex = headers.findIndex(header => header.includes('Bランク'));
    const aRankFeeIndex = headers.findIndex(header => header.includes('Aランク'));
    
    // 既存のfreightMaster.jsonを読み込む（存在する場合）
    let existingData = {};
    try {
      const existingJson = fs.readFileSync(OUTPUT_JSON_PATH, 'utf8');
      existingData = JSON.parse(existingJson);
    } catch (error) {
      console.log('既存のJSONファイルが見つからないか、読み込めません。新しいファイルを作成します。');
    }
    
    // データ行を処理
    const result = {};
    
    for (let i = 1; i < rows.length; i++) {
      // 空行をスキップ
      if (!rows[i].trim()) continue;
      
      const values = rows[i].split(',');
      const deliveryName = values[0].trim();
      
      // 納品先名をキーとして、正式名称、住所、電話番号、リードタイム、料金を値として設定
      if (deliveryName) {
        const phone = phoneIndex >= 0 && phoneIndex < values.length ? values[phoneIndex].trim() : "";
        const formalName = formalNameIndex >= 0 && formalNameIndex < values.length ? values[formalNameIndex].trim() : deliveryName;
        const leadTime = leadTimeIndex >= 0 && leadTimeIndex < values.length ? values[leadTimeIndex].trim() : "D+1";
        const bRankFee = bRankFeeIndex >= 0 && bRankFeeIndex < values.length ? parseInt(values[bRankFeeIndex].trim()) || 0 : 0;
        const aRankFee = aRankFeeIndex >= 0 && aRankFeeIndex < values.length ? parseInt(values[aRankFeeIndex].trim()) || 0 : 0;
        
        result[deliveryName] = {
          formalName: formalName,
          address: values[2].trim(), // 正式名称の列が追加されたので、住所は2列目になる
          leadTime: leadTime,
          bRankFee: bRankFee,
          aRankFee: aRankFee,
          phone: phone
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
  } catch (error) {
    console.error(`処理中にエラーが発生しました: ${error.message}`);
    process.exit(1);
  }
}

// スクリプトを実行
main();
