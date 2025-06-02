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
    const bRankFeeIndex = headers.findIndex(header => header === 'Bランク' || header.includes('Bランク（通常）'));
    const aRankFeeIndex = headers.findIndex(header => header === 'Aランク' || header.includes('Aランク（通常）'));
    const bRankFeeLowIndex = headers.findIndex(header => header.includes('Bランク5ケース以下') || header.includes('Bランク（5ケース以下）'));
    const aRankFeeLowIndex = headers.findIndex(header => header.includes('Aランク5ケース以下') || header.includes('Aランク（5ケース以下）'));
    const cRankFeeIndex = headers.findIndex(header => header === 'Cランク' || header.includes('Cランク（通常）'));
    const cRankFeeLowIndex = headers.findIndex(header => header.includes('Cランク5ケース以下') || header.includes('Cランク（5ケース以下）'));
    const dRankFeeIndex = headers.findIndex(header => header === 'Dランク' || header.includes('Dランク（通常）'));
    const dRankFeeLowIndex = headers.findIndex(header => header.includes('Dランク5ケース以下') || header.includes('Dランク（5ケース以下）'));
    
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
          deliveryName: deliveryName,
          formalName: formalName,
          address: values[2].trim(),
          leadTime: leadTime,
          cRankFee: cRankFeeIndex >= 0 && cRankFeeIndex < values.length ? parseInt(values[cRankFeeIndex].trim()) || 0 : 0,
          bRankFee: bRankFee,
          aRankFee: aRankFee,
          dRankFee: dRankFeeIndex >= 0 && dRankFeeIndex < values.length ? parseInt(values[dRankFeeIndex].trim()) || 0 : 0,
          cRankFeeLow: cRankFeeLowIndex >= 0 && cRankFeeLowIndex < values.length ? parseInt(values[cRankFeeLowIndex].trim()) || 0 : 0,
          bRankFeeLow: bRankFeeLowIndex >= 0 && bRankFeeLowIndex < values.length ? parseInt(values[bRankFeeLowIndex].trim()) || 0 : 0,
          aRankFeeLow: aRankFeeLowIndex >= 0 && aRankFeeLowIndex < values.length ? parseInt(values[aRankFeeLowIndex].trim()) || 0 : 0,
          dRankFeeLow: dRankFeeLowIndex >= 0 && dRankFeeLowIndex < values.length ? parseInt(values[dRankFeeLowIndex].trim()) || 0 : 0,
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
