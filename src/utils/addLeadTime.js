const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const ExcelJS = require('exceljs');

// JSONファイルのパス
const FREIGHT_MASTER_PATH = path.join(process.cwd(), 'data', 'freightMaster.json');

/**
 * JSONファイルからリードタイムマスタを読み込む
 * @returns {Object} - 配送先名をキー、リードタイム日数を値とするオブジェクト
 */
function loadLeadTimeMaster() {
  try {
    // JSONファイルを読み込む
    const jsonData = fs.readFileSync(FREIGHT_MASTER_PATH, 'utf8');
    const masterData = JSON.parse(jsonData);
    
    // データを処理
    const result = {};
    
    for (const [deliveryName, data] of Object.entries(masterData)) {
      if (data.leadTime) {
        // "D+1"や"D+2"から数値部分を抽出
        const leadTimeStr = data.leadTime.trim();
        const leadTimeDays = parseInt(leadTimeStr.replace('D+', ''), 10);
        
        if (!isNaN(leadTimeDays)) {
          result[deliveryName] = leadTimeDays;
        }
      }
    }
    
    return result;
  } catch (error) {
    console.error(`リードタイムマスタの読み込み中にエラーが発生しました: ${error.message}`);
    return {};
  }
}

/**
 * 日付が日曜日かどうかを判定する
 * @param {Date} date - 判定する日付
 * @returns {boolean} - 日曜日の場合はtrue、それ以外はfalse
 */
function isSunday(date) {
  return date.getDay() === 0; // 0は日曜日
}

const axios = require('axios');

// 祝日情報のキャッシュ
const holidayCache = {};

/**
 * 指定した年の祝日情報をAPIから取得する
 * @param {number} year - 取得する年
 * @returns {Promise<string[]>} - 祝日のリスト（YYYY/MM/DD形式）
 */
async function fetchHolidays(year) {
  // キャッシュに存在する場合はキャッシュから返す
  if (holidayCache[year]) {
    return holidayCache[year];
  }
  
  try {
    // APIから祝日情報を取得
    const response = await axios.get(`https://holidays-jp.github.io/api/v1/${year}/date.json`);
    const holidayData = response.data;
    
    // 祝日のリストを作成（YYYY/MM/DD形式）
    const holidays = Object.keys(holidayData).map(date => {
      // APIの日付形式は "YYYY-MM-DD" なので "YYYY/MM/DD" に変換
      return date.replace(/-/g, '/');
    });
    
    // キャッシュに保存
    holidayCache[year] = holidays;
    
    return holidays;
  } catch (error) {
    console.error(`祝日情報の取得に失敗しました: ${error.message}`);
    return []; // エラーの場合は空のリストを返す
  }
}

/**
 * 日付が祝日かどうかを判定する
 * @param {Date} date - 判定する日付
 * @returns {boolean} - 祝日の場合はtrue、それ以外はfalse
 */
async function isHoliday(date) {
  const year = date.getFullYear();
  const dateStr = formatDate(date);
  
  try {
    // 祝日情報を取得
    const holidays = await fetchHolidays(year);
    
    // 祝日リストに含まれているかチェック
    return holidays.includes(dateStr);
  } catch (error) {
    console.error(`祝日判定中にエラーが発生しました: ${error.message}`);
    return false; // エラーの場合は祝日ではないと判定
  }
}

/**
 * 日付に日数を加算する（日曜・祝日をスキップ）
 * @param {Date} date - 元の日付
 * @param {number} days - 加算する日数
 * @returns {Promise<Date>} - 加算後の日付
 */
async function addDays(date, days) {
  const result = new Date(date);
  let remainingDays = days;
  
  while (remainingDays > 0) {
    // 1日進める
    result.setDate(result.getDate() + 1);
    
    // 日曜日または祝日でない場合のみカウント
    if (!isSunday(result) && !(await isHoliday(result))) {
      remainingDays--;
    }
  }
  
  return result;
}

/**
 * 日付をYYYY/MM/DD形式に変換する
 * @param {Date} date - 日付オブジェクト
 * @returns {string} - YYYY/MM/DD形式の日付文字列
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

/**
 * 日付をMM/DD形式に変換する
 * @param {Date} date - 日付オブジェクト
 * @returns {string} - MM/DD形式の日付文字列
 */
function formatDateShort(date) {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${month}/${day}`;
}

/**
 * Excelの日付シリアル値をJavaScriptのDateオブジェクトに変換する
 * @param {number} serial - Excelの日付シリアル値
 * @returns {Date} - JavaScriptのDateオブジェクト
 */
function excelSerialDateToJSDate(serial) {
  // Excelの日付の開始日は1900年1月1日（ただし1900年はうるう年ではないのに
  // Excelは1900年2月29日が存在すると誤って扱っているため、1900年3月1日以降の日付は
  // シリアル値から1を引く必要がある）
  let days = serial - (serial > 60 ? 1 : 0);
  const date = new Date(1900, 0, 1);
  date.setDate(date.getDate() + days - 1);
  return date;
}

/**
 * 発注シートのA列の日付にリードタイムを加算する
 * @param {string} filePath - 発注シートのファイルパス
 * @returns {Promise<void>}
 */
async function addLeadTimeToOrderSheet(filePath) {
  try {
    console.log(`ファイル処理開始: ${filePath}`);
    
    // リードタイムマスタを読み込む
    const leadTimeMaster = loadLeadTimeMaster();
    console.log('リードタイムマスタを読み込みました。');
    console.log(leadTimeMaster);
    
    // xlsxライブラリを使用してExcelファイルを読み込む（データ構造の解析用）
    const workbookXlsx = xlsx.readFile(filePath);
    const sheetNameXlsx = workbookXlsx.SheetNames[0];
    const worksheetXlsx = workbookXlsx.Sheets[sheetNameXlsx];
    
    // C1セルから集荷日を取得
    const pickupDateCell = worksheetXlsx['C1'];
    let pickupDate = null;
    
    if (pickupDateCell) {
      console.log('C1セルの詳細情報:');
      console.log(`  型(t): ${pickupDateCell.t}`);
      console.log(`  値(v): ${pickupDateCell.v}`);
      console.log(`  表示形式(w): ${pickupDateCell.w}`);
      console.log(`  値の型: ${typeof pickupDateCell.v}`);
      
      if (pickupDateCell.t === 'd') {
        // 既にJavaScript Dateオブジェクトに変換されている場合
        pickupDate = pickupDateCell.v;
        console.log(`日付オブジェクトとして処理: ${formatDate(pickupDate)}`);
      } else if (pickupDateCell.w && pickupDateCell.w.includes('/')) {
        // 表示形式が日付の場合（例：2023/04/01）
        const parts = pickupDateCell.w.split('/');
        if (parts.length === 3) {
          // 年が2桁の場合は2000年代として解釈
          let year = parseInt(parts[2], 10);
          if (year < 100) {
            year = 2000 + year;
          }
          const month = parseInt(parts[0], 10) - 1; // JavaScriptの月は0から始まる
          const day = parseInt(parts[1], 10);
          pickupDate = new Date(year, month, day);
          console.log(`表示形式から処理: 元の形式=${pickupDateCell.w}, 変換後=${formatDate(pickupDate)}`);
        }
      } else if (typeof pickupDateCell.v === 'number') {
        // 数値の場合はExcelのシリアル値として扱い、日付に変換
        pickupDate = excelSerialDateToJSDate(pickupDateCell.v);
        console.log(`シリアル値から処理: ${pickupDateCell.v} -> ${formatDate(pickupDate)}`);
      }
    }
    
    if (!pickupDate) {
      console.error('C1セルから集荷日を取得できませんでした。処理を中止します。');
      return;
    }
    
    console.log(`集荷日: ${formatDate(pickupDate)}`);
    
    // ExcelJSを使用してExcelファイルを読み込む（編集用）
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];
    
    // 2行目以降のデータを取得（キー名付きで）
    const rawData = xlsx.utils.sheet_to_json(worksheetXlsx, { range: 1 });
    console.log(`データ行数: ${rawData.length}`);
    
    // 各行を処理
    for (let i = 0; i < rawData.length; i++) {
      const rowData = rawData[i];
      const rowIndex = i + 2; // 2行目から開始
      
      // お届け先名を取得
      const deliveryName = rowData['お届け先名'] || '';
      
      // A列のセルを取得
      const cell = worksheet.getCell(`A${rowIndex}`);
      
      // C1から取得した集荷日をA列に設定
      cell.value = pickupDate;
      
      if (deliveryName && leadTimeMaster[deliveryName]) {
        // リードタイムを加算
        const leadTimeDays = leadTimeMaster[deliveryName];
        const deliveryDate = await addDays(pickupDate, leadTimeDays);
        
        // A列に配送日（集荷日+リードタイム）をMM/DD形式で設定
        cell.value = formatDateShort(deliveryDate);
        
        console.log(`行 ${rowIndex}: ${deliveryName} - 集荷日 ${formatDate(pickupDate)} + リードタイム ${leadTimeDays}日 = 配送日 ${formatDate(deliveryDate)} (日曜・祝日はスキップ)`);
      } else {
        console.warn(`行 ${rowIndex}: お届け先「${deliveryName}」のリードタイム情報が見つかりません。集荷日をそのまま設定します。`);
      }
    }
    
    // ファイルを保存
    await workbook.xlsx.writeFile(filePath);
    console.log(`ファイルを更新しました: ${filePath}`);
    
  } catch (error) {
    console.error(`処理中にエラーが発生しました: ${error.message}`);
    console.error(error.stack);
    throw error;
  }
}

/**
 * メイン処理
 */
async function main() {
  try {
    const args = process.argv.slice(2);
    if (args.length < 1) {
      console.log('使用方法: node addLeadTime.js <発注シートのパス>');
      process.exit(1);
    }
    
    const filePath = args[0];
    await addLeadTimeToOrderSheet(filePath);
    
    console.log('処理が完了しました。');
  } catch (error) {
    console.error(`エラーが発生しました: ${error.message}`);
    process.exit(1);
  }
}

// スクリプトが直接実行された場合
if (require.main === module) {
  main();
}

module.exports = { addLeadTimeToOrderSheet };
