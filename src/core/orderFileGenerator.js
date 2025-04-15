const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const xlsx = require('xlsx'); // xlsxライブラリも使用（データ構造の解析用）
const { printFile } = require('../utils/printFile'); // 印刷機能を追加

// 発注用テンプレートファイルのパス
const TEMPLATE_FILE_PATH = path.join(process.cwd(), 'templates', 'saito_irai.xlsx');
// お届け先マスタファイルのパス
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

const axios = require('axios');

/**
 * 日付が日曜日かどうかを判定する
 * @param {Date} date - 判定する日付
 * @returns {boolean} - 日曜日の場合はtrue、それ以外はfalse
 */
function isSunday(date) {
  return date.getDay() === 0; // 0は日曜日
}

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
 * @returns {Promise<boolean>} - 祝日の場合はtrue、それ以外はfalse
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
 * 日付文字列を年、月、日に分解する
 * @param {string} dateStr - YYYY/MM/DD形式の日付文字列
 * @returns {Object} - 年、月、日を含むオブジェクト
 */
function splitDate(dateStr) {
  const parts = dateStr.split('/');
  return {
    year: parts[0],
    month: parts[1],
    day: parts[2]
  };
}

/**
 * お届け先マスタから情報を取得する
 * @param {string} name - お届け先名
 * @returns {Object} - 電話番号と住所を含むオブジェクト
 */
function getDeliveryInfo(name) {
  try {
    const masterData = JSON.parse(fs.readFileSync(FREIGHT_MASTER_PATH, 'utf8'));
    if (masterData[name]) {
      return masterData[name];
    }
    console.warn(`警告: お届け先「${name}」の情報がマスタに見つかりません。`);
    return { phone: '', address: '' };
  } catch (error) {
    console.error(`お届け先マスタの読み込みエラー: ${error.message}`);
    return { phone: '', address: '' };
  }
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
 * 受注データから値だけを抽出する
 * @param {string} orderFilePath - 受注ファイルのパス
 * @returns {Object} - 集荷日と受注データの配列
 */
function extractOrderData(orderFilePath) {
  try {
    // xlsxライブラリを使用して受注データの構造を解析（値だけを取得）
    const orderWorkbookXlsx = xlsx.readFile(orderFilePath, { cellStyles: false, cellFormula: false });
    const sheetNameXlsx = orderWorkbookXlsx.SheetNames[0];
    const worksheetXlsx = orderWorkbookXlsx.Sheets[sheetNameXlsx];
    
    // C1セルから集荷日を取得
    const pickupDateCellXlsx = worksheetXlsx['C1'];
    let pickupDate = '';
    
    if (pickupDateCellXlsx) {
      if (pickupDateCellXlsx.t === 'd') {
        // 既にJavaScript Dateオブジェクトに変換されている場合
        pickupDate = formatDate(pickupDateCellXlsx.v);
      } else if (pickupDateCellXlsx.w && pickupDateCellXlsx.w.includes('/')) {
        // 表示形式が日付の場合（例：2023/04/01）
        const parts = pickupDateCellXlsx.w.split('/');
        if (parts.length === 3) {
          // 年が2桁の場合は2000年代として解釈
          let year = parseInt(parts[2], 10);
          if (year < 100) {
            year = 2000 + year;
          }
          pickupDate = `${year}/${parts[0].padStart(2, '0')}/${parts[1].padStart(2, '0')}`;
        } else {
          pickupDate = pickupDateCellXlsx.w;
        }
      } else if (typeof pickupDateCellXlsx.v === 'number') {
        // 数値の場合はExcelのシリアル値として扱い、日付に変換
        const date = xlsx.SSF.parse_date_code(pickupDateCellXlsx.v);
        // 年が2桁の場合は2000年代として解釈し、4桁に変換
        const year = date.y < 100 ? 2000 + date.y : date.y;
        pickupDate = `${year}/${String(date.m).padStart(2, '0')}/${String(date.d).padStart(2, '0')}`;
      } else {
        // その他の場合は文字列としてそのまま使用
        pickupDate = pickupDateCellXlsx.v.toString();
      }
    }
    
    // 2行目以降のデータを取得（キー名付きで）- 値だけを取得
    const rawData = xlsx.utils.sheet_to_json(worksheetXlsx, { range: 1 });
    
    return { pickupDate, rawData };
  } catch (error) {
    console.error(`受注データの抽出中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * 受注データから配送日ごとにグループ化したデータを生成する
 * @param {Array} rawData - 受注データの配列
 * @param {string} pickupDate - 集荷日（YYYY/MM/DD形式）
 * @param {Object} leadTimeMaster - リードタイムマスタ
 * @returns {Promise<Object>} - 配送日をキー、注文データの配列を値とするオブジェクト
 */
async function groupOrdersByDeliveryDate(rawData, pickupDate, leadTimeMaster) {
  // 集荷日をDateオブジェクトに変換
  const parts = pickupDate.split('/');
  const pickupDateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  
  // 配送日ごとにデータをグループ化
  const groupedData = {};
  
  for (const row of rawData) {
    const deliveryName = row['お届け先名'] || '';
    
    // リードタイムを取得して配送日を計算
    let deliveryDate = new Date(pickupDateObj); // デフォルトは集荷日と同じ（コピーを作成）
    let deliveryDateStr = pickupDate; // 文字列形式のデフォルト値
    
    if (deliveryName && leadTimeMaster[deliveryName]) {
      const leadTimeDays = leadTimeMaster[deliveryName];
      // 日曜・祝日をスキップしてリードタイムを加算
      deliveryDate = await addDays(pickupDateObj, leadTimeDays);
      deliveryDateStr = formatDate(deliveryDate);
      console.log(`${deliveryName} - 集荷日 ${pickupDate} + リードタイム ${leadTimeDays}日 = 配送日 ${deliveryDateStr} (日曜・祝日はスキップ)`);
    } else {
      console.warn(`警告: お届け先「${deliveryName}」のリードタイム情報が見つかりません。集荷日をそのまま使用します。`);
    }
    
    // 配送日ごとにデータをグループ化
    if (!groupedData[deliveryDateStr]) {
      groupedData[deliveryDateStr] = [];
    }
    
    // 注文データに配送日情報を追加
    const orderWithDeliveryDate = {
      ...row,
      deliveryDate,
      deliveryDateStr
    };
    
    groupedData[deliveryDateStr].push(orderWithDeliveryDate);
  }
  
  return groupedData;
}

/**
 * 指定された配送日の注文データから発注用ファイルを生成する
 * @param {Array} orders - 特定の配送日の注文データ配列
 * @param {string} pickupDate - 集荷日（YYYY/MM/DD形式）
 * @param {string} deliveryDateStr - 配送日（YYYY/MM/DD形式）
 * @returns {Promise<string>} - 生成された発注ファイルのパス
 */
async function createOrderFileForDeliveryDate(orders, pickupDate, deliveryDateStr) {
  try {
    // テンプレートファイルを新しいワークブックとしてコピー
    const fs = require('fs');
    const tempFilePath = path.join(process.cwd(), 'templates', 'temp_template.xlsx');
    
    // テンプレートファイルをコピー
    fs.copyFileSync(TEMPLATE_FILE_PATH, tempFilePath);
    
    // コピーしたテンプレートファイルを開く
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(tempFilePath);
    const worksheet = workbook.worksheets[0];
    
    // 本日の日付を設定
    const today = new Date();
    const todayStr = formatDate(today);
    const todayParts = splitDate(todayStr);
    
    // 集荷日を設定
    const pickupDateParts = splitDate(pickupDate);
    
    // G3セルに本日日付を設定（年月日の形式）
    worksheet.getCell('G3').value = `${todayParts.year}年${todayParts.month}月${todayParts.day}日`;
    
    // G4セルに集荷日を設定（年月日の形式）
    worksheet.getCell('G4').value = `${pickupDateParts.year}年${pickupDateParts.month}月${pickupDateParts.day}日`;
    
    // データを発注用フォーマットに変換
    let rowIndex = 10; // データ記入開始行
    
    // 10行目のスタイル情報を保存（新しい行のテンプレートとして使用）
    const templateRow = worksheet.getRow(10);
    
    for (const order of orders) {
      // 各項目を取得
      const temperatureZone = order['温度帯'] || '';
      const productCode = order['品番'] || '';
      const productName = order['品名'] || '';
      const rank = order['ランク'] || '';
      const weight = order['重量（kg）'] || '';
      const quantity = order['数量'] || '';
      const deliveryName = order['お届け先名'] || '';
      const deliveryDate = order.deliveryDate; // グループ化時に追加した配送日
      
      // お届け先名からマスタ情報を取得
      const deliveryInfo = getDeliveryInfo(deliveryName);
      
      // 現在の行を取得
      const currentRow = worksheet.getRow(rowIndex);
      
      // すべての行に対して、行の高さをテンプレート行と同じに設定
      currentRow.height = templateRow.height;
      
      // 10行目より後の行の場合は、新しい行を挿入し、テンプレート行のスタイルをコピー
      if (rowIndex > 10) {
        // 新しい行を挿入
        worksheet.insertRow(rowIndex, []);
        
        // テンプレート行のスタイルをコピー
        for (let colIndex = 1; colIndex <= 10; colIndex++) {
          const templateCell = templateRow.getCell(colIndex);
          const currentCell = currentRow.getCell(colIndex);
          
          // スタイルを直接コピー
          if (templateCell.style) {
            currentCell.style = JSON.parse(JSON.stringify(templateCell.style));
          }
          
          // フォントスタイルをコピー
          if (templateCell.font) {
            currentCell.font = JSON.parse(JSON.stringify(templateCell.font));
          }
          
          // 罫線スタイルをコピー
          if (templateCell.border) {
            currentCell.border = JSON.parse(JSON.stringify(templateCell.border));
          }
          
          // 塗りつぶしスタイルをコピー
          if (templateCell.fill) {
            currentCell.fill = JSON.parse(JSON.stringify(templateCell.fill));
          }
          
          // 配置スタイルをコピー
          if (templateCell.alignment) {
            currentCell.alignment = JSON.parse(JSON.stringify(templateCell.alignment));
          }
        }
      }
      
      // 値だけを設定
      currentRow.getCell(1).value = formatDateShort(deliveryDate); // A列に配送日をMM/DD形式で設定
      currentRow.getCell(2).value = temperatureZone; // B列 温度帯
      currentRow.getCell(3).value = productName; // C列 品名
      currentRow.getCell(4).value = rank; // D列 ランク
      currentRow.getCell(5).value = weight; // E列 重量（kg）
      currentRow.getCell(6).value = quantity; // F列 数量
      currentRow.getCell(7).value = deliveryInfo.formalName || deliveryName; // G列 お届け先名（正式名称）
      currentRow.getCell(8).value = deliveryInfo.phone; // H列 届け先電話番号
      currentRow.getCell(9).value = deliveryInfo.address; // I列 届け先住所
      
      // 行をコミット
      currentRow.commit();
      
      rowIndex++;
    }
    
    // 一時ファイルを削除
    try {
      fs.unlinkSync(tempFilePath);
    } catch (error) {
      console.warn(`一時ファイルの削除に失敗しました: ${error.message}`);
    }
    
    // 発注ファイル名を生成
    // 配送日ごとに発注ファイルを作成するため、配送日を使用してファイル名を生成
    const deliveryDateForFileName = deliveryDateStr.replace(/\//g, '');
    const outputFileName = `発注_${deliveryDateForFileName}.xlsx`;
    const outputFilePath = path.join(process.cwd(), 'invoices', outputFileName);
    
    // 発注ファイルを保存
    await workbook.xlsx.writeFile(outputFilePath);
    console.log(`発注ファイルを生成しました: ${outputFilePath}`);
    
    return outputFilePath;
  } catch (error) {
    console.error(`発注ファイルの生成中にエラーが発生しました: ${error.message}`);
    console.error(error.stack);
    throw error;
  }
}

/**
 * 受注データから発注用ファイルを生成する
 * @param {string} orderFilePath - 受注ファイルのパス
 * @returns {Promise<string[]>} - 生成された発注ファイルのパスの配列
 */
async function generateOrderFile(orderFilePath) {
  try {
    // 受注データから値だけを抽出
    const { pickupDate, rawData } = extractOrderData(orderFilePath);
    console.log(`集荷日: ${pickupDate}`);
    console.log(`受注データ数: ${rawData.length}`);
    
    // リードタイムマスタを読み込む
    const leadTimeMaster = loadLeadTimeMaster();
    console.log('リードタイムマスタを読み込みました。');
    
    // 配送日ごとにデータをグループ化
    const groupedData = await groupOrdersByDeliveryDate(rawData, pickupDate, leadTimeMaster);
    const deliveryDates = Object.keys(groupedData);
    console.log(`配送日の種類: ${deliveryDates.length}種類`);
    
    // 各配送日ごとに発注ファイルを生成
    const outputFilePaths = [];
    for (const deliveryDateStr of deliveryDates) {
      const orders = groupedData[deliveryDateStr];
      console.log(`配送日 ${deliveryDateStr} の注文数: ${orders.length}`);
      
      const outputFilePath = await createOrderFileForDeliveryDate(orders, pickupDate, deliveryDateStr);
      outputFilePaths.push(outputFilePath);
    }
    
    return outputFilePaths;
  } catch (error) {
    console.error(`発注ファイルの生成中にエラーが発生しました: ${error.message}`);
    console.error(error.stack);
    throw error;
  }
}

/**
 * 発注ファイル生成処理のメイン関数
 * @param {string} orderFilePath - 受注ファイルのパス
 * @param {boolean} printAfterGeneration - 生成後に印刷するかどうか（デフォルトはtrue）
 * @returns {Promise<string[]>} - 生成された発注ファイルのパスの配列
 */
async function processOrderFile(orderFilePath, printAfterGeneration = true) {
  try {
    console.log(`受注ファイルの処理を開始します: ${orderFilePath}`);
    const outputFilePaths = await generateOrderFile(orderFilePath);
    
    // 発注ファイル生成後に印刷処理を実行
    if (printAfterGeneration) {
      for (const filePath of outputFilePaths) {
        try {
          console.log(`発注ファイルの印刷を開始します: ${filePath}`);
          await printFile(filePath);
          console.log(`${filePath} の印刷処理が完了しました`);
        } catch (printError) {
          console.error(`印刷処理中にエラーが発生しました: ${printError.message}`);
          // 印刷エラーがあっても処理は続行し、次のファイルの印刷に進む
        }
      }
    }
    
    return outputFilePaths;
  } catch (error) {
    console.error(`処理中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

// コマンドラインから直接実行された場合の処理
if (require.main === module) {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.log('使用方法: node orderFileGenerator.js <受注ファイルパス>');
    process.exit(1);
  }
  
  const orderFilePath = args[0];
  
  processOrderFile(orderFilePath)
    .then(outputFilePaths => {
      console.log(`処理が完了しました。発注ファイル数: ${outputFilePaths.length}`);
      outputFilePaths.forEach((filePath, index) => {
        console.log(`発注ファイル ${index + 1}: ${filePath}`);
      });
    })
    .catch(err => {
      console.error(`エラーが発生しました: ${err.message}`);
      process.exit(1);
    });
}

module.exports = { processOrderFile };
