const xlsx = require('xlsx');

// 受注ファイルのパス
const orderFilePath = './orders/依頼書フォーマット_テスト.xlsx';

// Excelファイルを読み込む
const workbook = xlsx.readFile(orderFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// シートの内容を表示
const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
console.log('シート名:', sheetName);
console.log('データ構造:');
data.forEach((row, index) => {
  console.log(`行 ${index + 1}:`, row);
});

// C1セルの内容を確認
const c1Cell = worksheet['C1'];
console.log('\nC1セルの内容:');
console.log('値:', c1Cell ? c1Cell.v : 'なし');
console.log('タイプ:', c1Cell ? c1Cell.t : 'なし');
console.log('表示形式:', c1Cell ? c1Cell.w : 'なし');

// 2行目以降のデータを取得
const rowData = xlsx.utils.sheet_to_json(worksheet, { range: 1 });
console.log('\n2行目以降のデータ:');
rowData.forEach((row, index) => {
  console.log(`データ ${index + 1}:`, row);
  console.log('  Object.values:', Object.values(row));
});
