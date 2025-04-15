/**
 * 指定されたファイルを印刷するユーティリティ
 */
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const { PRINTER_CONFIG } = require('../../config/config');

/**
 * 指定されたファイルをプリンタで印刷する
 * @param {string} filePath - 印刷するファイルのパス
 * @returns {Promise<void>} - 印刷処理の完了を表すPromise
 */
async function printFile(filePath) {
  try {
    // 設定ファイルから印刷設定を取得
    const { PRINTER_NAME, ENABLE_PRINTING, COPIES } = PRINTER_CONFIG;
    
    // 印刷が無効になっている場合は処理をスキップ
    if (!ENABLE_PRINTING) {
      console.log('印刷機能は無効になっています。設定ファイルで有効にしてください。');
      return;
    }
    
    // ファイルが存在するか確認
    if (!fs.existsSync(filePath)) {
      throw new Error(`印刷対象ファイルが見つかりません: ${filePath}`);
    }
    
    // ファイルの絶対パスを取得
    const absolutePath = path.resolve(filePath);
    
    // PowerShellスクリプトを構築
    // 印刷後にExcelを閉じるスクリプト
    const psScript = `
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $false
      $workbook = $excel.Workbooks.Open("${absolutePath}")
      $workbook.PrintOut()
      Start-Sleep -Seconds 2
      $workbook.Close($false)
      $excel.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
    `;
    
    // PowerShellコマンドを構築
    const command = `powershell -Command "${psScript}"`;
    
    console.log(`印刷コマンドを実行します（印刷後にExcelを閉じます）`);
    
    // コマンドを実行
    return new Promise((resolve, reject) => {
      exec(command, (error, stdout, stderr) => {
        if (error) {
          console.error(`印刷中にエラーが発生しました: ${error.message}`);
          console.error(`stderr: ${stderr}`);
          reject(error);
          return;
        }
        
        console.log(`ファイルを印刷しました: ${filePath}`);
        console.log(`使用プリンタ: ${PRINTER_NAME}`);
        console.log(`印刷部数: ${COPIES}`);
        
        if (stdout) {
          console.log(`stdout: ${stdout}`);
        }
        
        resolve();
      });
    });
  } catch (error) {
    console.error(`印刷処理中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

module.exports = { printFile };
