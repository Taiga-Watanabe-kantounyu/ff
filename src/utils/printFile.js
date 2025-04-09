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
    
    // PowerShellコマンドを構築
    // Start-Process: 指定されたプログラムを起動
    // -FilePath: 起動するプログラムのパス
    // -Verb Print: 印刷アクションを実行
    // -ArgumentList: プログラムに渡す引数
    const command = `powershell -Command "Start-Process -FilePath '${absolutePath}' -Verb Print -ArgumentList '/d:${PRINTER_NAME}'"`;
    
    console.log(`印刷コマンドを実行します: ${command}`);
    
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
