const fs = require('fs');
const path = require('path');

/**
 * HTMLテンプレートを読み込み、データで置換して返す
 * @param {string} templatePath - テンプレートファイルのパス
 * @param {Object} data - 置換するデータ
 * @returns {string} - 置換後のHTML文字列
 */
function renderTemplate(templatePath, data) {
  try {
    // テンプレートファイルを読み込む
    let templateContent = fs.readFileSync(templatePath, 'utf8');
    
    // シンプルな置換（{{キー}}形式）
    Object.keys(data).forEach(key => {
      const regex = new RegExp(`{{${key}}}`, 'g');
      templateContent = templateContent.replace(regex, data[key]);
    });
    
    // テーブル行の動的生成（<!-- BEGIN items -->...<!-- END items -->形式）
    if (data.itemsHtml) {
      const itemsRegex = /<!-- BEGIN items -->([\s\S]*?)<!-- END items -->/;
      const itemsMatch = templateContent.match(itemsRegex);
      
      if (itemsMatch && itemsMatch[1]) {
        templateContent = templateContent.replace(itemsMatch[0], data.itemsHtml);
      }
    }
    
    return templateContent;
  } catch (error) {
    console.error(`テンプレートのレンダリング中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * レンダリングしたHTMLをファイルに保存する
 * @param {string} htmlContent - 保存するHTML内容
 * @param {string} outputPath - 出力ファイルパス
 * @returns {string} - 保存したファイルの絶対パス
 */
function saveRenderedHtml(htmlContent, outputPath) {
  try {
    // 出力ディレクトリが存在しない場合は作成
    const outputDir = path.dirname(outputPath);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    // HTMLを保存
    fs.writeFileSync(outputPath, htmlContent, 'utf8');
    
    return path.resolve(outputPath);
  } catch (error) {
    console.error(`HTMLファイルの保存中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

module.exports = {
  renderTemplate,
  saveRenderedHtml
};
