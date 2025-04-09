const fs = require('fs');
const path = require('path');
const readline = require('readline');

// お届け先マスタファイルのパス
const DELIVERY_MASTER_PATH = path.join(process.cwd(), 'data', 'freightMaster.json');

/**
 * お届け先マスタを読み込む
 * @returns {Object} - お届け先マスタデータ
 */
function loadDeliveryMaster() {
  try {
    if (fs.existsSync(DELIVERY_MASTER_PATH)) {
      const data = fs.readFileSync(DELIVERY_MASTER_PATH, 'utf8');
      return JSON.parse(data);
    }
    return {};
  } catch (error) {
    console.error(`お届け先マスタの読み込みエラー: ${error.message}`);
    return {};
  }
}

/**
 * お届け先マスタを保存する
 * @param {Object} masterData - お届け先マスタデータ
 */
function saveDeliveryMaster(masterData) {
  try {
    fs.writeFileSync(DELIVERY_MASTER_PATH, JSON.stringify(masterData, null, 2));
    console.log('お届け先マスタを保存しました。');
  } catch (error) {
    console.error(`お届け先マスタの保存エラー: ${error.message}`);
  }
}

/**
 * お届け先情報を追加または更新する
 * @param {string} name - お届け先名
 * @param {string} phone - 電話番号
 * @param {string} address - 住所
 * @param {string} leadTime - リードタイム（例: "D+1"）
 * @param {number} bRankFee - Bランク料金
 * @param {number} aRankFee - Aランク料金
 */
function addOrUpdateDeliveryInfo(name, phone, address, leadTime = "D+1", bRankFee = 0, aRankFee = 0) {
  const masterData = loadDeliveryMaster();
  masterData[name] = { 
    address, 
    leadTime, 
    bRankFee, 
    aRankFee, 
    phone 
  };
  saveDeliveryMaster(masterData);
  console.log(`お届け先「${name}」の情報を追加/更新しました。`);
}

/**
 * お届け先情報を削除する
 * @param {string} name - お届け先名
 */
function deleteDeliveryInfo(name) {
  const masterData = loadDeliveryMaster();
  if (masterData[name]) {
    delete masterData[name];
    saveDeliveryMaster(masterData);
    console.log(`お届け先「${name}」の情報を削除しました。`);
  } else {
    console.log(`お届け先「${name}」は見つかりませんでした。`);
  }
}

/**
 * お届け先マスタの一覧を表示する
 */
function listDeliveryMaster() {
  const masterData = loadDeliveryMaster();
  const entries = Object.entries(masterData);
  
  if (entries.length === 0) {
    console.log('お届け先マスタにデータがありません。');
    return;
  }
  
  console.log('===== お届け先マスタ一覧 =====');
  entries.forEach(([name, info], index) => {
    console.log(`${index + 1}. ${name}`);
    console.log(`   電話番号: ${info.phone}`);
    console.log(`   住所: ${info.address}`);
    console.log(`   リードタイム: ${info.leadTime}`);
    console.log(`   Bランク料金: ${info.bRankFee}`);
    console.log(`   Aランク料金: ${info.aRankFee}`);
    console.log('----------------------------');
  });
}

/**
 * お届け先情報を検索する
 * @param {string} keyword - 検索キーワード
 */
function searchDeliveryInfo(keyword) {
  const masterData = loadDeliveryMaster();
  const entries = Object.entries(masterData);
  const results = entries.filter(([name, info]) => {
    return name.includes(keyword) || 
           info.phone.includes(keyword) || 
           info.address.includes(keyword);
  });
  
  if (results.length === 0) {
    console.log(`キーワード「${keyword}」に一致するお届け先情報は見つかりませんでした。`);
    return;
  }
  
  console.log(`===== 検索結果: "${keyword}" =====`);
  results.forEach(([name, info], index) => {
    console.log(`${index + 1}. ${name}`);
    console.log(`   電話番号: ${info.phone}`);
    console.log(`   住所: ${info.address}`);
    console.log(`   リードタイム: ${info.leadTime}`);
    console.log(`   Bランク料金: ${info.bRankFee}`);
    console.log(`   Aランク料金: ${info.aRankFee}`);
    console.log('----------------------------');
  });
}

/**
 * インタラクティブモードでお届け先マスタを管理する
 */
async function interactiveMode() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  
  const question = (query) => new Promise((resolve) => rl.question(query, resolve));
  
  try {
    while (true) {
      console.log('\n===== お届け先マスタ管理 =====');
      console.log('1. 一覧表示');
      console.log('2. 追加/更新');
      console.log('3. 削除');
      console.log('4. 検索');
      console.log('0. 終了');
      
      const choice = await question('操作を選択してください (0-4): ');
      
      switch (choice) {
        case '0':
          console.log('プログラムを終了します。');
          rl.close();
          return;
          
        case '1':
          listDeliveryMaster();
          break;
          
        case '2':
          const name = await question('お届け先名: ');
          const phone = await question('電話番号: ');
          const address = await question('住所: ');
          const leadTime = await question('リードタイム (例: D+1): ');
          const bRankFee = await question('Bランク料金: ');
          const aRankFee = await question('Aランク料金: ');
          addOrUpdateDeliveryInfo(name, phone, address, leadTime, parseInt(bRankFee), parseInt(aRankFee));
          break;
          
        case '3':
          const deleteName = await question('削除するお届け先名: ');
          deleteDeliveryInfo(deleteName);
          break;
          
        case '4':
          const keyword = await question('検索キーワード: ');
          searchDeliveryInfo(keyword);
          break;
          
        default:
          console.log('無効な選択です。もう一度お試しください。');
      }
    }
  } catch (error) {
    console.error(`エラーが発生しました: ${error.message}`);
  } finally {
    rl.close();
  }
}

// コマンドラインから直接実行された場合の処理
if (require.main === module) {
  const args = process.argv.slice(2);
  
  if (args.length === 0) {
    // 引数がない場合はインタラクティブモードを起動
    interactiveMode();
  } else {
    const command = args[0];
    
    switch (command) {
      case 'list':
        listDeliveryMaster();
        break;
        
      case 'add':
        if (args.length < 4) {
          console.log('使用方法: node deliveryMasterManager.js add <お届け先名> <電話番号> <住所> [リードタイム] [Bランク料金] [Aランク料金]');
          process.exit(1);
        }
        const leadTime = args.length > 4 ? args[4] : "D+1";
        const bRankFee = args.length > 5 ? parseInt(args[5]) : 0;
        const aRankFee = args.length > 6 ? parseInt(args[6]) : 0;
        addOrUpdateDeliveryInfo(args[1], args[2], args[3], leadTime, bRankFee, aRankFee);
        break;
        
      case 'delete':
        if (args.length < 2) {
          console.log('使用方法: node deliveryMasterManager.js delete <お届け先名>');
          process.exit(1);
        }
        deleteDeliveryInfo(args[1]);
        break;
        
      case 'search':
        if (args.length < 2) {
          console.log('使用方法: node deliveryMasterManager.js search <キーワード>');
          process.exit(1);
        }
        searchDeliveryInfo(args[1]);
        break;
        
      default:
        console.log('使用方法:');
        console.log('  node deliveryMasterManager.js                     - インタラクティブモード');
        console.log('  node deliveryMasterManager.js list                - 一覧表示');
        console.log('  node deliveryMasterManager.js add <名前> <電話> <住所> [リードタイム] [Bランク料金] [Aランク料金] - 追加/更新');
        console.log('  node deliveryMasterManager.js delete <名前>       - 削除');
        console.log('  node deliveryMasterManager.js search <キーワード> - 検索');
        process.exit(1);
    }
  }
}

module.exports = {
  loadDeliveryMaster,
  addOrUpdateDeliveryInfo,
  deleteDeliveryInfo,
  listDeliveryMaster,
  searchDeliveryInfo
};
