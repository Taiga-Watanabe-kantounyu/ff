const fs = require('fs').promises;
const fsSync = require('fs');
const path = require('path');
const process = require('process');
const xlsx = require('xlsx');
const {authenticate} = require('@google-cloud/local-auth');
const { processExcelAndUpdateSheet } = require('./excelProcessor');
const { processOrderFile } = require('./orderFileGenerator');
const {google} = require('googleapis');
const config = require('../../config/config');
const { sendErrorMail } = require('../utils/sendErrorMail');

// If modifying these scopes, delete token.json.
const SCOPES = [
  'https://www.googleapis.com/auth/gmail.readonly',
  'https://www.googleapis.com/auth/gmail.send',
  'https://mail.google.com/'
];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.join(process.cwd(), 'config', 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'config', 'creds.json');
const LAST_CHECK_PATH = path.join(process.cwd(), 'data', 'last_mail_check.json');

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * Serializes credentials to a file compatible with GoogleAUth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * Load or request or authorization to call APIs.
 *
 */
async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}
// you can also use "me" instead of an email addy

/**
 * Watches a mailbox for changes
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
async function listMessages(auth, query) {
    const gmail = google.gmail({version: 'v1', auth});
    const res = await gmail.users.messages.list({
        userId: 'me',
        q: query,
    });
    return res.data.messages || [];
}

async function getMessage(auth, messageId) {
    const gmail = google.gmail({version: 'v1', auth});
    const res = await gmail.users.messages.get({
        userId: 'me',
        id: messageId,
    });
    return res.data;
}

/**
 * 同名のファイルが存在する場合に一意のファイル名を生成する
 * @param {string} originalPath - 元のファイルパス
 * @param {string} messageId - メッセージID
 * @param {string} internalDate - 内部日付
 * @returns {string} - 一意のファイルパス
 */
async function generateUniqueFilePath(originalPath, messageId, internalDate) {
    const dir = path.dirname(originalPath);
    const ext = path.extname(originalPath);
    const baseName = path.basename(originalPath, ext);
    
    // メッセージIDの最後の8文字を使用
    const shortMessageId = messageId.slice(-8);
    
    // 内部日付からタイムスタンプを生成（ミリ秒の最後の6桁）
    const timestamp = internalDate.slice(-6);
    
    // 一意のファイル名を生成
    const uniqueFileName = `${baseName}_${shortMessageId}_${timestamp}${ext}`;
    const uniqueFilePath = path.join(dir, uniqueFileName);
    
    return uniqueFilePath;
}

async function processAttachments(auth, message) {
    const gmail = google.gmail({version: 'v1', auth});
    const parts = message.payload.parts;
    for (const part of parts) {
        if (part.filename && part.body && part.body.attachmentId) {
            const attachment = await gmail.users.messages.attachments.get({
                userId: 'me',
                messageId: message.id,
                id: part.body.attachmentId,
            });
            const data = attachment.data.data;
            const buffer = Buffer.from(data, 'base64');
            
            // 元のファイルパスを生成
            const originalFilePath = path.join(process.cwd(), 'orders', part.filename);
            
            // 一意のファイルパスを生成
            const uniqueFilePath = await generateUniqueFilePath(originalFilePath, message.id, message.internalDate);
            
            // ファイルを保存
            await fs.writeFile(uniqueFilePath, buffer);
            console.log(`Attachment saved as: ${uniqueFilePath}`);
            
            // Excelファイルを処理し、処理結果を取得
            const processed = await processExcelAndUpdateSheet(uniqueFilePath, message.id, message.internalDate);
            
            // 処理が実行された場合のみ、発注用ファイルを生成し、印刷する
            if (processed) {
                try {
                    const orderFilePaths = await processOrderFile(uniqueFilePath, true); // 第2引数のtrueは印刷を有効にする
                    console.log(`発注用ファイルを生成し、印刷しました: ${orderFilePaths.length}件`);
                    orderFilePaths.forEach((filePath, index) => {
                        console.log(`発注ファイル ${index + 1}: ${filePath}`);
                    });
                    sendErrorMail(
                      '【発注書印刷完了通知】',
                      `発注書の印刷が完了しました。\n件数: ${orderFilePaths.length}\nファイル一覧:\n${orderFilePaths.join('\n')}`,
                      undefined,
                      globalAuth
                    );
                } catch (error) {
                    console.error(`発注用ファイル生成・印刷中にエラーが発生しました: ${error.message}`);
                }
            }
        }
    }
}

async function readExcelFile(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);
    console.log(data);
}

/**
 * 最後にメールチェックを行った時間を取得する
 * @returns {string} - 最後にメールチェックを行った時間（YYYY/MM/DD形式）
 */
async function getLastCheckTime() {
    try {
        // ファイルが存在するか確認
        if (!fsSync.existsSync(LAST_CHECK_PATH)) {
            // 存在しない場合は初期値として30日前の日付を設定
            const defaultDate = new Date();
            defaultDate.setDate(defaultDate.getDate() - 30);
            const defaultDateStr = formatDateForQuery(defaultDate);
            
            await fs.writeFile(LAST_CHECK_PATH, JSON.stringify({ lastCheck: defaultDateStr }));
            console.log(`初期の最終チェック日時を設定しました: ${defaultDateStr}`);
            return defaultDateStr;
        }
        
        // ファイルから最後のチェック時間を読み込む
        const content = await fs.readFile(LAST_CHECK_PATH, 'utf8');
        const data = JSON.parse(content);
        console.log(`前回のメールチェック日時: ${data.lastCheck}`);
        return data.lastCheck;
    } catch (err) {
        console.error('最終チェック時間の取得中にエラーが発生しました:', err);
        // エラーが発生した場合は30日前の日付を返す
        const defaultDate = new Date();
        defaultDate.setDate(defaultDate.getDate() - 30);
        return formatDateForQuery(defaultDate);
    }
}

/**
 * 最後にメールチェックを行った時間を更新する
 */
async function updateLastCheckTime() {
    try {
        const now = new Date();
        const dateStr = formatDateForQuery(now);
        await fs.writeFile(LAST_CHECK_PATH, JSON.stringify({ lastCheck: dateStr }));
        console.log(`最終チェック日時を更新しました: ${dateStr}`);
    } catch (err) {
        console.error('最終チェック時間の更新中にエラーが発生しました:', err);
    }
}

/**
 * 日付をGmailクエリ用の形式（YYYY/MM/DD）に変換する
 * @param {Date} date - 変換する日付
 * @returns {string} - YYYY/MM/DD形式の日付文字列
 */
function formatDateForQuery(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}/${month}/${day}`;
}

async function watchUser(auth) {
    const gmail = google.gmail({version: 'v1', auth});
    const res = await gmail.users.watch({
        userId: 'me',
        topicName: 'projects/ff01-455323/topics/ff01',
        labelIds: ['INBOX']
    });

    console.log('Watch response:', res.data);

    // Set up a loop to continuously check for new messages
    setInterval(async () => {
        try {
            // 前回のチェック時間を取得
            const lastCheckTime = await getLastCheckTime();
            
            // 前回のチェック時間以降のメールを検索するクエリを作成
            const query = `from:${config.MAIL_SEARCH_CONFIG.FROM_EMAIL} after:${lastCheckTime}`;
            console.log(`メール検索クエリ: ${query}`);
            
            const messages = await listMessages(auth, query);
            console.log(`${messages.length}件の新しいメールを検出しました`);
            
            for (const message of messages) {
                const fullMessage = await getMessage(auth, message.id);
                await processAttachments(auth, fullMessage);
            }
            
            // 現在の時間を最後のチェック時間として記録
            await updateLastCheckTime();
            
            // 請求書生成が有効な場合、設定に応じて請求書を生成
        } catch (error) {
            console.error(`メールチェック中にエラーが発生しました: ${error.message}`);
            sendErrorMail('【Gmail監視エラー】メールチェック中にエラー', `${error.stack || error}`, undefined, globalAuth);
        }
    }, 60000); // Check every 60 seconds
}

async function stopWatchUser(auth) {
    const gmail = google.gmail({version: 'v1', auth});
    const res = await gmail.users.stop ({
        userId: 'me'
    });

    console.log(res);
}

process.on('uncaughtException', (err) => {
  console.error('uncaughtException:', err);
  sendErrorMail('【Gmail監視エラー】uncaughtException発生', `${err.stack || err}`, undefined, globalAuth);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('unhandledRejection:', reason);
  sendErrorMail('【Gmail監視エラー】unhandledRejection発生', `${reason.stack || reason}`, undefined, globalAuth);
});

let globalAuth = null;

authorize().then(async (auth) => {
  globalAuth = auth;
  await watchUser(auth);
}).catch(async (err) => {
  console.error(err);
  await sendErrorMail('【Gmail監視エラー】起動時エラー', `${err.stack || err}`, undefined, globalAuth);
});
