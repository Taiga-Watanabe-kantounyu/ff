// src/utils/sendErrorMail.js
const nodemailer = require('nodemailer');
const { google } = require('googleapis');
const config = require('../../config/config');
const NOTIFY_MAIL_CONFIG = config.NOTIFY_MAIL_CONFIG;

/**
 * 通知メールを送信する
 * @param {string} subject - メール件名（省略時はconfigのデフォルト件名）
 * @param {string} text - メール本文
 * @param {string} [to] - 宛先（省略時はconfigのTO）
 * @param {OAuth2Client} [oauth2Client] - google.auth.OAuth2インスタンス（省略時はSMTPパスワード認証）
 * @returns {Promise<void>}
 */
async function sendErrorMail(subject, text, to, oauth2Client) {
  if (!NOTIFY_MAIL_CONFIG.ENABLE_NOTIFICATION) {
    console.log('通知メール送信は無効化されています');
    return;
  }
  const mailTo = to || NOTIFY_MAIL_CONFIG.TO;
  const mailSubject = subject || NOTIFY_MAIL_CONFIG.SUBJECT;

    if (oauth2Client) {
      // OAuth2を使ってGmail APIで送信
      const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
      const encodedSubject = `=?UTF-8?B?${Buffer.from(mailSubject).toString('base64')}?=`;
      const messageLines = [
      `From: ${NOTIFY_MAIL_CONFIG.USER}`,
      `To: ${mailTo}`,
      `Subject: ${encodedSubject}`,
      'MIME-Version: 1.0',
      'Content-Type: text/plain; charset=UTF-8',
      'Content-Transfer-Encoding: 7bit',
      '',
      text
    ];
    const rawMessage = Buffer.from(messageLines.join('\r\n'))
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');
    try {
      await gmail.users.messages.send({
        userId: 'me',
        requestBody: {
          raw: rawMessage
        }
      });
      console.log('通知メールをGmail API経由で送信しました');
    } catch (err) {
      console.error('Gmail APIによる通知メール送信に失敗:', err);
    }
  } else {
    // SMTPパスワード認証で送信
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: NOTIFY_MAIL_CONFIG.USER,
        pass: NOTIFY_MAIL_CONFIG.PASS
      }
    });
    const mailOptions = {
      from: NOTIFY_MAIL_CONFIG.USER,
      to: mailTo,
      subject: mailSubject,
      text
    };
    try {
      await transporter.sendMail(mailOptions);
      console.log('通知メールをSMTP経由で送信しました');
    } catch (err) {
      console.error('SMTPによる通知メール送信に失敗:', err);
    }
  }
}

module.exports = { sendErrorMail };
