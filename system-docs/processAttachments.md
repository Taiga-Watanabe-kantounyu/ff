# processAttachments 関数の詳細

`processAttachments`関数は、Gmail APIを使用してメールの添付ファイルを取得し、ローカルに保存するためのものです。

## 関数の流れ

1. **Gmail APIの初期化**: 
   ```javascript
   const gmail = google.gmail({version: 'v1', auth});
   ```
   Gmail APIを初期化し、認証情報を使用してAPIにアクセスします。

2. **メールのパーツを取得**: 
   ```javascript
   const parts = message.payload.parts;
   ```
   メールのペイロードからパーツを取得します。メールは複数のパーツに分かれており、各パーツは本文や添付ファイルを含むことがあります。

3. **添付ファイルの処理**: 
   ```javascript
   for (const part of parts) { ... }
   ```
   各パーツをループし、添付ファイルがあるかどうかを確認します。
   ```javascript
   if (part.filename && part.body && part.body.attachmentId) { ... }
   ```
   ファイル名、ボディ、添付ファイルIDが存在する場合に処理を行います。

4. **添付ファイルの取得**: 
   ```javascript
   const attachment = await gmail.users.messages.attachments.get({ ... });
   ```
   Gmail APIを使用して添付ファイルを取得します。`attachmentId`を使用して特定の添付ファイルを取得します。

5. **データのデコード**: 
   ```javascript
   const data = attachment.data.data;
   const buffer = Buffer.from(data, 'base64');
   ```
   取得した添付ファイルのデータを取得し、Base64エンコードされたデータをデコードします。

6. **ファイルの保存**: 
   ```javascript
   await fs.writeFile(path.join(process.cwd(), part.filename), buffer);
   ```
   デコードしたデータをローカルファイルとして保存します。ファイル名は添付ファイルの名前を使用します。

7. **保存完了のログ**: 
   ```javascript
   console.log(`Attachment ${part.filename} saved.`);
   ```
   ファイルが保存されたことをログに出力します。

この関数は、メールの各パーツを確認し、添付ファイルがあればそれを取得してローカルに保存する役割を果たします。
