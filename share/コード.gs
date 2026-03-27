/**
 * 📋 請求書自動管理システム — メインスクリプト
 *
 * Gmailに届くPDF請求書を自動で:
 *   1. Google Driveに年月別フォルダで保存
 *   2. Gemini APIでメーカー名・税抜金額を抽出
 *   3. スプレッドシートに一覧記入（縦=メーカー、横=月）
 *
 * 使い方:
 *   - 手動実行: processInvoices() を実行
 *   - 自動実行: setupTrigger() でトリガー設定
 */

// ===== メイン関数 =====

/**
 * メイン処理: Gmailから請求書メールを検索し、PDF抽出→Drive保存→Gemini解析→シート記入
 */
function processInvoices() {
  // API Key チェック
  if (!GEMINI_API_KEY || GEMINI_API_KEY === 'ここにAPIキーを貼り付け') {
    throw new Error('❗ 設定.gs の GEMINI_API_KEY を設定してください。\nGoogle AI Studio: https://aistudio.google.com/apikey');
  }

  // スプレッドシート・フォルダの自動作成
  ensureSpreadsheet_();
  ensureDriveFolder_();

  // 処理済みラベルを取得（なければ作成）
  var label = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!label) label = GmailApp.createLabel(PROCESSED_LABEL);

  // Gmail検索（処理済みラベルを除外 + 開始日以降）
  var query = GMAIL_QUERY + ' -label:' + PROCESSED_LABEL + ' after:' + START_DATE.replace(/\//g, '/');
  var threads = GmailApp.search(query, 0, BATCH_SIZE);

  if (threads.length === 0) {
    Logger.log('✅ 新しい請求書メールはありません');
    return;
  }

  Logger.log('📧 ' + threads.length + '件のメールスレッドを処理します');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var processedCount = 0;
  var errorCount = 0;
  var logSheet = getOrCreateLogSheet_(ss);

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();

    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      var pdfs = extractPdfsFromMessage_(message);

      if (pdfs.length === 0) continue;

      // メールの受信日から年月を取得
      var mailDate = message.getDate();
      var yearMonth = Utilities.formatDate(mailDate, 'Asia/Tokyo', 'yyyy/MM');

      for (var p = 0; p < pdfs.length; p++) {
        try {
          // 1. Driveに保存
          var savedFile = saveToDrive_(pdfs[p], yearMonth);

          // 2. Gemini APIで解析
          var result = analyzeWithGemini_(pdfs[p]);

          if (result && result.makerName && result.amount > 0) {
            // 3. スプレッドシートに記入
            updateSheet_(sheet, result.makerName, result.amount, yearMonth);

            // ログ記録
            logResult_(logSheet, mailDate, message.getSubject(), result.makerName, result.amount, yearMonth, '成功');

            // Driveのファイル名をメーカー名に更新
            savedFile.setName(result.makerName + '_請求書_' + yearMonth.replace('/', '') + '.pdf');

            processedCount++;
            Logger.log('✅ ' + result.makerName + ' ¥' + result.amount.toLocaleString() + ' (' + yearMonth + ')');
          } else {
            // 解析失敗
            logResult_(logSheet, mailDate, message.getSubject(), '不明', 0, yearMonth, '解析失敗: ' + JSON.stringify(result));
            errorCount++;
            Logger.log('⚠️ 解析失敗: ' + message.getSubject());
          }
        } catch (e) {
          errorCount++;
          logResult_(logSheet, mailDate, message.getSubject(), 'エラー', 0, yearMonth, e.toString());
          Logger.log('❌ エラー: ' + e.toString());
        }
      }
    }

    // 処理済みラベルを付与
    threads[t].addLabel(label);
  }

  Logger.log('🎉 完了: 成功 ' + processedCount + '件, エラー ' + errorCount + '件');
}


// ===== PDF抽出 =====

/**
 * メールからPDF添付ファイルを抽出
 * @param {GmailMessage} message - メールメッセージ
 * @return {Blob[]} PDFファイルの配列
 */
function extractPdfsFromMessage_(message) {
  var attachments = message.getAttachments();
  var pdfs = [];
  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    if (att.getContentType() === 'application/pdf' ||
        att.getName().toLowerCase().endsWith('.pdf')) {
      pdfs.push(att.copyBlob());
    }
  }
  return pdfs;
}


// ===== Google Drive保存 =====

/**
 * PDFをDriveの年月別フォルダに保存
 * @param {Blob} pdfBlob - PDFデータ
 * @param {string} yearMonth - 年月（例: "2026/02"）
 * @return {File} 保存されたファイル
 */
function saveToDrive_(pdfBlob, yearMonth) {
  var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  // 年月フォルダを取得 or 作成（例: "2026年02月"）
  var parts = yearMonth.split('/');
  var folderName = parts[0] + '年' + parts[1] + '月';
  var folders = rootFolder.getFoldersByName(folderName);
  var monthFolder;
  if (folders.hasNext()) {
    monthFolder = folders.next();
  } else {
    monthFolder = rootFolder.createFolder(folderName);
  }

  // ファイル保存（一時的な名前、後でメーカー名にリネーム）
  var fileName = pdfBlob.getName() || 'invoice_' + new Date().getTime() + '.pdf';
  return monthFolder.createFile(pdfBlob).setName(fileName);
}


// ===== Gemini API解析 =====

/**
 * Gemini APIでPDFを解析し、メーカー名と税抜金額を抽出
 * @param {Blob} pdfBlob - PDFデータ
 * @return {Object} { makerName: string, amount: number } or null
 */
function analyzeWithGemini_(pdfBlob) {
  var base64 = Utilities.base64Encode(pdfBlob.getBytes());

  var apiBase = (typeof GEMINI_API_URL !== 'undefined' && GEMINI_API_URL) ? GEMINI_API_URL : 'https://generativelanguage.googleapis.com/v1beta/models/';
  var url = apiBase + GEMINI_MODEL + ':generateContent?key=' + GEMINI_API_KEY;

  var payload = {
    contents: [{
      parts: [
        {
          inlineData: {
            mimeType: 'application/pdf',
            data: base64
          }
        },
        {
          text: 'この請求書PDFから以下の情報を抽出してJSON形式で返してください。\n\n' +
                '1. makerName: 請求元の会社名（発行者/売り手の名前）。「株式会社」「(株)」「有限会社」などは除いた短い名前にしてください。\n' +
                '2. amount: 税抜金額（数値のみ、カンマなし）。税抜が不明な場合は合計金額を使ってください。\n\n' +
                '必ず以下のJSON形式のみを返してください（説明文不要）:\n' +
                '{"makerName": "会社名", "amount": 12345}\n\n' +
                'これが請求書でない場合は {"makerName": null, "amount": 0} を返してください。'
        }
      ]
    }],
    generationConfig: {
      temperature: 0.1
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var status = response.getResponseCode();

  if (status !== 200) {
    Logger.log('Gemini APIエラー: ' + status + ' ' + response.getContentText());
    return null;
  }

  var json = JSON.parse(response.getContentText());
  var text = '';

  try {
    text = json.candidates[0].content.parts[0].text;
  } catch (e) {
    Logger.log('Geminiレスポンス解析エラー: ' + e);
    return null;
  }

  // JSONを抽出（コードブロックで囲まれている場合にも対応）
  text = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();

  try {
    var result = JSON.parse(text);
    return {
      makerName: result.makerName ? String(result.makerName).trim() : null,
      amount: parseInt(result.amount) || 0
    };
  } catch (e) {
    Logger.log('JSON解析エラー: ' + text);
    return null;
  }
}


// ===== スプレッドシート操作 =====

/**
 * スプレッドシートの該当セルに金額を記入
 * @param {Sheet} sheet - シート
 * @param {string} makerName - メーカー名
 * @param {number} amount - 金額
 * @param {string} yearMonth - 年月（例: "2026/02"）
 */
function updateSheet_(sheet, makerName, amount, yearMonth) {
  var row = findOrCreateRow_(sheet, makerName);
  var col = findOrCreateColumn_(sheet, yearMonth);

  // 既存の値がある場合は加算（同じメーカーから複数請求書がある場合）
  var existing = sheet.getRange(row, col).getValue();
  var newValue = (existing ? Number(existing) : 0) + amount;
  sheet.getRange(row, col).setValue(newValue);
}

/**
 * メーカー名の行を検索、なければ新規追加
 * @param {Sheet} sheet - シート
 * @param {string} makerName - メーカー名
 * @return {number} 行番号
 */
function findOrCreateRow_(sheet, makerName) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();

  // A列からメーカー名を検索
  for (var i = 1; i < values.length; i++) { // 0行目はヘッダー
    if (String(values[i][0]).trim() === makerName) return i + 1;
  }

  // なければ最終行に追加
  var newRow = lastRow + 1;
  sheet.getRange(newRow, 1).setValue(makerName);
  return newRow;
}

/**
 * 年月の列を検索、なければ新規追加
 * @param {Sheet} sheet - シート
 * @param {string} yearMonth - 年月（例: "2026/02"）
 * @return {number} 列番号
 */
function findOrCreateColumn_(sheet, yearMonth) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var values = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 1行目から年月を検索（日付オブジェクトに変換されている場合にも対応）
  for (var i = 1; i < values.length; i++) { // 0列目はメーカー名
    if (normalizeYearMonth_(values[i]) === yearMonth) return i + 1;
  }

  // なければ最終列に追加（プレーンテキスト形式で書き込み）
  var newCol = lastCol + 1;
  var cell = sheet.getRange(1, newCol);
  cell.setNumberFormat('@'); // プレーンテキスト形式を強制
  cell.setValue(yearMonth);

  // 列を時系列順にソート（B列以降）
  sortColumnsByDate_(sheet);

  // ソート後に再検索
  lastCol = sheet.getLastColumn();
  values = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var j = 1; j < values.length; j++) {
    if (normalizeYearMonth_(values[j]) === yearMonth) return j + 1;
  }
  return newCol;
}

/**
 * セルの値をyyyy/MM形式に正規化（日付オブジェクト対応）
 * @param {*} value - セルの値（日付オブジェクト or 文字列）
 * @return {string} yyyy/MM形式の文字列
 */
function normalizeYearMonth_(value) {
  if (!value) return '';
  // 日付オブジェクトの場合
  if (value instanceof Date) {
    var y = value.getFullYear();
    var m = ('0' + (value.getMonth() + 1)).slice(-2);
    return y + '/' + m;
  }
  // 文字列の場合はそのまま返す
  return String(value).trim();
}


/**
 * ヘッダー行の年月列を時系列順にソート
 */
function sortColumnsByDate_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol <= 2) return; // B列以降が2つ未満ならソート不要

  var lastRow = Math.max(sheet.getLastRow(), 1);
  var headerRange = sheet.getRange(1, 2, 1, lastCol - 1);
  var headers = headerRange.getValues()[0];

  // ヘッダーとインデックスのペアを作成
  var pairs = [];
  for (var i = 0; i < headers.length; i++) {
    pairs.push({ header: headers[i], index: i });
  }

  // 年月でソート
  pairs.sort(function(a, b) {
    return String(a.header).localeCompare(String(b.header));
  });

  // ソートが必要か確認
  var needsSort = false;
  for (var k = 0; k < pairs.length; k++) {
    if (pairs[k].index !== k) { needsSort = true; break; }
  }
  if (!needsSort) return;

  // データ全体を取得してソート順で並べ替え
  var allData = sheet.getRange(1, 2, lastRow, lastCol - 1).getValues();
  var sorted = [];
  for (var r = 0; r < allData.length; r++) {
    var newRow = [];
    for (var c = 0; c < pairs.length; c++) {
      newRow.push(allData[r][pairs[c].index]);
    }
    sorted.push(newRow);
  }
  sheet.getRange(1, 2, lastRow, lastCol - 1).setValues(sorted);
}


// ===== 初期セットアップ =====

/**
 * スプレッドシートを自動作成（まだない場合）
 */
function ensureSpreadsheet_() {
  if (SPREADSHEET_ID) {
    try {
      SpreadsheetApp.openById(SPREADSHEET_ID);
      return;
    } catch (e) {
      Logger.log('⚠️ 指定のスプレッドシートが見つかりません。新規作成します。');
    }
  }

  var ss = SpreadsheetApp.create('📋 請求書管理');
  SPREADSHEET_ID = ss.getId();

  // プロパティに保存（次回以降に使用）
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', SPREADSHEET_ID);

  var sheet = ss.getActiveSheet();
  sheet.setName(SHEET_NAME);

  // ヘッダーのA1セルにラベル
  sheet.getRange(1, 1).setValue('メーカー名');
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.setColumnWidth(1, 200);

  // ヘッダー行のスタイル
  sheet.getRange(1, 1, 1, 1).setBackground('#1a237e').setFontColor('#ffffff');

  Logger.log('📋 スプレッドシート作成: ' + ss.getUrl());
  Logger.log('💡 設定.gs の SPREADSHEET_ID に以下を設定してください: ' + SPREADSHEET_ID);
}

/**
 * Driveフォルダを自動作成（まだない場合）
 */
function ensureDriveFolder_() {
  if (DRIVE_FOLDER_ID) {
    try {
      DriveApp.getFolderById(DRIVE_FOLDER_ID);
      return;
    } catch (e) {
      Logger.log('⚠️ 指定のフォルダが見つかりません。新規作成します。');
    }
  }

  var folder = DriveApp.createFolder('📁 ' + FOLDER_NAME);
  DRIVE_FOLDER_ID = folder.getId();

  // プロパティに保存
  PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', DRIVE_FOLDER_ID);

  Logger.log('📁 フォルダ作成: ' + folder.getUrl());
  Logger.log('💡 設定.gs の DRIVE_FOLDER_ID に以下を設定してください: ' + DRIVE_FOLDER_ID);
}

/**
 * スクリプトプロパティから保存済みIDを復元
 * （グローバル変数が空の場合、プロパティから自動復元）
 */
function restoreIds_() {
  var props = PropertiesService.getScriptProperties();
  if (!SPREADSHEET_ID) {
    SPREADSHEET_ID = props.getProperty('SPREADSHEET_ID') || '';
  }
  if (!DRIVE_FOLDER_ID) {
    DRIVE_FOLDER_ID = props.getProperty('DRIVE_FOLDER_ID') || '';
  }
}

// 起動時に自動復元
restoreIds_();


// ===== ログ管理 =====

/**
 * ログ用シートを取得 or 作成
 */
function getOrCreateLogSheet_(ss) {
  var logSheet = ss.getSheetByName('処理ログ');
  if (!logSheet) {
    logSheet = ss.insertSheet('処理ログ');
    logSheet.getRange(1, 1, 1, 6).setValues([['処理日時', 'メール件名', 'メーカー名', '金額', '年月', 'ステータス']]);
    logSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e3f2fd');
    logSheet.setFrozenRows(1);
  }
  return logSheet;
}

/**
 * 処理結果をログシートに記録
 */
function logResult_(logSheet, date, subject, makerName, amount, yearMonth, status) {
  logSheet.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    subject,
    makerName,
    amount,
    yearMonth,
    status
  ]);
}


// ===== トリガー管理 =====

/**
 * 月1回の自動実行トリガーを設定（毎月1日の朝9時）
 */
function setupTrigger() {
  // 既存トリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processInvoices') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 新規トリガー: 毎月1日 9:00〜10:00
  ScriptApp.newTrigger('processInvoices')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  Logger.log('⏰ 自動実行トリガー設定完了: 毎月1日 9:00〜10:00');
}

/**
 * トリガーを解除
 */
function removeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processInvoices') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log('🗑 ' + removed + '件のトリガーを解除しました');
}


// ===== ユーティリティ =====

/**
 * テスト用: 直近のメール1件だけ処理（テスト実行に使用）
 */
function testOneEmail() {
  if (!GEMINI_API_KEY || GEMINI_API_KEY === 'ここにAPIキーを貼り付け') {
    throw new Error('❗ 設定.gs の GEMINI_API_KEY を設定してください');
  }

  ensureSpreadsheet_();
  ensureDriveFolder_();

  var threads = GmailApp.search(GMAIL_QUERY + ' after:' + START_DATE.replace(/\//g, '/'), 0, 1);
  if (threads.length === 0) {
    Logger.log('PDF添付メールが見つかりません');
    return;
  }

  var message = threads[0].getMessages()[0];
  var pdfs = extractPdfsFromMessage_(message);

  if (pdfs.length === 0) {
    Logger.log('PDFが見つかりません: ' + message.getSubject());
    return;
  }

  Logger.log('📧 テスト対象: ' + message.getSubject());
  Logger.log('📎 PDF: ' + pdfs[0].getName());

  var result = analyzeWithGemini_(pdfs[0]);
  Logger.log('🤖 Gemini結果: ' + JSON.stringify(result));

  if (result && result.makerName) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var yearMonth = Utilities.formatDate(message.getDate(), 'Asia/Tokyo', 'yyyy/MM');
    updateSheet_(sheet, result.makerName, result.amount, yearMonth);
    Logger.log('✅ シート記入完了: ' + result.makerName + ' ¥' + result.amount);
  }
}

/**
 * 処理済みラベルをリセット（全てのメールを再処理したい場合）
 */
function resetProcessedLabel() {
  var label = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!label) {
    Logger.log('ラベルが見つかりません');
    return;
  }

  var threads = label.getThreads();
  for (var i = 0; i < threads.length; i++) {
    threads[i].removeLabel(label);
  }
  Logger.log('🔄 ' + threads.length + '件のスレッドからラベルを解除しました');
}
