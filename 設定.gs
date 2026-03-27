// ============================================================
// 請求書自動管理システム — 設定ファイル
// ============================================================

/**
 * 設定値をまとめて返す
 * ※ 初回セットアップ時に setupSystem() を実行すると自動設定される
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    // Gemini API Key（Google AI Studioで取得）
    GEMINI_API_KEY: props.getProperty('GEMINI_API_KEY') || '',

    // Google Drive 保存先フォルダID
    DRIVE_FOLDER_ID: props.getProperty('DRIVE_FOLDER_ID') || '',

    // スプレッドシートID（このスプレッドシート自体）
    SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),

    // Gmail検索クエリ（請求書メールを特定する条件）
    SEARCH_QUERY: 'has:attachment filename:pdf label:請求書 -label:処理済み',

    // 処理済みラベル名
    PROCESSED_LABEL: '処理済み',

    // シート名
    SHEET_NAME: '請求書一覧',
  };
}

/**
 * 初回セットアップ — メニューから実行
 * Driveフォルダ作成、Gmailラベル作成、シート初期化を行う
 */
function setupSystem() {
  const ui = SpreadsheetApp.getUi();

  // 1. Gemini API Key を入力
  const apiKeyResult = ui.prompt(
    '🔑 Gemini API Key',
    'Google AI Studio (https://aistudio.google.com/apikey) で取得したAPIキーを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  if (apiKeyResult.getSelectedButton() !== ui.Button.OK || !apiKeyResult.getResponseText().trim()) {
    ui.alert('セットアップを中断しました');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKeyResult.getResponseText().trim());

  // 2. Google Driveにフォルダ作成
  let folder;
  const folders = DriveApp.getFoldersByName('請求書');
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder('請求書');
  }
  PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', folder.getId());

  // 3. Gmailラベル作成
  let label = GmailApp.getUserLabelByName('請求書');
  if (!label) label = GmailApp.createLabel('請求書');
  let processedLabel = GmailApp.getUserLabelByName('処理済み');
  if (!processedLabel) GmailApp.createLabel('処理済み');

  // 4. シート初期化
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('請求書一覧');
  if (!sheet) {
    sheet = ss.insertSheet('請求書一覧');
    sheet.getRange('A1').setValue('メーカー名');
    sheet.getRange('A1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
    sheet.setColumnWidth(1, 200);
  }

  ui.alert('✅ セットアップ完了！\n\n' +
    '📁 Driveフォルダ: 「請求書」\n' +
    '🏷️ Gmailラベル: 「請求書」「処理済み」\n' +
    '📊 シート: 「請求書一覧」\n\n' +
    '使い方:\n' +
    '1. Gmailで請求書メールに「請求書」ラベルを付ける\n' +
    '2. メニュー「📄 請求書管理」→「▶ 請求書を処理」を実行');
}
