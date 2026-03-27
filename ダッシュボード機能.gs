/**
 * 📋 請求書管理ダッシュボード — Web App用関数
 *
 * このファイルをGASプロジェクトに追加してください。
 * デプロイ: 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」→ URL取得
 */

// ===== Web App エントリポイント =====

/**
 * Web Appのメインページを表示
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('ダッシュボード')
    .setTitle('📋 請求書管理ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ===== ダッシュボード用データ取得 =====

/**
 * ダッシュボードに表示するデータを取得
 * @return {Object} 統計情報、月別データ、テーブルデータ
 */
function getDashboardData() {
  restoreIds_();

  if (!SPREADSHEET_ID) return null;

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return null;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return { totalProcessed: 0, totalAmount: 0, totalMakers: 0, latestMonth: '-', monthly: [], topMakers: [], tableData: [] };

    // テーブルデータ全体を取得
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // ヘッダー（月のリスト）
    var months = [];
    for (var c = 1; c < allData[0].length; c++) {
      var mv = normalizeYearMonth_(allData[0][c]);
      if (mv) months.push({ col: c, month: mv });
    }

    // メーカーごとの合計を計算
    var makers = [];
    var totalAmount = 0;
    var totalProcessed = 0;

    for (var r = 1; r < allData.length; r++) {
      var name = String(allData[r][0]).trim();
      if (!name) continue;

      var makerTotal = 0;
      for (var mc = 1; mc < allData[r].length; mc++) {
        var v = Number(allData[r][mc]) || 0;
        if (v > 0) {
          makerTotal += v;
          totalProcessed++;
        }
      }
      totalAmount += makerTotal;
      makers.push({ name: name, amount: makerTotal });
    }

    // トップ5メーカー
    makers.sort(function(a, b) { return b.amount - a.amount; });
    var topMakers = makers.slice(0, 5);

    // 月別合計
    var monthly = [];
    for (var mi = 0; mi < months.length; mi++) {
      var monthTotal = 0;
      for (var ri = 1; ri < allData.length; ri++) {
        monthTotal += Number(allData[ri][months[mi].col]) || 0;
      }
      monthly.push({ month: months[mi].month, amount: monthTotal });
    }

    // テーブルデータ（表示用に整形）
    var tableData = [];
    var headerRow = ['メーカー名'];
    for (var hi = 0; hi < months.length; hi++) headerRow.push(months[hi].month);
    tableData.push(headerRow);

    for (var tr = 1; tr < allData.length; tr++) {
      var rowName = String(allData[tr][0]).trim();
      if (!rowName) continue;
      var row = [rowName];
      for (var tc = 0; tc < months.length; tc++) {
        row.push(Number(allData[tr][months[tc].col]) || 0);
      }
      tableData.push(row);
    }

    return {
      totalProcessed: totalProcessed,
      totalAmount: totalAmount,
      totalMakers: makers.length,
      latestMonth: months.length > 0 ? months[months.length - 1].month : '-',
      monthly: monthly,
      topMakers: topMakers,
      tableData: tableData
    };
  } catch (e) {
    Logger.log('ダッシュボード データ取得エラー: ' + e);
    return null;
  }
}


/**
 * 最新のログを取得
 * @return {Array} ログの配列
 */
function getRecentLogs() {
  restoreIds_();
  if (!SPREADSHEET_ID) return [];

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName('処理ログ');
    if (!logSheet) return [];

    var lastRow = logSheet.getLastRow();
    if (lastRow < 2) return [];

    // 最新20件を取得
    var startRow = Math.max(2, lastRow - 19);
    var data = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 6).getValues();

    var logs = [];
    for (var i = data.length - 1; i >= 0; i--) {
      logs.push({
        date: String(data[i][0]),
        subject: String(data[i][1]),
        maker: String(data[i][2]),
        amount: Number(data[i][3]) || 0,
        month: String(data[i][4]),
        status: String(data[i][5])
      });
    }
    return logs;
  } catch (e) {
    return [];
  }
}


// ===== ダッシュボードからのアクション =====

/**
 * ダッシュボードから処理実行
 */
function runProcessFromDashboard() {
  processInvoices();
  return '処理が完了しました';
}

/**
 * ダッシュボードからテスト実行
 */
function runTestFromDashboard() {
  testOneEmail();
  return 'テスト完了';
}

/**
 * スプレッドシートのURLを取得
 */
function getSheetUrl() {
  restoreIds_();
  if (!SPREADSHEET_ID) return null;
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID).getUrl();
  } catch (e) { return null; }
}

/**
 * DriveフォルダのURLを取得
 */
function getDriveUrl() {
  restoreIds_();
  if (!DRIVE_FOLDER_ID) return null;
  try {
    return DriveApp.getFolderById(DRIVE_FOLDER_ID).getUrl();
  } catch (e) { return null; }
}
