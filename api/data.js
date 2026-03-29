/**
 * データ取得API — ダッシュボード用
 * GET /api/data
 * 請求書・領収書の両方のスプレッドシートからデータを取得
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  try {
    // 両方のスプレッドシートを検索
    const [invoiceSearch, receiptSearch] = await Promise.all([
      googleApi(token,
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent("name='📋 請求書管理' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false")}&fields=files(id)`
      ),
      googleApi(token,
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent("name='📋 領収書管理' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false")}&fields=files(id)`
      )
    ]);

    const invoiceSheetId = (invoiceSearch.files && invoiceSearch.files.length > 0) ? invoiceSearch.files[0].id : null;
    const receiptSheetId = (receiptSearch.files && receiptSearch.files.length > 0) ? receiptSearch.files[0].id : null;

    // データ取得（メインシート + 明細シート）
    const [invoiceData, receiptData, invoiceDetails, receiptDetails] = await Promise.all([
      invoiceSheetId ? googleApi(token, `https://sheets.googleapis.com/v4/spreadsheets/${invoiceSheetId}/values/A:ZZ`) : { values: [] },
      receiptSheetId ? googleApi(token, `https://sheets.googleapis.com/v4/spreadsheets/${receiptSheetId}/values/A:ZZ`) : { values: [] },
      invoiceSheetId ? googleApi(token, `https://sheets.googleapis.com/v4/spreadsheets/${invoiceSheetId}/values/${encodeURIComponent('明細')}!A:F`).catch(() => ({ values: [] })) : { values: [] },
      receiptSheetId ? googleApi(token, `https://sheets.googleapis.com/v4/spreadsheets/${receiptSheetId}/values/${encodeURIComponent('明細')}!A:F`).catch(() => ({ values: [] })) : { values: [] }
    ]);

    // 統計計算関数
    function calcStats(values) {
      if (!values || values.length < 2) return { totalProcessed: 0, totalAmount: 0, totalMakers: 0, latestMonth: '-', monthly: [], topMakers: [], tableData: values || [] };

      const months = [];
      for (let c = 1; c < values[0].length; c++) {
        if (values[0][c]) months.push({ col: c, month: String(values[0][c]) });
      }

      let totalAmount = 0, totalProcessed = 0;
      const makers = [];

      for (let r = 1; r < values.length; r++) {
        const name = (values[r][0] || '').trim();
        if (!name) continue;
        let makerTotal = 0;
        for (let c = 1; c < (values[r] || []).length; c++) {
          const v = parseInt(values[r][c]) || 0;
          if (v > 0) { makerTotal += v; totalProcessed++; }
        }
        totalAmount += makerTotal;
        makers.push({ name, amount: makerTotal });
      }

      makers.sort((a, b) => b.amount - a.amount);

      const monthly = months.map(m => {
        let total = 0;
        for (let r = 1; r < values.length; r++) {
          total += parseInt((values[r] || [])[m.col]) || 0;
        }
        return { month: m.month, amount: total };
      });

      return {
        totalProcessed,
        totalAmount,
        totalMakers: makers.length,
        latestMonth: months.length > 0 ? months[months.length - 1].month : '-',
        monthly,
        topMakers: makers.slice(0, 5),
        tableData: values
      };
    }

    const invoice = calcStats(invoiceData.values || []);
    const receipt = calcStats(receiptData.values || []);

    // Driveフォルダのリンクを取得
    const driveFolderSearch = await googleApi(token,
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent("name='📂 請求書・領収書管理' and mimeType='application/vnd.google-apps.folder' and trashed=false")}&fields=files(id)`
    );
    const driveFolderId = (driveFolderSearch.files && driveFolderSearch.files.length > 0) ? driveFolderSearch.files[0].id : null;

    res.json({
      invoice: { ...invoice, sheetUrl: invoiceSheetId ? `https://docs.google.com/spreadsheets/d/${invoiceSheetId}` : '', details: (invoiceDetails.values || []) },
      receipt: { ...receipt, sheetUrl: receiptSheetId ? `https://docs.google.com/spreadsheets/d/${receiptSheetId}` : '', details: (receiptDetails.values || []) },
      driveUrl: driveFolderId ? `https://drive.google.com/drive/folders/${driveFolderId}` : '',
      user: { email: session.email, name: session.name, picture: session.picture }
    });
  } catch (e) {
    console.error('データ取得エラー:', e);
    res.status(500).json({ error: e.message });
  }
};
