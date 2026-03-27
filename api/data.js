/**
 * データ取得API — ダッシュボード用
 * GET /api/data
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  try {
    // スプレッドシートを検索
    const searchRes = await googleApi(token,
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent("name='📋 請求書管理' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false")}&fields=files(id)`
    );
    if (!searchRes.files || searchRes.files.length === 0) {
      return res.json({ totalProcessed: 0, totalAmount: 0, totalMakers: 0, latestMonth: '-', monthly: [], topMakers: [], tableData: [], user: { email: session.email, name: session.name, picture: session.picture } });
    }

    const sheetId = searchRes.files[0].id;

    // データ取得
    const dataRes = await googleApi(token,
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A:ZZ`
    );
    const values = dataRes.values || [];
    if (values.length < 2) {
      return res.json({ totalProcessed: 0, totalAmount: 0, totalMakers: 0, latestMonth: '-', monthly: [], topMakers: [], tableData: values, user: { email: session.email, name: session.name, picture: session.picture } });
    }

    // 統計計算
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

    // 月別合計
    const monthly = months.map(m => {
      let total = 0;
      for (let r = 1; r < values.length; r++) {
        total += parseInt((values[r] || [])[m.col]) || 0;
      }
      return { month: m.month, amount: total };
    });

    res.json({
      totalProcessed,
      totalAmount,
      totalMakers: makers.length,
      latestMonth: months.length > 0 ? months[months.length - 1].month : '-',
      monthly,
      topMakers: makers.slice(0, 5),
      tableData: values,
      sheetUrl: `https://docs.google.com/spreadsheets/d/${sheetId}`,
      user: { email: session.email, name: session.name, picture: session.picture }
    });
  } catch (e) {
    console.error('データ取得エラー:', e);
    res.status(500).json({ error: e.message });
  }
};
