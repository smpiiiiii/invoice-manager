/**
 * メーカー名マージAPI
 * POST /api/merge-makers
 * 複数のメーカー名を1つに統合する
 * body: { mode: 'invoice'|'receipt', targetName: '統合先名', sourceNames: ['統合元1', '統合元2'] }
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  const { mode, targetName, sourceNames } = req.body;
  if (!mode || !targetName || !sourceNames || !sourceNames.length) {
    return res.status(400).json({ error: 'パラメータ不足: mode, targetName, sourceNames が必要です' });
  }

  try {
    // スプレッドシートを特定
    const sheetLabel = mode === 'receipt' ? '領収書' : '請求書';
    const sheetName = `📋 ${sheetLabel}管理`;
    const searchRes = await googleApi(token,
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${sheetName}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`)}&fields=files(id)`
    );

    if (!searchRes.files || searchRes.files.length === 0) {
      return res.status(404).json({ error: `${sheetLabel}のスプレッドシートが見つかりません` });
    }

    const sheetId = searchRes.files[0].id;

    // 現在のデータを取得
    const dataRes = await googleApi(token,
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A:ZZ`
    );
    const values = dataRes.values || [['メーカー名']];

    // 統合元の行インデックスを特定
    const sourceIndices = [];
    let targetIdx = -1;

    for (let r = 1; r < values.length; r++) {
      const name = (values[r] && values[r][0] || '').trim();
      if (!name) continue;
      if (name === targetName) {
        targetIdx = r;
      }
      if (sourceNames.includes(name) && name !== targetName) {
        sourceIndices.push(r);
      }
    }

    // 統合先がなければ最初のソース行を統合先にリネーム
    if (targetIdx === -1 && sourceIndices.length > 0) {
      targetIdx = sourceIndices.shift();
      values[targetIdx][0] = targetName;
    }

    if (targetIdx === -1) {
      return res.status(400).json({ error: '統合対象が見つかりません' });
    }

    // 統合元の金額を統合先にマージ
    let mergedCount = 0;
    for (const srcIdx of sourceIndices) {
      const srcRow = values[srcIdx];
      for (let c = 1; c < srcRow.length; c++) {
        const srcVal = parseInt(srcRow[c]) || 0;
        if (srcVal > 0) {
          // 統合先の行を必要な長さに拡張
          while (values[targetIdx].length <= c) values[targetIdx].push('');
          const existingVal = parseInt(values[targetIdx][c]) || 0;
          values[targetIdx][c] = existingVal + srcVal;
          mergedCount++;
        }
      }
      // 統合元の行をクリア（名前も含め）
      values[srcIdx] = [''];
    }

    // 空行を除去して書き戻し
    const cleanedValues = [values[0]]; // ヘッダー保持
    for (let r = 1; r < values.length; r++) {
      const name = (values[r] && values[r][0] || '').trim();
      if (name) cleanedValues.push(values[r]);
    }

    // シートをクリアしてから書き戻し
    await googleApi(token,
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A:ZZ:clear`,
      { method: 'POST', body: '{}' }
    );
    await googleApi(token,
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1?valueInputOption=USER_ENTERED`,
      { method: 'PUT', body: JSON.stringify({ values: cleanedValues }) }
    );

    res.json({
      success: true,
      message: `${sourceNames.length}件 → 「${targetName}」に統合しました（${mergedCount}セル統合）`,
      targetName,
      merged: sourceNames.length,
      cells: mergedCount
    });
  } catch (e) {
    console.error('マージエラー:', e);
    res.status(500).json({ error: e.message });
  }
};
