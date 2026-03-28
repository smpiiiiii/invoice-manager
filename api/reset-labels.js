/**
 * 仕分け済ラベルリセット + 両方のスプレッドシートクリア
 * POST /api/reset-labels
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  const labelName = '仕分け済';

  try {
    // ラベルを検索
    const labels = await googleApi(token, 'https://gmail.googleapis.com/gmail/v1/users/me/labels');
    const label = (labels.labels || []).find(l => l.name === labelName);
    if (!label) {
      return res.json({ success: true, message: `「${labelName}」ラベルが見つかりません`, removed: 0 });
    }

    // そのラベルがついたメールを検索
    const gmailRes = await googleApi(token,
      `https://gmail.googleapis.com/gmail/v1/users/me/messages?q=${encodeURIComponent(`label:${labelName}`)}&maxResults=50`
    );
    const messageIds = (gmailRes.messages || []).map(m => m.id);

    // 各メールからラベルを外す
    let removed = 0;
    for (const msgId of messageIds) {
      await googleApi(token,
        `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/modify`,
        { method: 'POST', body: JSON.stringify({ removeLabelIds: [label.id] }) }
      );
      removed++;
    }

    // 両方のスプレッドシートをクリア
    let sheetsCleared = 0;
    for (const modeLabel of ['請求書', '領収書']) {
      const sheetName = `📋 ${modeLabel}管理`;
      const sheetSearch = await googleApi(token,
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${sheetName}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`)}&fields=files(id)`
      );
      if (sheetSearch.files && sheetSearch.files.length > 0) {
        const sheetId = sheetSearch.files[0].id;
        // 1. 全データをクリア
        await googleApi(token,
          `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A:ZZ:clear`,
          { method: 'POST', body: '{}' }
        );
        // 2. ヘッダー行だけ書き戻し
        await googleApi(token,
          `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1?valueInputOption=RAW`,
          { method: 'PUT', body: JSON.stringify({ values: [['メーカー名']] }) }
        );
        sheetsCleared++;
      }
    }

    res.json({ success: true, message: `${removed}件のラベル解除 + ${sheetsCleared}シートクリア`, removed });
  } catch (e) {
    console.error('ラベルリセットエラー:', e);
    res.status(500).json({ error: e.message });
  }
};
