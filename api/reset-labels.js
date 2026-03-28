/**
 * 仕分け済ラベルリセット + 両スプレッドシートクリア + Driveフォルダ削除
 * POST /api/reset-labels
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  // 新旧すべてのラベル名
  const labelNames = ['仕分け済', '請求書処理済', '領収書処理済'];

  try {
    // 全ラベルを検索
    const labels = await googleApi(token, 'https://gmail.googleapis.com/gmail/v1/users/me/labels');
    let totalRemoved = 0;

    for (const labelName of labelNames) {
      const label = (labels.labels || []).find(l => l.name === labelName);
      if (!label) continue;

      // そのラベルがついたメールを検索して外す
      const gmailRes = await googleApi(token,
        `https://gmail.googleapis.com/gmail/v1/users/me/messages?q=${encodeURIComponent(`label:${labelName}`)}&maxResults=100`
      );
      const messageIds = (gmailRes.messages || []).map(m => m.id);

      for (const msgId of messageIds) {
        await googleApi(token,
          `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/modify`,
          { method: 'POST', body: JSON.stringify({ removeLabelIds: [label.id] }) }
        );
        totalRemoved++;
      }
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
        await googleApi(token,
          `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A:ZZ:clear`,
          { method: 'POST', body: '{}' }
        );
        await googleApi(token,
          `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1?valueInputOption=RAW`,
          { method: 'PUT', body: JSON.stringify({ values: [['メーカー名']] }) }
        );
        sheetsCleared++;
      }
    }

    // Driveフォルダを削除（📂 請求書・領収書管理 + 旧フォルダ）
    let foldersDeleted = 0;
    const folderNames = ['📂 請求書・領収書管理', '📁 請求書', '📁 領収書'];
    for (const folderName of folderNames) {
      const folderSearch = await googleApi(token,
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id)`
      );
      if (folderSearch.files && folderSearch.files.length > 0) {
        for (const f of folderSearch.files) {
          await googleApi(token,
            `https://www.googleapis.com/drive/v3/files/${f.id}`,
            { method: 'DELETE' }
          );
          foldersDeleted++;
        }
      }
    }

    res.json({
      success: true,
      message: `${totalRemoved}件ラベル解除 + ${sheetsCleared}シートクリア + ${foldersDeleted}フォルダ削除`,
      removed: totalRemoved
    });
  } catch (e) {
    console.error('リセットエラー:', e);
    res.status(500).json({ error: e.message });
  }
};
