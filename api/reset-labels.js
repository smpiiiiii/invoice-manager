/**
 * フルリセットAPI
 * POST /api/reset-labels
 * 1. Gmailの「仕分け済」ラベルを全メールから削除 + ラベル自体も削除
 * 2. スプレッドシートの処理ログ・明細シートをクリア
 * 3. Driveの請求書・領収書フォルダ内のファイルを全削除
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  const results = { labels: '', sheet: '', drive: '' };

  try {
    // === 1. Gmailラベルリセット ===
    const labelsRes = await googleApi(token,
      'https://gmail.googleapis.com/gmail/v1/users/me/labels'
    );
    const labels = labelsRes.labels || [];
    const targetLabel = labels.find(l => l.name === '仕分け済');

    if (targetLabel) {
      const labelId = targetLabel.id;
      // ラベル付きメールを全取得
      let allMessageIds = [];
      let pageToken = '';
      do {
        const url = `https://gmail.googleapis.com/gmail/v1/users/me/messages?labelIds=${labelId}&maxResults=100${pageToken ? '&pageToken=' + pageToken : ''}`;
        const msgRes = await googleApi(token, url);
        allMessageIds = allMessageIds.concat((msgRes.messages || []).map(m => m.id));
        pageToken = msgRes.nextPageToken || '';
      } while (pageToken);

      // バッチでラベル削除
      if (allMessageIds.length > 0) {
        for (let i = 0; i < allMessageIds.length; i += 1000) {
          const batch = allMessageIds.slice(i, i + 1000);
          await googleApi(token,
            'https://gmail.googleapis.com/gmail/v1/users/me/messages/batchModify',
            { method: 'POST', body: JSON.stringify({ ids: batch, removeLabelIds: [labelId] }) }
          );
        }
      }
      // ラベル自体を削除
      await googleApi(token,
        `https://gmail.googleapis.com/gmail/v1/users/me/labels/${labelId}`,
        { method: 'DELETE' }
      );
      results.labels = `${allMessageIds.length}件のラベル削除`;
    } else {
      results.labels = 'ラベルなし';
    }

    // === 2. スプレッドシートのデータクリア ===
    // 「📋 請求書管理」スプレッドシートを検索
    const sheetSearch = await googleApi(token,
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent("name='📋 請求書管理' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false")}&fields=files(id)`
    );
    if (sheetSearch.files && sheetSearch.files.length > 0) {
      const sheetId = sheetSearch.files[0].id;
      // スプレッドシートのシート情報を取得
      const sheetInfo = await googleApi(token,
        `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties`
      );
      const sheets = sheetInfo.sheets || [];
      for (const s of sheets) {
        const sheetTitle = s.properties.title;
        // ヘッダー行だけ残してデータをクリア
        try {
          await googleApi(token,
            `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(sheetTitle)}!A2:Z?valueInputOption=USER_ENTERED`,
            { method: 'PUT', body: JSON.stringify({ values: [] }) }
          );
        } catch (e) {
          // シートが空の場合はエラーを無視
          console.log(`シート「${sheetTitle}」クリアスキップ:`, e.message);
        }
      }
      // batchClearで全シートのデータをクリア（ヘッダー以外）
      const ranges = sheets.map(s => `${s.properties.title}!A2:Z10000`);
      if (ranges.length > 0) {
        await googleApi(token,
          `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchClear`,
          { method: 'POST', body: JSON.stringify({ ranges }) }
        );
      }
      results.sheet = `${sheets.length}シートのデータクリア`;
    } else {
      results.sheet = 'スプレッドシートなし';
    }

    // === 3. Driveフォルダの中身を削除 ===
    const folderName = '📂 請求書・領収書管理';
    const folderSearch = await googleApi(token,
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id)`
    );
    if (folderSearch.files && folderSearch.files.length > 0) {
      const folderId = folderSearch.files[0].id;
      // フォルダ内の全ファイル・サブフォルダを再帰的に取得して削除
      let deletedCount = 0;
      async function deleteAllInFolder(parentId) {
        let pt = '';
        do {
          const filesRes = await googleApi(token,
            `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`'${parentId}' in parents and trashed=false`)}&fields=files(id,mimeType,name)&pageSize=100${pt ? '&pageToken=' + pt : ''}`
          );
          for (const file of (filesRes.files || [])) {
            if (file.mimeType === 'application/vnd.google-apps.folder') {
              // サブフォルダ内も先に削除
              await deleteAllInFolder(file.id);
            }
            await googleApi(token,
              `https://www.googleapis.com/drive/v3/files/${file.id}`,
              { method: 'DELETE' }
            );
            deletedCount++;
          }
          pt = filesRes.nextPageToken || '';
        } while (pt);
      }
      await deleteAllInFolder(folderId);
      results.drive = `${deletedCount}件のファイル/フォルダ削除`;
    } else {
      results.drive = 'Driveフォルダなし';
    }

    res.json({
      success: true,
      message: `リセット完了: Gmail ${results.labels} / スプシ ${results.sheet} / Drive ${results.drive}`
    });
  } catch (e) {
    console.error('リセットエラー:', e);
    res.status(500).json({ error: e.message });
  }
};
