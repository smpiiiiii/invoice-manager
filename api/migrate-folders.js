/**
 * Driveフォルダ移行API
 * POST /api/migrate-folders
 * 旧形式フォルダ（例: "202603 請求書"）から新形式（"2026年 請求書/3月/"）に移行
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  try {
    // 親フォルダを検索
    const parentName = '📂 請求書・領収書管理';
    const parentRes = await googleApi(token,
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${parentName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id)`
    );

    if (!parentRes.files || parentRes.files.length === 0) {
      return res.json({ success: true, message: '親フォルダが見つかりません。処理不要です。', moved: 0 });
    }

    const parentId = parentRes.files[0].id;

    // 親フォルダ内の全サブフォルダを取得
    const foldersRes = await googleApi(token,
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`'${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id,name)&pageSize=100`
    );

    const folders = foldersRes.files || [];
    let movedCount = 0;
    let foldersMigrated = [];

    // 旧形式フォルダを検出（"YYYYMM 請求書" or "YYYYMM 領収書"）
    const oldPattern = /^(\d{4})(\d{2})\s+(請求書|領収書)$/;

    for (const folder of folders) {
      const match = folder.name.match(oldPattern);
      if (!match) continue;

      const year = parseInt(match[1]);
      const month = parseInt(match[2]);
      const docType = match[3];

      // 新形式の年フォルダを取得or作成
      const yearFolderName = `${year}年 ${docType}`;
      let yearFolderId;
      const yearSearch = await googleApi(token,
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${yearFolderName}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id)`
      );
      if (yearSearch.files && yearSearch.files.length > 0) {
        yearFolderId = yearSearch.files[0].id;
      } else {
        const created = await googleApi(token, 'https://www.googleapis.com/drive/v3/files', {
          method: 'POST', body: JSON.stringify({ name: yearFolderName, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] })
        });
        yearFolderId = created.id;
      }

      // 月フォルダを取得or作成
      const monthFolderName = `${month}月`;
      let monthFolderId;
      const monthSearch = await googleApi(token,
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${monthFolderName}' and '${yearFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id)`
      );
      if (monthSearch.files && monthSearch.files.length > 0) {
        monthFolderId = monthSearch.files[0].id;
      } else {
        const created = await googleApi(token, 'https://www.googleapis.com/drive/v3/files', {
          method: 'POST', body: JSON.stringify({ name: monthFolderName, mimeType: 'application/vnd.google-apps.folder', parents: [yearFolderId] })
        });
        monthFolderId = created.id;
      }

      // 旧フォルダ内のファイルを新フォルダに移動
      const filesRes = await googleApi(token,
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`'${folder.id}' in parents and trashed=false`)}&fields=files(id,name)&pageSize=100`
      );

      const files = filesRes.files || [];
      for (const file of files) {
        // ファイルの親フォルダを変更（旧→新）
        await googleApi(token,
          `https://www.googleapis.com/drive/v3/files/${file.id}?addParents=${monthFolderId}&removeParents=${folder.id}`,
          { method: 'PATCH', body: JSON.stringify({}) }
        );
        movedCount++;
      }

      // 旧フォルダをゴミ箱に移動（空になったので）
      if (files.length > 0 || true) {
        await googleApi(token,
          `https://www.googleapis.com/drive/v3/files/${folder.id}`,
          { method: 'PATCH', body: JSON.stringify({ trashed: true }) }
        );
      }

      foldersMigrated.push({
        old: folder.name,
        new: `${yearFolderName}/${monthFolderName}`,
        files: files.length
      });
    }

    res.json({
      success: true,
      message: `${movedCount}件のファイルを移行しました（${foldersMigrated.length}フォルダ）`,
      moved: movedCount,
      details: foldersMigrated
    });

  } catch (e) {
    console.error('フォルダ移行エラー:', e);
    res.status(500).json({ error: e.message });
  }
};
