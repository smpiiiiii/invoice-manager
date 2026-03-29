/**
 * 空フォルダ削除API
 * POST /api/cleanup-folders
 * Drive内の「📂 請求書・領収書管理」配下の空フォルダを検出・削除
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  try {
    // 親フォルダを検索
    const parentSearch = await googleApi(token,
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent("name='📂 請求書・領収書管理' and mimeType='application/vnd.google-apps.folder' and trashed=false")}&fields=files(id,name)`
    );

    if (!parentSearch.files || parentSearch.files.length === 0) {
      return res.json({ success: true, message: '管理フォルダが見つかりません', deleted: 0 });
    }

    const parentId = parentSearch.files[0].id;
    const deleted = [];
    const kept = [];

    // 親フォルダ直下のサブフォルダを全取得
    const subFolders = await googleApi(token,
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`'${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id,name)&pageSize=100`
    );

    for (const folder of (subFolders.files || [])) {
      // そのフォルダの中身を確認
      const contents = await googleApi(token,
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`'${folder.id}' in parents and trashed=false`)}&fields=files(id,name,mimeType)&pageSize=100`
      );

      const files = (contents.files || []);
      
      if (files.length === 0) {
        // 空フォルダ → 削除
        await googleApi(token,
          `https://www.googleapis.com/drive/v3/files/${folder.id}`,
          { method: 'DELETE' }
        );
        deleted.push(folder.name);
      } else {
        // 中にサブフォルダだけの場合、その中も確認
        const hasRealFiles = files.some(f => f.mimeType !== 'application/vnd.google-apps.folder');
        
        if (!hasRealFiles) {
          // サブフォルダのみ → 各サブフォルダも確認
          let allSubEmpty = true;
          const emptySubFolders = [];
          
          for (const subFolder of files) {
            const subContents = await googleApi(token,
              `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`'${subFolder.id}' in parents and trashed=false`)}&fields=files(id)&pageSize=1`
            );
            if ((subContents.files || []).length === 0) {
              emptySubFolders.push(subFolder);
            } else {
              allSubEmpty = false;
            }
          }
          
          // 空のサブフォルダを削除
          for (const esf of emptySubFolders) {
            await googleApi(token,
              `https://www.googleapis.com/drive/v3/files/${esf.id}`,
              { method: 'DELETE' }
            );
            deleted.push(`${folder.name}/${esf.name}`);
          }
          
          // 全サブフォルダが空だった場合、親フォルダも削除
          if (allSubEmpty && emptySubFolders.length === files.length) {
            await googleApi(token,
              `https://www.googleapis.com/drive/v3/files/${folder.id}`,
              { method: 'DELETE' }
            );
            deleted.push(folder.name);
          } else {
            kept.push({ name: folder.name, items: files.length - emptySubFolders.length });
          }
        } else {
          kept.push({ name: folder.name, items: files.length });
        }
      }
    }

    res.json({
      success: true,
      message: `${deleted.length}個の空フォルダを削除しました`,
      deleted,
      kept
    });
  } catch (e) {
    console.error('フォルダ整理エラー:', e);
    res.status(500).json({ error: e.message });
  }
};
