/**
 * ラベルリセットAPI
 * POST /api/reset-labels
 * Gmailの「仕分け済」ラベルを全メールから削除し、ラベル自体も削除
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  try {
    // 「仕分け済」ラベルを検索
    const labelsRes = await googleApi(token,
      'https://gmail.googleapis.com/gmail/v1/users/me/labels'
    );
    const labels = labelsRes.labels || [];
    const targetLabel = labels.find(l => l.name === '仕分け済');

    if (!targetLabel) {
      return res.json({ success: true, message: '「仕分け済」ラベルは存在しません' });
    }

    const labelId = targetLabel.id;

    // ラベルが付いたメールを全取得
    let allMessageIds = [];
    let pageToken = '';
    do {
      const url = `https://gmail.googleapis.com/gmail/v1/users/me/messages?labelIds=${labelId}&maxResults=100${pageToken ? '&pageToken=' + pageToken : ''}`;
      const msgRes = await googleApi(token, url);
      const messages = msgRes.messages || [];
      allMessageIds = allMessageIds.concat(messages.map(m => m.id));
      pageToken = msgRes.nextPageToken || '';
    } while (pageToken);

    // バッチでラベルを削除
    if (allMessageIds.length > 0) {
      // 1000件ずつバッチ処理
      for (let i = 0; i < allMessageIds.length; i += 1000) {
        const batch = allMessageIds.slice(i, i + 1000);
        await googleApi(token,
          'https://gmail.googleapis.com/gmail/v1/users/me/messages/batchModify',
          {
            method: 'POST',
            body: JSON.stringify({
              ids: batch,
              removeLabelIds: [labelId]
            })
          }
        );
      }
    }

    // ラベル自体を削除
    await googleApi(token,
      `https://gmail.googleapis.com/gmail/v1/users/me/labels/${labelId}`,
      { method: 'DELETE' }
    );

    res.json({
      success: true,
      message: `${allMessageIds.length}件のメールからラベルを削除し、「仕分け済」ラベルを削除しました`
    });
  } catch (e) {
    console.error('ラベルリセットエラー:', e);
    res.status(500).json({ error: e.message });
  }
};
