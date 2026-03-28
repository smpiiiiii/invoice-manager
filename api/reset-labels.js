/**
 * 処理済みラベルをリセットするAPI
 * POST /api/reset-labels
 * body: { mode: 'invoice' | 'receipt' }
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;

  let body = {};
  try { body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {}); } catch(e) {}
  const mode = body.mode === 'receipt' ? 'receipt' : 'invoice';
  const labelName = mode === 'receipt' ? '領収書処理済' : '請求書処理済';

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

    res.json({ success: true, message: `${removed}件のメールから「${labelName}」ラベルを解除しました`, removed });
  } catch (e) {
    console.error('ラベルリセットエラー:', e);
    res.status(500).json({ error: e.message });
  }
};
