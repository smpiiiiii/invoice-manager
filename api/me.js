/**
 * ユーザー情報取得API
 * GET /api/me
 */
const { getSession } = require('./helpers');

module.exports = (req, res) => {
  const session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  res.json({
    email: session.email,
    name: session.name,
    picture: session.picture,
    refreshToken: session.refresh_token || null
  });
};
