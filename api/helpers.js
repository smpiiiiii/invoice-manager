/**
 * 共通ヘルパー関数
 */
const cookie = require('cookie');

/**
 * CookieからセッションデータをパースしてGoogleトークンを返す
 */
function getSession(req) {
  const cookies = cookie.parse(req.headers.cookie || '');
  if (!cookies.inv_session) return null;
  try {
    return JSON.parse(Buffer.from(cookies.inv_session, 'base64').toString());
  } catch (e) {
    return null;
  }
}

/**
 * トークンが期限切れの場合リフレッシュ
 */
async function refreshTokenIfNeeded(session) {
  if (!session || !session.refresh_token) return session;
  if (Date.now() < session.expiry - 60000) return session; // まだ有効

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: process.env.GOOGLE_CLIENT_ID,
      client_secret: process.env.GOOGLE_CLIENT_SECRET,
      refresh_token: session.refresh_token,
      grant_type: 'refresh_token'
    }).toString()
  });

  const data = await res.json();
  if (data.access_token) {
    session.access_token = data.access_token;
    session.expiry = Date.now() + (data.expires_in * 1000);
  }
  return session;
}

/**
 * 認証付きGoogle APIリクエスト
 */
async function googleApi(token, url, options = {}) {
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json',
      ...(options.headers || {})
    }
  });
  return res.json();
}

module.exports = { getSession, refreshTokenIfNeeded, googleApi };
