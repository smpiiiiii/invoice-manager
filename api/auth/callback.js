/**
 * Google OAuth認証 — コールバック
 * 認可コード → アクセストークン取得 → Cookie保存 → ダッシュボードにリダイレクト
 */
const { encryptSession } = require('../session-crypto');

module.exports = async (req, res) => {
  const { code, error } = req.query;

  if (error) {
    res.writeHead(302, { Location: '/?error=auth_denied' });
    return res.end();
  }

  if (!code) {
    res.writeHead(302, { Location: '/?error=no_code' });
    return res.end();
  }

  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const redirectUri = process.env.OAUTH_REDIRECT_URI ||
    `https://${process.env.VERCEL_URL}/api/auth/callback`;

  try {
    // 認可コードをトークンに交換
    const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        code,
        client_id: clientId,
        client_secret: clientSecret,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code'
      }).toString()
    });

    const tokenData = await tokenRes.json();

    if (tokenData.error) {
      console.error('トークン取得エラー:', tokenData);
      res.writeHead(302, { Location: '/?error=token_failed' });
      return res.end();
    }

    // ユーザー情報を取得
    const userRes = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: 'Bearer ' + tokenData.access_token }
    });
    const userInfo = await userRes.json();

    // セッションデータを構築
    const sessionData = {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expiry: Date.now() + (tokenData.expires_in * 1000),
      email: userInfo.email,
      name: userInfo.name,
      picture: userInfo.picture
    };

    // AES暗号化してCookieに保存（SESSION_SECRET未設定時はBase64フォールバック）
    const encoded = encryptSession(sessionData);

    // HttpOnly Cookie設定（7日間、secure）
    const cookie = `inv_session=${encoded}; Path=/; HttpOnly; SameSite=Lax; Max-Age=604800` +
      (process.env.VERCEL_URL ? '; Secure' : '');

    res.writeHead(302, {
      'Set-Cookie': cookie,
      Location: '/dashboard'
    });
    res.end();
  } catch (e) {
    console.error('OAuth error:', e);
    res.writeHead(302, { Location: '/?error=server_error' });
    res.end();
  }
};
