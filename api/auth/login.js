/**
 * Google OAuth認証 — ログイン開始
 * ユーザーをGoogleの認可ページにリダイレクト
 */
module.exports = (req, res) => {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const redirectUri = process.env.VERCEL_URL
    ? `https://${process.env.VERCEL_URL}/api/auth/callback`
    : 'http://localhost:3000/api/auth/callback';

  // 本番URLの場合は固定ドメインを使用
  const finalRedirect = process.env.OAUTH_REDIRECT_URI || redirectUri;

  const scopes = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.labels',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/spreadsheets',
    'openid',
    'email',
    'profile'
  ].join(' ');

  const authUrl = 'https://accounts.google.com/o/oauth2/v2/auth?' + new URLSearchParams({
    client_id: clientId,
    redirect_uri: finalRedirect,
    response_type: 'code',
    scope: scopes,
    access_type: 'offline',
    prompt: 'consent'
  }).toString();

  res.writeHead(302, { Location: authUrl });
  res.end();
};
