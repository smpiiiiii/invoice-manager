/**
 * Google OAuth認証 — ログイン開始
 * WebView（アプリ内ブラウザ）を検出し、外部ブラウザへ誘導
 * 通常ブラウザの場合はGoogleの認可ページにリダイレクト
 */
module.exports = (req, res) => {
  const ua = (req.headers['user-agent'] || '').toLowerCase();

  // WebView（アプリ内ブラウザ）検出
  // LINE, Instagram, Facebook, Twitter, TikTok, WeChat 等のアプリ内ブラウザを検出
  const isWebView = /line\//i.test(ua) ||
    /fbav|fban/i.test(ua) ||
    /instagram/i.test(ua) ||
    /twitter|x\.com/i.test(ua) ||
    /tiktok/i.test(ua) ||
    /wechat|micromessenger/i.test(ua) ||
    /wv\)/.test(ua) ||  // Android WebView
    (/iphone|ipad/.test(ua) && !/safari/.test(ua)); // iOS WebView （Safariなし）

  if (isWebView) {
    // WebViewの場合：外部ブラウザで開くよう案内するHTMLを返す
    const appUrl = process.env.OAUTH_REDIRECT_URI
      ? process.env.OAUTH_REDIRECT_URI.replace('/api/auth/callback', '/api/auth/login')
      : `https://${process.env.VERCEL_URL || 'localhost:3000'}/api/auth/login`;

    const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ブラウザで開いてください</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Hiragino Sans", "Yu Gothic", sans-serif;
    background: linear-gradient(135deg, #0a0e1a 0%, #1a1f3a 100%);
    color: #e2e8f0;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  }
  .card {
    background: rgba(30, 41, 59, 0.95);
    border: 1px solid rgba(56, 189, 248, 0.2);
    border-radius: 16px;
    padding: 32px 24px;
    max-width: 380px;
    width: 100%;
    text-align: center;
    box-shadow: 0 8px 32px rgba(0,0,0,0.4);
  }
  .icon { font-size: 48px; margin-bottom: 16px; }
  h1 { font-size: 18px; font-weight: 700; margin-bottom: 12px; color: #f1f5f9; }
  p { font-size: 13px; color: #94a3b8; line-height: 1.6; margin-bottom: 20px; }
  .btn {
    display: block;
    width: 100%;
    padding: 14px;
    background: linear-gradient(135deg, #38bdf8, #0284c7);
    color: #fff;
    font-size: 15px;
    font-weight: 700;
    border: none;
    border-radius: 10px;
    text-decoration: none;
    cursor: pointer;
    margin-bottom: 12px;
  }
  .url-box {
    background: #0f172a;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 10px;
    font-size: 11px;
    color: #38bdf8;
    word-break: break-all;
    margin-bottom: 12px;
  }
  .copy-btn {
    background: #334155;
    color: #e2e8f0;
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    font-size: 12px;
    cursor: pointer;
  }
  .steps { text-align: left; font-size: 12px; color: #94a3b8; margin-top: 16px; }
  .steps li { margin-bottom: 6px; }
</style>
</head>
<body>
<div class="card">
  <div class="icon">🌐</div>
  <h1>外部ブラウザで開いてください</h1>
  <p>アプリ内ブラウザではGoogleログインができません。<br>Chrome や Safari で開いてください。</p>
  <div class="url-box" id="urlBox">${appUrl}</div>
  <button class="copy-btn" onclick="navigator.clipboard.writeText('${appUrl}').then(function(){document.querySelector('.copy-btn').textContent='✅ コピーしました！'})">📋 URLをコピー</button>
  <ol class="steps">
    <li>上のURLをコピー</li>
    <li>Chrome / Safari を開く</li>
    <li>アドレスバーにペーストしてアクセス</li>
  </ol>
</div>
</body>
</html>`;

    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    return res.end(html);
  }

  // 通常ブラウザ：Google OAuth認可ページにリダイレクト
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
    'https://www.googleapis.com/auth/drive',
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
