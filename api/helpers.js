/**
 * 共通ヘルパー関数
 */
const cookie = require('cookie');
const { decryptSession } = require('./session-crypto');

/**
 * CookieからセッションデータをパースしてGoogleトークンを返す
 * AES暗号化セッション(enc:プレフィックス)と旧Base64形式の両方に対応
 */
function getSession(req) {
  const cookies = cookie.parse(req.headers.cookie || '');
  if (!cookies.inv_session) return null;
  const session = decryptSession(cookies.inv_session);
  if (!session) {
    console.error('セッションの復号化に失敗');
  }
  return session;
}

/**
 * トークンが期限切れの場合リフレッシュ
 * リフレッシュ失敗時もセッションを返す（古いトークンでリトライ可能にする）
 */
async function refreshTokenIfNeeded(session) {
  if (!session || !session.refresh_token) return session;
  if (Date.now() < session.expiry - 60000) return session; // まだ有効

  try {
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

    if (!res.ok) {
      const errText = await res.text();
      console.error(`トークンリフレッシュ失敗 (${res.status}):`, errText.substring(0, 200));
      // リフレッシュ失敗でも古いトークンで再試行の余地を残す
      return session;
    }

    const data = await res.json();
    if (data.access_token) {
      session.access_token = data.access_token;
      session.expiry = Date.now() + (data.expires_in * 1000);
    } else if (data.error) {
      console.error('トークンリフレッシュエラー:', data.error, data.error_description);
    }
  } catch (e) {
    console.error('トークンリフレッシュで例外:', e.message);
  }

  return session;
}

/**
 * 認証付きGoogle APIリクエスト
 * HTTPエラー時は詳細なメッセージ付きの例外をスロー
 * DELETE等のボディなしレスポンスにも対応
 */
async function googleApi(token, url, options = {}) {
  const method = (options.method || 'GET').toUpperCase();

  // URLを短くしてログ用に整形
  const shortUrl = url.length > 80 ? url.substring(0, 80) + '...' : url;

  let res;
  try {
    res = await fetch(url, {
      ...options,
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json',
        ...(options.headers || {})
      }
    });
  } catch (networkError) {
    // ネットワークエラー（DNS解決失敗、接続タイムアウト等）
    console.error(`[GoogleAPI] ネットワークエラー ${method} ${shortUrl}:`, networkError.message);
    throw new Error(`Google APIへの接続に失敗しました: ${networkError.message}`);
  }

  // 204 No Content（DELETE等で返ることがある）
  if (res.status === 204) {
    return {};
  }

  // レスポンスボディを取得
  const responseText = await res.text();

  // 成功レスポンス
  if (res.ok) {
    try {
      return JSON.parse(responseText);
    } catch (e) {
      // JSONパース失敗だが成功ステータスの場合は空オブジェクトを返す
      if (responseText.length === 0) return {};
      console.warn(`[GoogleAPI] JSONパース失敗 ${method} ${shortUrl} (${res.status}):`, responseText.substring(0, 100));
      return {};
    }
  }

  // エラーレスポンス — 詳細なログとメッセージ
  let errorMessage = `Google API エラー (${res.status})`;
  try {
    const errorData = JSON.parse(responseText);
    if (errorData.error) {
      const gErr = typeof errorData.error === 'object'
        ? (errorData.error.message || errorData.error.status || JSON.stringify(errorData.error))
        : errorData.error;
      errorMessage = `Google API ${res.status}: ${gErr}`;
    }
  } catch (e) {
    errorMessage = `Google API ${res.status}: ${responseText.substring(0, 150)}`;
  }

  console.error(`[GoogleAPI] ${method} ${shortUrl} → ${res.status}:`, errorMessage);

  // 401 Unauthorized — トークン期限切れの可能性
  if (res.status === 401) {
    throw new Error('Googleの認証が無効です。再ログインしてください。');
  }

  // 403 Forbidden — 権限不足
  if (res.status === 403) {
    throw new Error('この操作の権限がありません。Googleアカウントの権限を確認してください。');
  }

  // 404 Not Found
  if (res.status === 404) {
    throw new Error('リソースが見つかりません（削除済みまたはアクセス不可）。');
  }

  // 429 Too Many Requests
  if (res.status === 429) {
    throw new Error('APIリクエストが多すぎます。しばらく待ってから再試行してください。');
  }

  // その他のエラー
  throw new Error(errorMessage);
}

module.exports = { getSession, refreshTokenIfNeeded, googleApi };

