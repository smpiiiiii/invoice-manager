/**
 * セッション暗号化ヘルパー
 * AES-256-GCMでセッションデータを暗号化・復号化
 * 環境変数 SESSION_SECRET（32バイト以上）を鍵として使用
 * SESSION_SECRETが未設定の場合はBase64エンコードにフォールバック
 */
const crypto = require('crypto');

const ALGORITHM = 'aes-256-gcm';
const IV_LENGTH = 12;
const TAG_LENGTH = 16;

/**
 * 暗号化鍵を取得（環境変数から32バイトのキーを導出）
 */
function getKey() {
  const secret = process.env.SESSION_SECRET;
  if (!secret) return null;
  // SHA-256で常に32バイトのキーを生成
  return crypto.createHash('sha256').update(secret).digest();
}

/**
 * セッションデータを暗号化してエンコード
 * SESSION_SECRETがない場合はBase64フォールバック
 */
function encryptSession(data) {
  const json = JSON.stringify(data);
  const key = getKey();

  if (!key) {
    // フォールバック: Base64エンコード（後方互換性）
    return Buffer.from(json).toString('base64');
  }

  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);
  let encrypted = cipher.update(json, 'utf8', 'hex');
  encrypted += cipher.final('hex');
  const tag = cipher.getAuthTag();

  // iv + tag + encrypted をBase64エンコード
  const combined = Buffer.concat([iv, tag, Buffer.from(encrypted, 'hex')]);
  return 'enc:' + combined.toString('base64');
}

/**
 * 暗号化されたセッションを復号化
 * enc: プレフィックスがない場合はBase64フォールバック
 */
function decryptSession(encoded) {
  if (!encoded) return null;

  try {
    // 暗号化されたデータ
    if (encoded.startsWith('enc:')) {
      const key = getKey();
      if (!key) {
        console.error('SESSION_SECRETが設定されていません（暗号化セッションの復号不可）');
        return null;
      }

      const combined = Buffer.from(encoded.substring(4), 'base64');
      const iv = combined.subarray(0, IV_LENGTH);
      const tag = combined.subarray(IV_LENGTH, IV_LENGTH + TAG_LENGTH);
      const encrypted = combined.subarray(IV_LENGTH + TAG_LENGTH);

      const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
      decipher.setAuthTag(tag);
      let decrypted = decipher.update(encrypted, undefined, 'utf8');
      decrypted += decipher.final('utf8');
      return JSON.parse(decrypted);
    }

    // フォールバック: Base64デコード（旧形式との互換性）
    return JSON.parse(Buffer.from(encoded, 'base64').toString());
  } catch (e) {
    console.error('セッション復号化失敗:', e.message);
    return null;
  }
}

module.exports = { encryptSession, decryptSession };
