/**
 * 請求書/領収書処理API — メイン処理実行
 * POST /api/process
 * ユーザーのGmailからPDF/PNG/JPG添付を検索→Drive保存→Gemini解析→Sheets記入
 * bodyパラメータ: mode = 'invoice'(請求書) | 'receipt'(領収書)
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;
  const geminiKey = process.env.GEMINI_API_KEY;

  // モード判定（請求書 or 領収書）
  let body = {};
  try { body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {}); } catch(e) {}
  const mode = body.mode === 'receipt' ? 'receipt' : 'invoice';
  const modeLabel = mode === 'receipt' ? '領収書' : '請求書';
  const processedLabel = mode === 'receipt' ? '領収書処理済' : '請求書処理済';

  try {
    // 1. スプレッドシート・フォルダの確保
    const { sheetId, folderId } = await ensureResources(token, session.email, modeLabel);

    // 2. Gmail検索（モード別クエリ）
    let query;
    if (mode === 'receipt') {
      // 領収書モード: 添付付き + 本文に領収書/注文/購入系キーワードがあるメール
      query = `-label:${processedLabel} after:2026/03/01 {has:attachment (filename:pdf OR filename:png OR filename:jpg OR filename:jpeg) subject:(領収書 OR 領収 OR 注文確認 OR 購入 OR ご利用明細 OR receipt OR order)}`;
    } else {
      // 請求書モード: 添付ファイルがあるメール
      query = `has:attachment (filename:pdf OR filename:png OR filename:jpg OR filename:jpeg) -label:${processedLabel} after:2026/03/01`;
    }
    const gmailRes = await googleApi(token,
      `https://gmail.googleapis.com/gmail/v1/users/me/messages?q=${encodeURIComponent(query)}&maxResults=15`
    );

    const messageIds = (gmailRes.messages || []).map(m => m.id);
    if (messageIds.length === 0) {
      return res.json({ success: true, message: `新しい${modeLabel}メールはありません`, processed: 0, errors: 0 });
    }

    // 処理済みラベルを取得or作成
    const labelId = await getOrCreateLabel(token, processedLabel);

    let processed = 0, errors = 0;
    const results = [];

    for (const msgId of messageIds) {
      try {
        // メール詳細を取得
        const msg = await googleApi(token,
          `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}?format=full`
        );

        const mailDate = new Date(parseInt(msg.internalDate));
        const yearMonth = mailDate.getFullYear() + '/' + String(mailDate.getMonth() + 1).padStart(2, '0');
        const subject = getHeader(msg, 'Subject') || '(件名なし)';
        const fromAddr = getHeader(msg, 'From') || '';

        // 添付ファイルを抽出（PDF/PNG/JPG）
        const attachParts = getDocAttachments(msg);

        let hasProcessedSomething = false;

        // A. 添付ファイルがあれば解析
        for (const part of attachParts) {
          const attData = await googleApi(token,
            `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/attachments/${part.body.attachmentId}`
          );
          const fileBase64 = attData.data.replace(/-/g, '+').replace(/_/g, '/');
          const ext = getFileExtension(part.filename);
          const mimeType = getMimeType(ext);

          // Driveに保存
          const monthFolderId = await getOrCreateMonthFolder(token, folderId, yearMonth);
          const fileName = part.filename || `${modeLabel}.${ext}`;
          const driveFile = await uploadToDrive(token, monthFolderId, fileName, fileBase64, mimeType);

          // Gemini解析
          const analysis = await analyzeWithGemini(geminiKey, fileBase64, mimeType, mode);

          if (analysis && analysis.makerName && analysis.amount > 0) {
            await updateSheet(token, sheetId, analysis.makerName, analysis.amount, yearMonth);
            await googleApi(token,
              `https://www.googleapis.com/drive/v3/files/${driveFile.id}`,
              { method: 'PATCH', body: JSON.stringify({ name: `${analysis.makerName}_${modeLabel}_${yearMonth.replace('/', '')}.${ext}` }) }
            );
            results.push({ maker: analysis.makerName, amount: analysis.amount, month: yearMonth });
            processed++;
            hasProcessedSomething = true;
          }
        }

        // B. 領収書モード: 添付なし or 添付から見つからなかった場合、メール本文を解析
        if (mode === 'receipt' && !hasProcessedSomething) {
          const bodyText = extractEmailBody(msg);
          if (bodyText && bodyText.length > 20) {
            const analysis = await analyzeBodyWithGemini(geminiKey, subject, fromAddr, bodyText);
            if (analysis && analysis.makerName && analysis.amount > 0) {
              await updateSheet(token, sheetId, analysis.makerName, analysis.amount, yearMonth);
              results.push({ maker: analysis.makerName, amount: analysis.amount, month: yearMonth, source: 'メール本文' });
              processed++;
              hasProcessedSomething = true;
            }
          }
        }

        if (!hasProcessedSomething && attachParts.length > 0) errors++;

        // 処理済みラベルを付与
        await googleApi(token,
          `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/modify`,
          { method: 'POST', body: JSON.stringify({ addLabelIds: [labelId] }) }
        );
      } catch (e) {
        console.error('メール処理エラー:', e.message);
        errors++;
      }
    }

    res.json({ success: true, message: `処理完了: ${processed}件成功, ${errors}件エラー`, processed, errors, results });
  } catch (e) {
    console.error('処理エラー:', e);
    res.status(500).json({ error: e.message });
  }
};

// === ヘルパー ===

function getHeader(msg, name) {
  const h = (msg.payload.headers || []).find(h => h.name.toLowerCase() === name.toLowerCase());
  return h ? h.value : '';
}

/**
 * メール本文テキストを抽出（text/plain優先、なければHTMLからタグ除去）
 */
function extractEmailBody(msg) {
  let textBody = '';
  let htmlBody = '';

  function walk(p) {
    if (p.mimeType === 'text/plain' && p.body && p.body.data) {
      textBody += Buffer.from(p.body.data, 'base64').toString('utf-8');
    }
    if (p.mimeType === 'text/html' && p.body && p.body.data) {
      htmlBody += Buffer.from(p.body.data, 'base64').toString('utf-8');
    }
    if (p.parts) p.parts.forEach(walk);
  }
  walk(msg.payload);

  if (textBody.length > 0) return textBody.substring(0, 3000);

  // HTMLからタグを除去
  if (htmlBody.length > 0) {
    const stripped = htmlBody
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&yen;/g, '¥')
      .replace(/\s+/g, ' ')
      .trim();
    return stripped.substring(0, 3000);
  }

  return '';
}

/**
 * メール本文テキストをGeminiで解析して領収書情報を抽出
 */
async function analyzeBodyWithGemini(apiKey, subject, fromAddr, bodyText) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const prompt = `以下はメールの件名・差出人・本文です。これが購入・注文・決済・領収に関するメールかどうか判定し、該当する場合は情報を抽出してください。

件名: ${subject}
差出人: ${fromAddr}
本文:
${bodyText}

以下をJSON形式のみで返してください:
1. makerName: 店舗名・サービス名・会社名（短い名前）
2. amount: 支払い金額（税込、数値のみ）

JSON形式のみ返してください: {"makerName": "会社名", "amount": 12345}
購入・注文・決済・領収に関するメールでない場合: {"makerName": null, "amount": 0}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1 }
  };

  const gemRes = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (!gemRes.ok) return null;
  const data = await gemRes.json();
  try {
    let text = data.candidates[0].content.parts[0].text;
    text = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
    const result = JSON.parse(text);
    return { makerName: result.makerName ? String(result.makerName).trim() : null, amount: parseInt(result.amount) || 0 };
  } catch (e) { return null; }
}

// 対応ファイル形式
const SUPPORTED_EXTENSIONS = ['.pdf', '.png', '.jpg', '.jpeg'];

function getDocAttachments(msg) {
  const parts = [];
  function walk(p) {
    if (p.filename && p.body && p.body.attachmentId) {
      const lower = p.filename.toLowerCase();
      if (SUPPORTED_EXTENSIONS.some(ext => lower.endsWith(ext))) {
        parts.push(p);
      }
    }
    if (p.parts) p.parts.forEach(walk);
  }
  walk(msg.payload);
  return parts;
}

function getFileExtension(filename) {
  if (!filename) return 'pdf';
  const lower = filename.toLowerCase();
  if (lower.endsWith('.png')) return 'png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'jpg';
  return 'pdf';
}

function getMimeType(ext) {
  switch (ext) {
    case 'png': return 'image/png';
    case 'jpg': case 'jpeg': return 'image/jpeg';
    default: return 'application/pdf';
  }
}

async function getOrCreateLabel(token, labelName) {
  const labels = await googleApi(token, 'https://gmail.googleapis.com/gmail/v1/users/me/labels');
  const existing = (labels.labels || []).find(l => l.name === labelName);
  if (existing) return existing.id;

  const created = await googleApi(token, 'https://gmail.googleapis.com/gmail/v1/users/me/labels', {
    method: 'POST', body: JSON.stringify({ name: labelName, labelListVisibility: 'labelShow', messageListVisibility: 'show' })
  });
  return created.id;
}

async function ensureResources(token, email, modeLabel = '請求書') {
  // スプレッドシートを検索or作成
  const sheetName = `📋 ${modeLabel}管理`;
  let sheetId = '';
  const searchRes = await googleApi(token,
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${sheetName}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`)}&fields=files(id)`
  );
  if (searchRes.files && searchRes.files.length > 0) {
    sheetId = searchRes.files[0].id;
  } else {
    const created = await googleApi(token, 'https://www.googleapis.com/drive/v3/files', {
      method: 'POST', body: JSON.stringify({ name: sheetName, mimeType: 'application/vnd.google-apps.spreadsheet' })
    });
    sheetId = created.id;
    // ヘッダー行を設定
    await googleApi(token,
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1?valueInputOption=RAW`,
      { method: 'PUT', body: JSON.stringify({ values: [['メーカー名']] }) }
    );
  }

  // Driveフォルダ検索or作成
  const folderName = `📁 ${modeLabel}`;
  let folderId = '';
  const folderRes = await googleApi(token,
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id)`
  );
  if (folderRes.files && folderRes.files.length > 0) {
    folderId = folderRes.files[0].id;
  } else {
    const created = await googleApi(token, 'https://www.googleapis.com/drive/v3/files', {
      method: 'POST', body: JSON.stringify({ name: folderName, mimeType: 'application/vnd.google-apps.folder' })
    });
    folderId = created.id;
  }

  return { sheetId, folderId };
}

async function getOrCreateMonthFolder(token, parentId, yearMonth) {
  const parts = yearMonth.split('/');
  const folderName = parts[0] + '年' + parts[1] + '月';

  const searchRes = await googleApi(token,
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${folderName}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id)`
  );
  if (searchRes.files && searchRes.files.length > 0) return searchRes.files[0].id;

  const created = await googleApi(token, 'https://www.googleapis.com/drive/v3/files', {
    method: 'POST', body: JSON.stringify({ name: folderName, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] })
  });
  return created.id;
}

async function uploadToDrive(token, folderId, fileName, base64Data, mimeType = 'application/pdf') {
  const boundary = 'invoice_boundary';
  const metadata = JSON.stringify({ name: fileName, parents: [folderId] });
  const body = `--${boundary}\r\nContent-Type: application/json\r\n\r\n${metadata}\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\nContent-Transfer-Encoding: base64\r\n\r\n${base64Data}\r\n--${boundary}--`;

  const uploadRes = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': `multipart/related; boundary=${boundary}`
    },
    body
  });
  return uploadRes.json();
}

async function analyzeWithGemini(apiKey, fileBase64, mimeType = 'application/pdf', mode = 'invoice') {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  // モードに応じたプロンプト
  const prompt = mode === 'receipt'
    ? 'この領収書ファイルから以下をJSON形式で返してください。\n1. makerName: 発行元の会社名・店名（株式会社等は除いた短い名前）\n2. amount: 金額（税込、数値のみ）\n\nJSON形式のみ返してください: {"makerName": "会社名", "amount": 12345}\n領収書でない場合: {"makerName": null, "amount": 0}'
    : 'この請求書ファイルから以下をJSON形式で返してください。\n1. makerName: 請求元の会社名（株式会社等は除いた短い名前）\n2. amount: 税抜金額（数値のみ）\n\nJSON形式のみ返してください: {"makerName": "会社名", "amount": 12345}\n請求書でない場合: {"makerName": null, "amount": 0}';

  const payload = {
    contents: [{
      parts: [
        { inlineData: { mimeType, data: fileBase64 } },
        { text: prompt }
      ]
    }],
    generationConfig: { temperature: 0.1 }
  };

  const gemRes = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (!gemRes.ok) return null;
  const data = await gemRes.json();
  try {
    let text = data.candidates[0].content.parts[0].text;
    text = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
    const result = JSON.parse(text);
    return { makerName: result.makerName ? String(result.makerName).trim() : null, amount: parseInt(result.amount) || 0 };
  } catch (e) { return null; }
}

async function updateSheet(token, sheetId, makerName, amount, yearMonth) {
  // 既存データを取得
  const dataRes = await googleApi(token,
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A:ZZ`
  );
  const values = dataRes.values || [['メーカー名']];

  // メーカー行を検索
  let rowIdx = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i] && values[i][0] === makerName) { rowIdx = i; break; }
  }

  // 年月列を検索
  let colIdx = -1;
  for (let j = 1; j < values[0].length; j++) {
    if (values[0][j] === yearMonth) { colIdx = j; break; }
  }

  // なければ追加
  if (colIdx === -1) {
    colIdx = values[0].length;
    values[0].push(yearMonth);
  }
  if (rowIdx === -1) {
    rowIdx = values.length;
    values.push([makerName]);
  }

  // 行を必要な長さに拡張
  while (values[rowIdx].length <= colIdx) values[rowIdx].push('');
  const existing = parseInt(values[rowIdx][colIdx]) || 0;
  values[rowIdx][colIdx] = existing + amount;

  // 書き戻し
  await googleApi(token,
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1?valueInputOption=RAW`,
    { method: 'PUT', body: JSON.stringify({ values }) }
  );
}
