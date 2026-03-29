/**
 * 統合仕分け処理API
 * POST /api/process
 * メールを一括検索 → Geminiが請求書/領収書/該当なしを自動判定 → それぞれのスプシ・Driveに保存
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

// Vercelタイムアウトを60秒に延長
module.exports.maxDuration = 60;

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;
  const geminiKey = process.env.GEMINI_API_KEY;

  try {
    // 1. 両方のスプレッドシート・フォルダを確保
    const invoiceRes = await ensureResources(token, session.email, '請求書');
    const receiptRes = await ensureResources(token, session.email, '領収書');
    // 親フォルダ確保（YYYYMM フォルダの親）
    const driveFolderId = await ensureParentFolder(token);

    // 2. Gmail統合検索（仕分け済ラベル除外）
    const processedLabel = '仕分け済';
    const query = `-label:${processedLabel} after:2026/03/01 (has:attachment OR subject:領収 OR subject:注文 OR subject:購入 OR subject:請求 OR subject:キャンセル OR subject:返品 OR subject:返金 OR subject:取消 OR subject:receipt OR subject:order OR subject:invoice OR subject:cancel)`;
    const gmailRes = await googleApi(token,
      `https://gmail.googleapis.com/gmail/v1/users/me/messages?q=${encodeURIComponent(query)}&maxResults=10`
    );
    const totalRemaining = gmailRes.resultSizeEstimate || 0;

    const messageIds = (gmailRes.messages || []).map(m => m.id);
    if (messageIds.length === 0) {
      return res.json({ success: true, message: '新しいメールはありません', processed: 0, errors: 0, debug: { query, found: 0, totalRemaining: 0 } });
    }

    // 仕分け済ラベルを取得or作成
    const labelId = await getOrCreateLabel(token, processedLabel);

    let processed = 0, errors = 0;
    const results = [];
    const debugLogs = [];

    // レート制限対策
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));

    // 429リトライ付きGemini呼び出し
    async function callGeminiWithRetry(fn, maxRetries = 1) {
      for (let attempt = 0; attempt <= maxRetries; attempt++) {
        const result = await fn();
        if (result && result._error && result._error.includes('429')) {
          console.log(`429エラー、${5 * (attempt + 1)}秒待機してリトライ`);
          await sleep(5000 * (attempt + 1));
          continue;
        }
        return result;
      }
      return { _error: 'レート制限超過（リトライ失敗）' };
    }

    for (const msgId of messageIds) {
      try {
        const msg = await googleApi(token,
          `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}?format=full`
        );

        const mailDate = new Date(parseInt(msg.internalDate));
        const yearMonth = mailDate.getFullYear() + '/' + String(mailDate.getMonth() + 1).padStart(2, '0');
        const yearMonthCompact = mailDate.getFullYear() + String(mailDate.getMonth() + 1).padStart(2, '0');
        const subject = getHeader(msg, 'Subject') || '(件名なし)';
        const fromAddr = getHeader(msg, 'From') || '';
        const attachParts = getDocAttachments(msg);

        const dateStr = mailDate.getFullYear() + '/' + String(mailDate.getMonth() + 1).padStart(2, '0') + '/' + String(mailDate.getDate()).padStart(2, '0');
        const logEntry = { subject: subject.substring(0, 60), from: fromAddr.substring(0, 40), date: dateStr, msgId: msg.threadId || msgId, attachments: attachParts.length, status: '', type: '' };
        let hasProcessedSomething = false;
        let hitRateLimit = false;

        // A. 添付ファイルがあれば解析（タイプ自動判定）
        for (const part of attachParts) {
          const attData = await googleApi(token,
            `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/attachments/${part.body.attachmentId}`
          );
          const fileBase64 = attData.data.replace(/-/g, '+').replace(/_/g, '/');
          const ext = getFileExtension(part.filename);
          const mimeType = getMimeType(ext);

          const analysis = await callGeminiWithRetry(() => analyzeFileWithGemini(geminiKey, fileBase64, mimeType));

          if (analysis && analysis._error) {
            logEntry.status = `⚠️ ${analysis._error.substring(0, 50)}`;
            hitRateLimit = true;
            break;
          }

          if (analysis && analysis.makerName && analysis.amount > 0 && analysis.type !== 'none') {
            const docType = analysis.type === 'invoice' ? '請求書' : '領収書';
            const targetSheet = analysis.type === 'invoice' ? invoiceRes.sheetId : receiptRes.sheetId;

            // タイプ別月フォルダに保存
            const monthFolderName = `${yearMonthCompact} ${docType}`;
            const monthFolderId = await getOrCreateTypedMonthFolder(token, driveFolderId, monthFolderName);
            const fileName = part.filename || `${docType}.${ext}`;
            const driveFile = await uploadToDrive(token, monthFolderId, fileName, fileBase64, mimeType);

            // ファイル名変更
            await googleApi(token,
              `https://www.googleapis.com/drive/v3/files/${driveFile.id}`,
              { method: 'PATCH', body: JSON.stringify({ name: `${analysis.makerName}_${docType}_${yearMonthCompact}.${ext}` }) }
            );

            await updateSheet(token, targetSheet, analysis.makerName, analysis.amount, yearMonth);
            results.push({ maker: analysis.makerName, amount: analysis.amount, month: yearMonth, type: docType });
            logEntry.status = `${docType === '請求書' ? '📄' : '🧾'} ${docType} → ${analysis.makerName} ¥${analysis.amount.toLocaleString()}`;
            logEntry.type = analysis.type;
            processed++;
            hasProcessedSomething = true;
          } else if (analysis && analysis.type === 'cancel' && analysis.makerName) {
            // キャンセル処理: 両方のシートから削除試行
            const removed1 = await removeFromSheet(token, invoiceRes.sheetId, analysis.makerName, yearMonth);
            const removed2 = await removeFromSheet(token, receiptRes.sheetId, analysis.makerName, yearMonth);
            const removedAmt = removed1 + removed2;
            if (removedAmt > 0) {
              results.push({ maker: analysis.makerName, amount: -removedAmt, month: yearMonth, type: 'キャンセル' });
              logEntry.status = `❌ キャンセル → ${analysis.makerName} -¥${removedAmt.toLocaleString()}`;
            } else {
              logEntry.status = `❌ キャンセル → ${analysis.makerName}（元データなし）`;
            }
            logEntry.type = 'cancel';
            processed++;
            hasProcessedSomething = true;
          } else {
            logEntry.status = '⏭️ 該当なし';
          }
        }

        // B. 添付なし or 添付で該当なし → メール本文を解析
        if (!hasProcessedSomething && !hitRateLimit) {
          const bodyText = extractEmailBody(msg);
          if (bodyText && bodyText.length > 20) {
            const analysis = await callGeminiWithRetry(() => analyzeBodyWithGemini(geminiKey, subject, fromAddr, bodyText));

            if (analysis && analysis._error) {
              logEntry.status = `⚠️ ${analysis._error.substring(0, 50)}`;
              hitRateLimit = true;
            } else if (analysis && analysis.makerName && analysis.amount > 0 && analysis.type !== 'none') {
              const docType = analysis.type === 'invoice' ? '請求書' : '領収書';
              const targetSheet = analysis.type === 'invoice' ? invoiceRes.sheetId : receiptRes.sheetId;

              await updateSheet(token, targetSheet, analysis.makerName, analysis.amount, yearMonth);
              results.push({ maker: analysis.makerName, amount: analysis.amount, month: yearMonth, type: docType, source: 'メール本文' });
              logEntry.status = `${docType === '請求書' ? '📄' : '🧾'} ${docType} → ${analysis.makerName} ¥${analysis.amount.toLocaleString()}`;
              logEntry.type = analysis.type;
              processed++;
              hasProcessedSomething = true;
            } else if (analysis && analysis.type === 'cancel' && analysis.makerName) {
              const removed1 = await removeFromSheet(token, invoiceRes.sheetId, analysis.makerName, yearMonth);
              const removed2 = await removeFromSheet(token, receiptRes.sheetId, analysis.makerName, yearMonth);
              const removedAmt = removed1 + removed2;
              if (removedAmt > 0) {
                results.push({ maker: analysis.makerName, amount: -removedAmt, month: yearMonth, type: 'キャンセル', source: 'メール本文' });
                logEntry.status = `❌ キャンセル → ${analysis.makerName} -¥${removedAmt.toLocaleString()}`;
              } else {
                logEntry.status = `❌ キャンセル → ${analysis.makerName}（元データなし）`;
              }
              logEntry.type = 'cancel';
              processed++;
              hasProcessedSomething = true;
            } else {
              logEntry.status = '⏭️ 該当なし';
            }
          } else if (!logEntry.status) {
            logEntry.status = '⏭️ 本文なし';
          }
        }

        if (!hasProcessedSomething && !hitRateLimit && attachParts.length > 0) errors++;
        debugLogs.push(logEntry);

        // レート制限でなければ仕分け済ラベル付与
        if (!hitRateLimit) {
          await googleApi(token,
            `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/modify`,
            { method: 'POST', body: JSON.stringify({ addLabelIds: [labelId] }) }
          );
        }
      } catch (e) {
        console.error('メール処理エラー:', e.message);
        debugLogs.push({ subject: '???', status: `❌ ${e.message.substring(0, 50)}` });
        errors++;
      }
    }

    res.json({ success: true, message: `処理完了: ${processed}件成功, ${errors}件エラー`, processed, errors, results, debug: { query, found: messageIds.length, totalRemaining, logs: debugLogs } });
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
 * メール本文をGeminiで解析（タイプ自動判定付き）
 */
async function analyzeBodyWithGemini(apiKey, subject, fromAddr, bodyText) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const prompt = `以下はメールの件名・差出人・本文です。このメールが「請求書」「領収書（購入・注文・決済・領収）」または「キャンセル・返金」のどれかに該当するか判定し、情報を抽出してください。

件名: ${subject}
差出人: ${fromAddr}
本文:
${bodyText}

以下をJSON形式のみで返してください:
1. type: "invoice"(請求書) or "receipt"(領収書・購入・注文・決済) or "cancel"(キャンセル・返金・取消) or "none"(該当なし)
2. makerName: 店舗名・サービス名・会社名（短い名前）
3. amount: 金額（請求書なら税抜き、領収書なら税込み、キャンセルなら0、数値のみ）

JSON形式のみ返してください: {"type": "receipt", "makerName": "会社名", "amount": 12345}
※見積書・概算・クーポン・お知らせのみのメールは除外。
※注文・商品のキャンセル・返品・返金のメールは type:"cancel" で返してください。
該当なしの場合: {"type": "none", "makerName": null, "amount": 0}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1, thinkingConfig: { thinkingBudget: 0 } }
  };

  const gemRes = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (!gemRes.ok) {
    const errText = await gemRes.text();
    return { _error: `API ${gemRes.status}: ${errText.substring(0, 200)}` };
  }
  const data = await gemRes.json();
  try {
    let text = data.candidates[0].content.parts[0].text;
    text = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
    const result = JSON.parse(text);
    return {
      type: result.type || 'none',
      makerName: result.makerName ? String(result.makerName).trim() : null,
      amount: parseInt(result.amount) || 0
    };
  } catch (e) {
    return { _error: `Parse失敗: ${e.message}`, _raw: JSON.stringify(data).substring(0, 300) };
  }
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

/**
 * スプレッドシートを検索or作成
 */
async function ensureResources(token, email, modeLabel = '請求書') {
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
    await googleApi(token,
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1?valueInputOption=USER_ENTERED`,
      { method: 'PUT', body: JSON.stringify({ values: [['メーカー名']] }) }
    );
  }
  return { sheetId };
}

/**
 * 親フォルダ確保（請求書・領収書管理）
 */
async function ensureParentFolder(token) {
  const folderName = '📂 請求書・領収書管理';
  const folderRes = await googleApi(token,
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id)`
  );
  if (folderRes.files && folderRes.files.length > 0) return folderRes.files[0].id;

  const created = await googleApi(token, 'https://www.googleapis.com/drive/v3/files', {
    method: 'POST', body: JSON.stringify({ name: folderName, mimeType: 'application/vnd.google-apps.folder' })
  });
  return created.id;
}

/**
 * タイプ別月フォルダ（例: "202603 請求書"）
 */
async function getOrCreateTypedMonthFolder(token, parentId, folderName) {
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

/**
 * 添付ファイルをGeminiで解析（タイプ自動判定付き）
 */
async function analyzeFileWithGemini(apiKey, fileBase64, mimeType = 'application/pdf') {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const prompt = `このファイルを分析して以下を判定してください:
1. type: "invoice"(請求書) または "receipt"(領収書) または "none"(該当なし)
2. makerName: 発行元の会社名・店名（株式会社等は除いた短い名前）
3. amount: 金額（請求書なら税抜き金額、領収書なら税込み金額、数値のみ）

JSON形式のみ返してください: {"type": "invoice", "makerName": "会社名", "amount": 12345}
※見積書・見積もり・査定・概算は除外。実際の請求書または領収書のみ対象。
該当なしの場合: {"type": "none", "makerName": null, "amount": 0}`;

  const payload = {
    contents: [{
      parts: [
        { inlineData: { mimeType, data: fileBase64 } },
        { text: prompt }
      ]
    }],
    generationConfig: { temperature: 0.1, thinkingConfig: { thinkingBudget: 0 } }
  };

  const gemRes = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (!gemRes.ok) {
    const errText = await gemRes.text();
    return { _error: `API ${gemRes.status}: ${errText.substring(0, 200)}` };
  }
  const data = await gemRes.json();
  try {
    let text = data.candidates[0].content.parts[0].text;
    text = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
    const result = JSON.parse(text);
    return {
      type: result.type || 'none',
      makerName: result.makerName ? String(result.makerName).trim() : null,
      amount: parseInt(result.amount) || 0
    };
  } catch (e) {
    return { _error: `Parse: ${e.message}`, _raw: JSON.stringify(data).substring(0, 300) };
  }
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
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1?valueInputOption=USER_ENTERED`,
    { method: 'PUT', body: JSON.stringify({ values }) }
  );
}

/**
 * キャンセル処理: スプシからメーカーの該当月の金額を削除し、削除額を返す
 * 部分一致でメーカー名を検索（「BAYCREW'S」で「BAYCREW'S STORE」もマッチ）
 */
async function removeFromSheet(token, sheetId, makerName, yearMonth) {
  const dataRes = await googleApi(token,
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A:ZZ`
  );
  const values = dataRes.values || [['メーカー名']];
  const searchName = makerName.toLowerCase().replace(/[\s'’　]/g, '');

  // メーカー行を部分一致で検索
  let rowIdx = -1;
  for (let i = 1; i < values.length; i++) {
    if (!values[i] || !values[i][0]) continue;
    const sheetName = values[i][0].toLowerCase().replace(/[\s'’　]/g, '');
    // 完全一致 or どちらかが含まれる
    if (sheetName === searchName || sheetName.includes(searchName) || searchName.includes(sheetName)) {
      rowIdx = i;
      break;
    }
  }
  if (rowIdx === -1) return 0;

  // 年月列を検索
  let colIdx = -1;
  for (let j = 1; j < values[0].length; j++) {
    if (values[0][j] === yearMonth) { colIdx = j; break; }
  }
  if (colIdx === -1) return 0;

  const existing = parseInt(values[rowIdx][colIdx]) || 0;
  if (existing <= 0) return 0;

  // 金額を0にする
  values[rowIdx][colIdx] = '';

  await googleApi(token,
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1?valueInputOption=USER_ENTERED`,
    { method: 'PUT', body: JSON.stringify({ values }) }
  );

  return existing;
}
