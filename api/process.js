/**
 * 請求書処理API — メイン処理実行
 * POST /api/process
 * ユーザーのGmailからPDF請求書を検索→Drive保存→Gemini解析→Sheets記入
 */
const { getSession, refreshTokenIfNeeded, googleApi } = require('./helpers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let session = getSession(req);
  if (!session) return res.status(401).json({ error: 'ログインが必要です' });
  session = await refreshTokenIfNeeded(session);
  const token = session.access_token;
  const geminiKey = process.env.GEMINI_API_KEY;

  try {
    // 1. スプレッドシート・フォルダの確保
    const { sheetId, folderId } = await ensureResources(token, session.email);

    // 2. Gmail検索（PDF添付、未処理）
    const query = 'has:attachment filename:pdf -label:請求書処理済 after:2026/02/01';
    const gmailRes = await googleApi(token,
      `https://gmail.googleapis.com/gmail/v1/users/me/messages?q=${encodeURIComponent(query)}&maxResults=10`
    );

    const messageIds = (gmailRes.messages || []).map(m => m.id);
    if (messageIds.length === 0) {
      return res.json({ success: true, message: '新しい請求書メールはありません', processed: 0, errors: 0 });
    }

    // 処理済みラベルを取得or作成
    const labelId = await getOrCreateLabel(token, '請求書処理済');

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

        // PDF添付ファイルを抽出
        const pdfParts = getPdfAttachments(msg);
        if (pdfParts.length === 0) continue;

        for (const part of pdfParts) {
          // 添付ファイルデータを取得
          const attData = await googleApi(token,
            `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/attachments/${part.body.attachmentId}`
          );
          const pdfBase64 = attData.data.replace(/-/g, '+').replace(/_/g, '/');

          // Driveに保存
          const monthFolderId = await getOrCreateMonthFolder(token, folderId, yearMonth);
          const fileName = part.filename || 'invoice.pdf';
          const driveFile = await uploadToDrive(token, monthFolderId, fileName, pdfBase64);

          // Gemini解析
          const analysis = await analyzeWithGemini(geminiKey, pdfBase64);

          if (analysis && analysis.makerName && analysis.amount > 0) {
            // スプレッドシートに記入
            await updateSheet(token, sheetId, analysis.makerName, analysis.amount, yearMonth);

            // Driveファイルをリネーム
            await googleApi(token,
              `https://www.googleapis.com/drive/v3/files/${driveFile.id}`,
              { method: 'PATCH', body: JSON.stringify({ name: `${analysis.makerName}_請求書_${yearMonth.replace('/', '')}.pdf` }) }
            );

            results.push({ maker: analysis.makerName, amount: analysis.amount, month: yearMonth });
            processed++;
          } else {
            errors++;
          }
        }

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

function getPdfAttachments(msg) {
  const parts = [];
  function walk(p) {
    if (p.filename && p.filename.toLowerCase().endsWith('.pdf') && p.body && p.body.attachmentId) {
      parts.push(p);
    }
    if (p.parts) p.parts.forEach(walk);
  }
  walk(msg.payload);
  return parts;
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

async function ensureResources(token, email) {
  // スプレッドシートを検索or作成
  let sheetId = '';
  const searchRes = await googleApi(token,
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent("name='📋 請求書管理' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false")}&fields=files(id)`
  );
  if (searchRes.files && searchRes.files.length > 0) {
    sheetId = searchRes.files[0].id;
  } else {
    const created = await googleApi(token, 'https://www.googleapis.com/drive/v3/files', {
      method: 'POST', body: JSON.stringify({ name: '📋 請求書管理', mimeType: 'application/vnd.google-apps.spreadsheet' })
    });
    sheetId = created.id;
    // ヘッダー行を設定
    await googleApi(token,
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1?valueInputOption=RAW`,
      { method: 'PUT', body: JSON.stringify({ values: [['メーカー名']] }) }
    );
  }

  // Driveフォルダ検索or作成
  let folderId = '';
  const folderRes = await googleApi(token,
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent("name='📁 請求書' and mimeType='application/vnd.google-apps.folder' and trashed=false")}&fields=files(id)`
  );
  if (folderRes.files && folderRes.files.length > 0) {
    folderId = folderRes.files[0].id;
  } else {
    const created = await googleApi(token, 'https://www.googleapis.com/drive/v3/files', {
      method: 'POST', body: JSON.stringify({ name: '📁 請求書', mimeType: 'application/vnd.google-apps.folder' })
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

async function uploadToDrive(token, folderId, fileName, base64Data) {
  const boundary = 'invoice_boundary';
  const metadata = JSON.stringify({ name: fileName, parents: [folderId] });
  const body = `--${boundary}\r\nContent-Type: application/json\r\n\r\n${metadata}\r\n--${boundary}\r\nContent-Type: application/pdf\r\nContent-Transfer-Encoding: base64\r\n\r\n${base64Data}\r\n--${boundary}--`;

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

async function analyzeWithGemini(apiKey, pdfBase64) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{
      parts: [
        { inlineData: { mimeType: 'application/pdf', data: pdfBase64 } },
        { text: 'この請求書PDFから以下をJSON形式で返してください。\n1. makerName: 請求元の会社名（株式会社等は除いた短い名前）\n2. amount: 税抜金額（数値のみ）\n\nJSON形式のみ返してください: {"makerName": "会社名", "amount": 12345}\n請求書でない場合: {"makerName": null, "amount": 0}' }
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
