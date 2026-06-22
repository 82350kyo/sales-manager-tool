// =====================================================
// 売上管理ツール GASバックエンド（シンプル版）
// スプレッドシートにバインドして使用（Extensions → Apps Script）
// =====================================================

const SHEET_NAME = '売上報告';
const LINE_GROUP_ID = 'C87ef2d287f760158419805a3887ac7f9';

// =====================================================
// LINE一括インポート用（doPost）
// =====================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload && payload.type === 'bulkImport') {
      const result = bulkImportToSheet(payload.rows);
      return respond(result);
    }
    return respond({ ok: false, error: 'unknown type' });
  } catch (err) {
    return respond({ ok: false, error: err.toString() });
  }
}

function bulkImportToSheet(rows) {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(SHEET_NAME + ' シートが見つかりません');

  const data = rows.map((d, i) => {
    // timestamp: 日付 + インデックス（秒）で同日内の順序を保持
    const ts = d.timestamp ? new Date(d.timestamp) : new Date(d.date + 'T00:00:00');
    return [
      ts,               // A: タイムスタンプ
      d.date || '',     // B: 日時
      d.memberName || '',  // C: 担当者
      d.category || '',    // D: 導線
      d.lineName || '',    // E: 顧客名(LINE名)
      d.result || '',      // F: 結果
      d.product || '',     // G: 商品
      Number(d.amount) || 0,   // H: 売上金額
      Number(d.payment) || 0,  // I: 着金
      d.note || '',        // J: 属性
      d.recording || '',   // K: 録画
      '',                  // L: 着金日
      d.id || '',          // M: ID
      '', '', '', '', '', '', '', 0  // N〜U: 詳細情報・追加決済額
    ];
  });

  // 既存の最終行に一括追記
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);

  // タイムスタンプ降順でソート（新しい日付が上）
  const newLastRow = sheet.getLastRow();
  if (newLastRow > 2) {
    sheet.getRange(2, 1, newLastRow - 1, sheet.getLastColumn())
      .sort({ column: 1, ascending: false });
  }

  return { ok: true, imported: data.length };
}

function doGet(e) {
  try {
    const payload = (e && e.parameter && e.parameter.payload) ? JSON.parse(e.parameter.payload) : null;

    if (payload && payload.type === 'saleReport') {
      writeToSheet(payload);
      sendLineNotification(payload);
      return respond({ ok: true });
    }

    if (payload && payload.type === 'deleteSale') {
      const result = deleteFromSheet(payload);
      return respond(result);
    }

    if (payload && payload.type === 'updateResult') {
      const result = updateSheetRow(payload);
      return respond(result);
    }

    // payloadなし → 全行返す（同期用）
    return respond(getAllRows());

  } catch (err) {
    return respond({ ok: false, error: err.toString() });
  }
}

// =====================================================
// 売上報告タブに1行書き込む
// =====================================================
function writeToSheet(d) {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(SHEET_NAME + ' シートが見つかりません');

  const ci = d.customerInfo || {};
  const row = [
    new Date(),                                // タイムスタンプ
    d.date || '',                              // 日時
    d.memberName || '',                        // 担当者
    d.category || '',                          // 導線
    d.lineName || '',                          // 顧客名(LINE名)
    d.result || '',                            // 結果
    d.product || '',                           // 商品
    Number(d.amount) || 0,                     // 売上金額
    Number(d.payment) || 0,                    // 着金
    d.note || '',                              // メモ
    d.recording || '',                         // 録画
    d.paymentDate || '',                       // 着金日
    d.id || '',                                // ID
    ci.upsell || '',                           // アップセル見込み
    ci.rejection || '',                        // 断り理由
    ci.financial || '',                        // 資金状況
    ci.barrier || '',                          // 行動の障壁
    ci.approachMonths || '',                   // アプローチ目安
    ci.nextApproachDate || '',                 // 次回アプローチ予定日
    ci.customerMemo || ci.secondNote || '',    // 備考／所感
    Number(d.additionalAmount) || 0,           // 追加決済額
  ];
  // 末尾に追加してからタイムスタンプ降順でソート（最新が常に一番上）
  sheet.appendRow(row);
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .sort({ column: 1, ascending: false });
  }
}

// =====================================================
// LINEグループに通知送信
// =====================================================
function sendLineNotification(d) {
  if (!d.text) return;
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({
      to: LINE_GROUP_ID,
      messages: [{ type: 'text', text: d.text }]
    })
  });
}

// =====================================================
// 売上報告タブの全データを返す（同期用）
// =====================================================
function getAllRows() {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      obj[h] = (v instanceof Date) ? v.toISOString() : v;
    });
    return obj;
  });
}

// =====================================================
// JSON レスポンス生成
// =====================================================
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// IDで行を検索して削除
// =====================================================
function deleteFromSheet(payload) {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'error', message: 'データがありません' };

  const id = String(payload.id || '');
  const rows = sheet.getRange(2, 1, lastRow - 1, 20).getValues();

  // ① ID列（M列=インデックス12）で完全一致検索
  for (let i = 0; i < rows.length; i++) {
    const rowId = String(rows[i][12] || '').trim();
    if (rowId && rowId === id) {
      sheet.deleteRow(i + 2);
      return { status: 'ok', deleted: i + 2 };
    }
  }

  // ② IDで見つからない場合は 日時＋LINE名＋導線 で照合（sheet-row-N IDの場合）
  const targetDate = String(payload.date || '').slice(0, 10);
  const targetLine = String(payload.lineName || '').trim();
  const targetCat  = String(payload.category || '').trim();

  if (targetLine || targetCat) {
    for (let i = 0; i < rows.length; i++) {
      const rowDate = rows[i][1] instanceof Date
        ? Utilities.formatDate(rows[i][1], 'Asia/Tokyo', 'yyyy-MM-dd')
        : String(rows[i][1] || '').slice(0, 10);
      const rowLine = String(rows[i][4] || '').trim();
      const rowCat  = String(rows[i][3] || '').trim();
      if (rowDate === targetDate && rowLine === targetLine && rowCat === targetCat) {
        sheet.deleteRow(i + 2);
        return { status: 'ok', deleted: i + 2 };
      }
    }
  }

  return { status: 'error', message: '該当行が見つかりませんでした (ID: ' + id + ')' };
}

// =====================================================
// IDで行を検索して結果・商品・金額などを更新
// =====================================================
function updateSheetRow(payload) {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'error', message: 'データがありません' };

  const id = String(payload.id || '');
  const rows = sheet.getRange(2, 1, lastRow - 1, 20).getValues();
  let targetRow = -1;

  // ① ID列（M列=インデックス12）で完全一致検索
  for (let i = 0; i < rows.length; i++) {
    const rowId = String(rows[i][12] || '').trim();
    if (rowId && rowId === id) { targetRow = i + 2; break; }
  }

  // ② IDで見つからない場合は 日時＋LINE名＋導線 で照合
  if (targetRow === -1) {
    const targetDate = String(payload.date || '').slice(0, 10);
    const targetLine = String(payload.lineName || '').trim();
    const targetCat  = String(payload.category || '').trim();
    if (targetLine || targetCat) {
      for (let i = 0; i < rows.length; i++) {
        const rowDate = rows[i][1] instanceof Date
          ? Utilities.formatDate(rows[i][1], 'Asia/Tokyo', 'yyyy-MM-dd')
          : String(rows[i][1] || '').slice(0, 10);
        const rowLine = String(rows[i][4] || '').trim();
        const rowCat  = String(rows[i][3] || '').trim();
        if (rowDate === targetDate && rowLine === targetLine && rowCat === targetCat) {
          targetRow = i + 2; break;
        }
      }
    }
  }

  if (targetRow === -1) {
    return { status: 'error', message: '該当行が見つかりませんでした (ID: ' + id + ')' };
  }

  // 結果(F=6)・商品(G=7)・売上金額(H=8)・着金(I=9)・着金日(L=12) を更新
  sheet.getRange(targetRow, 6).setValue(payload.newResult || '');
  sheet.getRange(targetRow, 7).setValue(payload.newProduct || '');
  sheet.getRange(targetRow, 8).setValue(Number(payload.newAmount) || 0);
  sheet.getRange(targetRow, 9).setValue(Number(payload.newPayment) || 0);
  if (payload.newPaymentDate) {
    sheet.getRange(targetRow, 12).setValue(payload.newPaymentDate);
  }

  return { status: 'ok', updated: targetRow };
}

// =====================================================
// 既存データを一括ソート（エディタから手動実行用）
// タイムスタンプ降順（新しい順）に並び替える
// =====================================================
function sortSheet() {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .sort({ column: 1, ascending: false });
  }
  Logger.log('ソート完了：' + (lastRow - 1) + '行を並び替えました');
}
