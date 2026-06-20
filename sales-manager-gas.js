// =====================================================
// 売上管理ツール GASバックエンド（シンプル版）
// スプレッドシートにバインドして使用（Extensions → Apps Script）
// =====================================================

const SHEET_NAME = '売上報告';
const LINE_GROUP_ID = 'C87ef2d287f760158419805a3887ac7f9';

function doGet(e) {
  try {
    const payload = (e && e.parameter && e.parameter.payload) ? JSON.parse(e.parameter.payload) : null;

    if (payload && payload.type === 'saleReport') {
      writeToSheet(payload);
      sendLineNotification(payload);
      return respond({ ok: true });
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
