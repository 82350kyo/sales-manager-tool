// =====================================================
// 売上管理ツール GASバックエンド（シンプル版）
// スプレッドシートにバインドして使用（Extensions → Apps Script）
// =====================================================

const SHEET_NAME = '売上報告';
const LINE_GROUP_ID = 'C87ef2d287f760158419805a3887ac7f9';

// =====================================================
// 毎朝リマインド通知（タイムトリガーで自動実行）
// =====================================================
function sendDailyReminder() {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const rows = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

  const targets = [];
  for (const row of rows) {
    const raw = row[18]; // S列: 次回アプローチ予定日
    if (!raw) continue;
    const nextDate = raw instanceof Date
      ? Utilities.formatDate(raw, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(raw).slice(0, 10);
    if (nextDate !== today) continue;

    // 初回面談日（B列）を M/d 形式に
    const rawDate = row[1];
    let dateStr = '';
    if (rawDate instanceof Date) {
      dateStr = Utilities.formatDate(rawDate, 'Asia/Tokyo', 'M/d');
    } else {
      const parts = String(rawDate).slice(0, 10).split('-');
      if (parts.length === 3) dateStr = `${parseInt(parts[1])}/${parseInt(parts[2])}`;
    }

    const recording = String(row[10] || '').trim();

    targets.push({
      lineName:  String(row[4] || '').trim(),   // E列: 顧客LINE名
      member:    String(row[2] || '').trim(),   // C列: 担当者
      dateStr:   dateStr,                        // B列: 初回面談日
      category:  String(row[3] || '').trim(),   // D列: 導線
      recording: (recording && recording !== 'なし' && recording !== '-') ? recording : '',
    });
  }

  if (targets.length === 0) return;

  const todayLabel = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d');
  let msg = `📅 本日のアプローチリスト（${todayLabel}）\n`;
  msg += '━━━━━━━━━━━━━━\n';

  for (const t of targets) {
    msg += `担当：${t.member}\n`;
    msg += `👤 ${t.lineName}\n`;
    msg += `初回面談：${t.dateStr}\n`;
    msg += `導線：${t.category}\n`;
    if (t.recording) msg += `録画：${t.recording}\n`;
    msg += '━━━━━━━━━━━━━━\n';
  }

  msg += `\n計${targets.length}件`;
  sendLineMessage(msg);
}

// LINE グループにメッセージ送信（汎用）
function sendLineMessage(text) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({
      to: LINE_GROUP_ID,
      messages: [{ type: 'text', text: text }]
    })
  });
}

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
    // 共有データ保存（プロジェクト・動線・目標など）
    if (payload && payload.type === 'saveSharedData') {
      const result = saveSharedData(payload.data);
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

    if (payload && payload.type === 'updateSaleDetail') {
      const result = updateSaleDetailRow(payload);
      return respond(result);
    }

    // リマインドテスト送信
    if (payload && payload.type === 'testReminder') {
      sendDailyReminder();
      return respond({ ok: true, message: 'テスト送信しました' });
    }

    // 担当者名の一括変更（削除・修正時にスプシと同期）
    if (payload && payload.type === 'renameMember') {
      const result = renameMemberInSheet(payload.oldName, payload.newName);
      return respond(result);
    }

    // 共有データ取得（プロジェクト・動線・目標など）
    if (payload && payload.type === 'getSharedData') {
      return respond(getSharedData());
    }

    // 共有データ保存
    if (payload && payload.type === 'saveSharedData') {
      return respond(saveSharedData(payload.data));
    }

    // 録画URL一括取得
    if (payload && payload.type === 'getAllRecordings') {
      return respond(getAllRecordings());
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
// IDで行を検索して全フィールドを更新（詳細編集モーダル用）
// =====================================================
function updateSaleDetailRow(payload) {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'error', message: 'データがありません' };

  const id = String(payload.id || '');
  const rows = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
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

  // 各列を更新（列番号は1始まり）
  sheet.getRange(targetRow, 2).setValue(payload.date || '');          // B列: 面談日
  sheet.getRange(targetRow, 3).setValue(payload.memberName || '');    // C列: 担当者
  sheet.getRange(targetRow, 4).setValue(payload.category || '');      // D列: 導線
  sheet.getRange(targetRow, 5).setValue(payload.lineName || '');      // E列: 顧客LINE名
  sheet.getRange(targetRow, 6).setValue(payload.result || '');        // F列: 結果
  sheet.getRange(targetRow, 7).setValue(payload.product || '');       // G列: 商品
  sheet.getRange(targetRow, 8).setValue(Number(payload.amount) || 0); // H列: 売上金額
  sheet.getRange(targetRow, 9).setValue(Number(payload.payment) || 0);// I列: 着金
  // J列(10): 属性は変更しない
  sheet.getRange(targetRow, 11).setValue(payload.recording || '');    // K列: 録画URL
  if (payload.paymentDate) {
    sheet.getRange(targetRow, 12).setValue(payload.paymentDate);      // L列: 着金日
  }
  // M列(13): IDは変更しない
  sheet.getRange(targetRow, 14).setValue(payload.upsell || '');          // N列: アップセル見込み
  sheet.getRange(targetRow, 15).setValue(payload.rejection || '');        // O列: 断り理由
  sheet.getRange(targetRow, 16).setValue(payload.financial || '');        // P列: 資金状況
  sheet.getRange(targetRow, 17).setValue(payload.barrier || '');          // Q列: 行動の障壁
  sheet.getRange(targetRow, 18).setValue(payload.approachMonths || '');   // R列: アプローチ目安
  sheet.getRange(targetRow, 19).setValue(payload.nextApproachDate || ''); // S列: 次回アプローチ予定日
  sheet.getRange(targetRow, 20).setValue(payload.customerMemo || '');     // T列: 備考/所感

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

// =====================================================
// 担当者名の一括変更（削除・修正時にスプシと同期）
// =====================================================
function renameMemberInSheet(oldName, newName) {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, updated: 0 };

  const col = 3; // C列: 担当者
  let updated = 0;
  for (let i = 2; i <= lastRow; i++) {
    const cell = sheet.getRange(i, col);
    const val = String(cell.getValue() || '').replace(/[（(][^）)]*[）)]/g, '').trim();
    if (val === oldName.trim()) {
      cell.setValue(newName);
      updated++;
    }
  }
  return { ok: true, updated: updated };
}

// =====================================================
// 共有データの読み書き（プロジェクト・動線・目標など）
// シート「ツール共有データ」のA1にJSON保存
// =====================================================
function getSharedData() {
  try {
    const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
    let sheet = ss.getSheetByName('ツール共有データ');
    if (!sheet) return { ok: true, data: null };
    const val = sheet.getRange('A1').getValue();
    if (!val) return { ok: true, data: null };
    return { ok: true, data: JSON.parse(val) };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function saveSharedData(d) {
  try {
    const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
    let sheet = ss.getSheetByName('ツール共有データ');
    if (!sheet) sheet = ss.insertSheet('ツール共有データ');
    sheet.getRange('A1').setValue(JSON.stringify(d));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// =====================================================
// 録画URL一括取得（スプレッドシートのID→録画URLマップ）
// =====================================================
function getAllRecordings() {
  try {
    const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, data: [] };
    const rows = sheet.getRange(2, 1, lastRow - 1, 13).getValues(); // A〜M列
    const result = [];
    for (const row of rows) {
      const id        = String(row[12] || '').trim(); // M列: ID
      const recording = String(row[10] || '').trim(); // K列: 録画URL
      if (id && recording && recording !== 'なし' && recording !== '-') {
        result.push({ id, recording });
      }
    }
    return { ok: true, data: result };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// =====================================================
// 週次ランキング通知（毎週月曜8時 タイムトリガーで自動実行）
// =====================================================
function sendWeeklyRanking() {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName('売上報告');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 直近7日間（月曜8時トリガー時 = 前週月曜〜昨日日曜）
  const today = new Date();
  const lastMon = new Date(today); lastMon.setDate(today.getDate() - 7); lastMon.setHours(0,0,0,0);
  const lastSun = new Date(today); lastSun.setDate(today.getDate() - 1); lastSun.setHours(23,59,59,999);
  const startStr = Utilities.formatDate(lastMon, 'Asia/Tokyo', 'M/d');
  const endStr   = Utilities.formatDate(lastSun, 'Asia/Tokyo', 'M/d');

  const rows = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

  // 業種判定（D列の導線文字列からAI・物販を識別）
  function getBusType(cat) {
    const c = String(cat || '');
    if (c.includes('(AI)') || c.includes('（AI）') ||
        ['チョーさん','イーサン','浩志さん','的場'].some(kw => c.includes(kw))) return 'AI';
    if (c.includes('(物販)') || c.includes('（物販）') || c.includes('物販')) return '物販';
    return '';
  }

  // 先週の行を抽出し担当者別に集計
  const stats = {}; // { name: { busyo: {...}, ai: {...}, all: {...} } }

  rows.forEach(row => {
    const rawDate = row[1]; // B列 面談日
    if (!rawDate) return;
    const d = rawDate instanceof Date ? rawDate : new Date(rawDate);
    if (isNaN(d) || d < lastMon || d > lastSun) return;

    const name    = String(row[2] || '').replace(/[（(][^）)]*[）)]/g, '').trim(); // C列 担当者
    const cat     = String(row[3] || ''); // D列 導線
    const result  = String(row[5] || ''); // F列 結果
    const amount  = Number(String(row[7] || '0').replace(/[^0-9]/g, '')) || 0; // H列 売上金額
    const payment = Number(String(row[8] || '0').replace(/[^0-9]/g, '')) || 0; // I列 着金
    const bus     = getBusType(cat);
    if (!name) return;

    const keys = ['all'];
    if (bus) keys.push(bus);
    keys.forEach(key => {
      if (!stats[name]) stats[name] = {};
      if (!stats[name][key]) stats[name][key] = { apo: 0, cancel: 0, keiyaku: 0, seiyakuAmt: 0, chakkin: 0 };
      const s = stats[name][key];
      s.apo++;
      if (result === 'キャンセル') s.cancel++;
      if (result === '成約' || result.startsWith('成約')) {
        s.keiyaku++;
        s.seiyakuAmt += amount;
        s.chakkin += payment;
      }
    });
  });

  // 総合ランキング生成（着金額順、全員）
  const entries = Object.entries(stats)
    .filter(([, v]) => v['all'])
    .map(([name, v]) => {
      const s = v['all'];
      const mendan = s.apo - s.cancel;
      const rate = mendan > 0 ? (s.keiyaku / mendan * 100).toFixed(1) : '-';
      const tanka = s.keiyaku > 0 ? Math.round(s.seiyakuAmt / s.keiyaku) : null;
      return { name, chakkin: s.chakkin, mendan, keiyaku: s.keiyaku, rate, tanka };
    })
    .sort((a, b) => b.chakkin - a.chakkin);

  const fmt = n => n >= 10000 ? `¥${(n/10000).toFixed(n%10000===0?0:1)}万` : `¥${n.toLocaleString()}`;
  const rankLines = entries.length === 0 ? '（データなし）' : entries.map((e, i) => {
    return [
      `${i+1}位 ${e.name}`,
      `・着金額：${fmt(e.chakkin)}`,
      `・面談実施数：${e.mendan}件`,
      `・成約数：${e.keiyaku}件`,
      `・成約率：${e.rate}%`,
      `・平均成約単価：${e.tanka !== null ? fmt(e.tanka) : '-'}`,
    ].join('\n');
  }).join('\n\n');

  const msg = [
    `📊 先週の営業成績ランキング（${startStr}〜${endStr}）`,
    `🏆 週間ランキング 物販＋AI`,
    `━━━━━━━━━━━━━━━━`,
    ``,
    rankLines,
    ``,
    `🔗 詳細はツールで確認`,
    `https://82350kyo.github.io/sales-manager-tool/sales-manager.html`,
  ].join('\n');

  const LINE_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify({ to: LINE_GROUP_ID, messages: [{ type: 'text', text: msg }] })
  });
}

// =====================================================
// 週次ランキングトリガーのセットアップ（初回のみ手動実行）
// =====================================================
function setupWeeklyRankingTrigger() {
  // 既存の同名トリガーを削除してから再作成（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendWeeklyRanking') ScriptApp.deleteTrigger(t);
    if (t.getHandlerFunction() === 'sendWeeklyRankingImage') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendWeeklyRankingImage')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
}

// =====================================================
// 週次ランキング通知（画像版・OpenAI DALL-E 3使用）
// sendWeeklyRanking の上位互換。毎週月曜8時自動実行。
// 事前準備: スクリプトプロパティに OPENAI_API_KEY を設定すること
// =====================================================
function sendWeeklyRankingImage() {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const today = new Date();
  const lastMon = new Date(today); lastMon.setDate(today.getDate() - 7); lastMon.setHours(0,0,0,0);
  const lastSun = new Date(today); lastSun.setDate(today.getDate() - 1); lastSun.setHours(23,59,59,999);
  const startStr = Utilities.formatDate(lastMon, 'Asia/Tokyo', 'M/d');
  const endStr   = Utilities.formatDate(lastSun, 'Asia/Tokyo', 'M/d');

  const rows = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

  function getBusType(cat) {
    const c = String(cat || '');
    if (c.includes('(AI)') || c.includes('（AI）') ||
        ['チョーさん','イーサン','浩志さん','的場'].some(kw => c.includes(kw))) return 'AI';
    if (c.includes('(物販)') || c.includes('（物販）') || c.includes('物販')) return '物販';
    return '';
  }

  const stats = {};
  rows.forEach(row => {
    const rawDate = row[1];
    if (!rawDate) return;
    const d = rawDate instanceof Date ? rawDate : new Date(rawDate);
    if (isNaN(d) || d < lastMon || d > lastSun) return;
    const name    = String(row[2] || '').replace(/[（(][^）)]*[）)]/g, '').trim();
    const cat     = String(row[3] || '');
    const result  = String(row[5] || '');
    const amount  = Number(String(row[7] || '0').replace(/[^0-9]/g, '')) || 0;
    const payment = Number(String(row[8] || '0').replace(/[^0-9]/g, '')) || 0;
    const bus     = getBusType(cat);
    if (!name) return;
    const keys = ['all'];
    if (bus) keys.push(bus);
    keys.forEach(key => {
      if (!stats[name]) stats[name] = {};
      if (!stats[name][key]) stats[name][key] = { apo: 0, cancel: 0, keiyaku: 0, seiyakuAmt: 0, chakkin: 0 };
      const s = stats[name][key];
      s.apo++;
      if (result === 'キャンセル') s.cancel++;
      if (result === '成約' || result.startsWith('成約')) {
        s.keiyaku++;
        s.seiyakuAmt += amount;
        s.chakkin += payment;
      }
    });
  });

  const entries = Object.entries(stats)
    .filter(([, v]) => v['all'])
    .map(([name, v]) => {
      const s = v['all'];
      const mendan = s.apo - s.cancel;
      const rate = mendan > 0 ? (s.keiyaku / mendan * 100).toFixed(1) : '-';
      const tanka = s.keiyaku > 0 ? Math.round(s.seiyakuAmt / s.keiyaku) : null;
      return { name, chakkin: s.chakkin, mendan, keiyaku: s.keiyaku, rate, tanka };
    })
    .sort((a, b) => b.chakkin - a.chakkin);

  if (entries.length === 0) {
    sendWeeklyRanking();
    return;
  }

  const fmt = n => n >= 10000 ? `${(n/10000).toFixed(n%10000===0?0:1)}万円` : `${n.toLocaleString()}円`;
  const medals = ['🥇', '🥈', '🥉'];
  const rankText = entries.map((e, i) =>
    `${medals[i] || `${i+1}位`} ${e.name}：${fmt(e.chakkin)}（成約${e.keiyaku}件・成約率${e.rate}%）`
  ).join('\n');

  const prompt = `日本語の営業チーム週間売上ランキングカード画像を作成。

タイトル：「週間売上ランキング ${startStr}〜${endStr}」

ランキング：
${rankText}

デザイン：濃紺背景にゴールドアクセント、白テキスト、1位〜3位にトロフィーアイコン、各メンバーの順位・名前・着金額・成約件数を大きく表示、プロフェッショナルな企業スタイル、1080×1080ピクセル正方形。`;

  const openaiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  const openaiRes = UrlFetchApp.fetch('https://api.openai.com/v1/images/generations', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + openaiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      model: 'dall-e-3',
      prompt: prompt,
      n: 1,
      size: '1024x1024',
      quality: 'hd',
      response_format: 'url'
    }),
    muteHttpExceptions: true
  });

  const openaiData = JSON.parse(openaiRes.getContentText());
  if (!openaiData.data || !openaiData.data[0] || !openaiData.data[0].url) {
    sendWeeklyRanking(); // 画像生成失敗時はテキスト版にフォールバック
    return;
  }

  const imageUrl = openaiData.data[0].url;
  const lineToken = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + lineToken
    },
    payload: JSON.stringify({
      to: LINE_GROUP_ID,
      messages: [
        {
          type: 'image',
          originalContentUrl: imageUrl,
          previewImageUrl: imageUrl
        },
        {
          type: 'text',
          text: `🔗 詳細はツールで確認\nhttps://82350kyo.github.io/sales-manager-tool/sales-manager.html`
        }
      ]
    }),
    muteHttpExceptions: true
  });
}
