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

    // 【修正F】外部スクリプト（import_line_data.py）からのLINE一括インポート。
    // このエンドポイントはチーム利用パスワード(authHash)ではなく、専用トークン(importToken)で
    // 保護する（import_line_data.pyはチームのパスワードを知らない外部のスクリプトのため）。
    if (payload && payload.type === 'bulkImport') {
      if (!_requireImportToken(payload.importToken)) {
        return respond({ ok: false, error: _importTokenErrorReason(payload.importToken) });
      }
      const result = bulkImportToSheet(payload.rows);
      return respond(result);
    }
    // 【修正I】状態変更系・認証ハッシュを伴うエンドポイントはPOST(doPost)専用に統一した。
    // doGet側には同名の分岐を残さず、GET経由で来た場合は明示的にunauthorized(post_required)を
    // 返すようにしている（authHashがURL/履歴/GAS実行ログに残る経路を塞ぐため）。
    if (payload && payload.type === 'saleReport') {
      { const _ac = _checkAuthHash(payload.authHash); if (!_ac.ok) return respond(_ac); }
      writeToSheet(payload);
      sendLineNotification(payload);
      return respond({ ok: true });
    }

    if (payload && payload.type === 'deleteSale') {
      { const _ac = _checkAuthHash(payload.authHash); if (!_ac.ok) return respond(_ac); }
      const result = deleteFromSheet(payload);
      return respond(result);
    }

    if (payload && payload.type === 'updateResult') {
      { const _ac = _checkAuthHash(payload.authHash); if (!_ac.ok) return respond(_ac); }
      const result = updateSheetRow(payload);
      return respond(result);
    }

    if (payload && payload.type === 'updateSaleDetail') {
      { const _ac = _checkAuthHash(payload.authHash); if (!_ac.ok) return respond(_ac); }
      const result = updateSaleDetailRow(payload);
      return respond(result);
    }

    // 担当者名の一括変更（削除・修正時にスプシと同期）
    if (payload && payload.type === 'renameMember') {
      { const _ac = _checkAuthHash(payload.authHash); if (!_ac.ok) return respond(_ac); }
      const result = renameMemberInSheet(payload.oldName, payload.newName);
      return respond(result);
    }

    // リマインドテスト送信（手動テスト用。定期トリガー本体のsendDailyReminder()自体は無関係）
    if (payload && payload.type === 'testReminder') {
      { const _ac = _checkAuthHash(payload.authHash); if (!_ac.ok) return respond(_ac); }
      sendDailyReminder();
      return respond({ ok: true, message: 'テスト送信しました' });
    }

    // 週次ランキングテスト送信（手動テスト用。定期トリガー本体のsendWeeklyRanking()自体は無関係）
    if (payload && payload.type === 'testWeeklyRanking') {
      { const _ac = _checkAuthHash(payload.authHash); if (!_ac.ok) return respond(_ac); }
      sendWeeklyRanking();
      return respond({ ok: true, message: '週次ランキングをテスト送信しました' });
    }

    // 【修正H/I】起動時・手動同期の売上データ全件取得。gviz直読み廃止に伴い一本化した窓口。
    if (payload && payload.type === 'getAllRows') {
      { const _ac = _checkAuthHash(payload.authHash); if (!_ac.ok) return respond(_ac); }
      return respond(getAllRows());
    }

    // 共有データ保存（プロジェクト・動線・目標など）
    // 【修正4】authHash(利用パスワードハッシュ)・adminHash(管理者パスワードハッシュ)を
    // 受け取り、saveSharedData内で認証チェックを行う（未設定時は無視される＝ブートストラップ許可）
    if (payload && payload.type === 'saveSharedData') {
      const result = saveSharedData(payload.data, payload.authHash, payload.adminHash);
      return respond(result);
    }
    // 【修正D】ハッシュをURL(GET)に残さないよう、getSharedData/verifyAuth/getAllRecordingsも
    // POSTで受け付ける（レスポンスはdoGet経由と同一のためロジックは共通関数に委譲）
    if (payload && payload.type === 'getSharedData') {
      return respond(getSharedData(payload.authHash));
    }
    if (payload && payload.type === 'verifyAuth') {
      return respond(verifyAuth(payload.role, payload.hash));
    }
    if (payload && payload.type === 'getAllRecordings') {
      { const _ac = _checkAuthHash(payload.authHash); if (!_ac.ok) return respond(_ac); }
      return respond(getAllRecordings());
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

    // 【修正I】状態変更系・認証ハッシュを伴うエンドポイントはPOST(doPost)専用に統一したため、
    // doGet側の同名の重複分岐は廃止した。GET経由でこれらのtypeが指定された場合は
    // 明示的に拒否する（authHashがURL/ブラウザ履歴/GAS実行ログに残る経路を完全に塞ぐため）。
    const postOnlyTypes = [
      'saleReport', 'deleteSale', 'updateResult', 'updateSaleDetail', 'renameMember',
      'testReminder', 'testWeeklyRanking', 'getAllRows',
      'getSharedData', 'saveSharedData', 'verifyAuth', 'getAllRecordings'
    ];
    if (payload && postOnlyTypes.indexOf(payload.type) !== -1) {
      return respond({ ok: false, error: 'post_required' });
    }

    // 【修正L2】payloadなし（旧形式）のフォールバックは廃止した。
    // 以前はここで ?authHash=xxx というクエリパラメータ付きのGETだけで売上データ全件
    // (getAllRows)を返してしまい、authHashがURL・ブラウザ履歴・GAS実行ログに残る経路に
    // なっていた（クライアント側はこの経路を既に使用していないことを確認済み）。
    // 売上データをGETで返す経路を完全に無くすため、doGetは常にpost_requiredを返す。
    return respond({ ok: false, error: 'post_required' });

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
// 【修正4】2段階パスワード認証のGAS側チェックを追加。
// ・authConfigのハッシュ値はクライアントへ絶対に返さない（hasAuth真偽値のみ通知）
// ・authConfigが設定済み(hasAuth:true)の状態では、getSharedData/saveSharedDataとも
//   authHash（利用パスワードのハッシュ）が保存済みusePasswordHashと一致しないと拒否する
// ・authConfig自体を書き換えるにはadminHash（管理者パスワードのハッシュ）の一致が必須
// ・authConfigがまだ一度も設定されていない場合（ブートストラップ）は無認証で許可する
// =====================================================

// 保存済みの生のauthConfig（ハッシュを含む）を取得する内部専用ヘルパー
// ※この関数の戻り値を絶対にそのままクライアントへ返さないこと（ハッシュ漏えい防止）
function _getRawSharedData() {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  let sheet = ss.getSheetByName('ツール共有データ');
  if (!sheet) return { sheet: null, parsed: null };
  const val = sheet.getRange('A1').getValue();
  if (!val) return { sheet: sheet, parsed: null };
  try {
    return { sheet: sheet, parsed: JSON.parse(val) };
  } catch (e) {
    return { sheet: sheet, parsed: null };
  }
}

// クライアントに返してよい安全な形（ハッシュ値を除去した形）に変換する
function _sanitizeAuthConfig(auth) {
  if (!auth || !auth.usePasswordHash) return { hasAuth: false };
  return { hasAuth: true, updatedAt: auth.updatedAt || 0 };
}

// =====================================================
// 【修正M】authConfig（2段階パスワード認証設定）の保存先をA1(共有データJSON)から
// スクリプトプロパティ 'AUTH_CONFIG' へ切り離す。
// -----------------------------------------------------
// 背景: 旧デプロイ/旧バージョンのクライアントタブが開いたままだと、60秒ごとの定期同期
// (saveSharedData)がA1セルを丸ごと上書きしてしまう。旧コードはauthConfigという概念自体を
// 知らないため、A1上書きのたびにパスワード設定が消えてしまう実害が本番で確認された。
// ScriptPropertiesはスプレッドシートのセル値(A1)とは完全に独立した保存領域であり、
// 同じApps ScriptプロジェクトであればA1を上書きするどのデプロイ/バージョンの実行からも
// 影響を受けない。そのためauthConfigの真の保存場所をここに移す。
// 【厳守】既存の売上データ・SHARED_KEYS(A1)の保存/同期ロジックには一切手を入れない。
// authConfigの参照元だけをA1からScriptPropertiesへ切り替える。
// =====================================================

// authConfigをScriptPropertiesから読み込む（無ければnull）。
// 【マイグレーション】ScriptPropertiesに無く、かつA1のJSONに旧形式のauthConfigが
// 残っている場合だけ、初回に一度ScriptPropertiesへ自動移行する（安全のため実装。
// 通常の本番運用ではA1に残っていないはずなので基本はno-op）。
function _readAuthConfig() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('AUTH_CONFIG');
  if (raw) {
    try {
      return JSON.parse(raw);
    } catch (e) {
      return null;
    }
  }

  // ScriptPropertiesに無い場合のみ、A1の旧形式authConfigを一度だけ移行する
  try {
    const { parsed } = _getRawSharedData();
    const legacyAuth = parsed && parsed.authConfig;
    if (legacyAuth && legacyAuth.usePasswordHash) {
      _writeAuthConfig(legacyAuth);
      return legacyAuth;
    }
  } catch (e) {
    // マイグレーション自体の失敗は致命的ではないため無視してnullを返す
  }
  return null;
}

// authConfigをScriptPropertiesへ書き込む。objがnull/undefinedならプロパティ自体を削除する。
function _writeAuthConfig(obj) {
  const props = PropertiesService.getScriptProperties();
  if (obj) {
    props.setProperty('AUTH_CONFIG', JSON.stringify(obj));
  } else {
    props.deleteProperty('AUTH_CONFIG');
  }
}

// 【修正C→J→L1】role別・固定エポック時間窓によるサーバー側レート制限（総当たり対策）。
// GASではIP単位の制御が困難なため、CacheServiceを使ったグローバル(全体共有)な時間窓でロックする。
// 【修正J】以前はrole共通の1本のキー・スライディングTTL方式だったため、外部から
// 「TTLの間隔より少し短い周期」で失敗を送り続けるだけでロックを無限に延長でき、
// チーム全員を恒久的にログイン不能にできてしまう可用性DoSの懸念があった。
// これを解消するため、(1) role(use/admin)別にキーを分離し、(2) 固定エポック時間窓方式
// （現在時刻を窓の長さで割った「窓番号」をキーに含める）に変更した。窓番号が変わると
// カウンタは必ず新規キーになりリセットされるため、失敗を送り続けても恒久ロックはできない。
// 【修正L1】このカウンタはverifyAuth専用ではなく、_checkAuthHash（getAllRows/getSharedData/
// saveSharedData/saleReport等の全データ系エンドポイントの認証ガード）とも共有する共通ヘルパーに
// した。verifyAuthを経由しない直接呼び出しでオフラインPBKDF2総当たりされる迂回経路を塞ぐため。
const AUTH_RATE_LIMIT_MAX_FAILS   = 10;   // この時間窓内でこの回数失敗したらロックする
const AUTH_RATE_LIMIT_WINDOW_SEC  = 300;  // 固定時間窓の長さ: 5分

// role別・固定エポック時間窓のキャッシュキーを生成する
function _authRateLimitKey(role) {
  const windowIndex = Math.floor(Date.now() / (AUTH_RATE_LIMIT_WINDOW_SEC * 1000));
  return 'auth_fail_' + (role === 'admin' ? 'admin' : 'use') + '_' + windowIndex;
}

// 現在この時間窓でロック中かどうかを返す（実際のハッシュ照合を行う前に必ず呼ぶこと）
function _isAuthRateLimited(role) {
  const cache = CacheService.getScriptCache();
  const failCount = Number(cache.get(_authRateLimitKey(role)) || '0');
  return failCount >= AUTH_RATE_LIMIT_MAX_FAILS;
}

// 認証結果(ok)に応じて失敗カウンタを記録/リセットする共通ヘルパー。
// 【重要】成功時(ok=true)は必ずカウンタをリセットする。正規ユーザーの通常操作
// （多数の正しいauthHashでのgetAllRows/saveSharedData等）でロックされないようにするため。
function _recordAuthResult(role, ok) {
  const cache = CacheService.getScriptCache();
  const cacheKey = _authRateLimitKey(role);
  if (ok) {
    cache.remove(cacheKey); // 成功したらこの時間窓の失敗カウンタをリセット
    return;
  }
  const failCount = Number(cache.get(cacheKey) || '0');
  // 固定時間窓方式: TTLは常に「現在の窓が終わるまでの残り秒数」にする（窓を延長しない）。
  // これにより、失敗を送り続けてもロックが次の窓の境界を超えて延びることはない。
  const windowMs = AUTH_RATE_LIMIT_WINDOW_SEC * 1000;
  const remainSec = Math.max(1, Math.ceil((windowMs - (Date.now() % windowMs)) / 1000));
  cache.put(cacheKey, String(failCount + 1), remainSec);
}

// 利用/管理者パスワードのハッシュをサーバー側で照合する（ハッシュ本体は一切返さない）
// role: 'use' | 'admin'、hash: クライアントが計算した候補ハッシュ
function verifyAuth(role, hash) {
  try {
    if (_isAuthRateLimited(role)) {
      return { ok: false, error: 'locked' };
    }

    // 【修正M】authConfigはA1ではなくScriptProperties(_readAuthConfig)から取得する
    const auth = _readAuthConfig();
    if (!auth || !auth.usePasswordHash) return { ok: false, error: 'not_configured' };
    const target = role === 'admin' ? auth.adminPasswordHash : auth.usePasswordHash;
    const matched = !!(target && hash && target === hash);

    _recordAuthResult(role, matched);
    return { ok: matched };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// 【修正B→L1→M】共通認証ガード。authConfig未設定(ブートストラップ)時のみ免除。
// 設定済みの場合はauthHashが保存済みusePasswordHashと一致しないとng（error種別付き）を返す。
// 【修正L1】verifyAuthと共通のレート制限(role='use')を適用する。これにより、正規のログイン
// 画面(verifyAuth)を経由せず、候補ハッシュを直接この関数のガード対象エンドポイントへ
// 投げ続けるオフライン総当たりを防ぐ（ロック中は実際の比較すら行わずlockedを返す）。
// 【修正M】authConfigはA1ではなくScriptProperties(_readAuthConfig)から取得する。
// 【重要】定期トリガー由来の自動実行（sendDailyReminder/sendWeeklyRankingの本体関数）は
// HTTP経由ではなく直接呼び出されるためこのガードの対象外。doGet/doPostの各ハンドラ冒頭で
// HTTP経由の呼び出しにだけ適用すること（自動通知を壊さないため）。
function _checkAuthHash(authHash) {
  const auth = _readAuthConfig();
  const hasAuth = !!(auth && auth.usePasswordHash);
  if (!hasAuth) return { ok: true }; // ブートストラップ: 誰も設定していないので無認証で許可

  if (_isAuthRateLimited('use')) {
    return { ok: false, error: 'locked' };
  }
  const matched = !!(authHash && authHash === auth.usePasswordHash);
  _recordAuthResult('use', matched);
  return matched ? { ok: true } : { ok: false, error: 'unauthorized' };
}

// 真偽値だけでよい既存の呼び出し箇所向けの後方互換ラッパー
function _requireAuth(authHash) {
  return _checkAuthHash(authHash).ok;
}

// 【修正F】bulkImport専用のトークン認証。チーム利用パスワード(usePasswordHash)とは別に、
// スクリプトプロパティ IMPORT_TOKEN に保存した固定トークンとの一致だけを見る。
// import_line_data.py 等、チームのパスワードを知らない外部スクリプトからの一括投入を
// 想定しているため、authConfig(利用パスワード)の設定有無には一切依存しない。
// 【重要】IMPORT_TOKENが未設定の場合は安全側に倒し、常に拒否する（無防備な公開を防ぐ）。
function _requireImportToken(importToken) {
  const correctToken = PropertiesService.getScriptProperties().getProperty('IMPORT_TOKEN');
  if (!correctToken) return false; // 未設定 → 誤って無防備に開かないよう常に拒否
  return !!(importToken && importToken === correctToken);
}

// _requireImportToken が false だった理由を、クライアント（呼び出し元スクリプト）が
// 判別できるようエラー種別を返す（トークン未設定なのか、単に不一致なのかを区別する）
function _importTokenErrorReason(importToken) {
  const correctToken = PropertiesService.getScriptProperties().getProperty('IMPORT_TOKEN');
  if (!correctToken) return 'import_token_not_configured';
  return 'unauthorized';
}

function getSharedData(authHash) {
  try {
    const { parsed } = _getRawSharedData();
    if (!parsed) return { ok: true, data: null, hasAuth: false };

    // 【修正M】authConfigはA1ではなくScriptProperties(_readAuthConfig)から取得する
    const auth = _readAuthConfig();
    const hasAuth = !!(auth && auth.usePasswordHash);

    // ブートストラップ: まだ誰も認証設定をしていない → 無認証で許可
    if (!hasAuth) {
      // 【修正M】旧タブ等の影響でA1のJSONに古いauthConfigが紛れ込んでいる可能性があるため、
      // レスポンスからは必ず除去してから返す（ハッシュが漏れる/混乱する経路を断つため）
      const safeBoot = Object.assign({}, parsed);
      delete safeBoot.authConfig;
      return { ok: true, data: safeBoot, hasAuth: false };
    }

    // 【修正L1】verifyAuthを経由しない直接呼び出しでの総当たりを防ぐため、レート制限を適用する。
    // ロック中は実際の比較すら行わずlockedを返す。
    if (_isAuthRateLimited('use')) {
      return { ok: false, error: 'locked', hasAuth: true };
    }

    // 認証必須モード: authHashが保存済みusePasswordHashと一致しない限りデータを返さない
    const authOk = !!(authHash && authHash === auth.usePasswordHash);
    _recordAuthResult('use', authOk); // 成功時はリセット、失敗時のみカウントする
    if (!authOk) {
      return { ok: false, error: 'unauthorized', hasAuth: true };
    }

    // 一致 → データを返すが、authConfigはScriptProperties由来のサニタイズ済み情報に差し替える
    // （A1のJSONに万一authConfigが混ざっていても、ここで必ず上書き・除去される）
    const safe = Object.assign({}, parsed);
    safe.authConfig = _sanitizeAuthConfig(auth);
    return { ok: true, data: safe, hasAuth: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function saveSharedData(d, authHash, adminHash) {
  try {
    const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
    let sheet = ss.getSheetByName('ツール共有データ');
    if (!sheet) sheet = ss.insertSheet('ツール共有データ');

    // 【修正M】existingAuthはA1ではなくScriptProperties(_readAuthConfig)から取得する。
    // これにより、旧デプロイ/旧タブがA1を丸ごと上書きしても、認可判定の基準となる
    // 「本当のauthConfig」はScriptProperties側にあるため影響を受けない。
    const existingAuth = _readAuthConfig();
    const hasAuth = !!(existingAuth && existingAuth.usePasswordHash);

    if (hasAuth) {
      // 【修正A】トップレベルの書き込みゲート：authHashが利用パスワードハッシュと一致 "または"
      // adminHashが管理者パスワードハッシュと一致すれば許可する。
      // （利用パスワード変更の直後は、新しいauthHashがまだサーバー保存済みの値と一致しないため、
      //   このOR条件が無いと管理者操作であっても保存が永久に拒否されてしまう問題があった）
      // 【修正L1】use/adminそれぞれ独立にレート制限する。ロック中の役割は比較すら行わない。
      const authAttempted  = !!authHash;
      const adminAttempted = !!adminHash;
      const useLocked   = authAttempted  && _isAuthRateLimited('use');
      const adminLocked = adminAttempted && _isAuthRateLimited('admin');

      const authOk  = authAttempted  && !useLocked   && authHash  === existingAuth.usePasswordHash;
      const adminOk = adminAttempted && !adminLocked && !!existingAuth.adminPasswordHash && adminHash === existingAuth.adminPasswordHash;
      const overallOk = authOk || adminOk;

      // 記録は「試行された役割」ごとに行う。全体として成功した場合は、たとえ一方が
      // 不一致でも失敗として記録しない（利用パスワード変更直後の正規の管理者操作等で、
      // 一時的なauthHash不一致がuseロールの失敗として不必要に積み上がるのを防ぐため）。
      if (authAttempted  && !useLocked)   _recordAuthResult('use',   overallOk || authOk);
      if (adminAttempted && !adminLocked) _recordAuthResult('admin', overallOk || adminOk);

      if (!overallOk) {
        if (useLocked || adminLocked) return { ok: false, error: 'locked' };
        return { ok: false, error: 'unauthorized' };
      }
    }

    // 送られてきたデータに含まれるauthConfigをどう扱うか判定する
    // 【重要】既存の売上データ等(SHARED_KEYS由来のフィールド)の保存は、authConfigの
    // 可否判定に関わらず常に継続する。authConfigの更新可否だけを個別に判定すること。
    const incomingAuth = d && d.authConfig;
    let finalAuth = existingAuth || null;

    if (incomingAuth) {
      if (!hasAuth) {
        // ブートストラップ: まだ誰も設定していない → 無条件で初回登録を許可
        finalAuth = incomingAuth;
      } else if (JSON.stringify(incomingAuth) === JSON.stringify(existingAuth)) {
        // 内容が変わっていない通常の定期pushはadminHash不要でそのまま維持
        finalAuth = existingAuth;
      } else if (adminHash && existingAuth.adminPasswordHash && adminHash === existingAuth.adminPasswordHash) {
        // 管理者パスワードのハッシュが一致した場合のみ、authConfigの更新を許可する
        finalAuth = incomingAuth;
      } else {
        // 管理者確認が取れない場合はauthConfigの更新だけを拒否し、既存値を維持する
        // （他の通常データの保存は継続させ、売上データ等が巻き込まれて失われないようにする）
        finalAuth = existingAuth;
      }
    }

    // 【修正M】authConfigが実際に変わった場合だけScriptPropertiesへ永続化する
    // （ブートストラップでの初回登録、または管理者確認済みの変更のときのみ変化する）
    if (JSON.stringify(finalAuth) !== JSON.stringify(existingAuth)) {
      _writeAuthConfig(finalAuth);
    }

    // 【修正M・最重要】A1(共有データJSON)にはauthConfigを一切含めない。
    // 旧デプロイ/旧タブがこのA1セルを丸ごと上書きしても、authConfigの実体は
    // ScriptProperties側にあるため無傷のまま保たれる。
    const toSave = Object.assign({}, d);
    delete toSave.authConfig;
    sheet.getRange('A1').setValue(JSON.stringify(toSave));
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
// 週次ランキング通知（毎週木曜8時 タイムトリガーで自動実行）
// =====================================================
function sendWeeklyRanking() {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName('売上報告');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 直近7日間（木曜8時トリガー時 = 先週木曜〜今週水曜）
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

  // OpenAI DALL-E 3 で画像生成を試みる
  const openaiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  let imageUrl = null;

  if (openaiKey && entries.length > 0) {
    const medals = ['🥇', '🥈', '🥉'];
    const rankText = entries.map((e, i) =>
      `${medals[i] || `${i+1}位`} ${e.name}：${fmt(e.chakkin)}（成約${e.keiyaku}件・成約率${e.rate}%）`
    ).join('\n');

    const prompt = `日本語の営業チーム週間売上ランキングカード画像を作成。

タイトル：「週間売上ランキング ${startStr}〜${endStr}」

ランキング：
${rankText}

デザイン：濃紺背景にゴールドアクセント、白テキスト、各メンバーの順位・名前・着金額・成約件数を明確に表示、プロフェッショナルな企業スタイル、1080×1080ピクセル正方形。`;

    try {
      const openaiRes = UrlFetchApp.fetch('https://api.openai.com/v1/images/generations', {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + openaiKey, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ model: 'dall-e-3', prompt: prompt, n: 1, size: '1024x1024', quality: 'hd', response_format: 'url' }),
        muteHttpExceptions: true
      });
      const openaiData = JSON.parse(openaiRes.getContentText());
      if (openaiData.data && openaiData.data[0] && openaiData.data[0].url) {
        imageUrl = openaiData.data[0].url;
      }
    } catch (e) {
      // 画像生成失敗 → テキストにフォールバック
    }
  }

  if (imageUrl) {
    // 画像をLINEに送信
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
      payload: JSON.stringify({
        to: LINE_GROUP_ID,
        messages: [
          { type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl },
          { type: 'text', text: `🔗 詳細はツールで確認\nhttps://82350kyo.github.io/sales-manager-tool/sales-manager.html` }
        ]
      })
    });
  } else {
    // テキストで送信（フォールバック）
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
      payload: JSON.stringify({ to: LINE_GROUP_ID, messages: [{ type: 'text', text: msg }] })
    });
  }
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
  ScriptApp.newTrigger('sendWeeklyRanking')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(8)
    .create();
}

// デバッグ用：スプレッドシートの日付データを確認する
function debugRanking() {
  const ss = SpreadsheetApp.openById('1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw');
  const sheet = ss.getSheetByName('売上報告');
  const lastRow = sheet.getLastRow();
  Logger.log('lastRow: ' + lastRow);

  const today = new Date();
  const lastMon = new Date(today); lastMon.setDate(today.getDate() - 7); lastMon.setHours(0,0,0,0);
  const lastSun = new Date(today); lastSun.setDate(today.getDate() - 1); lastSun.setHours(23,59,59,999);
  Logger.log('today: ' + today);
  Logger.log('range: ' + lastMon + ' ~ ' + lastSun);

  const rows = sheet.getRange(2, 1, Math.min(10, lastRow - 1), 5).getValues();
  rows.forEach((row, i) => {
    const rawDate = row[1];
    const d = rawDate instanceof Date ? rawDate : new Date(rawDate);
    Logger.log('row' + i + ' B=' + rawDate + ' parsed=' + d + ' inRange=' + (!isNaN(d) && d >= lastMon && d <= lastSun) + ' C=' + row[2]);
  });

  const openaiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  Logger.log('openaiKey set: ' + !!openaiKey);
}

// =====================================================
// 週次ランキング通知（画像版・OpenAI DALL-E 3使用）
// sendWeeklyRanking の上位互換。毎週木曜8時自動実行。
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

// =====================================================
// 【管理者ユーティリティ】authConfig（2段階パスワード認証設定）だけをリセットする
// -----------------------------------------------------
// 使い方: Apps Scriptエディタの関数選択プルダウンで resetAuthConfig を選び、
// 「実行(Run)」ボタンを押すだけでよい。doGet/doPostからは一切呼び出さず、外部にも
// 公開していない（Webアプリ経由では実行できない）ため、再デプロイは不要。
//
// 【重要・破壊防止】この関数は authConfig（ScriptPropertiesの'AUTH_CONFIG'、および
// 念のためA1に紛れ込んでいる場合はそのauthConfigキー）だけを削除する。
// 売上データ・プロジェクト・動線・商品・目標・メンバー等、他の共有データキーには
// 一切触れない。A1のJSONパースに失敗した場合はA1側の書き換えだけをスキップする。
//
// 【実行後の注意】他のメンバーがツールを開いたままだと、そのタブの定期同期(push)で
// 古いauthConfigがまた書き戻されてしまう可能性がある。実行前には他のメンバーに
// ツールのタブを閉じてもらい、実行後はできるだけ早く自分自身が「初回セットアップ」
// 画面からパスワードを再設定すること。
//
// 【修正O】authConfigの正式な保存場所は、その後のバージョンでA1(共有データJSON)から
// スクリプトプロパティ 'AUTH_CONFIG'（PropertiesService.getScriptProperties()）へ
// 移行済み。以前のこの関数はA1のauthConfigキーしか消していなかったため実質効かなく
// なっていた問題を修正し、ScriptProperties側の'AUTH_CONFIG'を確実に削除するように
// した（A1側の掃除も安全のため引き続き行う）。
// =====================================================
function resetAuthConfig() {
  const SPREADSHEET_ID = '1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw';
  const SHEET_NAME_SHARED = 'ツール共有データ';

  try {
    // ---- (1) 本来の保存場所であるScriptPropertiesの'AUTH_CONFIG'を削除する ----
    const props = PropertiesService.getScriptProperties();
    const hadScriptProp = props.getProperty('AUTH_CONFIG') !== null;
    if (hadScriptProp) {
      props.deleteProperty('AUTH_CONFIG');
    }

    // ---- (2) 念のため、A1(共有データJSON)に旧形式のauthConfigが残っていれば掃除する ----
    let hadA1AuthConfig = false;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_SHARED);
    if (!sheet) {
      Logger.log('シート「' + SHEET_NAME_SHARED + '」が見つからないため、A1側の確認はスキップしました。');
    } else {
      const raw = sheet.getRange('A1').getValue();
      if (!raw) {
        Logger.log('A1セルは空です（A1側にauthConfigは残っていません）。');
      } else {
        let parsed = null;
        try {
          parsed = JSON.parse(raw);
        } catch (e) {
          // JSONパースに失敗した場合はA1側の書き換えだけをスキップする（破壊防止）
          Logger.log('エラー: A1セルのJSONパースに失敗したため、A1側は書き換えずスキップしました。詳細: ' + e.toString());
        }
        if (parsed) {
          hadA1AuthConfig = Object.prototype.hasOwnProperty.call(parsed, 'authConfig');
          if (hadA1AuthConfig) {
            // authConfigキーだけを削除する（他のキー＝共有データは一切変更しない）
            delete parsed.authConfig;
            sheet.getRange('A1').setValue(JSON.stringify(parsed));
          } else {
            Logger.log('A1側にauthConfigキーはありませんでした（既にクリーンです）。');
          }
        }
      }
    }

    // ---- (3) 実行結果をログ出力 ----
    if (!hadScriptProp && !hadA1AuthConfig) {
      Logger.log('authConfigはScriptProperties・A1のどちらにも存在しませんでした（既に未設定です）。');
      return;
    }
    Logger.log('✅ authConfig(ScriptProperties)を削除しました。');
    Logger.log('削除前にScriptPropertiesの AUTH_CONFIG は存在していました: ' + hadScriptProp);
    Logger.log('削除前にA1側にもauthConfigキーが存在していました: ' + hadA1AuthConfig);
    Logger.log('次にツールを開くと「初回セットアップ」画面が表示されます。');
  } catch (e) {
    Logger.log('予期しないエラーが発生しました。詳細: ' + e.toString());
  }
}
