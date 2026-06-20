// ============================================
// 売上管理 GAS バックエンド（完全版）
// スプレッドシートID
const SS_ID = '1BMptkze_WyYL6TRG5Jzugy8aYT-AL4F4O7gnmFPKXRw';

// LINE設定（既存の営業グループ通知用）
// セキュリティ対応：トークンはスクリプトプロパティから取得
const LINE_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
// ============================================


// ============================================
// HTTPリクエスト受信（GET）
// ============================================
function doGet(e) {
  try {
    const payload = e && e.parameter && e.parameter.payload;
    if (payload) {
      const data = JSON.parse(payload);
      const result = handleAction(data);
      return buildResponse(result);
    }
    // payloadなしの場合は全行を返す
    const rows = getAllRows();
    return buildResponse(rows);
  } catch (err) {
    return buildResponse({ error: err.toString() });
  }
}

// JSON形式でレスポンスを返す
function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// シート取得（売上報告タブを使用）
function getSheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  return ss.getSheetByName('売上報告');
}


// ============================================
// 全行データを取得してオブジェクト配列で返す
// ============================================
function getAllRows() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => {
      const val = data[i][j];
      if (val instanceof Date) {
        row[h] = val.toISOString();
      } else {
        row[h] = val !== undefined && val !== null ? val : '';
      }
    });
    rows.push(row);
  }
  return rows;
}


// ============================================
// アクションの振り分け
// ============================================
function handleAction(data) {
  if (!data || !data.type) return { error: 'typeが指定されていません' };
  if (data.type === 'updateResult') return updateRow(data);
  if (data.type === 'addSale' || data.type === 'saleReport') return addRow(data);
  return { error: '不明なtype: ' + data.type };
}


// ============================================
// 既存行の更新
// ============================================
function updateRow(data) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idCol     = headers.indexOf('ID');
  const nameCol   = headers.indexOf('顧客名(LINE名)');
  const dateCol   = headers.indexOf('日時');
  const resultCol = headers.indexOf('結果');
  const amtCol    = headers.indexOf('売上金額');
  const payCol    = headers.indexOf('着金');
  const payDCol   = headers.indexOf('着金日');
  const prodCol   = headers.indexOf('商品');

  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    const rowId = String(values[i][idCol] || '').trim();
    // IDで一致検索
    if (data.id && rowId === String(data.id).trim()) {
      targetRow = i;
      break;
    }
    // IDがない場合は顧客名＋日付で検索
    if (!data.id || rowId === '') {
      const rowName = String(values[i][nameCol] || '').trim();
      const rowDate = values[i][dateCol];
      const rowDateStr = rowDate instanceof Date
        ? rowDate.toISOString().slice(0, 10)
        : String(rowDate || '').slice(0, 10);
      const dataDateStr = String(data.date || '').slice(0, 10);
      if (rowName === String(data.lineName || '').trim() && rowDateStr === dataDateStr) {
        targetRow = i;
        break;
      }
    }
  }

  if (targetRow === -1) return { success: false, message: '対象行が見つかりませんでした' };

  const rowNum = targetRow + 1;
  if (resultCol >= 0) sheet.getRange(rowNum, resultCol + 1).setValue(data.newResult || '');
  if (amtCol  >= 0 && data.newAmount      !== undefined) sheet.getRange(rowNum, amtCol  + 1).setValue(data.newAmount);
  if (payCol  >= 0 && data.newPayment     !== undefined) sheet.getRange(rowNum, payCol  + 1).setValue(data.newPayment);
  if (payDCol >= 0 && data.newPaymentDate !== undefined) sheet.getRange(rowNum, payDCol + 1).setValue(data.newPaymentDate);
  if (prodCol >= 0 && data.newProduct     !== undefined) sheet.getRange(rowNum, prodCol + 1).setValue(data.newProduct);

  return { success: true, message: '更新しました（行' + rowNum + '）' };
}


// ============================================
// 新規行の追加（customerInfo 対応版）
// ============================================
function addRow(data) {
  const sheet = getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const newRow = new Array(headers.length).fill('');

  // ヘッダー名をキーにして値をセットするユーティリティ
  const set = (key, val) => {
    const idx = headers.indexOf(key);
    if (idx >= 0) newRow[idx] = val !== undefined && val !== null ? val : '';
  };

  // 基本情報
  set('タイムスタンプ',   new Date().toISOString());
  set('担当者',          data.memberName  || '');
  set('日時',            data.date        || '');
  set('売上金額',        data.amount      || 0);
  set('導線',            data.category    || '');
  set('顧客名(LINE名)',  data.lineName    || '');
  set('結果',            data.result      || '');
  set('着金',            data.payment     || 0);
  set('着金日',          data.paymentDate || '');
  set('商品',            data.product     || '');
  set('ID',              data.id          || '');

  // 顧客情報（customerInfo オブジェクトが存在する場合のみ書き込み）
  // ヘッダーに列が存在しない場合は set() 内で自動スキップされる
  const ci = data.customerInfo;
  if (ci) {
    set('アップセル見込み',       ci.upsell          || '');
    set('断り理由',               ci.rejection       || '');
    set('資金状況',               ci.financial       || '');
    set('行動の障壁',             ci.barrier         || '');
    set('アプローチ目安',         ci.approachMonths  || '');
    set('次回アプローチ予定日',   ci.nextApproachDate || '');
    // 備考／所感：成約時は customerMemo、未成約時は secondNote を使用
    set('備考／所感',             ci.customerMemo || ci.secondNote || '');
  }

  sheet.appendRow(newRow);
  SpreadsheetApp.flush();

  // saleReport の場合はLINEグループへ通知
  if (data.type === 'saleReport' && data.text) {
    sendLineMessage(data.text);
  }

  return { success: true, message: '追加しました' };
}


// ============================================
// HTTPリクエスト受信（POST）
// ============================================
function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);
    const event = json.events[0];
    if (event) {
      const groupId = event.source.groupId || event.source.roomId || event.source.userId;
      // GASの実行ログでグループIDを確認するためのログ出力
      console.log('受信イベント種別: ' + event.type);
      console.log('グループID: ' + groupId);
    }
  } catch (err) {
    console.error(err);
  }
  return ContentService.createTextOutput('OK');
}


// ============================================
// LINE メッセージ送信（リプライ）
// ============================================
function replyLineMessage(replyToken, text) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: text }]
  };
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}


// ============================================
// LINE メッセージ送信（営業グループへのプッシュ）
// ============================================
function sendLineMessage(text) {
  const GROUP_ID = 'C87ef2d287f760158419805a3887ac7f9';
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: GROUP_ID,
    messages: [{ type: 'text', text: text }]
  };
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}


// ============================================
// 毎朝の自動アプローチ通知
// スクリプトプロパティから管理用LINEトークンとグループIDを取得して通知する
// ============================================
function sendApproachNotification() {
  // スクリプトプロパティから管理用LINEの認証情報を取得
  const props       = PropertiesService.getScriptProperties();
  const mgmtToken   = props.getProperty('MGMT_LINE_TOKEN');
  const mgmtGroupId = props.getProperty('MGMT_GROUP_ID');

  // トークンが未設定の場合は何もしない
  if (!mgmtToken) {
    console.log('MGMT_LINE_TOKEN が未設定のため、アプローチ通知をスキップしました。');
    return;
  }

  // 本日の日付文字列（JST）を "YYYY/MM/DD" 形式で取得
  const now      = new Date();
  const jstOffset = 9 * 60 * 60 * 1000; // UTC+9
  const jstNow   = new Date(now.getTime() + jstOffset);
  const todayStr = Utilities.formatDate(jstNow, 'Asia/Tokyo', 'yyyy/MM/dd');

  // スプレッドシート全行を取得
  const sheet  = getSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headers        = values[0].map(h => String(h).trim());
  const colApproach    = headers.indexOf('次回アプローチ予定日');
  const colMember      = headers.indexOf('担当者');
  const colDate        = headers.indexOf('日時');
  const colName        = headers.indexOf('顧客名(LINE名)');
  const colCategory    = headers.indexOf('導線');
  const colResult      = headers.indexOf('結果');
  const colRejection   = headers.indexOf('断り理由');
  const colFinancial   = headers.indexOf('資金状況');
  const colBarrier     = headers.indexOf('行動の障壁');
  const colUpsell      = headers.indexOf('アップセル見込み');
  const colMemo        = headers.indexOf('備考／所感');

  // 「次回アプローチ予定日」列が存在しない場合は終了
  if (colApproach < 0) {
    console.log('「次回アプローチ予定日」列が見つかりません。');
    return;
  }

  // 本日がアプローチ予定日の行を抽出
  const targets = [];
  for (let i = 1; i < values.length; i++) {
    const raw = values[i][colApproach];
    let dateStr = '';
    if (raw instanceof Date) {
      // スプレッドシートのDate型の場合はJSTでフォーマット
      dateStr = Utilities.formatDate(raw, 'Asia/Tokyo', 'yyyy/MM/dd');
    } else {
      // 文字列の場合はそのまま正規化（ハイフン区切り→スラッシュ区切り）
      dateStr = String(raw || '').trim().replace(/-/g, '/');
      // "YYYY/M/D" → "YYYY/MM/DD" に正規化
      const parts = dateStr.split('/');
      if (parts.length === 3) {
        dateStr = parts[0] + '/' + parts[1].padStart(2, '0') + '/' + parts[2].padStart(2, '0');
      }
    }
    if (dateStr === todayStr) {
      targets.push(values[i]);
    }
  }

  // 該当行がなければ通知不要
  if (targets.length === 0) {
    console.log('本日(' + todayStr + ')のアプローチ予定はありません。');
    return;
  }

  // セルの値を安全に取得するユーティリティ
  const get = (row, col) => col >= 0 ? String(row[col] || '').trim() : '';

  let messageText = '';

  if (targets.length === 1) {
    // ─── 1件フォーマット ───
    const r          = targets[0];
    const memberName = get(r, colMember);
    const dateVal    = r[colDate];
    const dateDisp   = dateVal instanceof Date
      ? Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd')
      : String(dateVal || '').replace(/-/g, '/');
    const name       = get(r, colName);
    const category   = get(r, colCategory);
    const result     = get(r, colResult);
    const rejection  = get(r, colRejection);
    const financial  = get(r, colFinancial);
    const barrier    = get(r, colBarrier);
    const memo       = get(r, colMemo);
    const approach   = todayStr;

    messageText += '🔔【アプローチ通知】\n\n';
    messageText += '👤 担当：' + memberName + '\n';
    messageText += '📅 面談日：' + dateDisp + '\n';
    messageText += '👥 顧客名：' + name + '\n';
    messageText += '📌 導線：' + category + '\n';
    messageText += '\n▼ 前回結果\n';
    messageText += '結果：' + result + '\n';
    if (rejection)  messageText += '断り理由：' + rejection  + '\n';
    if (financial)  messageText += '資金状況：'  + financial  + '\n';
    if (barrier)    messageText += '行動の障壁：' + barrier    + '\n';
    if (memo)       messageText += '\n💬 メモ：' + memo       + '\n';
    messageText += '\n👉 次回アプローチ予定日：' + approach + '（本日）';

  } else {
    // ─── 複数件フォーマット ───
    messageText += '🔔【アプローチ通知】本日 ' + targets.length + '件\n';
    const sep = '\n━━━━━━━━━━━━━━';

    targets.forEach((r, idx) => {
      const memberName = get(r, colMember);
      const name       = get(r, colName);
      const result     = get(r, colResult);
      const rejection  = get(r, colRejection);
      const financial  = get(r, colFinancial);
      const barrier    = get(r, colBarrier);
      const upsell     = get(r, colUpsell);
      const memo       = get(r, colMemo);

      messageText += sep + '\n';
      messageText += '① '.replace('①', ['①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩'][idx] || (idx + 1) + '.');
      messageText += ' 担当：' + memberName + ' / ' + name + '\n';
      messageText += '結果：' + result;
      if (rejection) messageText += ' ／ 断り理由：' + rejection;
      messageText += '\n';
      if (financial) messageText += '資金状況：' + financial;
      if (barrier)   messageText += (financial ? ' ／ ' : '') + '障壁：' + barrier;
      if (financial || barrier) messageText += '\n';
      if (upsell)    messageText += 'アップセル見込み：' + upsell + '\n';
      if (memo)      messageText += '💬 メモ：' + memo + '\n';
    });

    messageText += sep + '\n';
    messageText += '👉 本日がアプローチ予定日です！';
  }

  // 管理用LINEグループへ送信
  const url     = 'https://api.line.me/v2/bot/message/push';
  const linePayload = {
    to: mgmtGroupId,
    messages: [{ type: 'text', text: messageText }]
  };
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + mgmtToken
    },
    payload: JSON.stringify(linePayload)
  };

  try {
    UrlFetchApp.fetch(url, options);
    console.log('アプローチ通知を送信しました（' + targets.length + '件）');
  } catch (err) {
    console.error('LINE通知の送信に失敗しました: ' + err.toString());
  }
}


// ============================================
// 【初回セットアップ用】顧客情報ヘッダー列を追加する
// 一度だけ手動実行してください。既存列は変更しません。
// ============================================
function addCustomerInfoHeaders() {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  const newHeaders = [
    'アップセル見込み',
    '断り理由',
    '資金状況',
    '行動の障壁',
    'アプローチ目安',
    '次回アプローチ予定日',
    '備考／所感'
  ];

  let addedCount = 0;
  newHeaders.forEach(header => {
    if (!headers.includes(header)) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(header);
      headers.push(header);
      addedCount++;
      console.log('追加: ' + header);
    } else {
      console.log('スキップ（既存）: ' + header);
    }
  });

  SpreadsheetApp.flush();
  console.log('完了：' + addedCount + '列を追加しました。');
}


// ============================================
// 動作テスト用関数
// ============================================
function testWrite() {
  const sheet = getSheet();
  Logger.log('シート名: ' + sheet.getName());
  sheet.appendRow(['テスト', new Date(), 'テスト担当者', 'テスト導線', 'テスト顧客', '未成約', '', 0, 0, 'テストメモ']);
  SpreadsheetApp.flush();
  Logger.log('書き込み完了');
}


// ============================================
// 【設定方法】スクリプトプロパティの登録
// ============================================
//
// GASエディタの「プロジェクトの設定」→「スクリプトプロパティ」から
// 以下の2つのキーを登録してください：
//
//   キー名              値の例
//   ─────────────────────────────────────────
//   LINE_TOKEN         （営業グループ用LINEボットのチャンネルアクセストークン）
//   MGMT_LINE_TOKEN    （管理用LINEボットのチャンネルアクセストークン）
//   MGMT_GROUP_ID      （管理用LINEグループのグループID。例: C87ef2d...）
//
// ※ MGMT_LINE_TOKEN が未設定の場合、sendApproachNotification() は
//    何もせずに終了します（エラーにはなりません）。
//
// ============================================
// 【設定方法】トリガーの設定（毎朝8時に自動実行）
// ============================================
//
// GASエディタの「トリガー」（時計アイコン）→「トリガーを追加」から：
//
//   実行する関数       : sendApproachNotification
//   イベントのソース   : 時間主導型
//   時間ベースのトリガー: 日付ベースのタイマー
//   時刻               : 午前8時〜9時
//
// これで毎朝8時台に自動でアプローチ通知が送信されます。
// ============================================
