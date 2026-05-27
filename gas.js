/**
 * gr-car-schedule GAS バックエンド
 * バインド対象: 車屋予定管理 (1puu1IMlb07MjZSdcZBjw1iMMgFn-K-cuB3vrppSrues)
 *
 * デプロイ手順:
 *  1) Drive で 車屋予定管理 を開く
 *  2) 拡張機能 → Apps Script
 *  3) このファイル全体を貼り付けて保存
 *  4) Script Properties に SCHEDULE_TOKEN と ADMIN_PIN を設定
 *  5) meta_init を 1 回実行（認可ダイアログ承認 + シート 4 枚自動構築）
 *  6) デプロイ → ウェブアプリ → アクセス: 全員（匿名含む）
 */

// ====== Script Properties ヘルパ ======
function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

// ====== 共通レスポンス ======
function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(Object.assign({status: 'ok'}, data || {})))
    .setMimeType(ContentService.MimeType.JSON);
}

function error(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({status: 'error', message: msg}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== doPost エントリ ======
function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return error('他の操作と競合しました、再度お試しください');
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = String(body.action || '');
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // === ここに各アクション分岐を追加していく（Task 4 以降） ===
    if (action === 'meta_init') {
      return handleMetaInit_(ss, body);
    }

    return error('未知のアクション: ' + action);
  } catch (err) {
    return error(err.toString());
  } finally {
    lock.releaseLock();
  }
}

// ====== doGet (動作確認用) ======
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status: 'ok', service: 'gr-car-schedule', timestamp: new Date().toISOString()}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== トークン認証 ======
function checkToken_(body) {
  const expected = getProp_('SCHEDULE_TOKEN');
  if (!expected) throw new Error('SCHEDULE_TOKEN 未設定');
  if (String(body.token || '') !== expected) throw new Error('認証失敗');
}

function checkPin_(body) {
  const expected = getProp_('ADMIN_PIN');
  if (!expected) throw new Error('ADMIN_PIN 未設定');
  if (String(body.pin || '') !== expected) throw new Error('PIN 不一致');
}

// ====== シート定義 ======
const SCHEDULE_SHEET = '予定';
const SCHEDULE_HEADERS = [
  '予定ID', '開始日時', '終了予定日時',
  '担当スタッフ', '担当LINE_ID',
  '作業区分', '車両ID', '車両名', 'ナンバー',
  '場所・行先', 'メモ', '状態', '起点',
  '顧客名', '顧客LINE_ID', 'GRCMS_受付ID',  // Phase 2
  '写真URL',                                  // Phase 3
  '登録日時', '更新日時', '更新者'
];

const VEHICLE_SHEET = '車両マスタ';
const VEHICLE_HEADERS = [
  '車両ID', '車両名', 'ナンバー', '所有会社', '状態',
  '仕入日', '仕入額', '売却日', '売却額', 'メモ',
  '登録日時', '更新日時'
];

const STAFF_SHEET = 'スタッフマスタ';
const STAFF_HEADERS = [
  'スタッフID', '氏名', 'LINE_ID', '役職', '連絡先',
  '有効フラグ', '登録日時'
];

const OPLOG_SHEET = '操作ログ';
const OPLOG_HEADERS = [
  '日時', 'アクション', '対象', '操作者', '起点', '差分'
];

// ====== 列幅の確保 ======
function ensureColumns_(sheet, needed) {
  const current = sheet.getMaxColumns();
  if (current < needed) sheet.insertColumnsAfter(current, needed - current);
}

// ====== シート自動構築 ======
function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    ensureColumns_(sheet, headers.length);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    ensureColumns_(sheet, headers.length);
    const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    headers.forEach((h, i) => {
      if (currentHeaders[i] !== h) sheet.getRange(1, i + 1).setValue(h);
    });
  }
  return sheet;
}

function getOrCreateScheduleSheet_(ss) { return getOrCreateSheet_(ss, SCHEDULE_SHEET, SCHEDULE_HEADERS); }
function getOrCreateVehicleSheet_(ss)  { return getOrCreateSheet_(ss, VEHICLE_SHEET, VEHICLE_HEADERS); }
function getOrCreateStaffSheet_(ss)    { return getOrCreateSheet_(ss, STAFF_SHEET, STAFF_HEADERS); }
function getOrCreateOpLogSheet_(ss)    { return getOrCreateSheet_(ss, OPLOG_SHEET, OPLOG_HEADERS); }

// ====== 操作ログ ======
function logOperation_(ss, action, target, operator, source, diff) {
  try {
    const sheet = getOrCreateOpLogSheet_(ss);
    sheet.appendRow([
      new Date(),
      String(action || ''),
      String(target || ''),
      String(operator || ''),
      String(source || ''),
      diff ? JSON.stringify(diff) : ''
    ]);
  } catch (err) {
    Logger.log('logOperation_ failed: ' + err);
  }
}

// ====== ID 生成 ======
function generateScheduleId_() {
  const tz = Session.getScriptTimeZone();
  return 'S-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 1000);
}

function generateVehicleId_(ss) {
  const sheet = getOrCreateVehicleSheet_(ss);
  const last = sheet.getLastRow();
  return 'V-' + String(last).padStart(3, '0');
}

function generateStaffId_(ss) {
  const sheet = getOrCreateStaffSheet_(ss);
  const last = sheet.getLastRow();
  return 'M-' + String(last).padStart(3, '0');
}

// ====== 行を ID で検索 (returns {row, data}) ======
function findRowById_(sheet, idCol, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      return { row: i + 1, data: data[i] };
    }
  }
  return null;
}

// ====== 行を Headers + values で更新（fieldMap 経由） ======
function applyPartialUpdate_(sheet, rowNum, headers, currentRow, partial, fieldMap) {
  const updates = currentRow.slice();
  Object.keys(fieldMap).forEach(key => {
    if (partial[key] !== undefined) {
      const colIdx = headers.indexOf(fieldMap[key]);
      if (colIdx >= 0) updates[colIdx] = partial[key];
    }
  });
  const updColIdx = headers.indexOf('更新日時');
  if (updColIdx >= 0) updates[updColIdx] = new Date();
  sheet.getRange(rowNum, 1, 1, headers.length).setValues([updates]);
  return updates;
}
