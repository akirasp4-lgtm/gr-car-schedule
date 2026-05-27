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

    // === Action dispatch ===
    if (action === 'schedule_add')    return handleScheduleAdd_(ss, body);
    if (action === 'schedule_update') return handleScheduleUpdate_(ss, body);
    if (action === 'schedule_delete') return handleScheduleDelete_(ss, body);
    if (action === 'schedule_list')   return handleScheduleList_(ss, body);
    if (action === 'vehicle_add')     return handleVehicleAdd_(ss, body);
    if (action === 'vehicle_update')  return handleVehicleUpdate_(ss, body);
    if (action === 'vehicle_delete')  return handleVehicleDelete_(ss, body);
    if (action === 'vehicle_list')    return handleVehicleList_(ss, body);
    if (action === 'staff_add')           return handleStaffAdd_(ss, body);
    if (action === 'staff_update')        return handleStaffUpdate_(ss, body);
    if (action === 'staff_list')          return handleStaffList_(ss, body);
    if (action === 'staff_lookup_by_line') return handleStaffLookupByLine_(ss, body);
    if (action === 'meta_init')       return handleMetaInit_(ss, body);

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

// ====== 共通ヘルパ ======

/**
 * シート/JSON から来る「有効フラグ」のような値を boolean に正規化する。
 * true / 'TRUE' / 'true' / 1 / '1' → true
 * それ以外 (false / 'FALSE' / 'false' / 0 / '' / null / undefined) → false
 */
function normalizeBool_(v) {
  if (v === true || v === 1) return true;
  if (v === false || v === 0 || v === null || v === undefined) return false;
  const s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === '1' || s === 'YES';
}

/**
 * 「状態」が「キャンセル」かどうか判定する (trim + 大文字小文字を許容)。
 * スペース混入や全角差異への耐性を持たせる。
 */
function isCancelledStatus_(v) {
  if (v === null || v === undefined) return false;
  return String(v).trim() === 'キャンセル';
}

/**
 * 日時値を ISO 8601 (yyyy-MM-dd'T'HH:mm:ssXXX) 文字列に正規化する。
 * Date オブジェクトでも文字列でも受け取れる。パース不可なら空文字を返す。
 */
function normalizeDateTime_(v) {
  if (!v) return '';
  let d;
  if (v instanceof Date) {
    d = v;
  } else {
    const s = String(v).trim();
    if (!s) return '';
    // YYYY-MM-DD のみなら 00:00:00+09:00 を付ける
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      d = new Date(s + 'T00:00:00+09:00');
    } else {
      d = new Date(s);
    }
  }
  if (!d || isNaN(d.getTime())) {
    // パース不可。元の値を String 化して返す（後方互換）
    return String(v);
  }
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
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

// ====== Action: meta_init ======
function handleMetaInit_(ss, body) {
  checkToken_(body);
  const created = [];
  if (!ss.getSheetByName(SCHEDULE_SHEET)) created.push(SCHEDULE_SHEET);
  if (!ss.getSheetByName(VEHICLE_SHEET)) created.push(VEHICLE_SHEET);
  if (!ss.getSheetByName(STAFF_SHEET)) created.push(STAFF_SHEET);
  if (!ss.getSheetByName(OPLOG_SHEET)) created.push(OPLOG_SHEET);

  getOrCreateScheduleSheet_(ss);
  getOrCreateVehicleSheet_(ss);
  getOrCreateStaffSheet_(ss);
  getOrCreateOpLogSheet_(ss);

  // デフォルトシート Sheet1 が残っていれば削除（4 つ作成済みなら）
  const sheet1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
  if (sheet1 && ss.getSheets().length > 4) {
    try { ss.deleteSheet(sheet1); } catch (_) {}
  }

  logOperation_(ss, 'meta_init', '', 'system', 'admin', { created });
  return ok({ sheets_created: created });
}

// ====== Action: schedule_add ======
function handleScheduleAdd_(ss, body) {
  checkToken_(body);
  const ev = body.event || body;
  const id = generateScheduleId_();
  const now = new Date();
  const sheet = getOrCreateScheduleSheet_(ss);

  const row = SCHEDULE_HEADERS.map(h => {
    switch (h) {
      case '予定ID': return id;
      case '開始日時': return normalizeDateTime_(ev.time_start || ev.start_dt || '');
      case '終了予定日時': return normalizeDateTime_(ev.time_end || ev.end_dt || '');
      case '担当スタッフ': return String(ev.staffName || '');
      case '担当LINE_ID': return String(ev.staffLineId || '');
      case '作業区分': return String(ev.workType || 'その他');
      case '車両ID': return String(ev.vehicleId || '');
      case '車両名': return String(ev.vehicleName || '');
      case 'ナンバー': return String(ev.vehiclePlate || '');
      case '場所・行先': return String(ev.place || '');
      case 'メモ': return String(ev.memo || '');
      case '状態': return String(ev.status || '予約');
      case '起点': return String(ev.source || 'Web');
      case '顧客名': return String(ev.customerName || '');
      case '顧客LINE_ID': return String(ev.customerLineId || '');
      case 'GRCMS_受付ID': return String(ev.grcmsIntakeId || '');
      case '写真URL': return String(ev.photoUrls || '');
      case '登録日時': return now;
      case '更新日時': return now;
      case '更新者': return String(ev.updatedBy || ev.staffName || '');
      default: return '';
    }
  });

  sheet.appendRow(row);
  logOperation_(ss, 'schedule_add', id, ev.updatedBy || ev.staffName || '', ev.source || 'Web', ev);
  return ok({ id });
}

// ====== Action: schedule_update ======
function handleScheduleUpdate_(ss, body) {
  checkToken_(body);
  const ev = body.event || body;
  const id = String(ev.id || '');
  if (!id) return error('id が指定されていません');

  const sheet = getOrCreateScheduleSheet_(ss);
  const found = findRowById_(sheet, SCHEDULE_HEADERS.indexOf('予定ID'), id);
  if (!found) return error('対象が見つかりません: ' + id);

  const fieldMap = {
    time_start: '開始日時', time_end: '終了予定日時',
    staffName: '担当スタッフ', staffLineId: '担当LINE_ID',
    workType: '作業区分',
    vehicleId: '車両ID', vehicleName: '車両名', vehiclePlate: 'ナンバー',
    place: '場所・行先', memo: 'メモ', status: '状態',
    customerName: '顧客名', customerLineId: '顧客LINE_ID',
    grcmsIntakeId: 'GRCMS_受付ID', photoUrls: '写真URL',
    updatedBy: '更新者'
  };

  const updated = applyPartialUpdate_(sheet, found.row, SCHEDULE_HEADERS, found.data, ev, fieldMap);
  logOperation_(ss, 'schedule_update', id, ev.updatedBy || '', ev.source || 'Web', ev);
  return ok({ updated: id });
}

// ====== Action: schedule_delete (論理削除) ======
function handleScheduleDelete_(ss, body) {
  checkToken_(body);
  const id = String(body.id || '');
  if (!id) return error('id が指定されていません');

  const sheet = getOrCreateScheduleSheet_(ss);
  const found = findRowById_(sheet, SCHEDULE_HEADERS.indexOf('予定ID'), id);
  if (!found) return error('対象が見つかりません: ' + id);

  const statusCol = SCHEDULE_HEADERS.indexOf('状態') + 1;
  const updCol = SCHEDULE_HEADERS.indexOf('更新日時') + 1;
  sheet.getRange(found.row, statusCol).setValue('キャンセル');
  sheet.getRange(found.row, updCol).setValue(new Date());

  logOperation_(ss, 'schedule_delete', id, body.updatedBy || '', body.source || 'Web', { reason: 'logical_delete' });
  return ok({ cancelled: id });
}

// ====== Action: schedule_list ======
function handleScheduleList_(ss, body) {
  checkToken_(body);
  const sheet = getOrCreateScheduleSheet_(ss);
  const tz = Session.getScriptTimeZone();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ok({ rows: [], count: 0 });

  const headers = data[0];
  let rows = data.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, j) => {
      const v = r[j];
      if (v instanceof Date) {
        obj[h] = Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
      } else {
        obj[h] = (v === undefined || v === null) ? '' : String(v);
      }
    });
    return obj;
  });

  // Filters
  if (body.dateFrom) rows = rows.filter(r => r['開始日時'] && r['開始日時'] >= body.dateFrom);
  if (body.dateTo)   rows = rows.filter(r => r['開始日時'] && r['開始日時'] <= body.dateTo);
  if (body.staffId)  rows = rows.filter(r => r['担当LINE_ID'] === body.staffId || r['担当スタッフ'] === body.staffId);
  if (body.vehicleId) rows = rows.filter(r => r['車両ID'] === body.vehicleId);
  if (body.status)   rows = rows.filter(r => r['状態'] === body.status);
  if (body.excludeCancelled) rows = rows.filter(r => r['状態'] !== 'キャンセル');

  return ok({ rows, count: rows.length });
}

// ====== Action: vehicle_add ======
function handleVehicleAdd_(ss, body) {
  checkToken_(body);
  checkPin_(body);
  const ev = body.event || body;
  const id = generateVehicleId_(ss);
  const now = new Date();
  const sheet = getOrCreateVehicleSheet_(ss);

  const row = VEHICLE_HEADERS.map(h => {
    switch (h) {
      case '車両ID': return id;
      case '車両名': return String(ev.name || '');
      case 'ナンバー': return String(ev.plate || '');
      case '所有会社': return String(ev.ownerCompany || '');
      case '状態': return String(ev.status || '稼働中');
      case '仕入日': return String(ev.purchaseDate || '');
      case '仕入額': return Number(ev.purchasePrice) || '';
      case '売却日': return String(ev.sellDate || '');
      case '売却額': return Number(ev.sellPrice) || '';
      case 'メモ': return String(ev.memo || '');
      case '登録日時': return now;
      case '更新日時': return now;
      default: return '';
    }
  });

  sheet.appendRow(row);
  logOperation_(ss, 'vehicle_add', id, body.updatedBy || 'admin', 'admin', ev);
  return ok({ id });
}

// ====== Action: vehicle_update ======
function handleVehicleUpdate_(ss, body) {
  checkToken_(body);
  checkPin_(body);
  const ev = body.event || body;
  const id = String(ev.id || '');
  if (!id) return error('id が指定されていません');

  const sheet = getOrCreateVehicleSheet_(ss);
  const found = findRowById_(sheet, VEHICLE_HEADERS.indexOf('車両ID'), id);
  if (!found) return error('対象が見つかりません: ' + id);

  const fieldMap = {
    name: '車両名', plate: 'ナンバー', ownerCompany: '所有会社', status: '状態',
    purchaseDate: '仕入日', purchasePrice: '仕入額',
    sellDate: '売却日', sellPrice: '売却額',
    memo: 'メモ'
  };
  applyPartialUpdate_(sheet, found.row, VEHICLE_HEADERS, found.data, ev, fieldMap);
  logOperation_(ss, 'vehicle_update', id, body.updatedBy || 'admin', 'admin', ev);
  return ok({ updated: id });
}

// ====== Action: vehicle_delete (force 二段確認) ======
function handleVehicleDelete_(ss, body) {
  checkToken_(body);
  checkPin_(body);
  const id = String(body.id || '');
  if (!id) return error('id が指定されていません');

  // 予定で使用中かチェック
  const scheduleSheet = getOrCreateScheduleSheet_(ss);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const vehicleIdCol = SCHEDULE_HEADERS.indexOf('車両ID');
  const statusCol = SCHEDULE_HEADERS.indexOf('状態');
  let usageCount = 0;
  for (let i = 1; i < scheduleData.length; i++) {
    if (String(scheduleData[i][vehicleIdCol]) === id &&
        !isCancelledStatus_(scheduleData[i][statusCol])) {
      usageCount++;
    }
  }

  if (usageCount > 0 && !body.force) {
    return error('予定 ' + usageCount + ' 件で使用中です。force:true で強制削除できます');
  }

  // 物理削除
  const sheet = getOrCreateVehicleSheet_(ss);
  const found = findRowById_(sheet, VEHICLE_HEADERS.indexOf('車両ID'), id);
  if (!found) return error('対象が見つかりません: ' + id);
  sheet.deleteRow(found.row);

  logOperation_(ss, 'vehicle_delete', id, body.updatedBy || 'admin', 'admin', { force: !!body.force, usageCount });
  return ok({ deleted: id, hadUsage: usageCount });
}

// ====== Action: vehicle_list ======
function handleVehicleList_(ss, body) {
  checkToken_(body);
  const sheet = getOrCreateVehicleSheet_(ss);
  const tz = Session.getScriptTimeZone();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ok({ rows: [] });
  const headers = data[0];
  const rows = data.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, j) => {
      const v = r[j];
      if (v instanceof Date) {
        obj[h] = Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
      } else {
        obj[h] = (v === undefined || v === null) ? '' : String(v);
      }
    });
    return obj;
  });
  return ok({ rows });
}

// ====== Action: staff_add ======
function handleStaffAdd_(ss, body) {
  checkToken_(body);
  checkPin_(body);
  const ev = body.event || body;
  const id = generateStaffId_(ss);
  const now = new Date();
  const sheet = getOrCreateStaffSheet_(ss);

  const row = STAFF_HEADERS.map(h => {
    switch (h) {
      case 'スタッフID': return id;
      case '氏名': return String(ev.name || '');
      case 'LINE_ID': return String(ev.lineId || '');
      case '役職': return String(ev.role || 'スタッフ');
      case '連絡先': return String(ev.contact || '');
      case '有効フラグ': return ev.active !== false;
      case '登録日時': return now;
      default: return '';
    }
  });

  sheet.appendRow(row);
  invalidateStaffCache_();
  logOperation_(ss, 'staff_add', id, body.updatedBy || 'admin', 'admin', ev);
  return ok({ id });
}

// ====== Action: staff_update ======
function handleStaffUpdate_(ss, body) {
  checkToken_(body);
  checkPin_(body);
  const ev = body.event || body;
  const id = String(ev.id || '');
  if (!id) return error('id が指定されていません');

  const sheet = getOrCreateStaffSheet_(ss);
  const found = findRowById_(sheet, STAFF_HEADERS.indexOf('スタッフID'), id);
  if (!found) return error('対象が見つかりません: ' + id);

  const fieldMap = {
    name: '氏名', lineId: 'LINE_ID', role: '役職', contact: '連絡先', active: '有効フラグ'
  };
  applyPartialUpdate_(sheet, found.row, STAFF_HEADERS, found.data, ev, fieldMap);
  invalidateStaffCache_();
  logOperation_(ss, 'staff_update', id, body.updatedBy || 'admin', 'admin', ev);
  return ok({ updated: id });
}

// ====== Action: staff_list ======
function handleStaffList_(ss, body) {
  checkToken_(body);
  const sheet = getOrCreateStaffSheet_(ss);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ok({ rows: [] });
  const headers = data[0];
  const rows = data.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = (r[j] === null || r[j] === undefined) ? '' : (r[j] instanceof Date ? r[j].toISOString() : r[j]); });
    return obj;
  });
  return ok({ rows });
}

// ====== Action: staff_lookup_by_line (5分キャッシュ) ======
function handleStaffLookupByLine_(ss, body) {
  checkToken_(body);
  const lineId = String(body.lineId || '');
  if (!lineId) return ok({ isStaff: false });

  // 世代カウンタを含めたキャッシュキー (invalidateStaffCache_ が世代を上げると古いキャッシュは無効)
  const gen = _getStaffCacheGeneration_();
  const cacheKey = 'staff_v' + gen + '_' + lineId;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) return ok(JSON.parse(cached));

  const sheet = getOrCreateStaffSheet_(ss);
  const data = sheet.getDataRange().getValues();
  const lineIdCol = STAFF_HEADERS.indexOf('LINE_ID');
  const activeCol = STAFF_HEADERS.indexOf('有効フラグ');
  const idCol = STAFF_HEADERS.indexOf('スタッフID');
  const nameCol = STAFF_HEADERS.indexOf('氏名');
  const roleCol = STAFF_HEADERS.indexOf('役職');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][lineIdCol]) === lineId && normalizeBool_(data[i][activeCol])) {
      const result = {
        isStaff: true,
        id: String(data[i][idCol]),
        name: String(data[i][nameCol]),
        role: String(data[i][roleCol])
      };
      cache.put(cacheKey, JSON.stringify(result), 300);
      return ok(result);
    }
  }

  const negative = { isStaff: false };
  cache.put(cacheKey, JSON.stringify(negative), 300);
  return ok(negative);
}

// ====== スタッフキャッシュ世代カウンタ ======
function _getStaffCacheGeneration_() {
  const props = PropertiesService.getScriptProperties();
  const v = props.getProperty('STAFF_CACHE_GEN');
  return v ? Number(v) : 1;
}

/**
 * スタッフマスタが更新された時に世代カウンタを +1 する。
 * これにより既存の全 staff_v{N}_* キャッシュが事実上無効化される
 * (古いキーは TTL 5min で自然消滅)。
 */
function invalidateStaffCache_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const cur = Number(props.getProperty('STAFF_CACHE_GEN') || '1');
    props.setProperty('STAFF_CACHE_GEN', String(cur + 1));
  } catch (err) {
    Logger.log('invalidateStaffCache_ failed: ' + err);
  }
}

// ============================================================
// 手動テスト用関数（GAS エディタから実行）
// ============================================================

function test_all() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ⚠️ Safety: do not overwrite production Script Properties
  const props = PropertiesService.getScriptProperties();
  const existingToken = props.getProperty('SCHEDULE_TOKEN');
  const existingPin = props.getProperty('ADMIN_PIN');
  if (existingToken && existingToken !== 'test-token') {
    Logger.log('❌ ABORT: SCHEDULE_TOKEN already set to a non-test value. test_all は本番環境では実行不可。');
    return;
  }
  if (existingPin && existingPin !== '1203') {
    Logger.log('❌ ABORT: ADMIN_PIN already set to a non-test value. test_all は本番環境では実行不可。');
    return;
  }
  if (!existingToken) props.setProperty('SCHEDULE_TOKEN', 'test-token');
  if (!existingPin) props.setProperty('ADMIN_PIN', '1203');

  Logger.log('=== Test 1: meta_init ===');
  Logger.log(handleMetaInit_(ss, { token: 'test-token' }).getContent());

  Logger.log('=== Test 2: vehicle_add ===');
  const v = JSON.parse(handleVehicleAdd_(ss, {
    token: 'test-token', pin: '1203',
    name: 'テスト車両', plate: 'なにわ123あ4567', ownerCompany: 'GRHD'
  }).getContent());
  Logger.log(v);

  Logger.log('=== Test 3: staff_add ===');
  const s = JSON.parse(handleStaffAdd_(ss, {
    token: 'test-token', pin: '1203',
    name: 'テスト スタッフ', lineId: 'U_TEST_USER_001', role: 'スタッフ'
  }).getContent());
  Logger.log(s);

  Logger.log('=== Test 4: staff_lookup_by_line ===');
  Logger.log(handleStaffLookupByLine_(ss, { token: 'test-token', lineId: 'U_TEST_USER_001' }).getContent());

  Logger.log('=== Test 5: schedule_add ===');
  const sc = JSON.parse(handleScheduleAdd_(ss, {
    token: 'test-token',
    time_start: new Date().toISOString(),
    staffName: 'テスト スタッフ',
    staffLineId: 'U_TEST_USER_001',
    workType: '整備修理',
    vehicleId: v.id,
    vehicleName: 'テスト車両',
    vehiclePlate: 'なにわ123あ4567',
    memo: 'テスト予定',
    source: 'test'
  }).getContent());
  Logger.log(sc);

  Logger.log('=== Test 6: schedule_list ===');
  Logger.log(handleScheduleList_(ss, { token: 'test-token' }).getContent());

  Logger.log('=== Test 7: vehicle_delete (拒否) ===');
  Logger.log(handleVehicleDelete_(ss, { token: 'test-token', pin: '1203', id: v.id }).getContent());

  Logger.log('=== Test 8: vehicle_delete (force) ===');
  Logger.log(handleVehicleDelete_(ss, { token: 'test-token', pin: '1203', id: v.id, force: true }).getContent());

  Logger.log('=== All tests done. シートに残ったテストデータは手動で削除してください。 ===');
}
