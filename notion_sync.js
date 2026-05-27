/**
 * notion_sync.js (gr-car-schedule)
 *
 * Notion の GRCMS スタッフ / 車両 DB から定期的にデータを取得し、
 * Spreadsheet の スタッフマスタ / 車両マスタ シートを上書きする同期処理。
 *
 * ==================== Script Properties ====================
 *  NOTION_TOKEN              : Notion Integration token (secret_xxx...)
 *  NOTION_STAFF_DB_ID        : 13c7ae89-446b-4554-b99c-7f747c13b12c
 *  NOTION_VEHICLE_DB_ID      : e2ee67dc-737c-4e28-bcc2-8cd2eee174c4
 *  NOTION_ROLE_DB_ID         : a4af3a41-ce88-4e01-aeca-58e08127035b (将来用)
 *  NOTION_SYNC_ENABLED       : 'true' を設定すると同期実行。未設定or false なら no-op
 *
 * ==================== 初回セットアップ (後日) ====================
 *  1) Notion で Internal Integration 作成 (https://www.notion.so/profile/integrations)
 *     → secret_xxx... トークンを取得
 *  2) Notion の「🚗 GRCMS」ページを開く → 右上 ⋯ → 接続 → 上の Integration を選択
 *  3) Apps Script エディタ → Script Properties に上記 5 つを設定
 *     最後に NOTION_SYNC_ENABLED='true' を入れる
 *  4) testNotionConnection() を実行 (接続テスト)
 *  5) syncFromNotion() を手動実行 (初回同期)
 *  6) setupNotionSyncTrigger() を実行 (5 分おきトリガー登録)
 */

// 2025-09-03 以降の Notion API は data_sources エンドポイントを使う
const NOTION_API_VERSION = '2025-09-03';

// ==================== HTTP ヘルパ ====================

function _notionFetch_(method, endpoint, body) {
  const token = getProp_('NOTION_TOKEN');
  if (!token) throw new Error('NOTION_TOKEN 未設定');
  const url = 'https://api.notion.com/v1' + endpoint;
  const options = {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + token,
      'Notion-Version': NOTION_API_VERSION,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  if (body) options.payload = JSON.stringify(body);

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) {
    Logger.log('Notion API error: ' + code + ' ' + endpoint + ' body=' + String(text).substring(0, 300));
    throw new Error('Notion API ' + code + ': ' + String(text).substring(0, 200));
  }
  return JSON.parse(text);
}

/** data source を全件ページネーション取得 */
function _notionQueryAll_(dataSourceId, filter) {
  const results = [];
  let cursor = null;
  let guard = 0;
  do {
    if (++guard > 50) {
      Logger.log('_notionQueryAll_ ページネーション 50 周到達、中断');
      break;
    }
    const body = { page_size: 100 };
    if (filter) body.filter = filter;
    if (cursor) body.start_cursor = cursor;
    const res = _notionFetch_('POST', '/data_sources/' + dataSourceId + '/query', body);
    (res.results || []).forEach(r => results.push(r));
    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);
  return results;
}

// ==================== プロパティ値抽出 ====================

function _propValue_(prop) {
  if (!prop) return '';
  switch (prop.type) {
    case 'title':
      return (prop.title || []).map(t => t.plain_text).join('');
    case 'rich_text':
      return (prop.rich_text || []).map(t => t.plain_text).join('');
    case 'select':
      return prop.select ? prop.select.name : '';
    case 'multi_select':
      return (prop.multi_select || []).map(o => o.name).join(',');
    case 'checkbox':
      return !!prop.checkbox;
    case 'number':
      return (prop.number !== null && prop.number !== undefined) ? prop.number : '';
    case 'date':
      return prop.date ? (prop.date.start || '') : '';
    case 'email':
      return prop.email || '';
    case 'phone_number':
      return prop.phone_number || '';
    case 'unique_id':
      if (!prop.unique_id) return '';
      return (prop.unique_id.prefix ? prop.unique_id.prefix + '-' : '') + prop.unique_id.number;
    case 'relation':
      return (prop.relation || []).map(r => r.id).join(',');
    case 'created_time':
      return prop.created_time || '';
    case 'last_edited_time':
      return prop.last_edited_time || '';
    case 'rollup':
      if (!prop.rollup) return '';
      if (prop.rollup.type === 'array') {
        return (prop.rollup.array || []).map(item => _propValue_(item)).join(',');
      }
      if (prop.rollup.type === 'number') return prop.rollup.number || '';
      if (prop.rollup.type === 'date') return prop.rollup.date ? (prop.rollup.date.start || '') : '';
      return '';
    default:
      return '';
  }
}

// ==================== メイン同期エントリ ====================

/**
 * Time-based trigger から 5 分おきに呼ばれる。
 * NOTION_SYNC_ENABLED='true' のときだけ実際に同期。
 */
function syncFromNotion() {
  const enabled = getProp_('NOTION_SYNC_ENABLED');
  if (enabled !== 'true') {
    Logger.log('NOTION_SYNC_ENABLED != true、同期スキップ');
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('他の処理と競合、同期スキップ');
    return;
  }

  try {
    Logger.log('=== Notion → Spreadsheet 同期開始 ===');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const staffCount = syncStaffFromNotion_(ss);
    const vehicleCount = syncVehiclesFromNotion_(ss);
    invalidateStaffCache_(); // staff_lookup_by_line のキャッシュも世代カウンタを上げる
    Logger.log('=== 同期完了: スタッフ ' + staffCount + ' 件 / 車両 ' + vehicleCount + ' 件 ===');
  } catch (err) {
    Logger.log('同期エラー: ' + err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// ==================== スタッフ同期 ====================

function syncStaffFromNotion_(ss) {
  const dsId = getProp_('NOTION_STAFF_DB_ID');
  if (!dsId) throw new Error('NOTION_STAFF_DB_ID 未設定');

  // GRCMS 有効 = true のみ
  const filter = {
    property: 'GRCMS有効',
    checkbox: { equals: true }
  };
  const pages = _notionQueryAll_(dsId, filter);
  Logger.log('Notion スタッフ取得: ' + pages.length + ' 件');

  const sheet = getOrCreateStaffSheet_(ss);
  _clearDataRows_(sheet);

  const now = new Date();
  pages.forEach(page => {
    const p = page.properties || {};
    const lineId = _propValue_(p['LINE_user_id']);
    // LINE_user_id 空のスタッフは LINE 連携できないので除外
    if (!lineId) return;
    const row = STAFF_HEADERS.map(h => {
      switch (h) {
        case 'スタッフID': return _propValue_(p['スタッフID']);
        case '氏名':       return _propValue_(p['氏名']);
        case 'LINE_ID':    return lineId;
        case '役職':       return _propValue_(p['GRCMS役割']);
        case '連絡先':     return '';
        case '有効フラグ': return true;
        case '登録日時':   return now;
        default:          return '';
      }
    });
    sheet.appendRow(row);
  });
  return pages.length;
}

// ==================== 車両同期 ====================

function syncVehiclesFromNotion_(ss) {
  const dsId = getProp_('NOTION_VEHICLE_DB_ID');
  if (!dsId) throw new Error('NOTION_VEHICLE_DB_ID 未設定');

  // 全件 (PWA でフィルタ)
  const pages = _notionQueryAll_(dsId);
  Logger.log('Notion 車両取得: ' + pages.length + ' 件');

  const sheet = getOrCreateVehicleSheet_(ss);
  _clearDataRows_(sheet);

  const now = new Date();
  pages.forEach(page => {
    const p = page.properties || {};
    const row = VEHICLE_HEADERS.map(h => {
      switch (h) {
        case '車両ID':   return _propValue_(p['車両ID']);
        case '車両名':   return _propValue_(p['車両名']);
        case 'ナンバー': return _propValue_(p['ナンバー']);
        case '所有会社': return _propValue_(p['所有会社']);
        case '状態':     return _propValue_(p['状態']);
        case '仕入日':   return _propValue_(p['仕入日']);
        case '仕入額':   return _propValue_(p['仕入額']);
        case '売却日':   return _propValue_(p['売却日']);
        case '売却額':   return _propValue_(p['売却額']);
        case 'メモ':     return _propValue_(p['メモ']);
        case '登録日時': return now;
        case '更新日時': return now;
        default:        return '';
      }
    });
    sheet.appendRow(row);
  });
  return pages.length;
}

// ==================== ユーティリティ ====================

function _clearDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}

// ==================== セットアップ系 (手動実行) ====================

/** 5 分おきの同期トリガーを登録 */
function setupNotionSyncTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncFromNotion')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncFromNotion')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('✅ 5 分おき syncFromNotion トリガー登録完了');
}

/** トリガーを削除 (停止) */
function removeNotionSyncTrigger() {
  let cnt = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncFromNotion') {
      ScriptApp.deleteTrigger(t);
      cnt++;
    }
  });
  Logger.log('✅ トリガー削除: ' + cnt + ' 件');
}

/** 接続テスト: data source の存在確認だけ */
function testNotionConnection() {
  ['NOTION_TOKEN', 'NOTION_STAFF_DB_ID', 'NOTION_VEHICLE_DB_ID', 'NOTION_ROLE_DB_ID'].forEach(k => {
    const v = getProp_(k);
    Logger.log(k + ': ' + (v ? '✅ 設定済' : '❌ 未設定'));
  });
  const dsId = getProp_('NOTION_STAFF_DB_ID');
  if (!dsId) return;
  try {
    const res = _notionFetch_('GET', '/data_sources/' + dsId);
    const title = (res.title || []).map(t => t.plain_text).join('') || '(no title)';
    Logger.log('✅ スタッフ DB 接続 OK: ' + title);
  } catch (err) {
    Logger.log('❌ 接続 NG: ' + err);
  }
}

/** 手動で 1 回だけ同期を強制実行 (NOTION_SYNC_ENABLED 関係なく) */
function syncFromNotionForce() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('競合'); return; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sc = syncStaffFromNotion_(ss);
    const vc = syncVehiclesFromNotion_(ss);
    invalidateStaffCache_();
    Logger.log('✅ 強制同期完了: スタッフ ' + sc + ' / 車両 ' + vc);
  } finally {
    lock.releaseLock();
  }
}
