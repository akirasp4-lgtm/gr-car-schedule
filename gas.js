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
