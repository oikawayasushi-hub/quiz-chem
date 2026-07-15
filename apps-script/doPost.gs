/**
 * 小テスト結果 受信用 Apps Script (doPost)
 *
 * LockService により書き込みを直列化し、クラス全員が同時に送信しても
 * 行の上書き・取りこぼしが起きないようにしている。
 *
 * 【設置方法】
 * 1. 記録先スプレッドシートを開く → 拡張機能 → Apps Script
 * 2. 既存の doPost 関数をこのファイルの内容で置き換える
 *    (シート名や列の並びが既存と異なる場合は appendRow の中身を調整)
 * 3. 「デプロイ」→「デプロイを管理」→ 鉛筆アイコン →
 *    バージョン「新バージョン」を選んで「デプロイ」
 *    ※「新しいデプロイ」で作り直すとURLが変わり、全HTMLの
 *      SUBMIT_URL を差し替える必要が出るので注意
 */

// 記録先シート名(存在しなければ自動作成される)
const SHEET_NAME = '回答記録';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // 先客の書き込みが終わるまで最大30秒待つ(通常は一瞬で取得できる)
    lock.waitLock(30000);

    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['受信日時', '回答日時', '生徒ID', 'クイズ名', '得点', '満点', '合否', '間違えた問題']);
    }

    sheet.appendRow([
      new Date(),
      data.timestamp || '',
      data.studentId || '',
      data.quizName || '',
      data.score,
      data.total,
      data.passed || '',
      Array.isArray(data.missed) ? data.missed.join(' / ') : ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}
