/**
 * 小テスト結果 受信用 Apps Script (doPost)
 *
 * LockService により書き込みを直列化し、クラス全員が同時に送信しても
 * 行の上書き・取りこぼしが起きないようにしている。
 *
 * 【設置方法】
 * 1. 記録先スプレッドシートを開く → 拡張機能 → Apps Script
 * 2. 既存の doPost 関数をこのファイルの内容で置き換える
 * 3. 「デプロイ」→「デプロイを管理」→ 鉛筆アイコン →
 *    バージョン「新バージョン」を選んで「デプロイ」
 *    ※「新しいデプロイ」で作り直すとURLが変わり、全HTMLの
 *      SUBMIT_URL を差し替える必要が出るので注意
 */

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // 先客の書き込みが終わるまで最大30秒待つ(通常は一瞬で取得できる)
    lock.waitLock(30000);

    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Responses") || ss.insertSheet("Responses");

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["受信日時", "学籍番号", "クイズ名", "得点", "問題数", "合否", "誤答した問題", "クラス"]);
    }

    const studentId = String(data.studentId);
    const className = studentId.substring(1, 2);

    sheet.appendRow([
      new Date(),
      studentId,
      data.quizName,
      data.score,
      data.total,
      data.passed,
      (data.missed || []).join(" / "),
      className
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}
