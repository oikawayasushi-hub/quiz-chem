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

/**
 * 取り組み状況の読み取りAPI (取り組み状況.html から使用)
 *
 * GET ?studentId=1104 に対して、その生徒のクイズごとの合否と挑戦回数を返す。
 * プライバシー配慮のため、得点・日時などの詳細は返さない。
 * 例: {"ok":true,"quizzes":{"113 元素記号と元素名（原子番号1〜20）":{"passed":true,"attempts":3}}}
 */
function doGet(e) {
  const studentId = String((e.parameter && e.parameter.studentId) || '').trim();
  if (!/^(20(0[0-9]|10)|1[1-6](0[1-9]|[1-3][0-9]|40))$/.test(studentId)) {
    return jsonOut_({ ok: false, error: 'invalid studentId' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Responses');
  const quizzes = {};

  if (sheet && sheet.getLastRow() >= 2) {
    // 列: 受信日時, 学籍番号, クイズ名, 得点, 問題数, 合否, ...
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    rows.forEach(function (row) {
      if (String(row[1]) !== studentId) return;
      const name = String(row[2]);
      const passed = String(row[5]) === '合格';
      if (!quizzes[name]) quizzes[name] = { passed: false, attempts: 0 };
      quizzes[name].attempts++;
      if (passed) quizzes[name].passed = true;
    });
  }

  return jsonOut_({ ok: true, quizzes: quizzes });
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
