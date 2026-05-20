// ============================================================
// Google Apps Script — English Level Quiz → スプレッドシート保存
// ============================================================

const SPREADSHEET_ID = '1lesu0r5iEWe-IPkKUJE2U3ztiei4IFP-aOJbJ6WOgP4';
const SHEET_NAME = '入校テスト';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // 回答履歴のGoogle Docを作成
    const docUrl = createAnswerDoc(data);

    sheet.appendRow([
      data.datetime,   // A列: 日時
      data.name,       // B列: 名前
      data.score,      // C列: スコア
      data.cefr,       // D列: CEFRレベル
      data.levelName,  // E列: レベル名
      data.elapsed,    // F列: 所要時間
      docUrl           // G列: 回答履歴（Google Doc URL）
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createAnswerDoc(data) {
  const title = `[英語テスト] ${data.name} — ${data.datetime}`;
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  // タイトル
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);

  // サマリー
  body.appendParagraph('テスト結果サマリー').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`氏名: ${data.name}`);
  body.appendParagraph(`受験日時: ${data.datetime}`);
  body.appendParagraph(`スコア: ${data.score} / 10`);
  body.appendParagraph(`CEFRレベル: ${data.cefr}（${data.levelName}）`);
  body.appendParagraph(`所要時間: ${data.elapsed}`);

  // 回答履歴
  body.appendParagraph('回答の詳細').setHeading(DocumentApp.ParagraphHeading.HEADING1);

  data.answers.forEach(a => {
    const mark = a.result === '正解' ? '✓' : '✗';
    body.appendParagraph(`Q${a.no}. [${a.cefr}] ${a.type}  ${mark} ${a.result}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(a.question);
    body.appendParagraph(`あなたの回答: ${a.chosen}`);
    if (a.result === '不正解') {
      body.appendParagraph(`正解: ${a.correct}`);
    }
    body.appendParagraph('');
  });

  doc.saveAndClose();

  // 閲覧リンクを全員に公開
  const file = DriveApp.getFileById(doc.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return doc.getUrl();
}

function doGet(e) {
  return ContentService
    .createTextOutput('English Level Quiz Script is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}
