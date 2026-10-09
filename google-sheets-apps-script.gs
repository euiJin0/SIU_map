/**
 * @OnlyCurrentDoc
 * SIU METRO response collector for Google Sheets.
 * Setup: create/open the destination spreadsheet, Extensions > Apps Script,
 * paste this file's code, save, then Deploy > New deployment > Web app.
 * Execute as: Me. Access: Anyone. Put the deployed /exec URL in
 * GOOGLE_SHEETS_WEB_APP_URL in index.html.
 */
function doPost(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('응답') || spreadsheet.insertSheet('응답');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['응답 시각', '역 ID', '역 이름', '질문', '답변 유형', '답변']);
  }

  const response = JSON.parse(e.parameter.payload || '{}');
  sheet.appendRow([
    response.submittedAt || new Date().toISOString(),
    response.stationId || '',
    response.stationTitle || '',
    response.question || '',
    response.answerType || '',
    response.answer || ''
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('SIU METRO response collector is ready.');
}
