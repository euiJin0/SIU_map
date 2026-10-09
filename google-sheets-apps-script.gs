/**
 * @OnlyCurrentDoc
 * SIU METRO response collector for Google Sheets.
 * Setup: open the destination spreadsheet, Extensions > Apps Script,
 * paste this file's code, save, then deploy a new Web app version.
 * Execute as: Me. Access: Anyone.
 */
function doPost(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('대상 스프레드시트를 찾을 수 없습니다. 시트에서 Apps Script를 열어 배포하세요.');

  const response = JSON.parse(e.parameter.payload || '{}');
  const participantId = response.participantId || 'ID 없음';
  const submittedAt = response.submittedAt || new Date().toISOString();
  const rawSheet = spreadsheet.getSheetByName('응답') || spreadsheet.insertSheet('응답');
  const rawHeaders = ensureHeaders(rawSheet, ['응답 시각', '역 ID', '역 이름', '질문', '답변 유형', '답변', '참여자 ID']);

  // Keep the existing six response columns intact; the participant ID is added at the end.
  const rawRow = new Array(rawHeaders.length).fill('');
  setByHeader(rawRow, rawHeaders, '응답 시각', submittedAt);
  setByHeader(rawRow, rawHeaders, '역 ID', response.stationId || '');
  setByHeader(rawRow, rawHeaders, '역 이름', response.stationTitle || '');
  setByHeader(rawRow, rawHeaders, '질문', response.question || '');
  setByHeader(rawRow, rawHeaders, '답변 유형', response.answerType || '');
  setByHeader(rawRow, rawHeaders, '답변', asPlainText(response.answer || ''));
  setByHeader(rawRow, rawHeaders, '참여자 ID', participantId);
  rawSheet.appendRow(rawRow);

  updateParticipantSummary(spreadsheet, response, participantId, submittedAt);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, participantId: participantId }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureHeaders(sheet, requiredHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    return requiredHeaders.slice();
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  requiredHeaders.forEach(function (header) {
    if (headers.indexOf(header) === -1) {
      headers.push(header);
      sheet.getRange(1, headers.length).setValue(header);
    }
  });
  return headers;
}

function setByHeader(row, headers, header, value) {
  const column = headers.indexOf(header);
  if (column !== -1) row[column] = value;
}

function updateParticipantSummary(spreadsheet, response, participantId, submittedAt) {
  const summary = spreadsheet.getSheetByName('참여자별 보기') || spreadsheet.insertSheet('참여자별 보기');
  const stationId = response.stationId || 'other';
  const stationTitle = response.stationTitle || stationId;
  const stationColumn = stationTitle + ' (' + stationId + ')';
  const headers = ensureHeaders(summary, ['참여자 ID', '첫 응답', '최근 응답', stationColumn]);

  let rowNumber = 0;
  const lastRow = summary.getLastRow();
  if (lastRow >= 2) {
    const ids = summary.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === participantId) {
        rowNumber = i + 2;
        break;
      }
    }
  }

  if (rowNumber === 0) {
    rowNumber = summary.getLastRow() + 1;
    summary.getRange(rowNumber, 1, 1, headers.length).setValues([new Array(headers.length).fill('')]);
    summary.getRange(rowNumber, headers.indexOf('참여자 ID') + 1).setValue(participantId);
    summary.getRange(rowNumber, headers.indexOf('첫 응답') + 1).setValue(submittedAt);
  } else {
    const firstColumn = headers.indexOf('첫 응답') + 1;
    const existingFirst = summary.getRange(rowNumber, firstColumn).getValue();
    if (!existingFirst || toMillis(submittedAt) < toMillis(existingFirst)) {
      summary.getRange(rowNumber, firstColumn).setValue(submittedAt);
    }
  }

  const lastColumn = headers.indexOf('최근 응답') + 1;
  const existingLast = summary.getRange(rowNumber, lastColumn).getValue();
  if (!existingLast || toMillis(submittedAt) >= toMillis(existingLast)) {
    summary.getRange(rowNumber, lastColumn).setValue(submittedAt);
  }
  summary.getRange(rowNumber, headers.indexOf(stationColumn) + 1).setValue(asPlainText(response.answer || ''));
}

function toMillis(value) {
  const millis = value instanceof Date ? value.getTime() : Date.parse(String(value));
  return isNaN(millis) ? 0 : millis;
}

function asPlainText(value) {
  const text = String(value);
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function doGet() {
  return ContentService.createTextOutput('SIU METRO response collector is ready.');
}
