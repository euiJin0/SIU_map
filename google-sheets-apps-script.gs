/**
 * @OnlyCurrentDoc
 * SIU METRO response collector for Google Sheets.
 * Setup: open the destination spreadsheet, Extensions > Apps Script,
 * paste this file's code, save, then deploy a new Web app version.
 * Execute as: Me. Access: Anyone.
 */
const RAW_HEADERS = ['응답시각', '참여자 아이디', '역 이름', '질문', '답변 유형', '답변'];
const SUMMARY_HEADERS = ['참여자 아이디', '기기 키', '첫 응답', '최근 응답'];

function doPost(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('대상 스프레드시트를 찾을 수 없습니다. 시트에서 Apps Script를 열어 배포하세요.');
  const response = JSON.parse((e && e.parameter && e.parameter.payload) || '{}');
  const deviceKey = String(response.participantId || 'ID 없음');
  const submittedAt = response.submittedAt || new Date().toISOString();
  const summary = ensureSummarySchema(spreadsheet);
  const rawSheet = spreadsheet.getSheetByName('응답') || spreadsheet.insertSheet('응답');
  const participantId = getOrCreateParticipant(summary, deviceKey, submittedAt);
  migrateRawSheet(rawSheet, summary);

  rawSheet.appendRow([
    submittedAt,
    participantId,
    response.stationTitle || '',
    response.question || '',
    response.answerType || '',
    asPlainText(response.answer || '')
  ]);
  updateParticipantSummary(summary, response, participantId, submittedAt);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, participantId: participantId }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Existing summary rows are preserved. Old long IDs become short P001-style IDs;
// the original browser key remains in the hidden "기기 키" column for matching.
function ensureSummarySchema(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('참여자별 보기') || spreadsheet.insertSheet('참여자별 보기');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
    sheet.hideColumns(2);
    return sheet;
  }

  const oldValues = sheet.getDataRange().getValues();
  const oldHeaders = oldValues[0].map(String);
  const hasKey = oldHeaders.indexOf('기기 키') !== -1;
  const hasNewId = oldHeaders.indexOf('참여자 아이디') !== -1;
  if (hasKey && hasNewId && oldHeaders[0] === '참여자 아이디') return sheet;

  const stationHeaders = oldHeaders.filter(function (header) {
    return ['참여자 ID', '참여자 아이디', '첫 응답', '최근 응답', '기기 키'].indexOf(header) === -1;
  });
  const headers = SUMMARY_HEADERS.concat(stationHeaders);
  const rows = [];
  for (let i = 1; i < oldValues.length; i++) {
    const oldRow = oldValues[i];
    const oldIdIndex = oldHeaders.indexOf('참여자 아이디') !== -1 ? oldHeaders.indexOf('참여자 아이디') : oldHeaders.indexOf('참여자 ID');
    const oldId = oldIdIndex >= 0 ? String(oldRow[oldIdIndex] || '') : '';
    if (!oldId) continue;
    const compactId = formatParticipantId(rows.length + 1);
    const row = new Array(headers.length).fill('');
    row[0] = compactId;
    row[1] = hasKey ? oldRow[oldHeaders.indexOf('기기 키')] : oldId;
    const firstIndex = oldHeaders.indexOf('첫 응답');
    const lastIndex = oldHeaders.indexOf('최근 응답');
    if (firstIndex >= 0) row[2] = oldRow[firstIndex];
    if (lastIndex >= 0) row[3] = oldRow[lastIndex];
    stationHeaders.forEach(function (header, index) {
      const oldIndex = oldHeaders.indexOf(header);
      if (oldIndex >= 0) row[SUMMARY_HEADERS.length + index] = oldRow[oldIndex];
    });
    rows.push(row);
  }
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.hideColumns(2);
  return sheet;
}

// Reorders the existing response history once and preserves every available field.
function migrateRawSheet(sheet, summary) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, RAW_HEADERS.length).setValues([RAW_HEADERS]);
    return;
  }
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  if (RAW_HEADERS.every(function (header, i) { return headers[i] === header; }) && headers.length === RAW_HEADERS.length) return;
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const oldRow = values[i];
    const pick = function (names) {
      for (let n = 0; n < names.length; n++) {
        const index = headers.indexOf(names[n]);
        if (index >= 0) return oldRow[index];
      }
      return '';
    };
    const oldId = String(pick(['참여자 아이디', '참여자 ID']) || '');
    const compactId = oldId ? getOrCreateParticipant(summary, oldId, pick(['응답시각', '응답 시각'])) : '';
    rows.push([
      pick(['응답시각', '응답 시각']), compactId,
      pick(['역 이름']), pick(['질문']), pick(['답변 유형']), pick(['답변'])
    ]);
  }
  sheet.clearContents();
  sheet.getRange(1, 1, 1, RAW_HEADERS.length).setValues([RAW_HEADERS]);
  if (rows.length) sheet.getRange(2, 1, rows.length, RAW_HEADERS.length).setValues(rows);
}

function getOrCreateParticipant(summary, deviceKey, submittedAt) {
  const key = String(deviceKey || 'ID 없음');
  const lastRow = summary.getLastRow();
  if (lastRow >= 2) {
    const values = summary.getRange(2, 1, lastRow - 1, 4).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][1]) === key) return String(values[i][0]);
    }
  }
  const participantId = formatParticipantId(nextParticipantNumber(summary));
  const row = summary.getLastRow() + 1;
  summary.getRange(row, 1, 1, 4).setValues([[participantId, key, submittedAt || '', submittedAt || '']]);
  return participantId;
}

function nextParticipantNumber(summary) {
  const lastRow = summary.getLastRow();
  if (lastRow < 2) return 1;
  const ids = summary.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(function (row) { return row[0]; });
  return ids.reduce(function (next, id) {
    const match = /^P(\d+)$/.exec(String(id));
    return match ? Math.max(next, Number(match[1]) + 1) : next;
  }, 1);
}

function formatParticipantId(number) {
  return 'P' + String(number).padStart(3, '0');
}

function updateParticipantSummary(summary, response, participantId, submittedAt) {
  const stationId = response.stationId || 'other';
  const stationColumn = (response.stationTitle || stationId) + ' (' + stationId + ')';
  const lastColumn = summary.getLastColumn();
  const headers = summary.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  let stationIndex = headers.indexOf(stationColumn);
  if (stationIndex < 0) {
    stationIndex = headers.length;
    summary.getRange(1, stationIndex + 1).setValue(stationColumn);
  }
  const lastRow = summary.getLastRow();
  const ids = lastRow >= 2 ? summary.getRange(2, 1, lastRow - 1, 1).getDisplayValues() : [];
  let rowNumber = 0;
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === participantId) { rowNumber = i + 2; break; }
  }
  if (!rowNumber) return;
  const firstCell = summary.getRange(rowNumber, 3);
  const lastCell = summary.getRange(rowNumber, 4);
  const first = firstCell.getValue();
  if (!first || toMillis(submittedAt) < toMillis(first)) firstCell.setValue(submittedAt);
  const recent = lastCell.getValue();
  if (!recent || toMillis(submittedAt) >= toMillis(recent)) lastCell.setValue(submittedAt);
  summary.getRange(rowNumber, stationIndex + 1).setValue(asPlainText(response.answer || ''));
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
