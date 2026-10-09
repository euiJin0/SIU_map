/**
 * @OnlyCurrentDoc
 * SIU METRO response collector for Google Sheets.
 * Setup: open the destination spreadsheet, Extensions > Apps Script,
 * paste this file's code, save, then deploy a new Web app version.
 * Execute as: Me. Access: Anyone.
 */
const RAW_HEADERS = ['응답시각', '참여자 아이디', '역 이름', '질문', '답변 유형', '답변'];
const SUMMARY_HEADERS = ['참여자 아이디', '응답 시각', '기기 키'];
const SHEET_DATE_FORMAT = 'yyyy-mm-dd hh:mm';

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

  const responseRow = rawSheet.getLastRow() + 1;
  rawSheet.appendRow([
    toSheetDate(submittedAt),
    participantId,
    response.stationTitle || '',
    response.question || '',
    response.answerType || '',
    asPlainText(response.answer || '')
  ]);
  rawSheet.getRange(responseRow, 1).setNumberFormat(SHEET_DATE_FORMAT);
  updateParticipantSummary(summary, response, participantId, submittedAt);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, participantId: participantId }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Keep the compact participant ID and response time visible. "기기 키" stays hidden.
// Each answer column is explicitly labeled so its value cannot be mistaken for a station name.
function ensureSummarySchema(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('참여자별 보기') || spreadsheet.insertSheet('참여자별 보기');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
    sheet.hideColumns(3);
    return sheet;
  }

  const oldValues = sheet.getDataRange().getValues();
  const oldHeaders = oldValues[0].map(String);
  const hasKey = oldHeaders.indexOf('기기 키') !== -1;
  const hasNewId = oldHeaders.indexOf('참여자 아이디') !== -1;
  const answerHeaders = oldHeaders.slice(3);
  const schemaIsCurrent = hasKey && hasNewId &&
    oldHeaders[0] === SUMMARY_HEADERS[0] &&
    oldHeaders[1] === SUMMARY_HEADERS[1] &&
    oldHeaders[2] === SUMMARY_HEADERS[2] &&
    answerHeaders.every(function (header) { return header.indexOf('답변 | ') === 0; });

  if (!schemaIsCurrent) {
    const idIndex = oldHeaders.indexOf('참여자 아이디') !== -1
      ? oldHeaders.indexOf('참여자 아이디')
      : oldHeaders.indexOf('참여자 ID');
    const keyIndex = oldHeaders.indexOf('기기 키');
    const timeIndex = oldHeaders.indexOf('응답 시각');
    const recentIndex = oldHeaders.indexOf('최근 응답');
    const firstIndex = oldHeaders.indexOf('첫 응답');
    const oldAnswerHeaders = oldHeaders.filter(function (header) {
      return ['참여자 ID', '참여자 아이디', '응답 시각', '첫 응답', '최근 응답', '기기 키'].indexOf(header) === -1;
    });
    const newAnswerHeaders = oldAnswerHeaders.map(function (header) {
      return header.indexOf('답변 | ') === 0 ? header : '답변 | ' + header;
    });
    const headers = SUMMARY_HEADERS.concat(newAnswerHeaders);
    const rows = [];

    for (let i = 1; i < oldValues.length; i++) {
      const oldRow = oldValues[i];
      const oldId = idIndex >= 0 ? String(oldRow[idIndex] || '') : '';
      if (!oldId) continue;
      const row = new Array(headers.length).fill('');
      row[0] = formatParticipantId(rows.length + 1);
      row[1] = toSheetDate(
        timeIndex >= 0 ? oldRow[timeIndex] :
          (recentIndex >= 0 && oldRow[recentIndex] ? oldRow[recentIndex] :
            (firstIndex >= 0 ? oldRow[firstIndex] : ''))
      );
      row[2] = keyIndex >= 0 ? oldRow[keyIndex] : oldId;
      oldAnswerHeaders.forEach(function (oldHeader, index) {
        const oldColumn = oldHeaders.indexOf(oldHeader);
        if (oldColumn >= 0) row[SUMMARY_HEADERS.length + index] = oldRow[oldColumn];
      });
      rows.push(row);
    }

    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  if (sheet.getMaxColumns() >= 3) {
    sheet.showColumns(1, 2);
    sheet.hideColumns(3);
  }
  formatSummaryTimes(sheet);
  return sheet;
}

// Reorder legacy raw rows once, preserving their responses and converting timestamps to dates.
function migrateRawSheet(sheet, summary) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, RAW_HEADERS.length).setValues([RAW_HEADERS]);
    return;
  }
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const isCurrent = headers.length === RAW_HEADERS.length &&
    RAW_HEADERS.every(function (header, i) { return headers[i] === header; });
  if (isCurrent) {
    formatRawTimes(sheet);
    return;
  }

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
    const participantId = oldId ? getOrCreateParticipant(summary, oldId, pick(['응답시각', '응답 시각'])) : '';
    rows.push([
      toSheetDate(pick(['응답시각', '응답 시각'])),
      participantId,
      pick(['역 이름']),
      pick(['질문']),
      pick(['답변 유형']),
      pick(['답변'])
    ]);
  }
  sheet.clearContents();
  sheet.getRange(1, 1, 1, RAW_HEADERS.length).setValues([RAW_HEADERS]);
  if (rows.length) sheet.getRange(2, 1, rows.length, RAW_HEADERS.length).setValues(rows);
  formatRawTimes(sheet);
}

function getOrCreateParticipant(summary, deviceKey, submittedAt) {
  const key = String(deviceKey || 'ID 없음');
  const lastRow = summary.getLastRow();
  if (lastRow >= 2) {
    const values = summary.getRange(2, 1, lastRow - 1, Math.max(summary.getLastColumn(), 3)).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === key || String(values[i][2]) === key) return String(values[i][0]);
    }
  }
  const participantId = formatParticipantId(nextParticipantNumber(summary));
  const row = summary.getLastRow() + 1;
  summary.getRange(row, 1, 1, SUMMARY_HEADERS.length).setValues([
    [participantId, toSheetDate(submittedAt), key]
  ]);
  summary.getRange(row, 2).setNumberFormat(SHEET_DATE_FORMAT);
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
  const answerColumn = '답변 | ' + (response.stationTitle || stationId) + ' (' + stationId + ')';
  const headers = summary.getRange(1, 1, 1, summary.getLastColumn()).getValues()[0].map(String);
  let answerIndex = headers.indexOf(answerColumn);
  if (answerIndex < 0) {
    answerIndex = headers.length;
    summary.getRange(1, answerIndex + 1).setValue(answerColumn);
  }

  const lastRow = summary.getLastRow();
  const ids = lastRow >= 2 ? summary.getRange(2, 1, lastRow - 1, 1).getDisplayValues() : [];
  let rowNumber = 0;
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === participantId) { rowNumber = i + 2; break; }
  }
  if (!rowNumber) return;
  const responseTimeCell = summary.getRange(rowNumber, 2);
  const existingTime = responseTimeCell.getValue();
  if (!existingTime || toMillis(submittedAt) >= toMillis(existingTime)) {
    responseTimeCell.setValue(toSheetDate(submittedAt));
  }
  responseTimeCell.setNumberFormat(SHEET_DATE_FORMAT);
  summary.getRange(rowNumber, answerIndex + 1).setValue(asPlainText(response.answer || ''));
}

function formatSummaryTimes(sheet) {
  const rows = sheet.getLastRow() - 1;
  if (rows <= 0) return;
  const range = sheet.getRange(2, 2, rows, 1);
  const values = range.getValues().map(function (row) { return [toSheetDate(row[0])]; });
  range.setValues(values);
  range.setNumberFormat(SHEET_DATE_FORMAT);
}

function formatRawTimes(sheet) {
  const rows = sheet.getLastRow() - 1;
  if (rows <= 0) return;
  const range = sheet.getRange(2, 1, rows, 1);
  const values = range.getValues().map(function (row) { return [toSheetDate(row[0])]; });
  range.setValues(values);
  range.setNumberFormat(SHEET_DATE_FORMAT);
}

function toSheetDate(value) {
  if (value instanceof Date) return value;
  if (!value) return '';
  const date = new Date(value);
  return isNaN(date.getTime()) ? String(value) : date;
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
