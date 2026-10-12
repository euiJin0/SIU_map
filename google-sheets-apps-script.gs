/**
 * @OnlyCurrentDoc
 * SIU METRO response collector for Google Sheets.
 * Setup: paste this file into the spreadsheet's Apps Script project,
 * save, then update the existing Web app deployment to a new version.
 * Execute as: Me. Access: Anyone.
 */
const RAW_HEADERS = ['응답시각', '참여자 아이디', '역 이름', '질문', '답변 유형', '답변'];
const SELF_FIELDS = ['초중고 공통 관심사', '듣고 싶은 말', '진로기준', '랜덤뽑기'];
const STATION_FIELDS = [
  '답변 | 관심분야역',
  '답변 | 시니어역',
  '답변 | 환승센터역',
  '답변 | 경영 X 시비즈역',
  '답변 | 문화콘텐츠 X 시비즈역',
  '답변 | 사회복지 X 시비즈역',
  '답변 | 체육 X 시비즈역',
  '답변 | 부동산건설 X 시비즈역',
  '답변 | 교육 X 시비즈역'
];
const SUMMARY_HEADERS = ['참여자 아이디'].concat(SELF_FIELDS, ['응답 시각'], STATION_FIELDS);
const SHEET_DATE_FORMAT = 'yyyy-mm-dd hh:mm';

function doPost(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('대상 스프레드시트를 찾을 수 없습니다. 시트에서 Apps Script를 열어 배포하세요.');
  const response = JSON.parse((e && e.parameter && e.parameter.payload) || '{}');
  const participantId = String(response.participantId || '').trim();
  if (!participantId) throw new Error('참여자 이름을 입력해 주세요.');
  const submittedAt = response.submittedAt || new Date().toISOString();

  const summary = ensureSummarySchema(spreadsheet);
  const rawSheet = spreadsheet.getSheetByName('응답') || spreadsheet.insertSheet('응답');
  getOrCreateParticipant(summary, participantId, submittedAt);
  migrateRawSheet(rawSheet, summary);

  const rowNumber = rawSheet.getLastRow() + 1;
  rawSheet.appendRow([
    toSheetDate(submittedAt),
    participantId,
    response.stationTitle || '',
    response.question || '',
    response.answerType || '',
    asPlainText(response.answer || '')
  ]);
  rawSheet.getRange(rowNumber, 1).setNumberFormat(SHEET_DATE_FORMAT);
  updateParticipantSummary(summary, response, participantId, submittedAt);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, participantId: participantId }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSummarySchema(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('참여자별 보기') || spreadsheet.insertSheet('참여자별 보기');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
    sheet.getRange('F:F').setNumberFormat(SHEET_DATE_FORMAT);
    return sheet;
  }

  const values = sheet.getDataRange().getValues();
  const oldHeaders = values[0].map(String);
  const isCurrent = SUMMARY_HEADERS.every(function (header, index) { return oldHeaders[index] === header; }) &&
    oldHeaders.slice(SUMMARY_HEADERS.length).every(function (header) { return header.indexOf('답변 | ') === 0; });
  if (isCurrent) {
    formatSummaryTimes(sheet);
    return sheet;
  }

  const idIndex = oldHeaders.indexOf('참여자 아이디') >= 0 ? oldHeaders.indexOf('참여자 아이디') : oldHeaders.indexOf('참여자 ID');
  const keyIndex = oldHeaders.indexOf('기기 키');
  const responseTimeIndex = oldHeaders.indexOf('응답 시각');
  const recentIndex = oldHeaders.indexOf('최근 응답');
  const firstIndex = oldHeaders.indexOf('첫 응답');
  const oldAnswerHeaders = oldHeaders.filter(function (header) {
    return ['참여자 ID', '참여자 아이디', '응답 시각', '최근 응답', '첫 응답', '기기 키'].indexOf(header) === -1;
  });
  const extraHeaders = [];
  const oldHeaderMap = {};
  oldAnswerHeaders.forEach(function (oldHeader) {
    const legacyStation = mapLegacyStationHeader(oldHeader);
    const mappedHeader = legacyStation === '__DROP__' ? '' : (legacyStation || normalizeLegacyAnswerHeader(oldHeader));
    if (SUMMARY_HEADERS.indexOf(mappedHeader) >= 0) {
      oldHeaderMap[oldHeader] = mappedHeader;
    } else if (mappedHeader && STATION_FIELDS.indexOf(mappedHeader) >= 0) {
      oldHeaderMap[oldHeader] = mappedHeader;
    } else if (mappedHeader && extraHeaders.indexOf(mappedHeader) < 0) {
      extraHeaders.push(mappedHeader);
      oldHeaderMap[oldHeader] = mappedHeader;
    }
  });
  const headers = SUMMARY_HEADERS.concat(extraHeaders);
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const oldRow = values[r];
    const oldId = idIndex >= 0 ? String(oldRow[idIndex] || '').trim() : '';
    if (!oldId) continue;
    const row = new Array(headers.length).fill('');
    row[0] = oldId;
    SELF_FIELDS.forEach(function (field) {
      const oldColumn = oldHeaders.indexOf(field);
      if (oldColumn >= 0) row[1 + SELF_FIELDS.indexOf(field)] = oldRow[oldColumn];
    });
    const time = responseTimeIndex >= 0 ? oldRow[responseTimeIndex] :
      (recentIndex >= 0 && oldRow[recentIndex] ? oldRow[recentIndex] : (firstIndex >= 0 ? oldRow[firstIndex] : ''));
    row[1 + SELF_FIELDS.length] = toSheetDate(time);
    oldAnswerHeaders.forEach(function (oldHeader) {
      const oldColumn = oldHeaders.indexOf(oldHeader);
      const newColumn = headers.indexOf(oldHeaderMap[oldHeader]);
      if (oldColumn >= 0 && newColumn >= 0 && newColumn !== 0) row[newColumn] = oldRow[oldColumn];
    });
    rows.push(row);
  }

  sheet.clearContents();
  if (keyIndex >= 0 && keyIndex + 1 <= sheet.getMaxColumns()) {
    sheet.showColumns(keyIndex + 1);
    sheet.deleteColumn(keyIndex + 1);
  }
  sheet.showColumns(1, sheet.getMaxColumns());
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  formatSummaryTimes(sheet);
  return sheet;
}

function mapLegacyStationHeader(header) {
  const aliases = [
    [['(interest)', '관심분야역'], '답변 | 관심분야역'],
    [['(career)', '진로발견역', '자기탐색역'], '__DROP__'],
    [['(senior)', '시니어역'], '답변 | 시니어역'],
    [['(transfer)', '환승역', '환승센터역'], '답변 | 환승센터역'],
    [['(business)', '경영역', '경영 X 시비즈역'], '답변 | 경영 X 시비즈역'],
    [['(culture)', '문화콘텐츠역', '문화콘텐츠 X 시비즈역'], '답변 | 문화콘텐츠 X 시비즈역'],
    [['(welfare)', '사회복지역', '사회복지 X 시비즈역'], '답변 | 사회복지 X 시비즈역'],
    [['(sports)', '체육역', '체육 X 시비즈역'], '답변 | 체육 X 시비즈역'],
    [['(realestate)', '부동산건설역', '부동산건설 X 시비즈역'], '답변 | 부동산건설 X 시비즈역'],
    [['(education)', '교육역', '교육 X 시비즈역'], '답변 | 교육 X 시비즈역']
  ];
  for (let i = 0; i < aliases.length; i++) {
    if (aliases[i][0].some(function (alias) { return header.indexOf(alias) >= 0; })) return aliases[i][1];
  }
  return '';
}

function normalizeLegacyAnswerHeader(header) {
  if (header.indexOf('답변 | ') === 0) return header;
  if (/^(연결 테스트|Codex)/i.test(header)) return '';
  return header ? '답변 | ' + header : '';
}

function migrateRawSheet(sheet, summary) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, RAW_HEADERS.length).setValues([RAW_HEADERS]);
    return;
  }
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const isCurrent = headers.length === RAW_HEADERS.length &&
    RAW_HEADERS.every(function (header, index) { return headers[index] === header; });
  if (isCurrent) {
    formatRawTimes(sheet);
    return;
  }

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const oldRow = values[r];
    const pick = function (names) {
      for (let i = 0; i < names.length; i++) {
        const column = headers.indexOf(names[i]);
        if (column >= 0) return oldRow[column];
      }
      return '';
    };
    const oldId = String(pick(['참여자 아이디', '참여자 ID']) || '').trim();
    if (oldId && oldId !== 'ID 없음') getOrCreateParticipant(summary, oldId, pick(['응답시각', '응답 시각']));
    rows.push([
      toSheetDate(pick(['응답시각', '응답 시각'])),
      oldId,
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

function getOrCreateParticipant(sheet, participantId, submittedAt) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === participantId) return i + 2;
    }
  }
  const row = new Array(Math.max(sheet.getLastColumn(), SUMMARY_HEADERS.length)).fill('');
  row[0] = participantId;
  row[SUMMARY_HEADERS.indexOf('응답 시각')] = toSheetDate(submittedAt);
  const rowNumber = sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  sheet.getRange(rowNumber, SUMMARY_HEADERS.indexOf('응답 시각') + 1).setNumberFormat(SHEET_DATE_FORMAT);
  return rowNumber;
}

function updateParticipantSummary(sheet, response, participantId, submittedAt) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const participantColumn = headers.indexOf('참여자 아이디') + 1;
  const timeColumn = headers.indexOf('응답 시각') + 1;
  const ids = sheet.getLastRow() >= 2 ? sheet.getRange(2, participantColumn, sheet.getLastRow() - 1, 1).getDisplayValues() : [];
  let rowNumber = 0;
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === participantId) { rowNumber = i + 2; break; }
  if (!rowNumber) return;

  const oldTime = sheet.getRange(rowNumber, timeColumn).getValue();
  if (!oldTime || Date.parse(submittedAt) >= (oldTime instanceof Date ? oldTime.getTime() : Date.parse(String(oldTime)))) {
    sheet.getRange(rowNumber, timeColumn).setValue(toSheetDate(submittedAt));
  }
  sheet.getRange(rowNumber, timeColumn).setNumberFormat(SHEET_DATE_FORMAT);

  const answerColumn = response.summaryColumn || stationSummaryColumn(response.stationId);
  const answerColumnIndex = headers.indexOf(answerColumn);
  if (answerColumnIndex >= 0 && answerColumnIndex !== participantColumn - 1 && answerColumnIndex !== timeColumn - 1) {
    sheet.getRange(rowNumber, answerColumnIndex + 1).setValue(asPlainText(response.answer || ''));
  }
}

function stationSummaryColumn(stationId) {
  const columns = {
    interest: '답변 | 관심분야역',
    senior: '답변 | 시니어역',
    transfer: '답변 | 환승센터역',
    business: '답변 | 경영 X 시비즈역',
    culture: '답변 | 문화콘텐츠 X 시비즈역',
    welfare: '답변 | 사회복지 X 시비즈역',
    sports: '답변 | 체육 X 시비즈역',
    realestate: '답변 | 부동산건설 X 시비즈역',
    education: '답변 | 교육 X 시비즈역'
  };
  return columns[stationId] || '';
}

function formatSummaryTimes(sheet) {
  const rowCount = sheet.getLastRow() - 1;
  if (rowCount <= 0) return;
  const column = SUMMARY_HEADERS.indexOf('응답 시각') + 1;
  const range = sheet.getRange(2, column, rowCount, 1);
  range.setValues(range.getValues().map(function (row) { return [toSheetDate(row[0])]; }));
  range.setNumberFormat(SHEET_DATE_FORMAT);
}

function formatRawTimes(sheet) {
  const rowCount = sheet.getLastRow() - 1;
  if (rowCount <= 0) return;
  const range = sheet.getRange(2, 1, rowCount, 1);
  range.setValues(range.getValues().map(function (row) { return [toSheetDate(row[0])]; }));
  range.setNumberFormat(SHEET_DATE_FORMAT);
}

function toSheetDate(value) {
  if (value instanceof Date) return value;
  if (!value) return '';
  const date = new Date(value);
  return isNaN(date.getTime()) ? String(value) : date;
}

function asPlainText(value) {
  const text = String(value);
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function doGet() {
  return ContentService.createTextOutput('SIU METRO response collector is ready.');
}
