const NEWCOMER_SHEET_NAME = 'NewComer';
const NEWCOMER_NOTIFICATION_EMAIL = 'newtoguelph@guelphtamilsociety.com';

const NEWCOMER_HEADERS = [
  'Timestamp',
  'Full name',
  'Email address',
  'Phone number',
  'Arrival status',
  'Arrival date',
  'Current city or country',
  'Preferred language',
  'Family status',
  'Areas where help is needed',
  'Additional questions or information',
];

function setupNewComerSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let tab = sheet.getSheetByName(NEWCOMER_SHEET_NAME);
  if (!tab) tab = sheet.insertSheet(NEWCOMER_SHEET_NAME);
  tab.getRange(1, 1, 1, NEWCOMER_HEADERS.length).setValues([NEWCOMER_HEADERS]);
  tab.setFrozenRows(1);
  tab.getRange(1, 1, 1, NEWCOMER_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#8b1324')
    .setFontColor('#fff8e6');
  tab.autoResizeColumns(1, NEWCOMER_HEADERS.length);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = e && e.parameter ? e.parameter : {};
  const required = ['fullName', 'email', 'arrivalStatus', 'preferredLanguage', 'familyStatus'];
  const missing = required.filter((key) => !String(data[key] || '').trim());
  if (missing.length) return jsonResponse({ ok: false, error: 'Missing required fields.' });

  const areas = String(data.helpAreas || '').trim();
  const row = [
    new Date(),
    data.fullName,
    data.email,
    data.phone || '',
    data.arrivalStatus,
    data.arrivalDate || '',
    data.currentCity || '',
    data.preferredLanguage,
    data.familyStatus,
    areas,
    data.additionalInfo || '',
  ];

  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let tab = sheet.getSheetByName(NEWCOMER_SHEET_NAME);
  if (!tab) {
    setupNewComerSheet();
    tab = sheet.getSheetByName(NEWCOMER_SHEET_NAME);
  }
  tab.appendRow(row);

  const subject = 'New GTS newcomer connection request: ' + data.fullName;
  const body = NEWCOMER_HEADERS.map((header, index) => header + ': ' + row[index]).join('\n');
  MailApp.sendEmail(NEWCOMER_NOTIFICATION_EMAIL, subject, body);

  return jsonResponse({ ok: true });
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
