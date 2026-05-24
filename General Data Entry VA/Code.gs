/**
 * Data Entry with Script Form - Google Sheets backend
 * Bind this script to the spreadsheet named "Data Entry with Script Form"
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 */

var SPREADSHEET_NAME = 'Data Entry with Script Form';

var HEADERS = [
  'First Name',
  'Last Name',
  'Full Name',
  'Email Address',
  'PhoneNumber',
  'Mobile Number',
  'Date Of Birth',
  'Gender',
  'Company Name',
  'Job Title',
  'Department',
  'Website URL',
  'LinkedIn Profile',
  'Street Address',
  'City',
  'State / Province',
  'ZIP / Postal Code',
  'Country',
  'Customer ID',
  'Employee ID',
  'Order Number',
  'Invoice Number',
  'Database Entry ID',
  'Product Name',
  'Product SKU',
  'Quantity',
  'Price',
  'Payment Status',
  'Shipping Status',
  'Appointment Date',
  'Meeting Time',
  'Account Username',
  'Registration Date',
  'Subscription Status',
  'Tags / Categories',
  'Social Media Links',
  'Notes / Comments'
];

var HEADER_TO_KEY = {
  'First Name': 'firstName',
  'Last Name': 'lastName',
  'Full Name': 'fullName',
  'Email Address': 'email',
  'PhoneNumber': 'phone',
  'Mobile Number': 'mobile',
  'Date Of Birth': 'dob',
  'Gender': 'gender',
  'Company Name': 'companyName',
  'Job Title': 'jobTitle',
  'Department': 'department',
  'Website URL': 'website',
  'LinkedIn Profile': 'linkedin',
  'Street Address': 'streetAddress',
  'City': 'city',
  'State / Province': 'state',
  'ZIP / Postal Code': 'zip',
  'Country': 'country',
  'Customer ID': 'customerId',
  'Employee ID': 'employeeId',
  'Order Number': 'orderNumber',
  'Invoice Number': 'invoiceNumber',
  'Database Entry ID': 'databaseEntryId',
  'Product Name': 'productName',
  'Product SKU': 'productSku',
  'Quantity': 'quantity',
  'Price': 'price',
  'Payment Status': 'paymentStatus',
  'Shipping Status': 'shippingStatus',
  'Appointment Date': 'appointmentDate',
  'Meeting Time': 'meetingTime',
  'Account Username': 'accountUsername',
  'Registration Date': 'registrationDate',
  'Subscription Status': 'subscriptionStatus',
  'Tags / Categories': 'tags',
  'Social Media Links': 'socialLinks',
  'Notes / Comments': 'notes'
};

var DATE_KEYS = ['dob', 'appointmentDate', 'registrationDate'];
var TIME_KEYS = ['meetingTime'];
var NUMBER_KEYS = ['quantity', 'price'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Data Entry Pro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Data Entry')
    .addItem('Open Form', 'openFormSidebar')
    .addToUi();
  ensureSheetReady_();
}

function openFormSidebar() {
  var html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif;padding:12px;">' +
    'Use <b>Deploy &gt; New deployment &gt; Web app</b> to open the full data entry form.</p>'
  ).setWidth(280).setHeight(100);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active && active.getName() === SPREADSHEET_NAME) {
    return active;
  }
  return SpreadsheetApp.openByName(SPREADSHEET_NAME);
}

function getSheet_() {
  var ss = getSpreadsheet_();
  if (!ss) {
    throw new Error('Spreadsheet "' + SPREADSHEET_NAME + '" was not found.');
  }
  var sheet = ss.getSheets()[0];
  ensureHeaders_(sheet);
  return sheet;
}

function ensureSheetReady_() {
  try {
    getSheet_();
  } catch (e) {
    Logger.log(e.message);
  }
}

function ensureHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  var existing = lastCol > 0
    ? sheet.getRange(1, 1, 1, Math.max(lastCol, HEADERS.length)).getValues()[0]
    : [];

  var needsWrite = existing.length < HEADERS.length;
  if (!needsWrite) {
    for (var i = 0; i < HEADERS.length; i++) {
      if (String(existing[i] || '').trim() !== HEADERS[i]) {
        needsWrite = true;
        break;
      }
    }
  }

  if (needsWrite) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#e8eef7');
    sheet.setFrozenRows(1);
  }
}

function getRecords() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow, HEADERS.length).getValues();
  var records = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (isRowEmpty_(row)) {
      continue;
    }
    var record = rowToRecord_(row, i + 2);
    records.push(record);
  }

  records.sort(function(a, b) {
    return b._rowIndex - a._rowIndex;
  });

  return records;
}

function saveRecord(data) {
  if (!data) {
    throw new Error('No data received.');
  }

  var sheet = getSheet_();
  var payload = normalizeIncomingRecord_(data);
  validateRequired_(payload);

  var rowValues = recordToRow_(payload);
  var rowIndex = Number(data._rowIndex);
  var savedRowIndex;

  if (rowIndex && rowIndex >= 2 && rowIndex <= sheet.getLastRow()) {
    var existingId = String(sheet.getRange(rowIndex, indexOfHeader_('Database Entry ID') + 1).getValue() || '').trim();
    if (existingId && existingId === String(payload.databaseEntryId || '').trim()) {
      rowValues = recordToRow_(payload);
      sheet.getRange(rowIndex, 1, rowIndex, HEADERS.length).setValues([rowValues]);
      savedRowIndex = rowIndex;
    }
  }

  if (!savedRowIndex) {
    if (!payload.databaseEntryId) {
      payload.databaseEntryId = generateId_();
    }
    rowValues = recordToRow_(payload);
    sheet.appendRow(rowValues);
    savedRowIndex = sheet.getLastRow();
  }

  var saved = rowToRecord_(rowValues, savedRowIndex);
  return saved;
}

function deleteRecordById(databaseEntryId) {
  if (!databaseEntryId) {
    throw new Error('Record ID is required.');
  }

  var sheet = getSheet_();
  var idCol = indexOfHeader_('Database Entry ID') + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('Record not found.');
  }

  var ids = sheet.getRange(2, idCol, lastRow, idCol).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === String(databaseEntryId).trim()) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }

  throw new Error('Record not found.');
}

function rowToRecord_(row, rowIndex) {
  var record = {
    _rowIndex: rowIndex,
    id: '',
    tags: [],
    socialLinks: []
  };

  for (var c = 0; c < HEADERS.length; c++) {
    var header = HEADERS[c];
    var key = HEADER_TO_KEY[header];
    var value = row[c];

    if (key === 'tags') {
      record.tags = parseListField_(value);
      continue;
    }
    if (key === 'socialLinks') {
      record.socialLinks = parseListField_(value);
      continue;
    }

    record[key] = formatOutgoingValue_(value, key);
  }

  record.id = record.databaseEntryId || ('ROW-' + rowIndex);
  return record;
}

function recordToRow_(record) {
  var row = [];

  for (var i = 0; i < HEADERS.length; i++) {
    var header = HEADERS[i];
    var key = HEADER_TO_KEY[header];
    var value = record[key];

    if (key === 'tags') {
      row.push(Array.isArray(value) ? value.join('; ') : String(value || ''));
      continue;
    }
    if (key === 'socialLinks') {
      row.push(Array.isArray(value) ? value.join('; ') : String(value || ''));
      continue;
    }

    row.push(formatIncomingValue_(value, key));
  }

  return row;
}

function normalizeIncomingRecord_(data) {
  var record = {};
  var keys = Object.keys(HEADER_TO_KEY);

  for (var i = 0; i < keys.length; i++) {
    var key = HEADER_TO_KEY[keys[i]];
    if (data[key] !== undefined && data[key] !== null) {
      record[key] = data[key];
    } else {
      record[key] = '';
    }
  }

  if (Array.isArray(data.tags)) {
    record.tags = data.tags;
  } else if (typeof data.tags === 'string' && data.tags) {
    record.tags = parseListField_(data.tags);
  } else {
    record.tags = [];
  }

  if (Array.isArray(data.socialLinks)) {
    record.socialLinks = data.socialLinks;
  } else if (typeof data.socialLinks === 'string' && data.socialLinks) {
    record.socialLinks = parseListField_(data.socialLinks);
  } else {
    record.socialLinks = [];
  }

  record.fullName = [String(record.firstName || '').trim(), String(record.lastName || '').trim()]
    .filter(Boolean)
    .join(' ');

  if (!record.databaseEntryId) {
    record.databaseEntryId = generateId_();
  }

  return record;
}

function validateRequired_(record) {
  if (!String(record.firstName || '').trim()) {
    throw new Error('First Name is required.');
  }
  if (!String(record.lastName || '').trim()) {
    throw new Error('Last Name is required.');
  }
  if (!String(record.email || '').trim()) {
    throw new Error('Email Address is required.');
  }
  var email = String(record.email).trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Email Address is invalid.');
  }
}

function parseListField_(value) {
  if (Array.isArray(value)) {
    return value.map(function(item) { return String(item).trim(); }).filter(Boolean);
  }
  if (value === null || value === undefined || value === '') {
    return [];
  }
  return String(value).split(';').map(function(item) {
    return item.trim();
  }).filter(Boolean);
}

function formatOutgoingValue_(value, key) {
  if (value === null || value === undefined) {
    return '';
  }

  if (value instanceof Date) {
    if (DATE_KEYS.indexOf(key) !== -1) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    if (TIME_KEYS.indexOf(key) !== -1) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
    }
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }

  if (NUMBER_KEYS.indexOf(key) !== -1 && typeof value === 'number') {
    return value;
  }

  return String(value);
}

function formatIncomingValue_(value, key) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  if (NUMBER_KEYS.indexOf(key) !== -1) {
    var num = Number(value);
    return isNaN(num) ? '' : num;
  }

  return String(value);
}

function isRowEmpty_(row) {
  for (var i = 0; i < row.length; i++) {
    if (row[i] !== '' && row[i] !== null && row[i] !== undefined) {
      return false;
    }
  }
  return true;
}

function indexOfHeader_(header) {
  var idx = HEADERS.indexOf(header);
  if (idx === -1) {
    throw new Error('Missing header: ' + header);
  }
  return idx;
}

function generateId_() {
  return 'DB-' + new Date().getTime().toString(36).toUpperCase() + '-' +
    Math.random().toString(36).slice(2, 6).toUpperCase();
}
