// ============================================================
//  COME & SHINE Season 4 — Google Apps Script Web App
//  Paste this entire file into Google Apps Script editor,
//  then Deploy → New deployment → Web App
//  (Execute as: Me | Who has access: Anyone)
// ============================================================

var SHEET_NAME = 'Registrations';

var HEADERS = [
  'Reg ID', 'Name', 'Date of Birth', 'Gender',
  'Parent / Guardian', 'Phone', 'Division', 'Corps',
  'Competition', 'Photo', 'ID Card'
];

// ── GET — only used for nextserial (JSONP) ───────────────────
function doGet(e) {
  var params = e.parameter || {};
  var callback = params.callback || '';

  var result;
  try {
    result = { next: getNextSerial() };
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }

  var json = JSON.stringify(result);
  var output = callback
    ? ContentService.createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(json)
        .setMimeType(ContentService.MimeType.JSON);

  return output;
}

// ── POST — used for saving registration + photos ─────────────
function doPost(e) {
  var result;
  try {
    var payload = JSON.parse(e.postData.contents);
    result = saveRow(payload);
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Get next serial number ───────────────────────────────────
function getNextSerial() {
  var sheet = getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  return lastRow > 0 ? lastRow : 1;
}

// ── Save a registration row ──────────────────────────────────
function saveRow(p) {
  var sheet = getOrCreateSheet();

  var regId       = p.regId       || '';
  var name        = p.name        || '';
  var dob         = p.dob         || '';
  var gender      = p.gender      || '';
  var parent      = p.parent      || '';
  var phone       = p.phone       || '';
  var division    = p.division    || '';
  var corps       = p.corps       || '';
  var competition = p.competition || '';
  var photoB64    = p.photoBase64 || '';
  var idCardB64   = p.idCardBase64 || '';

  // Prevent exact duplicate regId
  if (regId) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === regId) {
        return { status: 'duplicate', regId: regId };
      }
    }
  }

  var newRow = sheet.getLastRow() + 1;

  sheet.appendRow([
    regId, name, dob, gender,
    parent, phone, division, corps,
    competition, '', ''   // cols 10 & 11 filled with images below
  ]);

  // Embed face photo as image in column 10
  if (photoB64) {
    try {
      var photoData  = photoB64.split(',')[1] || photoB64;
      var photoMime  = photoB64.startsWith('data:') ? photoB64.split(';')[0].split(':')[1] : 'image/jpeg';
      var photoBlob  = Utilities.newBlob(Utilities.base64Decode(photoData), photoMime, 'photo_' + regId);
      var photoThumb = DriveApp.createFile(photoBlob);
      photoThumb.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheet.getRange(newRow, 10).setFormula(
        '=IMAGE("https://drive.google.com/uc?id=' + photoThumb.getId() + '")'
      );
      sheet.setRowHeight(newRow, 90);
    } catch(imgErr) {
      sheet.getRange(newRow, 10).setValue('Upload error: ' + imgErr.message);
    }
  }

  // Embed ID card photo as image in column 11
  if (idCardB64) {
    try {
      var idData  = idCardB64.split(',')[1] || idCardB64;
      var idMime  = idCardB64.startsWith('data:') ? idCardB64.split(';')[0].split(':')[1] : 'image/jpeg';
      var idBlob  = Utilities.newBlob(Utilities.base64Decode(idData), idMime, 'idcard_' + regId);
      var idFile  = DriveApp.createFile(idBlob);
      idFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheet.getRange(newRow, 11).setFormula(
        '=IMAGE("https://drive.google.com/uc?id=' + idFile.getId() + '")'
      );
    } catch(idErr) {
      sheet.getRange(newRow, 11).setValue('Upload error: ' + idErr.message);
    }
  }

  return { status: 'ok', regId: regId };
}

// ── Helper: get or create the Registrations sheet ───────────
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);

    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1,  110);  // Reg ID
    sheet.setColumnWidth(2,  180);  // Name
    sheet.setColumnWidth(3,  110);  // DOB
    sheet.setColumnWidth(4,   80);  // Gender
    sheet.setColumnWidth(5,  180);  // Parent
    sheet.setColumnWidth(6,  110);  // Phone
    sheet.setColumnWidth(7,  200);  // Division
    sheet.setColumnWidth(8,  220);  // Corps
    sheet.setColumnWidth(9,  130);  // Competition
    sheet.setColumnWidth(10, 100);  // Photo
    sheet.setColumnWidth(11, 130);  // ID Card
  }

  return sheet;
}
