// ════════════════════════════════════════════════════════
//  Come & Shine Season 4 — Google Apps Script
//  Deploy as Web App: Execute as Me, Anyone can access
// ════════════════════════════════════════════════════════

var SHEET_ID         = '1IJu9xU8L6Ghyx-gh0Wx-JqH-o3VeJ96BhL1vZMFNGQU';
var GROQ_API_KEY     = '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';
var RZP_KEY_SECRET   = 'UkOrajFO6yYgMHlXzB8UHVk3'; // ← paste your Groq API key
var INDIVIDUAL_SHEET = 'Individual';
var EXCLUSIONS_SHEET = 'Exclusions';
var PAYMENTS_SHEET   = 'Payments';
var TAMBOURINE_SHEET = 'Tambourine';
var DRIVE_FOLDER_ID  = '1C53vENvJDsi48uK1uGZBku76Id8pLyA0';

// ── GET (JSONP) ──
function doGet(e) {
  var params   = (e && e.parameter) ? e.parameter : {};
  var action   = params.action   || 'nextSerial';
  var callback = params.callback || '';
  var result;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
    if (!ss) throw new Error('Could not open spreadsheet.');

    if (action === 'nextSerial') {
      var sheet = getOrCreateSheet(ss, INDIVIDUAL_SHEET, INDIVIDUAL_HEADERS);
      var last  = sheet.getLastRow();
      result    = { next: last <= 1 ? 1 : last };

    } else if (action === 'save') {
      result = saveIndividual(ss, params);

    } else if (action === 'tambourine') {
      result = saveTambourine(ss, params);

    } else if (action === 'receiveSms') {
      result = receiveSms(ss, params);

    } else if (action === 'getExclusions') {
      result = getAllExclusions(ss);

    } else if (action === 'getTambourineById') {
      result = getTambourineById(ss, params.refId || '');

    } else if (action === 'getById') {
      result = getById(ss, params.regId || '');

    } else if (action === 'search') {
      result = searchRegistration(ss, params.q || '');

    } else {
      result = { status: 'ok', action: action };
    }
  } catch(err) {
    result = { status: 'error', message: err.message };
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse(result);
}

// ── POST — accepts both application/json and text/plain (CORS bypass) ──
function doPost(e) {
  var body = {};
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    body = JSON.parse(raw);
  } catch(err) {
    return jsonResponse({ status: 'error', message: 'Invalid JSON: ' + err.message });
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
    var action = body.action;
    if (action === 'ocr')        return jsonResponse(runOcr(body));
    if (action === 'receiveSms')  return jsonResponse(receiveSms(ss, body));
    if (action === 'verifyUtr')   return jsonResponse(verifyUtr(ss, body.utr, body.amount));
    if (action === 'getLatestUtr') return jsonResponse(getLatestUtr(ss, body.amount));
    if (action === 'getExclusions') return jsonResponse(getAllExclusions(ss));
    if (action === 'checkByNameDob') return jsonResponse(checkByNameDob(ss, body));
    if (action === 'getById')        return jsonResponse(getById(ss, body.regId || ''));
    if (action === 'getTambourineById') return jsonResponse(getTambourineById(ss, body.refId || ''));
    if (action === 'save')       return jsonResponse(saveIndividual(ss, body));
    if (action === 'tambourine') return jsonResponse(saveTambourine(ss, body));
    if (action === 'savePhoto')  return jsonResponse(savePhoto(ss, body.regId, body.photo));
    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  } catch(err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── Test manually in editor ──
function testAccess() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('SUCCESS: ' + ss.getName());
  } catch(err) {
    Logger.log('ERROR: ' + err.message);
  }
}

// ── Headers: Reg ID, Name, DOB, Gender, Parent, Phone, Division, Corps, Competition, Photo, Registered At ──
var INDIVIDUAL_HEADERS = [
  'Reference Number', 'Name', 'Date of Birth', 'Gender',
  'Parent / Guardian', 'Phone', 'Division', 'Corps',
  'Competition', 'Photo', 'Timestamp', 'UTR / Transaction ID'
];
var TAMBOURINE_HEADERS = [
  'Reference Number', 'Team Name', 'Division', 'Corps', 'Leader', 'Phone', 'Member Count', 'Competition', 'Timeline',
  'Member1', 'Gender 1', 'Date of Birth 1', 'Photo1',
  'Member2', 'Gender2', 'Date of Birth 2', 'Photo2',
  'Member3', 'Gender3', 'Date of Birth 3', 'Photo3',
  'Member4', 'Gender4', 'Date of Birth 4', 'Photo4',
  'Member5', 'Gender5', 'Date of Birth 5', 'Photo5',
  'Member6', 'Gender6', 'Date of Birth 6', 'Photo6',
  'Member7', 'Gender7', 'Date of Birth 7', 'Photo7',
  'Fees'
];
var TAMBOURINE_MAX_MEMBERS = 7;

// ── Get or create sheet tab ──
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.setFrozenRows(1);
  }
  // Always keep the header row in sync with the current headers definition —
  // prevents column drift when the layout changes but the tab already existed.
  var range = sheet.getRange(1, 1, 1, headers.length);
  var current = range.getValues()[0];
  var needsUpdate = false;
  for (var i = 0; i < headers.length; i++) {
    if (String(current[i] || '') !== String(headers[i])) { needsUpdate = true; break; }
  }
  if (needsUpdate) {
    range.setValues([headers])
      .setFontWeight('bold')
      .setBackground('#1a0e2e')
      .setFontColor('#ffffff');
  }
  return sheet;
}

// ── Save individual registration ──
function saveIndividual(ss, d) {
  if (!d) return { status: 'error', message: 'No data received' };

  var name        = String(d.name        || '');
  var dob         = String(d.dob         || '');
  var gender      = String(d.gender      || '');
  var division    = String(d.division    || '');
  var corps       = String(d.corps       || '');
  var regId       = String(d.regId       || '');
  var parent      = String(d.parent      || '');
  var phone       = String(d.phone       || '');
  var competition = String(d.competition || '');
  var photo       = String(d.photo       || '');
  var utr         = String(d.utr         || '');

  // Log payment ID (Razorpay ID or UTR)
  Logger.log('Payment ID: ' + utr);

  var sheet = getOrCreateSheet(ss, INDIVIDUAL_SHEET, INDIVIDUAL_HEADERS);

  // Duplicate check
  var key  = makeKey([name, dob, gender, division, corps]);
  var rows = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row) continue;
    if (makeKey([row[1]||'', row[2]||'', row[3]||'', row[6]||'', row[7]||'']) === key) {
      return { status: 'duplicate', regId: row[0] };
    }
  }

  var now = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

  // Upload photo to Drive and store lh3 URL
  var photoData = '';
  if (photo && photo.length > 100) {
    try {
      var mimeType = 'image/jpeg';
      var b64Data = photo;
      if (photo.indexOf('data:') === 0 && photo.indexOf(',') > 0) {
        var parts = photo.split(',');
        b64Data = parts[1];
        var mimeMatch = parts[0].match(/:(.*?);/);
        if (mimeMatch) mimeType = mimeMatch[1];
      }
      var ext = mimeType.split('/')[1] || 'jpg';
      var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      var bytes = Utilities.base64Decode(b64Data);
      var blob = Utilities.newBlob(bytes, mimeType, regId + '_photo.' + ext);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoData = 'https://lh3.googleusercontent.com/d/' + file.getId();
    } catch(pe) {
      Logger.log('Photo upload failed: ' + pe.message);
      photoData = '';
    }
  }

  try {
    sheet.appendRow([regId, name, dob, gender, parent, phone, division, corps, competition, photoData, now, utr]);
  } catch(appendErr) {
    return { status: 'error', message: 'Failed to save row: ' + appendErr.message };
  }

  return { status: 'ok', regId: regId, next: sheet.getLastRow() };
}

// ── Save photo as base64 to Sheet row ──
function savePhoto(ss, regId, photo) {
  if (!regId || !photo) return { status: 'error', message: 'Missing regId or photo' };
  try {
    var photoData = photo;
    if (photo && photo.length > 100 && photo.indexOf('data:') === 0) {
      var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      var imgData = photo.split(',');
      var mimeType = imgData[0].match(/:(.*?);/)[1];
      var ext = mimeType.split('/')[1] || 'jpg';
      var bytes = Utilities.base64Decode(imgData[1]);
      var blob = Utilities.newBlob(bytes, mimeType, regId + '_photo.' + ext);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoData = 'https://lh3.googleusercontent.com/d/' + file.getId();
    }
    var sheet = ss.getSheetByName(INDIVIDUAL_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'error', message: 'Sheet not found' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(regId).trim()) {
        sheet.getRange(i + 1, 10).setValue(photoData);
        return { status: 'ok', photoUrl: photoData };
      }
    }
    return { status: 'error', message: 'RegId not found: ' + regId };
  } catch(err) {
    return { status: 'error', message: err.message };
  }
}

// ── Get registration by regId ──
function getById(ss, regId) {
  if (!regId) return { status: 'error', message: 'No regId provided' };
  var sheet = ss.getSheetByName(INDIVIDUAL_SHEET);
  if (!sheet) return { status: 'error', message: 'Sheet not found' };
  var rows = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row) continue;
    if (String(row[0]).trim() === String(regId).trim()) {
      return {
        status: 'ok',
        data: {
          refId:       String(row[0]  || ''),
          name:        String(row[1]  || ''),
          dob:         String(row[2]  || ''),
          gender:      String(row[3]  || ''),
          parent:      String(row[4]  || ''),
          phone:       String(row[5]  || ''),
          division:    String(row[6]  || ''),
          corps:       String(row[7]  || ''),
          competition: String(row[8]  || ''),
          photo:       String(row[9]  || ''),
          registeredAt:String(row[10] || '')
        }
      };
    }
  }
  return { status: 'error', message: 'Registration not found for ID: ' + regId };
}

// ── Save tambourine group ──
function saveTambourine(ss, d) {
  if (!d) return { status: 'error', message: 'No data received' };
  Logger.log('Tambourine Payment ID: ' + (d.utr||''));

  var sheet = getOrCreateSheet(ss, TAMBOURINE_SHEET, TAMBOURINE_HEADERS);
  var now   = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
  var refId = String(d.refId || '');

  // Parse members array
  var members = [];
  try {
    members = typeof d.members === 'string' ? JSON.parse(d.members) : (d.members || []);
  } catch(e) { members = []; }

  // Validate age 16-20 for each member (competition date: 29 Aug 2026)
  var compDate = new Date(2026, 7, 29); // month is 0-indexed: 7 = August
  for (var v = 0; v < members.length; v++) {
    var mDob = members[v] && members[v].dob;
    if (mDob) {
      var parts = mDob.split('/');
      if (parts.length === 3) {
        var bDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        var age = compDate.getFullYear() - bDate.getFullYear();
        var mDiff = compDate.getMonth() - bDate.getMonth();
        if (mDiff < 0 || (mDiff === 0 && compDate.getDate() < bDate.getDate())) age--;
        if (age < 16 || age > 20) {
          return { status: 'error', message: 'Member ' + (v+1) + ' (' + (members[v].name||'') + ') age is ' + age + ' — Tambourine Competition requires age 16 to 20.' };
        }
      }
    }
  }

  // Base row
  var row = [
    refId,
    String(d.teamName    || ''),
    String(d.division    || ''),
    String(d.corps       || ''),
    String(d.leader      || ''),
    String(d.phone       || ''),
    String(d.memberCount || ''),
    'Tambourine',
    now
  ];

  // Expand each member into 4 columns: Name, Gender, DOB, Photo (max 5 members)
  for (var i = 0; i < TAMBOURINE_MAX_MEMBERS; i++) {
    var m = members[i];
    if (m) {
      // Store photo as base64 data URI directly in sheet
      var photoData = (m.photo && m.photo.length > 100) ? m.photo : '';
      row.push(String(m.name   || ''));
      row.push(String(m.gender || ''));
      row.push(String(m.dob    || ''));
      row.push(photoData);
    } else {
      row.push('', '', '', '');
    }
  }

  // Fees — last column: store the UTR / transaction ID as proof of payment
  row.push(String(d.utr || d.fee || ''));

  sheet.appendRow(row);
  return { status: 'ok', refId: refId };
}

// ── Get tambourine group by reference number ──
function getTambourineById(ss, refId) {
  if (!refId) return { status: 'error', message: 'No reference number provided' };
  var sheet = ss.getSheetByName(TAMBOURINE_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return { status: 'error', message: 'No tambourine groups found' };

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row) continue;
    if (String(row[0]).trim() === String(refId).trim()) {
      var members = [];
      for (var m = 0; m < TAMBOURINE_MAX_MEMBERS; m++) {
        var base = 9 + (m * 4); // columns: refId,team,div,corps,leader,phone,count,comp,timeline = 9 cols
        var mName = row[base];
        if (mName) {
          members.push({
            name:   String(row[base]   || ''),
            gender: String(row[base+1] || ''),
            dob:    String(row[base+2] || ''),
            photo:  String(row[base+3] || '')
          });
        }
      }
      var feeCol = 9 + (TAMBOURINE_MAX_MEMBERS * 4); // Fees is the last column
      return {
        status: 'ok',
        data: {
          refId:       String(row[0] || ''),
          teamName:    String(row[1] || ''),
          division:    String(row[2] || ''),
          corps:       String(row[3] || ''),
          leader:      String(row[4] || ''),
          phone:       String(row[5] || ''),
          memberCount: String(row[6] || ''),
          competition: String(row[7] || ''),
          registeredAt:String(row[8] || ''),
          fee:         String(row[feeCol] || ''),
          members:     JSON.stringify(members)
        }
      };
    }
  }
  return { status: 'error', message: 'Group registration not found for reference: ' + refId };
}

// ── Search registration by name, phone or regId ──
function searchRegistration(ss, q) {
  if (!q) return { status: 'error', message: 'No query provided' };
  var sheet = ss.getSheetByName(INDIVIDUAL_SHEET);
  if (!sheet) return { status: 'results', results: [] };
  var rows = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [];
  var ql = q.toLowerCase().trim();
  var results = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row) continue;
    var regId  = String(row[0] || '').toLowerCase();
    var name   = String(row[1] || '').toLowerCase();
    var phone  = String(row[5] || '').toLowerCase();
    if (name.indexOf(ql) > -1 || phone.indexOf(ql) > -1 || regId.indexOf(ql) > -1) {
      results.push({
        regId:       String(row[0] || ''),
        name:        String(row[1] || ''),
        dob:         String(row[2] || ''),
        gender:      String(row[3] || ''),
        parent:      String(row[4] || ''),
        phone:       String(row[5] || ''),
        division:    String(row[6] || ''),
        corps:       String(row[7] || ''),
        competition: String(row[8] || ''),
        registeredAt:String(row[10]|| '')
      });
    }
  }
  return { status: 'ok', results: results };
}

// ── Get all top-10 exclusions from Exclusions sheet ──
function getAllExclusions(ss) {
  var sheet = ss.getSheetByName(EXCLUSIONS_SHEET);
  if (!sheet) {
    // Auto-create with headers
    sheet = ss.insertSheet(EXCLUSIONS_SHEET);
    sheet.appendRow(['Name', 'DOB (DD/MM/YYYY)', 'Exclude Competition Key']);
    sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1a0e2e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    // Add JOANN as first entry
    sheet.appendRow(['JOANN SAM ANDRINA', '09/08/2016', 'drawing2']);
    return { status: 'ok', exclusions: [{name:'JOANN SAM ANDRINA', dob:'09/08/2016', exclude:'drawing2'}] };
  }
  if (sheet.getLastRow() <= 1) return { status: 'ok', exclusions: [] };
  var rows = sheet.getDataRange().getValues();
  var exclusions = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row || !row[0]) continue;
    exclusions.push({
      name:    String(row[0] || '').trim().toUpperCase(),
      dob:     String(row[1] || '').trim(),
      exclude: String(row[2] || '').trim().toLowerCase()
    });
  }
  return { status: 'ok', exclusions: exclusions };
}

// ── Receive SMS from SMS Forwarder app and extract UTR ──
function receiveSms(ss, d) {
  // Open ss if not provided (e.g. called from doGet)
  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID); } catch(e) {}
  }
  if (!ss) return { status: 'error', message: 'Could not open spreadsheet' };

  // Handle various SMS Forwarder app field name formats
  // Handle all SMS forwarder app field formats including {{message}} template
  var smsBody = String(d.sms || d.message || d.body || d.text || d.content || d.msg || d['{{message}}'] || '');
  var from    = String(d.from || d.sender || d.address || d.phone || d['{{sender}}'] || '');
  // Clean up if template vars were not replaced
  if(smsBody==='{{message}}') smsBody='';
  if(from==='{{sender}}') from='';
  if (!smsBody) {
    Logger.log('receiveSms called with data: ' + JSON.stringify(d));
    return { status: 'error', message: 'No SMS body received. Fields: ' + Object.keys(d).join(',') };
  }

  // Extract UTR/transaction reference from SMS
  var utr = extractUtrFromSms(smsBody);
  if (!utr) return { status: 'ok', message: 'No UTR found in SMS', sms: smsBody };

  // Extract amount
  var amount = extractAmountFromSms(smsBody);

  // Store in Payments sheet
  var sheet = getOrCreatePaymentsSheet(ss);
  var now   = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

  // Check for duplicate UTR
  if (sheet.getLastRow() > 1) {
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === utr) {
        return { status: 'ok', message: 'UTR already recorded', utr: utr };
      }
    }
  }

  sheet.appendRow([utr, amount, from, smsBody, now, 'Pending']);
  return { status: 'ok', utr: utr, amount: amount, message: 'UTR recorded successfully' };
}

// ── Verify UTR exists in Payments sheet ──
function verifyUtr(ss, utr, expectedAmount) {
  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID); } catch(e) {}
  }
  if (!ss) return { status: 'error', message: 'Could not open spreadsheet' };
  if (!utr) return { status: 'error', message: 'No UTR provided' };
  utr = String(utr).trim().toUpperCase();

  var sheet = ss.getSheetByName(PAYMENTS_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { status: 'error', message: 'No payments recorded yet. Please try again or contact admin.' };
  }

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowUtr = String(rows[i][0]).trim().toUpperCase();
    if (rowUtr === utr) {
      var amount  = String(rows[i][1] || '');
      var used    = String(rows[i][5] || '');
      if (used === 'Used') {
        return { status: 'error', message: 'This UTR has already been used for another registration.' };
      }
      // Mark as used
      sheet.getRange(i + 1, 6).setValue('Used');
      return { status: 'ok', utr: utr, amount: amount, message: 'Payment verified!' };
    }
  }
  return { status: 'error', message: 'UTR not found. Please check the transaction ID or wait a few minutes for SMS to arrive.' };
}

// ── Extract UTR from bank SMS ──
function extractUtrFromSms(sms) {
  if (!sms || typeof sms !== 'string') return null;
  var s = sms.trim();
  if (!s) return null;

  // Common UTR/Ref patterns in Indian bank SMS (ICICI, SBI, HDFC, Axis etc.)
  // Note: - must be first or last in character class to avoid range errors
  var patterns = [
    /UTR[\s:#No.-]*([A-Z0-9]{10,22})/i,
    /UPI\s*Ref[\s:#No.-]*([A-Z0-9]{10,22})/i,
    /Ref\s*No[\s:.#-]*([A-Z0-9]{10,22})/i,
    /Ref\s*Num[\s:.#-]*([A-Z0-9]{10,22})/i,
    /Transaction\s*ID[\s:#-]*([A-Z0-9]{10,22})/i,
    /Txn\s*ID[\s:#-]*([A-Z0-9]{10,22})/i,
    /Txn\s*Ref[\s:#-]*([A-Z0-9]{10,22})/i,
    /Ref[\s:#-]+([A-Z0-9]{10,22})/i,
    /([0-9]{12})/
  ];

  for (var i = 0; i < patterns.length; i++) {
    try {
      var m = s.match(patterns[i]);
      if (m && m[1] && m[1].trim().length >= 10) {
        return m[1].trim().toUpperCase();
      }
    } catch(e) { continue; }
  }
  return null;
}

// ── Extract amount from bank SMS ──
function extractAmountFromSms(sms) {
  if (!sms || typeof sms !== 'string') return '';
  try {
    var m = sms.match(/(?:Rs\.?|INR|INR\s*Rs\.?|₹)\s*([\d,]+(?:\.\d{1,2})?)/i);
    return m ? m[1].replace(/,/g, '') : '';
  } catch(e) { return ''; }
}

// ── Get or create Payments sheet ──
function getOrCreatePaymentsSheet(ss) {
  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID); } catch(e) {}
  }
  if (!ss) throw new Error('Could not open spreadsheet for Payments sheet');
  var sheet = ss.getSheetByName(PAYMENTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PAYMENTS_SHEET);
    sheet.appendRow(['UTR / Transaction ID', 'Amount (₹)', 'From', 'SMS Message', 'Received At', 'Status']);
    sheet.getRange(1, 1, 1, 6)
      .setFontWeight('bold')
      .setBackground('#1a0e2e')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(4, 400);
  }
  return sheet;
}

// ── Get latest unverified UTR matching amount (for auto-polling) ──
function getLatestUtr(ss, amount) {
  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID); } catch(e) {}
  }
  var sheet = ss ? ss.getSheetByName(PAYMENTS_SHEET) : null;
  if (!sheet || sheet.getLastRow() <= 1) return { status: 'waiting' };

  var rows = sheet.getDataRange().getValues();
  var fiveMinAgo = new Date(Date.now() - 5 * 60 * 1000);

  // Find most recent Pending UTR matching the amount, received in last 5 minutes
  for (var i = rows.length - 1; i >= 1; i--) {
    var rowUtr    = String(rows[i][0] || '').trim();
    var rowAmount = String(rows[i][1] || '').trim();
    var rowStatus = String(rows[i][5] || '').trim();
    var rowTime   = rows[i][4];

    if (rowUtr && rowStatus === 'Pending') {
      // Check amount matches (if provided)
      if (amount && rowAmount && rowAmount !== amount) continue;
      // Check received recently
      var receivedAt = new Date(rowTime);
      if (!isNaN(receivedAt) && receivedAt < fiveMinAgo) continue;
      return { status: 'ok', utr: rowUtr, amount: rowAmount };
    }
  }
  return { status: 'waiting' };
}

// ── Check duplicate by name + dob only (panel 1 early check) ──
function checkByNameDob(ss, d) {
  if (!d || !d.name) return { status: 'ok' };
  var name = String(d.name || '').trim().toLowerCase().replace(/\s+/g,'');
  var dob  = String(d.dob  || '').trim().replace(/[^0-9]/g,'');
  var sheet = ss.getSheetByName(INDIVIDUAL_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return { status: 'ok' };
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row) continue;
    var rName = String(row[1]||'').trim().toLowerCase().replace(/\s+/g,'');
    var rDob  = String(row[2]||'').trim().replace(/[^0-9]/g,'');
    if (rName === name && rDob === dob) {
      return { status: 'duplicate', regId: String(row[0]), name: String(row[1]) };
    }
  }
  return { status: 'ok' };
}

// ── OCR via Groq API ──
// ── Extract DOB from raw OCR text using regex patterns ──
function enhanceDobFromRawText(txt) {
  if (!txt) return txt;
  try {
    // Try to parse existing JSON
    var parsed = null;
    var m = txt.match(/[{][\s\S]*[}]/);
    if (m) {
      try { parsed = JSON.parse(m[0]); } catch(e) { parsed = null; }
    }

    // If DOB already found, return as-is
    if (parsed && parsed.date_of_birth && parsed.date_of_birth !== 'null') return txt;

    // Scan raw text for date patterns
    var dob = extractDateFromText(txt);
    if (!dob) return txt;

    // Inject DOB into result
    if (parsed) {
      parsed.date_of_birth = dob;
      return JSON.stringify(parsed);
    } else {
      // Append to raw text
      return txt + '\n{"date_of_birth":"' + dob + '"}';
    }
  } catch(e) {
    return txt;
  }
}

// ── Find any date in text and normalise to DD/MM/YYYY ──
function extractDateFromText(txt) {
  if (!txt) return null;

  // Remove Issue Date segments so they are never mistaken for DOB
  var cleaned = txt.replace(/Issue\s*Date\s*:?\s*\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}/gi, '');

  var months = {jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',
                jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'};

  // Priority 1: date immediately following a DOB / Date of Birth label
  var m = cleaned.match(/(?:DOB|Date\s*of\s*Birth|D\.O\.B|Birth\s*Date)[^\d]{0,15}(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/i);
  if (m) return pad(m[1]) + '/' + pad(m[2]) + '/' + m[3];

  // Pattern 1: any remaining DD/MM/YYYY or DD-MM-YYYY
  m = cleaned.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
  if (m) return pad(m[1]) + '/' + pad(m[2]) + '/' + m[3];

  // Pattern 2: YYYY/MM/DD or YYYY-MM-DD
  m = cleaned.match(/(19|20)(\d{2})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (m) return pad(m[4]) + '/' + pad(m[3]) + '/' + m[1] + m[2];

  // Pattern 3: DD Mon YYYY (e.g. 15 Aug 2010, 02 Jan 2016)
  m = cleaned.match(/(\d{1,2})\s+([A-Za-z]{3,9})\s+((?:19|20)\d{2})/);
  if (m) {
    var mon = months[m[2].toLowerCase().slice(0,3)];
    if (mon) return pad(m[1]) + '/' + mon + '/' + m[3];
  }

  // Pattern 4: Mon DD, YYYY (e.g. Aug 15, 2010)
  m = cleaned.match(/([A-Za-z]{3,9})\s+(\d{1,2}),?\s+((?:19|20)\d{2})/);
  if (m) {
    var mon2 = months[m[1].toLowerCase().slice(0,3)];
    if (mon2) return pad(m[2]) + '/' + mon2 + '/' + m[3];
  }

  return null;
}

function pad(n) { return String(n).length === 1 ? '0' + n : String(n); }

function runOcr(d) {
  if (!d || !d.imageBase64) return { status: 'error', message: 'No image provided' };
  if (!GROQ_API_KEY || GROQ_API_KEY === 'YOUR_GROQ_API_KEY_HERE') {
    return { status: 'error', message: 'GROQ_API_KEY is not configured in Code.gs. Please set your Groq API key.' };
  }
  try {
    // Use different prompt for UTR screenshot vs ID card
    var isUtrMode = (d.mode === 'utr');
    var prompt;
    if (isUtrMode) {
      prompt = 'You are an OCR system for UPI payment screenshots. Extract the UTR/Transaction ID from this payment confirmation screenshot. Look for labels like UTR, UPI Ref, Transaction ID, Txn ID, Ref No, Reference Number. Return ONLY this JSON: { "utr": "the UTR or transaction ID value or null", "amount": "amount paid or null", "status": "Success/Failed/null" }. No markdown, no explanation.';
    } else {
      prompt = 'You are a meticulous OCR system for Indian government ID documents (Aadhaar, Voter ID, Passport, Driving Licence, PAN Card). Your SINGLE MOST IMPORTANT task is to accurately extract the exact Date of Birth printed on the card.\n\nCRITICAL — Indian Aadhaar cards often show MULTIPLE dates:\n- "Issue Date" (when the card was issued) — often printed sideways/vertically along the card edge — IGNORE this completely, it is NOT the date of birth.\n- "DOB" or "பிறந்த தேதி / DOB" or "जन्म तिथि / DOB" — this is the ACTUAL date of birth you must extract. It is usually next to the person\'s name and photo.\n- If you see both an Issue Date and a DOB label, ALWAYS use the value next to the DOB label, never the Issue Date.\n\nSTRICT RULES:\n1. Scan the ENTIRE card systematically for the field explicitly labeled DOB / Date of Birth / D.O.B / Birth Date / Born / जन्म तिथि / பிறந்த தேதி.\n2. Aadhaar cards are bilingual — the DOB line often looks like "பிறந்த நாள் / DOB : 14/09/2010" — extract only the date value after the colon.\n3. Convert ANY date format found (DD-MM-YYYY, DD.MM.YYYY, "15 Aug 2010") to DD/MM/YYYY strictly with leading zeros.\n4. Double-check digits are not confused: 1 vs 7, 0 vs 8, 3 vs 8, 5 vs 6. Re-verify before outputting.\n5. If text is blurry, rotated, partially cut off, or genuinely unreadable, set date_of_birth to null. NEVER guess.\n6. If this image is NOT a valid ID card, set error to "Not an ID card" and all fields to null.\n7. Verify the extracted year is realistic: between 1920 and 2026.\n\nReturn ONLY this JSON (no markdown, no explanation, no extra text):\n{\n  "full_name": "complete name exactly as printed or null",\n  "date_of_birth": "DD/MM/YYYY (from the DOB field only, never Issue Date) or null",\n  "gender": "Male, Female, Transgender, or null",\n  "id_type": "Aadhaar Card, Voter ID, Passport, Driving Licence, PAN Card, or Other",\n  "id_number": "last 4 digits masked as XXXX-XXXX-1234 or null",\n  "confidence": "high, medium, or low — your confidence in the date_of_birth reading"\n}';
    }

    var payload = JSON.stringify({
      model: 'meta-llama/llama-4-scout-17b-16e-instruct',
      max_tokens: 512,
      messages: [{
        role: 'user',
        content: [
          { type: 'image_url', image_url: { url: 'data:' + (d.mimeType||'image/jpeg') + ';base64,' + d.imageBase64 } },
          { type: 'text', text: prompt }
        ]
      }]
    });

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + GROQ_API_KEY },
      payload: payload,
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    if (responseCode !== 200) {
      Logger.log('Groq API error (' + responseCode + '): ' + responseText);
      return { status: 'error', message: 'OCR service error (' + responseCode + '). Please try again.' };
    }

    var json = JSON.parse(responseText);
    if (json.error) {
      Logger.log('Groq API error response: ' + JSON.stringify(json.error));
      return { status: 'error', message: json.error.message };
    }
    if (!json.choices || !json.choices[0]) {
      Logger.log('Groq API unexpected response: ' + responseText);
      return { status: 'error', message: 'Unexpected OCR response format.' };
    }
    var result = json.choices[0].message.content;

    // Extra DOB extraction: if model missed it, scan raw text for date patterns
    var enhanced = enhanceDobFromRawText(result);
    return { status: 'ok', result: enhanced };
  } catch(err) {
    return { status: 'error', message: err.message };
  }
}

// ── Migration: convert ALL photo formats to Drive lh3 URLs ──
// Run manually in Apps Script editor — select migratePhotosToBase64, click Run
function migratePhotosToBase64() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INDIVIDUAL_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log('No data to migrate.');
    return;
  }

  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var rows = sheet.getDataRange().getValues();
  var migrated = 0;
  var skipped = 0;
  var failed = 0;

  for (var i = 1; i < rows.length; i++) {
    var photoCell = rows[i][9];
    var regId = String(rows[i][0] || 'row' + (i+1));
    var cellStr = String(photoCell || '');

    // Already lh3 URL — skip
    if (cellStr.indexOf('lh3.googleusercontent.com') !== -1) {
      Logger.log('Row ' + (i+1) + ' (' + regId + '): already lh3, skipping.');
      skipped++;
      continue;
    }

    // Base64 data URI — upload to Drive, store lh3 URL
    if (cellStr.indexOf('data:image') === 0) {
      try {
        var imgData = cellStr.split(',');
        var mimeType = imgData[0].match(/:(.*?);/)[1];
        var ext = mimeType.split('/')[1] || 'jpg';
        var bytes = Utilities.base64Decode(imgData[1]);
        var blob = Utilities.newBlob(bytes, mimeType, regId + '_photo.' + ext);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var lh3Url = 'https://lh3.googleusercontent.com/d/' + file.getId();
        sheet.getRange(i + 1, 10).setValue(lh3Url);
        Logger.log('Row ' + (i+1) + ' (' + regId + '): base64 → lh3 OK');
        migrated++;
        Utilities.sleep(300);
      } catch (err) {
        Logger.log('Row ' + (i+1) + ' (' + regId + ') base64 failed: ' + err.message);
        failed++;
      }
      continue;
    }

    // Old drive.google.com URL — convert to lh3
    if (cellStr.indexOf('drive.google.com') !== -1) {
      try {
        var match = cellStr.match(/id=([a-zA-Z0-9_-]+)/);
        if (match && match[1]) {
          var existingFile = DriveApp.getFileById(match[1]);
          existingFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          var lh3Url = 'https://lh3.googleusercontent.com/d/' + match[1];
          sheet.getRange(i + 1, 10).setValue(lh3Url);
          Logger.log('Row ' + (i+1) + ' (' + regId + '): drive → lh3 OK');
          migrated++;
        }
      } catch (err) {
        Logger.log('Row ' + (i+1) + ' (' + regId + ') drive failed: ' + err.message);
        failed++;
      }
      continue;
    }

    Logger.log('Row ' + (i+1) + ' (' + regId + '): no photo, skipping.');
    skipped++;
  }

  Logger.log('Done. Migrated: ' + migrated + ', Skipped: ' + skipped + ', Failed: ' + failed);
}

// ── Helpers ──
function makeKey(arr) {
  if (!arr || !Array.isArray(arr)) return '';
  return arr.map(function(v) {
    return (v !== null && v !== undefined ? v : '').toString().trim().toLowerCase().replace(/\s+/g, '');
  }).join('|');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
