// Tripledger — Apps Script server code (MailApp OTP, rate-limit, attempt-limit)

var META_SHEET_NAME = 'TripMeta';
var TRIP_HEADERS = ['ID', 'Date', 'Description', 'Paid By', 'Amount', 'Split Among'];

// Replace with the file ID of YOUR logo image uploaded to Google Drive.
// Upload the PNG to Drive -> Share -> Anyone with the link -> copy the ID
// from the URL: https://drive.google.com/file/d/<THIS PART>/view
var DRA_LOGO_FILE_ID = '1DiLaApuBTVw8aB5qJr8JrHMo5UyqKDl5';

// OTP settings
var OTP_EXPIRY_MINUTES = 5;
var OTP_RESEND_INTERVAL_SECONDS = 60;
var OTP_MAX_ATTEMPTS = 5;
var ALLOW_DEBUG_OTP = false; // set true only for local testing

function ensureSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var metaSheet = ss.getSheetByName(META_SHEET_NAME);
  if (!metaSheet) {
    metaSheet = ss.insertSheet(META_SHEET_NAME);
    metaSheet.appendRow(['Code', 'Name', 'Members', 'CreatedAt']);
    metaSheet.setFrozenRows(1);
  }
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}

// Helper: SHA-256 hex
// NOTE: this is a HELPER function used internally by sendOtpEmail() and
// verifyOtp(). It always needs a real string argument, so it can't be
// run standalone from the Apps Script editor's "Run" button (which calls
// functions with zero arguments). The guard below just turns that into a
// clear message instead of a confusing native error. Use runManualTests()
// below if you want to sanity-check this from the editor.
function sha256Hex(text) {
  if (text === undefined || text === null) {
    throw new Error('sha256Hex needs a text argument — run this via generateOTP(), testMail(), or runManualTests(), not standalone.');
  }
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  return digest.map(function(b) {
    var v = b < 0 ? b + 256 : b;
    var s = v.toString(16);
    return s.length === 1 ? '0' + s : s;
  }).join('');
}

// Send OTP email using MailApp (free)
function sendOtpEmail(userEmail) {
  var otpCode = Math.floor(100000 + Math.random() * 900000).toString();
  var expiry = Date.now() + OTP_EXPIRY_MINUTES * 60000;
  var otpHash = sha256Hex(otpCode);
  var props = PropertiesService.getScriptProperties();

  var record = {
    otpHash: otpHash,
    expiry: expiry,
    attempts: 0,
    lastSent: Date.now()
  };
  props.setProperty('otp_' + userEmail, JSON.stringify(record));

  var subject = 'Tripledger — your verification code';
  var htmlBody = '<p>Your Tripledger verification code is:</p>' +
                 '<p style="font-size:20px;font-weight:700;color:#7c4dff;">' + otpCode + '</p>' +
                 '<p>This code will expire in ' + OTP_EXPIRY_MINUTES + ' minutes.</p>';

  try {
    MailApp.sendEmail(userEmail, subject, 'Your verification code for Tripledger.', {
      htmlBody: htmlBody,
      name: 'Tripledger'
    });
    return ALLOW_DEBUG_OTP ? otpCode : true;
  } catch (e) {
    Logger.log('MailApp error: ' + e.toString());
    return null;
  }
}

function generateOTP(userEmail) {
  ensureSheets();

  if (!userEmail || userEmail.indexOf('@') === -1) {
    return { success: false, message: 'Valid email required.' };
  }

  var props = PropertiesService.getScriptProperties();
  var key = 'otp_' + userEmail;
  var raw = props.getProperty(key);
  if (raw) {
    try {
      var rec = JSON.parse(raw);
      var since = Date.now() - (rec.lastSent || 0);
      if (since < (OTP_RESEND_INTERVAL_SECONDS * 1000)) {
        var wait = Math.ceil((OTP_RESEND_INTERVAL_SECONDS * 1000 - since) / 1000);
        return { success: false, message: 'Please wait ' + wait + 's before requesting a new code.' };
      }
    } catch (e) {}
  }

  var sent = sendOtpEmail(userEmail);
  if (!sent) {
    return { success: false, message: 'Failed to send OTP. Check MailApp permissions.' };
  }

  var resp = { success: true, message: 'OTP sent to your email.' };
  if (ALLOW_DEBUG_OTP && typeof sent === 'string') {
    resp.data = { otp: sent };
  }
  return resp;
}

function verifyOtp(userEmail, enteredOtp) {
  var props = PropertiesService.getScriptProperties();
  var key = 'otp_' + userEmail;
  var raw = props.getProperty(key);
  if (!raw) return 'INVALID';

  var data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    props.deleteProperty(key);
    return 'INVALID';
  }

  if (Date.now() > data.expiry) {
    props.deleteProperty(key);
    return 'EXPIRED';
  }

  if (data.attempts >= OTP_MAX_ATTEMPTS) {
    props.deleteProperty(key);
    return 'TOO_MANY_ATTEMPTS';
  }

  var enteredHash = sha256Hex(String(enteredOtp));
  if (enteredHash === data.otpHash) {
    props.deleteProperty(key);
    return 'SUCCESS';
  } else {
    data.attempts = (data.attempts || 0) + 1;
    if (data.attempts >= OTP_MAX_ATTEMPTS) {
      props.deleteProperty(key);
      return 'TOO_MANY_ATTEMPTS';
    } else {
      props.setProperty(key, JSON.stringify(data));
      return 'INVALID';
    }
  }
}

function verifyOTP(email, otp) {
  var result = verifyOtp(email, otp);
  if (result === 'SUCCESS') {
    return { success: true, message: 'Login successful.', data: { email: email } };
  } else if (result === 'EXPIRED') {
    return { success: false, message: 'That OTP expired. Please request a new one.' };
  } else if (result === 'TOO_MANY_ATTEMPTS') {
    return { success: false, message: 'Too many incorrect attempts. Please request a new code.' };
  } else {
    return { success: false, message: 'Invalid OTP. Please try again.' };
  }
}

/* -------------------------------
   Trip logic
------------------------------- */

function respond(success, message, data) {
  var payload = { success: success, message: message };
  if (data) payload.data = data;
  return payload;
}

function generateCode() {
  var chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  var code = '';
  for (var i = 0; i < 6; i++) code += chars[Math.floor(Math.random() * chars.length)];
  return code;
}

function parseMembers(raw) {
  try {
    var parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed;
  } catch (e) {}
  return [];
}

// NOTE: helper — needs (ss, code). Called internally by createTrip() and
// addExpense(). Don't run standalone from the editor dropdown; use
// runManualTests() instead if you want to sanity-check it.
function createTripSheet(ss, code) {
  if (!ss || !code) {
    throw new Error('createTripSheet needs (ss, code) — run this via createTrip() or runManualTests(), not standalone.');
  }
  var sheet = ss.getSheetByName(code);
  if (!sheet) {
    sheet = ss.insertSheet(code);
    sheet.appendRow(TRIP_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('B2:B').setNumberFormat('@');
  }
  return sheet;
}

function readExpenses(ss, code) {
  var sheet = ss.getSheetByName(code);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var expenses = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var dateVal = row[1];
    var dateStr = (dateVal instanceof Date)
      ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(dateVal);
    expenses.push({
      id: String(row[0]),
      date: dateStr,
      description: String(row[2]),
      paidBy: String(row[3]),
      amount: Number(row[4]),
      splitAmong: String(row[5]).split(',').map(function(s){ return s.trim(); }).filter(Boolean)
    });
  }
  return expenses;
}

function createTrip(email, tripData) {
  ensureSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = (tripData.name || '').trim();
  if (!name) return respond(false, 'Trip name required.');
  var code = generateCode();

  var meta = ss.getSheetByName(META_SHEET_NAME);
  var existing = meta.getDataRange().getValues();
  var sheetExists = function(c) { return ss.getSheetByName(c) !== null; };
  // Ensure code is unique in meta AND as a sheet name
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]) === code || sheetExists(code)) {
      code = generateCode();
      i = 1;
    }
  }

  var members = [{ name: email, email: email }];
  var createdAt = new Date().toISOString();
  var membersStr = JSON.stringify(members);
  meta.appendRow([code, name, membersStr, createdAt]);
  createTripSheet(ss, code);
  return respond(true, 'Trip created.', { code: code, name: name });
}

function listTrips(email) {
  ensureSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var meta = ss.getSheetByName(META_SHEET_NAME);
  var data = meta.getDataRange().getValues();
  var trips = [];
  for (var i = 1; i < data.length; i++) {
    var members = parseMembers(data[i][2]);
    var isMember = members.some(function(m){ return m.email === email; });
    if (isMember) trips.push({ code: String(data[i][0]), name: data[i][1], createdAt: data[i][3] });
  }
  return respond(true, 'Trips retrieved.', { trips: trips });
}

function getTrip(email, code) {
  ensureSheets();
  code = String(code || '').toUpperCase().trim();
  if (!code) return respond(false, 'Trip code required.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var meta = ss.getSheetByName(META_SHEET_NAME);
  var data = meta.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === code) {
      var members = parseMembers(data[i][2]);
      var isMember = members.some(function(m){ return m.email === email; });
      if (!isMember) return respond(false, 'You are not a member of this trip.');
      var expenses = readExpenses(ss, code);
      return respond(true, 'Trip found.', {
        code: code,
        name: data[i][1],
        members: members,
        createdAt: data[i][3],
        expenses: expenses
      });
    }
  }
  return respond(false, 'Trip not found.');
}

function joinTrip(email, code, memberName) {
  ensureSheets();
  code = String(code || '').toUpperCase().trim();
  email = String(email || '').trim().toLowerCase();
  memberName = String(memberName || '').trim();
  if (!code) return respond(false, 'Trip code required.');
  if (!email || email.indexOf('@') === -1) return respond(false, 'A valid email is required.');
  if (!memberName) return respond(false, 'Enter your name to join this trip.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var meta = ss.getSheetByName(META_SHEET_NAME);
  var data = meta.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === code) {
      var members = parseMembers(data[i][2]);
      var already = members.some(function(m) {
        return String(m.email || '').trim().toLowerCase() === email;
      });
      if (already) return respond(true, 'Already a member.', { code: code });

      // Claim a placeholder made by the trip owner instead of creating a duplicate.
      var placeholder = members.find(function(m) {
        return !m.email && String(m.name || '').trim().toLowerCase() === memberName.toLowerCase();
      });
      if (placeholder) {
        placeholder.email = email;
      } else {
        var nameTaken = members.some(function(m) {
          return String(m.name || '').trim().toLowerCase() === memberName.toLowerCase();
        });
        if (nameTaken) return respond(false, 'That member name is already in use.');
        members.push({ name: memberName, email: email });
      }
      meta.getRange(i+1, 3).setValue(JSON.stringify(members));
      return respond(true, 'Joined trip.', { code: code, members: members });
    }
  }
  return respond(false, 'Trip not found.');
}

function addExpense(email, code, expense) {
  ensureSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  code = String(code || '').toUpperCase().trim();
  if (!code || !expense) return respond(false, 'Invalid request.');
  var meta = ss.getSheetByName(META_SHEET_NAME);
  var data = meta.getDataRange().getValues();
  var rowIndex = -1;
  var members = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === code) {
      members = parseMembers(data[i][2]);
      var isMember = members.some(function(m){ return m.email === email; });
      if (!isMember) return respond(false, 'Not a member.');
      rowIndex = i;
      break;
    }
  }
  if (rowIndex === -1) return respond(false, 'Trip not found.');

  var desc = (expense.description || '').trim();
  var amt = parseFloat(expense.amount);
  var date = expense.date || (new Date()).toISOString().slice(0,10);
  var paidBy = (expense.paidBy || '').trim();
  var splitAmong = expense.splitAmong || [];

  if (!desc || !amt || amt <= 0 || !paidBy || splitAmong.length === 0) {
    return respond(false, 'Invalid expense data.');
  }

  var memberNames = members.map(function(m){ return m.name; });
  if (memberNames.indexOf(paidBy) === -1) return respond(false, 'PaidBy must be a member.');
  for (var j = 0; j < splitAmong.length; j++) {
    if (memberNames.indexOf(splitAmong[j]) === -1) return respond(false, 'Split among must be members.');
  }

  var id = Utilities.getUuid();
  var newExpense = {
    id: id,
    date: date,
    description: desc,
    paidBy: paidBy,
    amount: amt,
    splitAmong: splitAmong
  };

  var tripSheet = ss.getSheetByName(code);
  if (!tripSheet) tripSheet = createTripSheet(ss, code);
  tripSheet.appendRow([newExpense.id, newExpense.date, newExpense.description, newExpense.paidBy, newExpense.amount, newExpense.splitAmong.join(', ')]);
  return respond(true, 'Expense added.', { expense: newExpense });
}

function deleteExpense(email, code, expenseId) {
  ensureSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  code = String(code || '').toUpperCase().trim();
  if (!code || !expenseId) return respond(false, 'Invalid request.');
  var meta = ss.getSheetByName(META_SHEET_NAME);
  var data = meta.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === code) {
      var members = parseMembers(data[i][2]);
      var isMember = members.some(function(m){ return m.email === email; });
      if (!isMember) return respond(false, 'Not a member.');
      found = true;
      break;
    }
  }
  if (!found) return respond(false, 'Trip not found.');

  var tripSheet = ss.getSheetByName(code);
  if (!tripSheet) return respond(false, 'Trip sheet missing.');
  var rows = tripSheet.getDataRange().getValues();
  var rowToDelete = -1;
  for (var j = 1; j < rows.length; j++) {
    if (String(rows[j][0]) === expenseId) { rowToDelete = j + 1; break; }
  }
  if (rowToDelete === -1) return respond(false, 'Expense not found.');
  tripSheet.deleteRow(rowToDelete);
  return respond(true, 'Expense deleted.');
}

function addMember(email, code, name) {
  ensureSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  code = String(code || '').toUpperCase().trim();
  name = (name || '').trim();
  if (!code || !name) return respond(false, 'Invalid request.');
  var meta = ss.getSheetByName(META_SHEET_NAME);
  var data = meta.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === code) {
      var members = parseMembers(data[i][2]);
      var isMember = members.some(function(m){ return m.email === email; });
      if (!isMember) return respond(false, 'Not a member.');
      var exists = members.some(function(m){ return m.name === name; });
      if (exists) return respond(false, 'A member with that name already exists.');
      members.push({ name: name, email: null });
      meta.getRange(i+1, 3).setValue(JSON.stringify(members));
      return respond(true, 'Member added.', { members: members });
    }
  }
  return respond(false, 'Trip not found.');
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Tripledger by DRA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// Serves the Drive image to the webpage as a reliable data URL. This avoids
// Drive's viewer/permission redirects that break a normal <img src="...">.
function getDraLogoDataUrl() {
  var blob = DriveApp.getFileById(DRA_LOGO_FILE_ID).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function testMail() {
  MailApp.sendEmail('divyam11110@gmail.com', 'Test', 'This is a test.');
}

/* -------------------------------------------------------------
   RUN THIS FROM THE EDITOR to sanity-check the helper functions
   that need arguments (so you don't hit "Cannot read properties
   of undefined" by picking them directly from the Run dropdown).
   Select "runManualTests" in the function dropdown, click Run,
   then check Execution log.
------------------------------------------------------------- */
function runManualTests() {
  Logger.log('--- sha256Hex ---');
  Logger.log(sha256Hex('123456'));

  Logger.log('--- createTripSheet (creates/reuses a TEST tab) ---');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var testSheet = createTripSheet(ss, 'TEST01');
  Logger.log('Sheet name: ' + testSheet.getName());

  Logger.log('--- ensureSheets ---');
  ensureSheets();
  Logger.log('TripMeta exists: ' + (ss.getSheetByName(META_SHEET_NAME) !== null));

  Logger.log('All manual tests completed without errors.');
}
