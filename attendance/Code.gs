// ═══════════════════════════════════════════════════════════════════
//  Amorè N' More Studio — Combined Sheet Script
//  • POST  → Bookings tab   (existing behaviour, unchanged)
//  • GET   → Attendance tab (lookup user + record Time In / Time Out)
//
//  Deploy: Extensions → Apps Script → Deploy → Manage deployments
//          Edit the EXISTING deployment — no new URL needed.
//          The attendance page uses the same URL as the bookings page.
// ═══════════════════════════════════════════════════════════════════

// ── Shared config ────────────────────────────────────────────────────
var SPREADSHEET_ID = '10InXMuvd8pnqPeftniAKlg4fx6M3UgduDKxstn5Km-c';
var TIMEZONE       = 'Asia/Tokyo';
var RECORDS_TAB    = 'AttendanceRecords';
var USERS_TAB      = 'Users';          // columns: A = ID,  B = Name

function getSheet(tabName) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(tabName);
}

// ════════════════════════════════════════════════════════════════════
//  POST — Bookings  (original, unchanged)
// ════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var data = e.parameter;

    var sheet = ss.getSheetByName('Bookings');
    if (!sheet) {
      sheet = ss.insertSheet('Bookings');
      var headers = ['Submitted At', 'Name', 'Contact', 'Facebook Name',
                     'Service', 'Preferred Date', 'Time Slot', 'Notes'];
      sheet.appendRow(headers);

      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setFontColor('#F4EDE3');
      headerRange.setBackground('#7A1515');
      headerRange.setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, headers.length, 160);
    }

    var d = data.submitted ? new Date(data.submitted) : new Date();
    sheet.appendRow([
      d.toLocaleString('en-JP', { timeZone: TIMEZONE }),
      data.name     || '',
      data.contact  || '',
      data.fbName   || '',
      data.service  || '',
      data.date     || '',
      data.timeslot || '',
      data.notes    || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ════════════════════════════════════════════════════════════════════
//  GET — Attendance
//  ?action=lookup&id=XXX
//  ?action=record&id=XXX&name=XXX&purpose=XXX&type=in|out[&manualTimeIn=…]
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  var result = handleAttendance(e.parameter);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleAttendance(p) {
  try {
    if (p.action === 'lookup') return lookupUser(p.id);
    if (p.action === 'record') return recordAttendance(p);
    if (p.action === 'guest')  return recordGuest(p);
    // Plain browser hit — friendly message instead of an error
    return { info: 'Amorè N\' More — Sheet script is running.' };
  } catch (err) {
    return { error: err.toString() };
  }
}

// ── Look up a registered user by ID ──────────────────────────────────
function lookupUser(id) {
  if (!id) return { error: 'ID is required' };

  var sheet = getSheet(USERS_TAB);
  if (!sheet) return {
    error: 'Sheet "' + USERS_TAB + '" not found. ' +
           'Create it with columns: A = ID, B = Name'
  };

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(id).trim()) {
      return { ok: true, id: String(rows[i][0]).trim(), name: String(rows[i][1]).trim() };
    }
  }
  return { error: 'ID not registered: ' + id };
}

// ── Write a Time In or Time Out record ───────────────────────────────
function recordAttendance(p) {
  var id = p.id, name = p.name, purpose = p.purpose, type = p.type;
  if (!id || !name || !purpose || !type) return { error: 'Missing required fields' };

  var sheet = getSheet(RECORDS_TAB);
  if (!sheet) return { error: 'Sheet "' + RECORDS_TAB + '" not found' };

  ensureAttendanceHeaders(sheet);

  var now = new Date();
  var ts  = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  // ── Time In: always a new row ─────────────────────────────────────
  if (type === 'in') {
    sheet.appendRow([name, id, purpose, ts, '']);
    return { ok: true, message: 'Time In recorded', time: ts };
  }

  // ── Time Out: find the most recent open row for this ID ───────────
  if (type === 'out') {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = data.length - 1; i >= 0; i--) {
        if (String(data[i][1]).trim() === String(id).trim() && !data[i][4]) {
          sheet.getRange(i + 2, 5).setValue(ts);
          return { ok: true, message: 'Time Out recorded', time: ts, timeIn: data[i][3] };
        }
      }
    }

    // No open Time In — check if the caller supplied a manual one
    if (p.manualTimeIn) {
      var timeInStr = p.manualTimeIn.replace('T', ' ') + ':00';
      sheet.appendRow([name, id, purpose, timeInStr, ts]);
      return { ok: true, message: 'Attendance recorded with manual Time In', time: ts };
    }

    // Tell the front-end to ask the user for a Time In
    return { needTimeIn: true, message: 'No open Time In record found for this ID.' };
  }

  return { error: 'type must be "in" or "out"' };
}

// ── Guest record (matched by name, ID stored as "GUEST") ─────────────
function recordGuest(p) {
  var name = (p.name || '').trim(), purpose = p.purpose, type = p.type;
  if (!name || !purpose || !type) return { error: 'Missing required fields' };

  var sheet = getSheet(RECORDS_TAB);
  if (!sheet) return { error: 'Sheet "' + RECORDS_TAB + '" not found' };

  ensureAttendanceHeaders(sheet);

  var now = new Date();
  var ts  = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  if (type === 'in') {
    sheet.appendRow([name, 'GUEST', purpose, ts, '']);
    return { ok: true, message: 'Time In recorded', time: ts };
  }

  if (type === 'out') {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      // Match by name (case-insensitive) + GUEST marker, most recent open row
      for (var i = data.length - 1; i >= 0; i--) {
        var sameName = String(data[i][0]).trim().toLowerCase() === name.toLowerCase();
        var isGuest  = String(data[i][1]).trim() === 'GUEST';
        if (sameName && isGuest && !data[i][4]) {
          sheet.getRange(i + 2, 5).setValue(ts);
          return { ok: true, message: 'Time Out recorded', time: ts, timeIn: data[i][3] };
        }
      }
    }

    if (p.manualTimeIn) {
      var timeInStr = p.manualTimeIn.replace('T', ' ') + ':00';
      sheet.appendRow([name, 'GUEST', purpose, timeInStr, ts]);
      return { ok: true, message: 'Attendance recorded with manual Time In', time: ts };
    }

    return { needTimeIn: true, message: 'No open Time In record found for ' + name + '.' };
  }

  return { error: 'type must be "in" or "out"' };
}

// ── Write header row if AttendanceRecords sheet is blank ─────────────
function ensureAttendanceHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    var h = ['Name', 'ID', 'Purpose', 'Time In', 'Time Out'];
    sheet.appendRow(h);

    var r = sheet.getRange(1, 1, 1, h.length);
    r.setFontWeight('bold');
    r.setFontColor('#F4EDE3');
    r.setBackground('#7A1515');
    r.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, h.length, 180);
  }
}
