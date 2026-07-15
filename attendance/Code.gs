// ═══════════════════════════════════════════════════════════════════
//  Amorè N' More Studio — Combined Sheet Script
//  • POST  → Bookings tab
//  • GET   → Attendance (lookup · record · guest)
//            + Google Calendar event per visit
//            + Email notification on every check-in / check-out
// ═══════════════════════════════════════════════════════════════════

// ── Core config ──────────────────────────────────────────────────────
var SPREADSHEET_ID = '10InXMuvd8pnqPeftniAKlg4fx6M3UgduDKxstn5Km-c';
var TIMEZONE       = 'Asia/Tokyo';
var RECORDS_TAB    = 'AttendanceRecords';
var USERS_TAB      = 'Users';            // A = ID, B = Name

// ── Calendar & notifications ─────────────────────────────────────────
// CALENDAR_ID: 'primary'  uses the account's default calendar.
//              Or paste a specific calendar ID from Google Calendar settings.
var CALENDAR_ID   = 'primary';
var NOTIFY_EMAILS = ['the.studioeternelle@gmail.com', 'ms.burberryy@gmail.com'];

// ── Sheet columns ─────────────────────────────────────────────────────
//  A=Name  B=ID  C=Purpose  D=TimeIn  E=TimeOut  F=CalEventId (internal)

// ── Helpers ───────────────────────────────────────────────────────────
function getSheet(tabName) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(tabName);
}

function toDate(ts) {
  // ts is "yyyy-MM-dd HH:mm:ss"
  return new Date(String(ts).replace(' ', 'T'));
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
      var hr = sheet.getRange(1, 1, 1, headers.length);
      hr.setFontWeight('bold');
      hr.setFontColor('#F4EDE3');
      hr.setBackground('#7A1515');
      hr.setHorizontalAlignment('center');
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
//  GET — Attendance router
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
    return { info: 'Amorè N\' More — Sheet script is running.' };
  } catch (err) {
    return { error: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════════
//  User lookup
// ════════════════════════════════════════════════════════════════════
function lookupUser(id) {
  if (!id) return { error: 'ID is required' };
  var sheet = getSheet(USERS_TAB);
  if (!sheet) return { error: 'Sheet "' + USERS_TAB + '" not found. Create it with columns: A = ID, B = Name' };

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(id).trim()) {
      return { ok: true, id: String(rows[i][0]).trim(), name: String(rows[i][1]).trim() };
    }
  }
  return { error: 'ID not registered: ' + id };
}

// ════════════════════════════════════════════════════════════════════
//  Record — Registered user
// ════════════════════════════════════════════════════════════════════
function recordAttendance(p) {
  var id = p.id, name = p.name, purpose = p.purpose, type = p.type;
  if (!id || !name || !purpose || !type) return { error: 'Missing required fields' };

  var sheet = getSheet(RECORDS_TAB);
  if (!sheet) return { error: 'Sheet "' + RECORDS_TAB + '" not found' };
  ensureAttendanceHeaders(sheet);

  var now = new Date();
  var ts  = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  // ── Time In ──────────────────────────────────────────────────────
  if (type === 'in') {
    sheet.appendRow([name, id, purpose, ts, '', '']);
    var newRow  = sheet.getLastRow();
    var eventId = calCreate(name, id, purpose, ts);
    if (eventId) sheet.getRange(newRow, 6).setValue(eventId);
    notify('Time In', name, id, purpose, ts, '');
    return { ok: true, message: 'Time In recorded', time: ts };
  }

  // ── Time Out ─────────────────────────────────────────────────────
  if (type === 'out') {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][1]).trim() === String(id).trim() && !rows[i][4]) {
          sheet.getRange(i + 2, 5).setValue(ts);
          calUpdate(rows[i][5], rows[i][3], ts);
          notify('Time Out', name, id, purpose, rows[i][3], ts);
          return { ok: true, message: 'Time Out recorded', time: ts, timeIn: rows[i][3] };
        }
      }
    }

    if (p.manualTimeIn) {
      var timeInStr = p.manualTimeIn.replace('T', ' ') + ':00';
      var eventId2  = calCreate(name, id, purpose, timeInStr);
      if (eventId2) calUpdate(eventId2, timeInStr, ts);
      sheet.appendRow([name, id, purpose, timeInStr, ts, eventId2 || '']);
      notify('Check-In/Out (manual Time In)', name, id, purpose, timeInStr, ts);
      return { ok: true, message: 'Attendance recorded with manual Time In', time: ts };
    }

    return { needTimeIn: true, message: 'No open Time In record found for this ID.' };
  }

  return { error: 'type must be "in" or "out"' };
}

// ════════════════════════════════════════════════════════════════════
//  Record — Guest
// ════════════════════════════════════════════════════════════════════
function recordGuest(p) {
  var name = (p.name || '').trim(), purpose = p.purpose, type = p.type;
  if (!name || !purpose || !type) return { error: 'Missing required fields' };

  var sheet = getSheet(RECORDS_TAB);
  if (!sheet) return { error: 'Sheet "' + RECORDS_TAB + '" not found' };
  ensureAttendanceHeaders(sheet);

  var now = new Date();
  var ts  = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  // ── Time In ──────────────────────────────────────────────────────
  if (type === 'in') {
    sheet.appendRow([name, 'GUEST', purpose, ts, '', '']);
    var newRow  = sheet.getLastRow();
    var eventId = calCreate(name, 'GUEST', purpose, ts);
    if (eventId) sheet.getRange(newRow, 6).setValue(eventId);
    notify('Time In', name, 'GUEST', purpose, ts, '');
    return { ok: true, message: 'Time In recorded', time: ts };
  }

  // ── Time Out ─────────────────────────────────────────────────────
  if (type === 'out') {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        var sameName = String(rows[i][0]).trim().toLowerCase() === name.toLowerCase();
        var isGuest  = String(rows[i][1]).trim() === 'GUEST';
        if (sameName && isGuest && !rows[i][4]) {
          sheet.getRange(i + 2, 5).setValue(ts);
          calUpdate(rows[i][5], rows[i][3], ts);
          notify('Time Out', name, 'GUEST', purpose, rows[i][3], ts);
          return { ok: true, message: 'Time Out recorded', time: ts, timeIn: rows[i][3] };
        }
      }
    }

    if (p.manualTimeIn) {
      var timeInStr = p.manualTimeIn.replace('T', ' ') + ':00';
      var eventId2  = calCreate(name, 'GUEST', purpose, timeInStr);
      if (eventId2) calUpdate(eventId2, timeInStr, ts);
      sheet.appendRow([name, 'GUEST', purpose, timeInStr, ts, eventId2 || '']);
      notify('Check-In/Out (manual Time In)', name, 'GUEST', purpose, timeInStr, ts);
      return { ok: true, message: 'Attendance recorded with manual Time In', time: ts };
    }

    return { needTimeIn: true, message: 'No open Time In record found for ' + name + '.' };
  }

  return { error: 'type must be "in" or "out"' };
}

// ════════════════════════════════════════════════════════════════════
//  Calendar helpers
// ════════════════════════════════════════════════════════════════════

// Create event at Time In; end = Time In + 1h (updated when they check out)
function calCreate(name, id, purpose, timeInStr) {
  if (!CALENDAR_ID) return '';
  try {
    var cal   = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!cal) return '';
    var start = toDate(timeInStr);
    var end   = new Date(start.getTime() + 60 * 60 * 1000);
    var label = id === 'GUEST' ? name + ' (Guest)' : name + ' [' + id + ']';
    var event = cal.createEvent(purpose + ' — ' + label, start, end, {
      description:
        'Amorè N’ More Studio — Attendance\n' +
        'Name:    ' + name + '\n' +
        'ID:      ' + (id === 'GUEST' ? 'Guest' : id) + '\n' +
        'Purpose: ' + purpose + '\n' +
        'Time In: ' + timeInStr
    });
    return event.getId();
  } catch (e) {
    return '';
  }
}

// Set the actual end time once Time Out is recorded
function calUpdate(eventId, timeInStr, timeOutStr) {
  if (!CALENDAR_ID || !eventId) return;
  try {
    var cal   = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!cal) return;
    var event = cal.getEventById(eventId);
    if (!event) return;
    event.setEndTime(toDate(timeOutStr));
  } catch (e) {}
}

// ════════════════════════════════════════════════════════════════════
//  Email notification
// ════════════════════════════════════════════════════════════════════
function notify(type, name, id, purpose, timeIn, timeOut) {
  if (!NOTIFY_EMAILS || !NOTIFY_EMAILS.length) return;
  try {
    var idLabel  = (id === 'GUEST') ? 'Guest' : id;
    var subject  = '\u{1F4CB} [' + type + '] ' + name + ' — ' + purpose;
    var body     =
      'Amorè N’ More Studio — Attendance Record\n' +
      '================================================\n' +
      'Type     : ' + type    + '\n' +
      'Name     : ' + name    + '\n' +
      'ID       : ' + idLabel + '\n' +
      'Purpose  : ' + purpose + '\n' +
      'Time In  : ' + timeIn  + '\n' +
      (timeOut ? 'Time Out : ' + timeOut + '\n' : '') +
      '\nView sheet → https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID;

    var htmlBody =
      '<div style="font-family:sans-serif;max-width:480px">' +
      '<h2 style="color:#7A1515;margin-bottom:4px">Amor&egrave; N&rsquo; More Studio</h2>' +
      '<p style="color:#7A6050;font-size:12px;margin-top:0;letter-spacing:2px">ATTENDANCE RECORD</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
      row('Type',     '<strong>' + type + '</strong>') +
      row('Name',     name) +
      row('ID',       idLabel) +
      row('Purpose',  purpose) +
      row('Time In',  timeIn) +
      (timeOut ? row('Time Out', timeOut) : '') +
      '</table>' +
      '<p style="margin-top:20px">' +
      '<a href="https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '" ' +
      'style="background:#7A1515;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;font-size:13px">' +
      'Open Attendance Sheet</a></p></div>';

    for (var i = 0; i < NOTIFY_EMAILS.length; i++) {
      MailApp.sendEmail({
        to:       NOTIFY_EMAILS[i],
        subject:  subject,
        body:     body,
        htmlBody: htmlBody
      });
    }
  } catch (e) {}
}

function row(label, value) {
  return '<tr><td style="padding:6px 10px;color:#7A6050;font-size:11px;letter-spacing:1px;text-transform:uppercase;width:90px">' +
         label + '</td><td style="padding:6px 10px;border-bottom:1px solid #f0e8d8">' + value + '</td></tr>';
}

// ════════════════════════════════════════════════════════════════════
//  Sheet setup
// ════════════════════════════════════════════════════════════════════
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
    // Column F holds the calendar event ID — keep it narrow and hidden-ish
    sheet.setColumnWidth(6, 220);
    sheet.hideColumns(6);
  }
}
