// ═══════════════════════════════════════════════════════════════════
//  Amorè N' More Studio — Combined Sheet Script
//
//  POST  → Bookings tab    + calendar event + email notification
//  GET   → Attendance tab  + calendar event + email notification
//          (lookup · record · guest)
// ═══════════════════════════════════════════════════════════════════

// ── Config ───────────────────────────────────────────────────────────
var SPREADSHEET_ID = '10InXMuvd8pnqPeftniAKlg4fx6M3UgduDKxstn5Km-c';
var TIMEZONE       = 'Asia/Tokyo';
var RECORDS_TAB    = 'AttendanceRecords';
var USERS_TAB      = 'Users';   // A = ID, B = Name

// Leave CALENDAR_ID blank to use the account's default calendar,
// or set it to a specific calendar's email address for a shared/studio calendar.
var CALENDAR_ID   = '';
var NOTIFY_EMAILS = ['the.studioeternelle@gmail.com', 'ms.burberryy@gmail.com'];

// Sheet columns for AttendanceRecords:
//  A=Name  B=ID  C=Purpose  D=TimeIn  E=TimeOut  F=CalEventId (hidden)

// ── Shared helpers ────────────────────────────────────────────────────
function getSheet(tabName) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(tabName);
}
function toDate(ts) {
  return new Date(String(ts).replace(' ', 'T'));
}
// CalendarApp.getCalendarById('primary') returns null in Apps Script.
// Use getDefaultCalendar() when no specific ID is set.
function getCal() {
  return CALENDAR_ID
    ? CalendarApp.getCalendarById(CALENDAR_ID)
    : CalendarApp.getDefaultCalendar();
}

// ════════════════════════════════════════════════════════════════════
//  Quick test — run this from the Apps Script editor (▶ Run)
//  to confirm calendar access works and see results in Execution log.
// ════════════════════════════════════════════════════════════════════
function testCalendar() {
  Logger.log('=== testCalendar START ===');
  try {
    var cal = getCal();
    Logger.log('Calendar: ' + (cal ? cal.getName() + ' (' + cal.getId() + ')' : 'NULL — no calendar found'));
    if (!cal) { Logger.log('FAIL: getCal() returned null'); return; }

    var now  = new Date();
    var end  = new Date(now.getTime() + 60 * 60 * 1000);
    var evt  = cal.createEvent('TEST — Amorè Studio Script Check', now, end,
                 { description: 'Auto-created by testCalendar(). Safe to delete.' });
    Logger.log('SUCCESS: event created → ' + evt.getId());
    Logger.log('Event title: ' + evt.getTitle());
    Logger.log('Start: ' + evt.getStartTime());
  } catch (err) {
    Logger.log('ERROR: ' + err.toString());
  }
  Logger.log('=== testCalendar END ===');
}

// ════════════════════════════════════════════════════════════════════
//  POST — Bookings
// ════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var data = e.parameter;

    var sheet = ss.getSheetByName('Bookings');
    if (!sheet) {
      sheet = ss.insertSheet('Bookings');
      var h  = ['Submitted At','Name','Contact','Facebook Name',
                 'Service','Preferred Date','Time Slot','Notes'];
      sheet.appendRow(h);
      var hr = sheet.getRange(1, 1, 1, h.length);
      hr.setFontWeight('bold');
      hr.setFontColor('#F4EDE3');
      hr.setBackground('#7A1515');
      hr.setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, h.length, 160);
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

    // ── Calendar event for this booking ──────────────────────────
    calCreateBooking(
      data.name    || '',
      data.contact || '',
      data.service || '',
      data.date    || '',
      data.timeslot|| '',
      data.notes   || ''
    );

    // ── Email notification ────────────────────────────────────────
    notifyBooking(
      data.name    || '',
      data.contact || '',
      data.fbName  || '',
      data.service || '',
      data.date    || '',
      data.timeslot|| '',
      data.notes   || ''
    );

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
//  Attendance — Registered user
// ════════════════════════════════════════════════════════════════════
function recordAttendance(p) {
  var id = p.id, name = p.name, purpose = p.purpose, type = p.type;
  if (!id || !name || !purpose || !type) return { error: 'Missing required fields' };

  var sheet = getSheet(RECORDS_TAB);
  if (!sheet) return { error: 'Sheet "' + RECORDS_TAB + '" not found' };
  ensureAttendanceHeaders(sheet);

  var now = new Date();
  var ts  = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  if (type === 'in') {
    sheet.appendRow([name, id, purpose, ts, '', '']);
    var newRow  = sheet.getLastRow();
    var eventId = calCreateAttendance(name, id, purpose, ts);
    if (eventId) sheet.getRange(newRow, 6).setValue(eventId);
    notifyAttendance('Time In', name, id, purpose, ts, '');
    return { ok: true, message: 'Time In recorded', time: ts };
  }

  if (type === 'out') {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][1]).trim() === String(id).trim() && !rows[i][4]) {
          sheet.getRange(i + 2, 5).setValue(ts);
          calUpdateAttendance(rows[i][5], rows[i][3], ts);
          notifyAttendance('Time Out', name, id, purpose, rows[i][3], ts);
          return { ok: true, message: 'Time Out recorded', time: ts, timeIn: rows[i][3] };
        }
      }
    }
    if (p.manualTimeIn) {
      var tin     = p.manualTimeIn.replace('T', ' ') + ':00';
      var evtId   = calCreateAttendance(name, id, purpose, tin);
      if (evtId) calUpdateAttendance(evtId, tin, ts);
      sheet.appendRow([name, id, purpose, tin, ts, evtId || '']);
      notifyAttendance('Check-In/Out (manual Time In)', name, id, purpose, tin, ts);
      return { ok: true, message: 'Attendance recorded with manual Time In', time: ts };
    }
    return { needTimeIn: true, message: 'No open Time In record found for this ID.' };
  }

  return { error: 'type must be "in" or "out"' };
}

// ════════════════════════════════════════════════════════════════════
//  Attendance — Guest
// ════════════════════════════════════════════════════════════════════
function recordGuest(p) {
  var name = (p.name || '').trim(), purpose = p.purpose, type = p.type;
  if (!name || !purpose || !type) return { error: 'Missing required fields' };

  var sheet = getSheet(RECORDS_TAB);
  if (!sheet) return { error: 'Sheet "' + RECORDS_TAB + '" not found' };
  ensureAttendanceHeaders(sheet);

  var now = new Date();
  var ts  = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  if (type === 'in') {
    sheet.appendRow([name, 'GUEST', purpose, ts, '', '']);
    var newRow  = sheet.getLastRow();
    var eventId = calCreateAttendance(name, 'GUEST', purpose, ts);
    if (eventId) sheet.getRange(newRow, 6).setValue(eventId);
    notifyAttendance('Time In', name, 'GUEST', purpose, ts, '');
    return { ok: true, message: 'Time In recorded', time: ts };
  }

  if (type === 'out') {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        var sameName = String(rows[i][0]).trim().toLowerCase() === name.toLowerCase();
        if (sameName && String(rows[i][1]).trim() === 'GUEST' && !rows[i][4]) {
          sheet.getRange(i + 2, 5).setValue(ts);
          calUpdateAttendance(rows[i][5], rows[i][3], ts);
          notifyAttendance('Time Out', name, 'GUEST', purpose, rows[i][3], ts);
          return { ok: true, message: 'Time Out recorded', time: ts, timeIn: rows[i][3] };
        }
      }
    }
    if (p.manualTimeIn) {
      var tin   = p.manualTimeIn.replace('T', ' ') + ':00';
      var evtId = calCreateAttendance(name, 'GUEST', purpose, tin);
      if (evtId) calUpdateAttendance(evtId, tin, ts);
      sheet.appendRow([name, 'GUEST', purpose, tin, ts, evtId || '']);
      notifyAttendance('Check-In/Out (manual Time In)', name, 'GUEST', purpose, tin, ts);
      return { ok: true, message: 'Attendance recorded with manual Time In', time: ts };
    }
    return { needTimeIn: true, message: 'No open Time In record found for ' + name + '.' };
  }

  return { error: 'type must be "in" or "out"' };
}

// ════════════════════════════════════════════════════════════════════
//  Calendar — Bookings
//  Creates a timed event if the Time Slot contains a parseable time,
//  otherwise falls back to an all-day event on the Preferred Date.
// ════════════════════════════════════════════════════════════════════
function calCreateBooking(name, contact, service, dateStr, timeSlot, notes) {
  if (!dateStr) return;
  try {
    var cal = getCal();
    if (!cal) return;

    var title = '🌸 ' + (service || 'Appointment') + ' — ' + (name || 'Client');
    var desc  =
      'Amorè N’ More Studio — Booking\n' +
      '============================================\n' +
      'Name      : ' + name      + '\n' +
      'Contact   : ' + contact   + '\n' +
      'Service   : ' + service   + '\n' +
      'Date      : ' + dateStr   + '\n' +
      'Time Slot : ' + timeSlot  + '\n' +
      (notes ? 'Notes     : ' + notes : '') + '\n' +
      '\nView sheet → https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID;

    // Try to parse date
    var base = new Date(dateStr);
    if (isNaN(base.getTime())) base = new Date(dateStr.replace(/-/g, '/'));
    if (isNaN(base.getTime())) return;

    // Try to extract a start time from timeSlot (e.g. "10:00 AM", "14:00", "2pm")
    var timeMatch = String(timeSlot).match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);
    if (timeMatch) {
      var h   = parseInt(timeMatch[1], 10);
      var m   = parseInt(timeMatch[2] || '0', 10);
      var mer = (timeMatch[3] || '').toLowerCase();
      if (mer === 'pm' && h < 12) h += 12;
      if (mer === 'am' && h === 12) h = 0;
      base.setHours(h, m, 0, 0);
      var end = new Date(base.getTime() + 2 * 60 * 60 * 1000); // 2-hour block
      cal.createEvent(title, base, end, { description: desc });
    } else {
      cal.createAllDayEvent(title, base, { description: desc });
    }
  } catch (e) {}
}

// ════════════════════════════════════════════════════════════════════
//  Calendar — Attendance
// ════════════════════════════════════════════════════════════════════
function calCreateAttendance(name, id, purpose, timeInStr) {
  try {
    var cal = getCal();
    if (!cal) return '';
    var start = toDate(timeInStr);
    var end   = new Date(start.getTime() + 60 * 60 * 1000); // +1h placeholder
    var label = id === 'GUEST' ? name + ' (Guest)' : name + ' [' + id + ']';
    var event = cal.createEvent(
      '📋 ' + purpose + ' — ' + label,
      start, end,
      {
        description:
          'Amorè N’ More Studio — Attendance\n' +
          'Name    : ' + name + '\n' +
          'ID      : ' + (id === 'GUEST' ? 'Guest' : id) + '\n' +
          'Purpose : ' + purpose + '\n' +
          'Time In : ' + timeInStr
      }
    );
    return event.getId();
  } catch (e) { return ''; }
}

function calUpdateAttendance(eventId, timeInStr, timeOutStr) {
  if (!eventId) return;
  try {
    var cal   = getCal();
    if (!cal) return;
    var event = cal.getEventById(eventId);
    if (event) event.setEndTime(toDate(timeOutStr));
  } catch (e) {}
}

// ════════════════════════════════════════════════════════════════════
//  Email — Bookings
// ════════════════════════════════════════════════════════════════════
function notifyBooking(name, contact, fbName, service, date, timeslot, notes) {
  if (!NOTIFY_EMAILS || !NOTIFY_EMAILS.length) return;
  try {
    var subject = '🌸 New Booking: ' + name + ' — ' + service;
    var plain   =
      'Amorè N’ More Studio — New Booking\n' +
      '================================================\n' +
      'Name          : ' + name     + '\n' +
      'Contact       : ' + contact  + '\n' +
      'Facebook Name : ' + fbName   + '\n' +
      'Service       : ' + service  + '\n' +
      'Preferred Date: ' + date     + '\n' +
      'Time Slot     : ' + timeslot + '\n' +
      (notes ? 'Notes         : ' + notes + '\n' : '') +
      '\nView sheet → https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID;

    var html =
      emailWrap('New Booking', [
        emailRow('Name',           name),
        emailRow('Contact',        contact),
        emailRow('Facebook Name',  fbName),
        emailRow('Service',        '<strong>' + service + '</strong>'),
        emailRow('Preferred Date', '<strong>' + date + '</strong>'),
        emailRow('Time Slot',      timeslot),
        notes ? emailRow('Notes', notes) : ''
      ], SPREADSHEET_ID);

    NOTIFY_EMAILS.forEach(function(addr) {
      MailApp.sendEmail({ to: addr, subject: subject, body: plain, htmlBody: html });
    });
  } catch (e) {}
}

// ════════════════════════════════════════════════════════════════════
//  Email — Attendance
// ════════════════════════════════════════════════════════════════════
function notifyAttendance(type, name, id, purpose, timeIn, timeOut) {
  if (!NOTIFY_EMAILS || !NOTIFY_EMAILS.length) return;
  try {
    var idLabel = id === 'GUEST' ? 'Guest' : id;
    var subject = '📋 [' + type + '] ' + name + ' — ' + purpose;
    var plain   =
      'Amorè N’ More Studio — Attendance\n' +
      '================================================\n' +
      'Type    : ' + type    + '\n' +
      'Name    : ' + name    + '\n' +
      'ID      : ' + idLabel + '\n' +
      'Purpose : ' + purpose + '\n' +
      'Time In : ' + timeIn  + '\n' +
      (timeOut ? 'Time Out: ' + timeOut + '\n' : '') +
      '\nView sheet → https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID;

    var html =
      emailWrap(type, [
        emailRow('Name',     name),
        emailRow('ID',       idLabel),
        emailRow('Purpose',  purpose),
        emailRow('Time In',  timeIn),
        timeOut ? emailRow('Time Out', timeOut) : ''
      ], SPREADSHEET_ID);

    NOTIFY_EMAILS.forEach(function(addr) {
      MailApp.sendEmail({ to: addr, subject: subject, body: plain, htmlBody: html });
    });
  } catch (e) {}
}

// ════════════════════════════════════════════════════════════════════
//  Email template helpers
// ════════════════════════════════════════════════════════════════════
function emailWrap(title, rows, sheetId) {
  return (
    '<div style="font-family:Georgia,serif;max-width:500px;margin:0 auto">' +
    '<div style="background:#7A1515;padding:20px 24px">' +
    '<p style="color:#E8D5B0;font-size:11px;letter-spacing:3px;text-transform:uppercase;margin:0">Amorè N’ More Studio</p>' +
    '<h2 style="color:#fff;margin:4px 0 0;font-size:20px">' + title + '</h2>' +
    '</div>' +
    '<div style="background:#fff;padding:24px">' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
    rows.filter(Boolean).join('') +
    '</table>' +
    '<p style="margin-top:24px">' +
    '<a href="https://docs.google.com/spreadsheets/d/' + sheetId + '" ' +
    'style="background:#7A1515;color:#fff;padding:10px 22px;border-radius:6px;' +
    'text-decoration:none;font-size:13px;letter-spacing:1px">Open Sheet →</a>' +
    '</p></div></div>'
  );
}

function emailRow(label, value) {
  return (
    '<tr>' +
    '<td style="padding:8px 12px 8px 0;color:#7A6050;font-size:11px;' +
    'letter-spacing:1px;text-transform:uppercase;white-space:nowrap;width:120px">' + label + '</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid #F4EDE3">' + (value || '—') + '</td>' +
    '</tr>'
  );
}

// ════════════════════════════════════════════════════════════════════
//  Installable trigger — fires when any cell in the Bookings sheet
//  is edited directly (including manual row additions).
//
//  Run installTriggers() ONCE from the Apps Script editor to register.
// ════════════════════════════════════════════════════════════════════
function onBookingSheetEdit(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'Bookings') return;

    var row = e.range.getRow();
    if (row <= 1) return; // skip header

    var data = sheet.getRange(row, 1, 1, 8).getValues()[0];
    var submittedAt = data[0]; // col A
    var name        = data[1]; // col B — must be filled
    var contact     = data[2]; // col C
    var fbName      = data[3]; // col D
    var service     = data[4]; // col E
    var dateRaw     = data[5]; // col F — may be a Date object
    var timeSlot    = data[6]; // col G
    var notes       = data[7]; // col H

    // Only process rows where Name is filled but Submitted At is still blank
    if (!name || submittedAt) return;

    // Stamp Submitted At first — acts as a write-lock against re-processing
    var ts = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(row, 1).setValue(ts);

    var dateStr = dateRaw instanceof Date
      ? Utilities.formatDate(dateRaw, TIMEZONE, 'yyyy-MM-dd')
      : String(dateRaw);

    calCreateBooking(name, contact, service, dateStr, String(timeSlot), notes);
    notifyBooking(name, contact, fbName, service, dateStr, String(timeSlot), notes);
  } catch (err) {}
}

// Run this function ONCE from the Apps Script editor (▶ Run → installTriggers)
// to register the onBookingSheetEdit trigger.
function installTriggers() {
  // Remove existing copies to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onBookingSheetEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('onBookingSheetEdit')
    .forSpreadsheet(SPREADSHEET_ID)
    .onEdit()
    .create();
  Logger.log('Trigger installed successfully.');
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
    sheet.setColumnWidth(6, 220);
    sheet.hideColumns(6); // hide the internal Cal Event ID column
  }
}
