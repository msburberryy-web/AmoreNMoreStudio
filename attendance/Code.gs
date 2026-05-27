// ═══════════════════════════════════════════════════════════════════
//  Amorè N' More Studio — Attendance System
//  Deploy: Extensions → Apps Script → Deploy → New deployment
//          Type: Web App | Execute as: Me | Access: Anyone
// ═══════════════════════════════════════════════════════════════════

const SS_ID        = '10InXMuvd8pnqPeftniAKlg4fx6M3UgduDKxstn5Km-c';
const RECORDS_TAB  = 'AttendanceRecords';
const USERS_TAB    = 'Users';           // columns: A=ID  B=Name
const TIMEZONE     = 'Asia/Tokyo';

// ── Entry point ─────────────────────────────────────────────────────
function doGet(e) {
  const result = handle(e.parameter);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle(p) {
  try {
    if (p.action === 'lookup') return lookupUser(p.id);
    if (p.action === 'record') return recordAttendance(p);
    return { error: 'Unknown action' };
  } catch (err) {
    return { error: err.toString() };
  }
}

// ── Lookup user by ID from Users sheet ──────────────────────────────
function lookupUser(id) {
  if (!id) return { error: 'ID is required' };

  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(USERS_TAB);
  if (!sheet) return {
    error: 'Sheet "' + USERS_TAB + '" not found. ' +
           'Create it with columns: A = ID, B = Name'
  };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {          // skip header row
    if (String(rows[i][0]).trim() === String(id).trim()) {
      return { ok: true, id: String(rows[i][0]).trim(), name: String(rows[i][1]).trim() };
    }
  }
  return { error: 'ID not registered: ' + id };
}

// ── Record Time In or Time Out ───────────────────────────────────────
function recordAttendance(p) {
  const { id, name, purpose, type } = p;
  if (!id || !name || !purpose || !type) return { error: 'Missing required fields' };

  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(RECORDS_TAB);
  if (!sheet) return { error: 'Sheet "' + RECORDS_TAB + '" not found' };

  ensureHeaders(sheet);

  const now = new Date();
  const ts  = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  // ── TIME IN ──────────────────────────────────────────────────────
  if (type === 'in') {
    sheet.appendRow([name, id, purpose, ts, '']);
    return { ok: true, message: 'Time In recorded', time: ts };
  }

  // ── TIME OUT ─────────────────────────────────────────────────────
  if (type === 'out') {
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      // Search from the most recent row upward for an open Time In
      for (let i = data.length - 1; i >= 0; i--) {
        const rowId    = String(data[i][1]).trim();
        const timeOut  = data[i][4];
        if (rowId === String(id).trim() && !timeOut) {
          sheet.getRange(i + 2, 5).setValue(ts);
          return { ok: true, message: 'Time Out recorded', time: ts, timeIn: data[i][3] };
        }
      }
    }

    // No open Time In found ─ check if caller supplied a manual one
    if (p.manualTimeIn) {
      const timeInStr = p.manualTimeIn.replace('T', ' ') + ':00';
      sheet.appendRow([name, id, purpose, timeInStr, ts]);
      return { ok: true, message: 'Attendance recorded with manual Time In', time: ts };
    }

    // Signal the front-end to ask for Time In
    return { needTimeIn: true, message: 'No open Time In found for this ID.' };
  }

  return { error: 'type must be "in" or "out"' };
}

// ── Write header row if sheet is blank ──────────────────────────────
function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Name', 'ID', 'Purpose', 'Time In', 'Time Out']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
}
