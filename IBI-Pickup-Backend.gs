/**
 * ============================================================================
 *  IBI DISPATCH LEDGER — Google Apps Script backend
 *  Shared source of truth for the IBI Pickup Request Manager and the IBI ERP.
 * ----------------------------------------------------------------------------
 *  The Pickup Manager POSTs dispatch records here (order ID + pickup date +
 *  platform); the ERP GETs them back to build day-wise dispatch reports.
 *  Data lives in a Google Sheet in your Google Drive, so it is shared across
 *  every device and browser — unlike the old per-browser localStorage.
 *
 *  DEPLOY (one time):
 *   1. Create a Google Sheet (any name, e.g. "IBI Dispatch Ledger").
 *   2. Extensions -> Apps Script. Delete the sample, paste this whole file.
 *   3. Save. Then Deploy -> New deployment -> type "Web app".
 *        - Description : IBI Dispatch Ledger
 *        - Execute as  : Me
 *        - Who has access : Anyone   (required so both sites can reach it)
 *   4. Click Deploy, authorise, and COPY the "Web app URL"
 *        (it looks like https://script.google.com/macros/s/AKfyc.../exec).
 *   5. Paste that URL into BOTH apps:
 *        - Pickup Manager : "Cloud Backend URL" field
 *        - ERP            : Reports -> Import Pickup Data -> Cloud URL field
 *   6. If you ever change this code, Deploy -> Manage deployments ->
 *        edit -> New version, so the same URL serves the new code.
 * ============================================================================
 */

var SHEET_NAME = 'Dispatches';
var HEADERS = ['OrderID', 'Date', 'Platform', 'Courier', 'UpdatedAt'];

/** Returns the data sheet, creating it (with headers, Date column as text). */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  // Force the Date column to PLAIN TEXT so "2026-06-01" never drifts by
  // timezone or gets reformatted by the spreadsheet.
  sh.getRange('B:B').setNumberFormat('@');
  return sh;
}

/** Normalise any cell value to a YYYY-MM-DD string. */
function toISO_(v) {
  if (v instanceof Date) {
    var y = v.getFullYear();
    var m = ('0' + (v.getMonth() + 1)).slice(-2);
    var d = ('0' + v.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(v).slice(0, 10);
}

function isValidDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(s); }

/** JSON response helper. */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * WRITE — the Pickup Manager posts { dispatches: [{orderID,date,platform,courier}] }.
 * Upserts by OrderID (last write wins). Body is sent as text/plain to avoid a
 * CORS preflight, so we parse e.postData.contents ourselves.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (err) { return json_({ ok: false, error: 'busy' }); }
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);

    // DELETE — { action:'delete', orderID:'…' }  or  { action:'delete', orderIDs:[…] }
    // Removes the matching row(s) so a mis-recorded order can be corrected.
    if (body.action === 'delete') {
      var delIds = body.orderIDs || (body.orderID ? [body.orderID] : []);
      var delSet = {};
      for (var k = 0; k < delIds.length; k++) delSet[String(delIds[k]).trim()] = true;
      var shD = getSheet_();
      var valsD = shD.getDataRange().getValues();
      var deleted = 0;
      // Delete bottom-up so row indexes stay valid as rows are removed.
      for (var r = valsD.length - 1; r >= 1; r--) {
        if (delSet[String(valsD[r][0])]) { shD.deleteRow(r + 1); deleted++; }
      }
      return json_({ ok: true, deleted: deleted, total: Math.max(0, shD.getLastRow() - 1) });
    }

    var dispatches = body.dispatches || [];
    var sh = getSheet_();
    var values = sh.getDataRange().getValues();

    // Index existing rows: OrderID -> 1-based row number.
    var idx = {};
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] !== '' && values[i][0] != null) idx[String(values[i][0])] = i + 1;
    }

    var now = new Date().toISOString();
    var added = 0, updated = 0, skipped = 0;
    var toAppend = [];

    for (var j = 0; j < dispatches.length; j++) {
      var rec = dispatches[j] || {};
      var id = String(rec.orderID || '').trim();
      var date = toISO_(rec.date || '');
      if (!id || !isValidDate_(date)) { skipped++; continue; }
      var platform = rec.platform || '';
      var courier = rec.courier || '';
      if (idx[id]) {
        sh.getRange(idx[id], 2, 1, 4).setValues([[date, platform, courier, now]]);
        updated++;
      } else {
        toAppend.push([id, date, platform, courier, now]);
        idx[id] = -1; // guard against duplicates within the same batch
        added++;
      }
    }
    if (toAppend.length) {
      sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, HEADERS.length).setValues(toAppend);
    }
    return json_({ ok: true, added: added, updated: updated, skipped: skipped, total: Math.max(0, sh.getLastRow() - 1) });
  } catch (err2) {
    return json_({ ok: false, error: String(err2) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * READ — returns every dispatch record.
 *   - With ?callback=fn  -> JSONP (used by the ERP; bypasses CORS entirely).
 *   - Without callback   -> plain JSON.
 * Optional ?since=YYYY-MM-DD filters to records on/after that date.
 */
function doGet(e) {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  var since = (e && e.parameter && e.parameter.since) ? String(e.parameter.since).slice(0, 10) : '';
  var dispatches = [];
  for (var i = 1; i < values.length; i++) {
    var id = values[i][0];
    if (id === '' || id == null) continue;
    var date = toISO_(values[i][1]);
    if (since && date < since) continue;
    dispatches.push({
      orderID: String(id),
      date: date,
      platform: String(values[i][2] || ''),
      courier: String(values[i][3] || '')
    });
  }
  var payload = { ok: true, app: 'IBI Dispatch Ledger', count: dispatches.length, dispatches: dispatches };

  var cb = (e && e.parameter && e.parameter.callback) ? String(e.parameter.callback) : '';
  if (cb && /^[A-Za-z_$][\w$]*$/.test(cb)) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}
