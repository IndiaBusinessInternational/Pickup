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

/**
 * Flipkart/Shopsy Order IDs also print glued to the Invoice No on the tax
 * invoice (e.g. OD337737259727708100-LWACJIC270000291). The trailing invoice
 * text is NOT part of the order key — it is the same single package — so we
 * collapse any "OD"+15-22 digits value down to that bare base. One order can
 * then never hold two rows. Non-Flipkart IDs are returned unchanged, so Meesho
 * 18-digit sub-orders (incl. their _1/_2 suffixes) and Amazon 3-7-7 IDs are
 * left exactly as-is.
 */
function normID_(id) {
  var s = String(id == null ? '' : id).trim();
  var m = s.match(/^OD\d{15,22}/);
  return m ? m[0] : s;
}

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

    // Index existing rows by NORMALISED OrderID (Flipkart collapsed to its
    // "OD…" base). rowsByKey keeps every physical row sharing a key so stale
    // suffixed Flipkart duplicates left by older builds can be merged into one.
    var idx = {};        // normKey -> 1-based row chosen as the live row
    var rowsByKey = {};  // normKey -> [1-based rows...]
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === '' || values[i][0] == null) continue;
      var k = normID_(values[i][0]);
      idx[k] = i + 1;
      (rowsByKey[k] = rowsByKey[k] || []).push(i + 1);
    }

    var now = new Date().toISOString();
    var added = 0, updated = 0, skipped = 0, merged = 0;
    var toAppend = [];
    var rowsToDelete = [];

    for (var j = 0; j < dispatches.length; j++) {
      var rec = dispatches[j] || {};
      var id = normID_(rec.orderID);
      var date = toISO_(rec.date || '');
      if (!id || !isValidDate_(date)) { skipped++; continue; }
      var platform = rec.platform || '';
      var courier = rec.courier || '';

      // Merge away any stale rows that share this base (e.g. a suffixed
      // Flipkart row alongside the bare one): keep the last, mark the rest
      // for deletion so the live row below is the only survivor.
      var sibs = rowsByKey[id] || [];
      if (sibs.length > 1) {
        for (var sIdx = 0; sIdx < sibs.length - 1; sIdx++) {
          rowsToDelete.push(sibs[sIdx]); merged++;
        }
        var live = sibs[sibs.length - 1];
        rowsByKey[id] = [live];
        idx[id] = live;
      }

      if (idx[id] && idx[id] > 0) {
        // Rewrite the OrderID cell too, in case the surviving row still holds
        // a suffixed value, then update the rest of the columns.
        sh.getRange(idx[id], 1, 1, HEADERS.length).setValues([[id, date, platform, courier, now]]);
        updated++;
      } else {
        toAppend.push([id, date, platform, courier, now]);
        idx[id] = -1; // guard against duplicates within the same batch
        rowsByKey[id] = [-1];
        added++;
      }
    }
    if (toAppend.length) {
      sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, HEADERS.length).setValues(toAppend);
    }
    // Delete merged-away duplicate rows bottom-up so row indexes stay valid.
    rowsToDelete.sort(function (a, b) { return b - a; });
    for (var dRow = 0; dRow < rowsToDelete.length; dRow++) sh.deleteRow(rowsToDelete[dRow]);

    return json_({ ok: true, added: added, updated: updated, skipped: skipped, merged: merged, total: Math.max(0, sh.getLastRow() - 1) });
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
