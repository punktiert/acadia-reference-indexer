/**
 * SheetIO.gs — persistence. Saves the uploaded .docx to Drive and writes submission rows to
 * the two master tabs in the exact CumInCAD column order, under a script lock, with
 * delete-then-append idempotency so a re-submission of the same paper ID replaces (not
 * duplicates) its rows.
 */

/** Ensure a tab exists with the exact header row. If a non-empty header is present it must
 *  match exactly — we never silently reorder an editor-customized layout. */
function ensureHeaders_(ss, tabName) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);
  var headers = HEADERS[tabName];
  var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var isEmpty = firstRow.every(function (c) { return c === '' || c === null; });
  if (isEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    for (var i = 0; i < headers.length; i++) {
      if (String(firstRow[i]).trim() !== headers[i]) {
        throw new Error('Tab "' + tabName + '" header does not match the expected CumInCAD layout at column ' +
          (i + 1) + ' (found "' + firstRow[i] + '", expected "' + headers[i] + '"). Fix the sheet header and retry.');
      }
    }
  }
  return sheet;
}

function colIndexOf_(tabName, header) {
  return HEADERS[tabName].indexOf(header) + 1;  // 1-based; 0 if not found
}

/** Delete every data row whose key column equals id (bottom-up so indices don't shift). */
function deleteRowsForId_(sheet, keyCol, id) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var keys = sheet.getRange(2, keyCol, last - 1, 1).getValues();
  var toDelete = [];
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === String(id)) toDelete.push(i + 2);
  }
  for (var d = toDelete.length - 1; d >= 0; d--) sheet.deleteRow(toDelete[d]);
  return toDelete.length;
}

function idExists_(sheet, keyCol, id) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var keys = sheet.getRange(2, keyCol, last - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) if (String(keys[i][0]) === String(id)) return true;
  return false;
}

/** A pure 4-digit year is written as a Number (2024 stored ints); ranges / "N/A" stay text. */
function normYear_(y) {
  var s = String(y == null ? '' : y).trim();
  return /^\d{4}$/.test(s) ? parseInt(s, 10) : s;
}

function buildTitleRow_(meta, title, authors) {
  var slots = (HEADERS.title_authors.length - 7) / 2;   // 8 AUTHOR FIRST/LAST pairs
  var row = new Array(HEADERS.title_authors.length);
  for (var i = 0; i < row.length; i++) row[i] = '';
  row[0] = meta.id;
  row[1] = meta.source;
  row[2] = DEFAULTS.SERIES;
  row[3] = String(meta.type || '').trim();
  row[4] = title || '';
  row[5] = meta.email || '';
  row[6] = '';   // SUMMARY (empty, as in 2024)
  for (var k = 0; k < slots; k++) {
    var a = authors[k];
    row[7 + 2 * k] = a && a.first ? a.first : '';
    row[8 + 2 * k] = a && a.last ? a.last : '';
  }
  return row;
}

function buildRefRows_(id, references) {
  return references.map(function (r) {
    return [id, r.authors || '', normYear_(r.year), r.title || '', r.source || ''];
  });
}

/** Save the uploaded docx into the configured Drive folder, named by paper ID. Replaces any
 *  prior upload for the same ID. Returns the Drive file id. */
function saveDocx_(blob, id, originalName) {
  var folder = getOrCreateUploadsFolder_();
  var clean = String(originalName || 'paper.docx').replace(/[\\\/:*?"<>|]+/g, '_');
  var name = id + '__' + clean;
  var existing = folder.getFilesByName(name);
  while (existing.hasNext()) existing.next().setTrashed(true);
  var file = folder.createFile(blob.setName(name));
  return file.getId();
}

/** Write one submission (1 title row + N reference rows). Locked + idempotent. */
function writeSubmission_(reviewed) {
  var cfg = getConfig();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = getMasterSpreadsheet_();
    var taTab = ensureHeaders_(ss, DEFAULTS.TAB_TITLE_AUTHORS);
    var refTab = ensureHeaders_(ss, DEFAULTS.TAB_REFERENCES);
    var id = reviewed.meta.id;

    if (cfg.ON_RESUBMIT === 'block') {
      if (idExists_(taTab, colIndexOf_(DEFAULTS.TAB_TITLE_AUTHORS, 'ID'), id)) {
        return { ok: false, error: 'duplicate_id', paperId: id };
      }
    } else {
      deleteRowsForId_(taTab, colIndexOf_(DEFAULTS.TAB_TITLE_AUTHORS, 'ID'), id);
      deleteRowsForId_(refTab, colIndexOf_(DEFAULTS.TAB_REFERENCES, 'paper'), id);
    }

    taTab.appendRow(buildTitleRow_(reviewed.meta, reviewed.title, reviewed.authors || []));

    var rows = buildRefRows_(id, reviewed.references || []);
    if (rows.length) {
      refTab.getRange(refTab.getLastRow() + 1, 1, rows.length, HEADERS.references.length).setValues(rows);
    }
    return { ok: true, paperId: id, titleRowsWritten: 1, refRowsWritten: rows.length };
  } finally {
    lock.releaseLock();
  }
}
