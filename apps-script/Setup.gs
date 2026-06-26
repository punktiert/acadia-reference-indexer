// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 Daniel Koehler / ACADIA
/**
 * Setup.gs — the editor-facing, no-code setup. Adds an "ACADIA Indexer" menu to the Sheet.
 *
 * The script is bound to the Sheet, so there are no IDs to find: it uses the active
 * spreadsheet and auto-creates the uploads folder. "First-time setup" builds the data tabs,
 * a friendly Settings tab, and an in-sheet "How to publish" guide.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ACADIA Indexer')
    .addItem('1. First-time setup', 'setup')
    .addItem('2. How to publish (get the author link)', 'showHelpTab')
    .addSeparator()
    .addItem('Run self-test', 'runTests')
    .addToUi();
}

/** The spreadsheet the script is bound to (with a stored-id fallback for the web app). */
function getMasterSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('No spreadsheet found. Open the Sheet and run "ACADIA Indexer ▸ First-time setup".');
}

/** Auto-created Drive folder for stored uploads (cached by id in Script Properties).
 *  Created in the SAME folder as the Sheet (so it sits right next to it), falling back to
 *  My Drive root if the Sheet's parent isn't writable under the granted Drive scope. */
function getOrCreateUploadsFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('UPLOADS_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* was deleted — recreate below */ }
  }
  var ss = getMasterSpreadsheet_();
  var name = 'ACADIA Indexer uploads — ' + ss.getName();
  var folder;
  try {
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    folder = (parents.hasNext() ? parents.next() : DriveApp.getRootFolder()).createFolder(name);
  } catch (e) {
    folder = DriveApp.createFolder(name);   // fallback: My Drive root
  }
  props.setProperty('UPLOADS_FOLDER_ID', folder.getId());
  writeSetting_('Uploads folder ID', folder.getId());
  return folder;
}

/** One-click setup: bind, build data tabs, Settings tab, uploads folder, and help tab. */
function setup() {
  var ss = getMasterSpreadsheet_();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  ensureHeaders_(ss, DEFAULTS.TAB_TITLE_AUTHORS);
  ensureHeaders_(ss, DEFAULTS.TAB_REFERENCES);
  ensureSettingsTab_(ss);
  getOrCreateUploadsFolder_();
  ensureHelpTab_(ss);

  ss.setActiveSheet(ss.getSheetByName(DEFAULTS.TAB_SETTINGS));
  ss.toast('Setup complete. Fill in the Settings tab, then open the "How to publish" tab.', 'ACADIA Indexer', 10);
}

function ensureSettingsTab_(ss) {
  var sh = ss.getSheetByName(DEFAULTS.TAB_SETTINGS);
  if (sh) return sh;                 // never clobber existing edits
  sh = ss.insertSheet(DEFAULTS.TAB_SETTINGS, 0);
  sh.getRange(1, 1, 1, 3).setValues([['Setting', 'Value (edit this column)', 'Notes']]);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, SETTINGS_ROWS.length, 3).setValues(SETTINGS_ROWS);
  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 520);
  sh.setColumnWidth(3, 380);
  sh.getRange(2, 1, SETTINGS_ROWS.length, 1).setFontWeight('bold');
  return sh;
}

/** Write a value into column B of the Settings row whose column A matches `label`. */
function writeSetting_(label, value) {
  var ss = getMasterSpreadsheet_();
  var sh = ss.getSheetByName(DEFAULTS.TAB_SETTINGS);
  if (!sh) return;
  var col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim().toLowerCase() === String(label).toLowerCase()) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
}

function showHelpTab() {
  var ss = getMasterSpreadsheet_();
  var sh = ss.getSheetByName(DEFAULTS.TAB_HELP);
  if (!sh) { setup(); sh = ss.getSheetByName(DEFAULTS.TAB_HELP); }
  if (sh) ss.setActiveSheet(sh);
}

function ensureHelpTab_(ss) {
  var name = DEFAULTS.TAB_HELP;
  var sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  var lines = [
    ['How to publish the indexer (one time) — then share the link with authors'],
    [''],
    ['STEP 1  Fill in the Settings tab — especially the three "Volume … SOURCE" rows with the exact CumInCAD wording.'],
    [''],
    ['STEP 2  Open the script editor:  top menu  Extensions ▸ Apps Script.'],
    [''],
    ['STEP 3  In the script editor, click the blue "Deploy" button (top right) ▸ New deployment.'],
    ['        • Click the gear ▸ choose "Web app".'],
    ['        • Description: ACADIA Reference Indexer'],
    ['        • Execute as: Me'],
    ['        • Who has access: Anyone   (so authors can open it without signing in)'],
    ['        • Click Deploy. Approve the permissions if asked (this is your own script).'],
    [''],
    ['STEP 4  Copy the "Web app URL" it shows you. THAT is the link you send to authors.'],
    [''],
    ['Future years: you do NOT repeat this. Just update the Settings tab (the 3 SOURCE rows),'],
    ['clear last year’s rows from the title_authors and references tabs, and reuse the same link.'],
    [''],
    ['If you ever change the CODE (not just settings): script editor ▸ Deploy ▸ Manage deployments ▸ edit ▸ New version.']
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  sh.setColumnWidth(1, 760);
  sh.getRange(1, 1, lines.length, 1).setWrap(true);
  return sh;
}
