// ===== ACADIA Reference Indexer — combined build. Edit the modular .gs files, re-run build.py. =====


// ---------- Config.gs ----------

// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 Daniel Koehler / ACADIA
/**
 * Config.gs — defaults + live config read from the in-sheet "Settings" tab.
 *
 * Editors NEVER edit this file. All deployment values (the per-volume SOURCE text, etc.) are
 * entered as plain cells in the Settings tab, which the "First-time setup" menu creates and
 * pre-fills. getConfig() reads that tab on each request and merges it over these defaults, so
 * the only thing IDs/folders depend on is the Sheet the script is bound to.
 */

var DEFAULTS = {
  TAB_TITLE_AUTHORS: 'title_authors',
  TAB_REFERENCES: 'references',
  TAB_SETTINGS: 'Settings',
  TAB_HELP: 'How to publish',

  SERIES: 'ACADIA',
  CONF_YEAR: '2026',
  ID_PREFIX: 'acadia26',   // ID = acadia26_v{vol}_{NN}
  ID_PAD: 2,
  MAX_AUTHORS: 8,
  ON_RESUBMIT: 'replace',  // 'replace' | 'block'
  TYPES: ['paper', 'project'],

  // v1=project, v2=paper. Pre-filled from 2026.acadia.org; the editor drops in the ISBN and
  // confirms the editor list on the Settings tab (neither is published on the site yet).
  VOLUMES: {
    '1': { type: 'project', sourceBase: 'ACADIA 2026: Humanism Recoded [Project Catalog of the 46th Annual Conference for the Association for Computer Aided Design in Architecture (ACADIA) ISBN <ISBN>]. Detroit. 22-24 October 2026. edited by <editors>.' },
    '2': { type: 'paper',   sourceBase: 'ACADIA 2026: Humanism Recoded [Proceedings of the 46th Annual Conference for the Association for Computer Aided Design in Architecture (ACADIA) ISBN <ISBN>]. Detroit. 22-24 October 2026. edited by <editors>.' }
  }
};

/** Exact column order for the two data tabs — MUST match the CumInCAD .xlsx layout. */
var HEADERS = {
  title_authors: [
    'ID', 'SOURCE', 'SERIES', 'TYPE', 'TITLE', 'MAIN CONTACT EMAIL', 'SUMMARY',
    'AUTHOR 1 FIRST', 'AUTHOR 1 LAST', 'AUTHOR 2 FIRST', 'AUTHOR 2 LAST',
    'AUTHOR 3 FIRST', 'AUTHOR 3 LAST', 'AUTHOR 4 FIRST', 'AUTHOR 4 LAST',
    'AUTHOR 5 FIRST', 'AUTHOR 5 LAST', 'AUTHOR 6 FIRST', 'AUTHOR 6 LAST',
    'AUTHOR 7 FIRST', 'AUTHOR 7 LAST', 'AUTHOR 8 FIRST', 'AUTHOR 8 LAST'
  ],
  references: ['paper', 'authors', 'year', 'title', 'source']
};

/** Settings tab layout: [label, default value, hint]. The setup menu writes these rows. */
var SETTINGS_ROWS = [
  ['Conference year', DEFAULTS.CONF_YEAR, 'e.g. 2025'],
  ['Paper ID prefix', DEFAULTS.ID_PREFIX, 'IDs become <prefix>_v{vol}_{number}, e.g. acadia26_v2_07'],
  ['Max authors', DEFAULTS.MAX_AUTHORS, 'CumInCAD allows up to 8'],
  ['On re-submit', DEFAULTS.ON_RESUBMIT, '"replace" = a re-upload overwrites that paper; "block" = refuse duplicates'],
  ['Volume 1 type', DEFAULTS.VOLUMES['1'].type, 'usually: project'],
  ['Volume 1 SOURCE', DEFAULTS.VOLUMES['1'].sourceBase, 'EXACT CumInCAD wording for volume 1 (replace the <…> parts)'],
  ['Volume 2 type', DEFAULTS.VOLUMES['2'].type, 'usually: paper'],
  ['Volume 2 SOURCE', DEFAULTS.VOLUMES['2'].sourceBase, 'EXACT CumInCAD wording for volume 2 (replace the <…> parts)'],
  ['Uploads folder ID', '', 'filled in automatically — do not edit']
];

/** Read the Settings tab and merge over DEFAULTS. Returns the live config object. */
function getConfig() {
  var cfg = JSON.parse(JSON.stringify(DEFAULTS));
  var ss = getMasterSpreadsheet_();
  var sh = ss.getSheetByName(DEFAULTS.TAB_SETTINGS);
  if (!sh) return cfg;

  var map = {};
  sh.getDataRange().getValues().forEach(function (r) {
    if (r[0] !== '' && r[0] != null) map[String(r[0]).trim().toLowerCase()] = r[1];
  });
  function get(label) {
    var v = map[label.toLowerCase()];
    return (v === '' || v == null) ? null : v;
  }

  if (get('Conference year')) cfg.CONF_YEAR = String(get('Conference year')).trim();
  if (get('Paper ID prefix')) cfg.ID_PREFIX = String(get('Paper ID prefix')).trim();
  if (get('Max authors')) cfg.MAX_AUTHORS = parseInt(get('Max authors'), 10) || cfg.MAX_AUTHORS;
  if (get('On re-submit')) cfg.ON_RESUBMIT = (String(get('On re-submit')).trim().toLowerCase() === 'block') ? 'block' : 'replace';
  Object.keys(cfg.VOLUMES).forEach(function (v) {
    var t = get('Volume ' + v + ' type'); if (t) cfg.VOLUMES[v].type = String(t).trim();
    var s = get('Volume ' + v + ' SOURCE'); if (s) cfg.VOLUMES[v].sourceBase = String(s).trim();
  });
  return cfg;
}

function buildPaperId(cfg, vol, number) {
  var n = String(parseInt(number, 10));
  while (n.length < cfg.ID_PAD) n = '0' + n;
  return cfg.ID_PREFIX + '_v' + String(vol) + '_' + n;
}

function sourceForVolume(cfg, vol) {
  var v = cfg.VOLUMES[String(vol)];
  return v ? v.sourceBase : '';
}

function typeForVolume(cfg, vol) {
  var v = cfg.VOLUMES[String(vol)];
  return v ? v.type : 'paper';
}


// ---------- Setup.gs ----------

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


// ---------- Docx.gs ----------

// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 Daniel Koehler / ACADIA
/**
 * Docx.gs — extract title / authors / keywords / references from a .docx by paragraph style.
 *
 * A .docx is a ZIP; we unzip in-memory (Utilities.unzip), read word/document.xml as UTF-8
 * (so smart quotes / en-dashes survive — avoids the 2024 mojibake bug), and walk paragraphs
 * with XmlService. References are the paragraphs styled "ACADIA-Reference"; a single such
 * paragraph can split across several runs (italic title / formatted year), so we concatenate
 * every w:t descendant per paragraph.
 */

var W = XmlService.getNamespace('http://schemas.openxmlformats.org/wordprocessingml/2006/main');

var HEADING_REF_RX = /^(references|bibliography|works cited)\s*$/i;
var TERMINATOR_RX = /^(image credits|acknowledg(e)?ments?|author bio|biograph|appendix|figure\s*\d|fig\.?\s*\d|table\s*\d)/i;

// Section labels that share the title's heading style but are NOT the title.
var SECTION_WORDS = {
  abstract: 1, references: 1, bibliography: 1, 'works cited': 1, keywords: 1, introduction: 1,
  background: 1, methodology: 1, methods: 1, results: 1, discussion: 1, conclusion: 1,
  conclusions: 1, acknowledgments: 1, acknowledgements: 1
};

/** Title fallback when ACADIA-Title isn't applied: the first ACADIA-HeaderA Not-Numbered
 *  paragraph (authors often style the title that way) that isn't a section label, else the
 *  first non-empty paragraph that isn't a section label. */
function fallbackTitle_(paras) {
  for (var pass = 0; pass < 2; pass++) {
    for (var i = 0; i < paras.length; i++) {
      var st = pStyleOf_(paras[i]) || '';
      var t = normalizeWs(paraText_(paras[i]));
      if (!t || SECTION_WORDS[t.toLowerCase()]) continue;
      if (pass === 0 && st.indexOf('ACADIA-HeaderANot') !== 0) continue;
      return t;
    }
  }
  return '';
}

function readDocumentXml_(docxBlob) {
  var files = Utilities.unzip(docxBlob.setContentType('application/zip'));
  var xml = null;
  for (var i = 0; i < files.length; i++) {
    if (files[i].getName() === 'word/document.xml') { xml = files[i].getDataAsString('UTF-8'); break; }
  }
  if (xml == null) throw new Error('That file is not a valid .docx (missing word/document.xml).');
  return XmlService.parse(xml);
}

function pStyleOf_(p) {
  var pPr = p.getChild('pPr', W);
  if (!pPr) return null;
  var pStyle = pPr.getChild('pStyle', W);
  if (!pStyle) return null;
  var a = pStyle.getAttribute('val', W);
  return a ? a.getValue() : null;
}

function collectText_(el, out) {
  var kids = el.getChildren();
  for (var i = 0; i < kids.length; i++) {
    var c = kids[i], name = c.getName();
    if (name === 't') out.push(c.getText());
    else if (name === 'tab') out.push('\t');
    else if (name === 'br' || name === 'cr') out.push(' ');
    else collectText_(c, out);
  }
}

function paraText_(p) {
  var out = [];
  collectText_(p, out);
  return out.join('');
}

function collectParagraphs_(el, out) {
  var kids = el.getChildren();
  for (var i = 0; i < kids.length; i++) {
    var c = kids[i];
    if (c.getName() === 'p') out.push(c);
    else collectParagraphs_(c, out);   // descend into tables / text boxes
  }
}

function stripKeywordsLabel_(text) {
  return text.replace(/^\s*keywords\s*[:\-]\s*/i, '').trim();
}

function mapParagraphs_(paras, res) {
  for (var i = 0; i < paras.length; i++) {
    var p = paras[i];
    var style = pStyleOf_(p);
    var text = normalizeWs(paraText_(p));
    if (!text) continue;
    res.styleCounts[style] = (res.styleCounts[style] || 0) + 1;
    if (style === 'ACADIA-Title') res.title = res.title ? res.title + ' ' + text : text;
    else if (style === 'ACADIA-Author') res.authorBlock = res.authorBlock ? res.authorBlock + '\n' + text : text;
    else if (style === 'ACADIA-Keywords') res.keywords = stripKeywordsLabel_(text);
    else if (style === 'ACADIA-Reference') res.rawRefs.push(text);
  }
}

function fallbackReferences_(paras) {
  var out = [], collecting = false;
  for (var i = 0; i < paras.length; i++) {
    var style = pStyleOf_(paras[i]) || '';
    var text = normalizeWs(paraText_(paras[i]));
    if (!text) continue;
    if (!collecting) { if (HEADING_REF_RX.test(text)) collecting = true; continue; }
    if (TERMINATOR_RX.test(text) ||
        style.indexOf('ACADIA-HeaderA') === 0 || style.indexOf('ACADIA-HeaderB') === 0 ||
        style.indexOf('Heading') === 0 ||
        style === 'ACADIA-FigureCaption' || style === 'ACADIA-TableCaption') break;
    out.push(text);
  }
  return out;
}

function newExtractResult_() {
  return { title: '', authorBlock: '', keywords: '', rawRefs: [], styleCounts: {}, usedFallback: false };
}

function extractDocx(docxBlob) {
  var doc = readDocumentXml_(docxBlob);
  var body = doc.getRootElement().getChild('body', W);
  var res = newExtractResult_();
  var topParas = body.getChildren('p', W);

  mapParagraphs_(topParas, res);                   // common case: top-level paragraphs

  if (!res.rawRefs.length) {                         // tables / text boxes
    var nested = [];
    collectParagraphs_(body, nested);
    var tmp = newExtractResult_();
    mapParagraphs_(nested, tmp);
    if (tmp.rawRefs.length) {
      res.rawRefs = tmp.rawRefs;
      res.title = res.title || tmp.title;
      res.authorBlock = res.authorBlock || tmp.authorBlock;
      res.keywords = res.keywords || tmp.keywords;
    }
  }

  if (!res.rawRefs.length) {                          // author didn't apply the Reference style
    res.rawRefs = fallbackReferences_(topParas);
    res.usedFallback = res.rawRefs.length > 0;
  }

  res.titleFallback = false;                          // author didn't apply ACADIA-Title
  if (!res.title) {
    res.title = fallbackTitle_(topParas);
    res.titleFallback = !!res.title;
  }

  res.diagnostics = {
    styledRefCount: res.styleCounts['ACADIA-Reference'] || 0,
    usedFallback: res.usedFallback,
    titleFound: !!res.title,
    titleFallback: res.titleFallback,
    authorFound: !!res.authorBlock
  };
  return res;
}


// ---------- RefParser.gs ----------

// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 Daniel Koehler / ACADIA
/**
 * RefParser.gs — split one Chicago reference string into {authors, year, title, source}.
 *
 * Faithful port of the Python prototype (prototype/acadia_index.py), which is regression-
 * tested against the real Examples corpus (6 template refs + Conway/Yang/Mostafavi). The
 * parser ANCHORS ON THE YEAR (never on capitalization) so it handles both Chicago variants:
 *   - author-date:        "Authors. YYYY. Title. Source."
 *   - notes-bibliography: "Authors. Title. Source, YYYY."  /  "... (YYYY): pages."
 * Anything uncertain is flagged so the author can fix it in the confirmation UI.
 */

// Tokens after which a period is NOT a sentence boundary (abbreviations + we also treat any
// single-letter initial as guarded).
var ABBR = (function () {
  var o = {};
  ('ed eds vol no nos pp p diss al inc ltd co corp st trans rev eg ie cf et jr sr dr prof ' +
   'mr mrs ms vs etc fig figs eq eqs ch chap repr').split(' ').forEach(function (w) { o[w] = 1; });
  return o;
})();

var QUOTED_SPAN = /[“"]([^”"]{2,})[”"]/;     // “...” or "..."
var NA_RX       = /^\s*(n\.?\s*a\.?|n\.?\s*d\.?|no date|forthcoming|in press)/i;
var ACCESSED_RX = /accessed/i;
var URL_RX      = /https?:\/\/\S+/i;

// Flags that strongly highlight a row as "needs review". Low-severity info flags (e.g. an
// unquoted book title, which is normal) are excluded so the highlight stays meaningful.
var REVIEW_FLAGS = {
  year_not_found: 1, verify_year: 1, verify_segmentation: 1, authors_empty: 1,
  title_missing: 1, source_empty: 1, verify_long_span: 1, verify_author_has_url: 1,
  empty: 1, verify_title_quotes: 1
};

function needsReview(flags) {
  for (var i = 0; i < flags.length; i++) if (REVIEW_FLAGS[flags[i]]) return true;
  return false;
}

function normalizeWs(s) {
  if (s == null) return '';
  return String(s).replace(/\s+/g, ' ').trim();
}

/** All 4-digit years 1500-2099 with optional range + paren detection, in order. */
function findYears(s) {
  var rx = /\b(?:1[5-9]|20)\d{2}(?:\s*[–—-]\s*\d{2,4})?\b/g;
  var out = [], m;
  while ((m = rx.exec(s)) !== null) {
    var start = m.index, end = m.index + m[0].length;
    var paren = (s.charAt(start - 1) === '(' && s.charAt(end) === ')');
    out.push({ text: m[0], start: start, end: end, paren: paren, year4: m[0].slice(0, 4) });
  }
  return out;
}

/** Pick the publication year: parenthesized (YYYY) -> early ". YYYY." -> last bare year. */
function chooseYear(years, s) {
  if (!years.length) return null;
  for (var i = 0; i < years.length; i++) if (years[i].paren) return years[i];
  var EARLY = 120;
  for (var j = 0; j < years.length; j++) {
    if (years[j].start <= EARLY) {
      var prev = s.substring(Math.max(0, years[j].start - 2), years[j].start);
      if (prev.charAt(prev.length - 1) === '.' || prev.slice(-2) === '. ') return years[j];
    }
  }
  return years[years.length - 1];
}

function isGuardedPeriod(prefix) {
  var m = /([A-Za-z][A-Za-z\.]*)\.$/.exec(prefix);
  if (!m) return false;
  var base = m[1].replace(/[^A-Za-z]/g, '');
  if (base.length === 1) return true;          // single-letter initial "A." "S."
  if (ABBR[base.toLowerCase()]) return true;    // "ed." "no." "et al." "diss."
  return false;
}

/** Split into sentence segments on a real ". " (or '." ') boundary, respecting quotes and
 *  abbreviations/initials. Returns [{t, a, b}] with char spans. */
function sentenceSegments(s) {
  var segs = [], inQ = false, start = 0, i = 0, n = s.length;
  var openClose = '“”‘’"\'';
  var closers = '”’"\'';
  while (i < n) {
    var ch = s.charAt(i);
    if (openClose.indexOf(ch) >= 0) { inQ = !inQ; i++; continue; }
    if (ch === '.' && !inQ) {
      var j = i + 1, close = 0;
      if (j < n && closers.indexOf(s.charAt(j)) >= 0) { close = 1; j++; }
      if (j >= n || s.charAt(j) === ' ') {
        var end = i + 1 + close;
        if (!isGuardedPeriod(s.substring(start, i + 1))) {
          var seg = s.substring(start, end).trim();
          if (seg) segs.push({ t: seg, a: start, b: end });
          start = end; i = end; continue;
        }
      }
    }
    i++;
  }
  var tail = s.substring(start).trim();
  if (tail) segs.push({ t: tail, a: start, b: n });
  return segs;
}

function isJustYear(segText, year) {
  var core = segText.replace(/[^0-9]/g, '');
  return segText.length <= 8 && core.slice(0, 4) === year.year4;
}

/** Author-date iff the chosen year stands right after the authors as a ". YYYY." token:
 *  (A) the year is its own 2nd segment, or (B) the year is the trailing token of segment 0
 *  (author block ended in an initial, so the year did not split into its own segment). */
function isAuthorDate(year, segs) {
  if (!year || !segs.length) return false;
  var ys = year.start;
  if (segs.length >= 2 && segs[1].a <= ys && ys < segs[1].b && isJustYear(segs[1].t, year)) return true;
  var s0 = segs[0];
  if (s0.a <= ys && ys < s0.b) {
    var trailing = s0.t.substring(ys - s0.a);
    if (/^\d{4}\b/.test(trailing) && trailing.length <= 8) return true;
  }
  return false;
}

function cleanAuthors(text) {
  var a = String(text).trim().replace(/^,+/, '').replace(/,+$/, '').trim();
  if (!/\b[A-Z]\.$/.test(a)) a = a.replace(/[.\s]+$/, '').trim();   // keep an initial's period
  a = a.replace(/,?\s*(ed\.|eds\.)$/, '').trim();
  return a;
}

function cleanTitle(text) {
  var t = String(text).trim();
  t = t.replace(/^[“”‘’"']+/, '').replace(/[“”‘’"']+$/, '').trim();
  t = t.replace(/\.+$/, '').trim();
  return t;
}

function cleanSource(text) {
  return String(text).trim().replace(/^\.+/, '').trim();
}

function parseReference(raw) {
  var s = normalizeWs(raw);
  var flags = [];
  if (!s) return { authors: '', year: '', title: '', source: '', flags: ['empty'] };

  var years = findYears(s);
  var year = chooseYear(years, s);
  var segs = sentenceSegments(s);
  var qm = QUOTED_SPAN.exec(s);

  // ---- year string ----
  var yearStr;
  if (year) {
    yearStr = year.text.slice(0, 4) + year.text.slice(4).replace(/\s*[–—-]\s*/g, '-');
  } else {
    yearStr = '';
    flags.push(NA_RX.test(s) ? 'year_nonstandard' : 'year_not_found');
  }

  var authorDate = isAuthorDate(year, segs);
  var authors, title, source;

  if (authorDate) {
    authors = cleanAuthors(s.substring(0, year.start));
    var after = year.end;
    if (qm && qm.index >= after) {
      title = cleanTitle(qm[0]);
      source = cleanSource(s.substring(qm.index + qm[0].length));
    } else {
      var tail = segs.filter(function (g) { return g.a >= after; });
      if (tail.length) {
        title = cleanTitle(tail[0].t);
        source = cleanSource(s.substring(tail[0].b));
      } else {
        title = cleanTitle(s.substring(after));
        source = '';
        flags.push('source_empty');
      }
      flags.push('title_unquoted_guess');
    }
  } else {
    if (qm) {
      authors = cleanAuthors(s.substring(0, qm.index));
      title = cleanTitle(qm[0]);
      source = cleanSource(s.substring(qm.index + qm[0].length));
    } else if (segs.length >= 2) {
      authors = cleanAuthors(segs[0].t);
      title = cleanTitle(segs[1].t);
      source = cleanSource(s.substring(segs[1].b));
      flags.push('title_unquoted_guess');
      if (segs.length < 3) flags.push('verify_segmentation');
    } else {
      authors = segs.length ? cleanAuthors(segs[0].t) : '';
      title = '';
      source = '';
      flags.push('verify_segmentation');
    }
  }

  // ---- confidence flags ----
  if (!authors) { authors = 'Unknown'; flags.push('authors_empty'); }
  if (!title) flags.push('title_missing');
  if (!source && flags.indexOf('source_empty') < 0) flags.push('source_empty');
  if (yearStr === '') {
    if (flags.indexOf('verify_year') < 0) flags.push('verify_year');
  } else if (!authorDate) {
    var same = 0;
    for (var k = 0; k < years.length; k++) if (years[k].year4 === year.year4) same++;
    if (ACCESSED_RX.test(s) && same >= 2) flags.push('verify_year');
  }
  if (URL_RX.test(authors)) flags.push('verify_author_has_url');
  if (title.length > 300 || authors.length > 200) flags.push('verify_long_span');

  // de-dup flags, keep order
  var seen = {}, out = [];
  for (var f = 0; f < flags.length; f++) if (!seen[flags[f]]) { seen[flags[f]] = 1; out.push(flags[f]); }

  return { authors: authors, year: yearStr, title: title, source: source, flags: out };
}


// ---------- NameSplit.gs ----------

// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 Daniel Koehler / ACADIA
/**
 * NameSplit.gs — split the ACADIA-Author block (display order, e.g. "Jane Q. Smith, John
 * Doe") into [{first, last, flags}] pairs for the AUTHOR n FIRST/LAST columns.
 *
 * Default rule: last whitespace token = LAST, the rest = FIRST. This is the inverse of the
 * 2024 bug (which lumped middle names into LAST). Every field is editable in the UI, and
 * genuinely ambiguous cases (particles, suffixes, single-token names) are flagged.
 */

var NAME_PARTICLES = (function () {
  var o = {};
  'van von de del della der da di la le bin al ter ten dos das'.split(' ').forEach(function (w) { o[w] = 1; });
  return o;
})();

var NAME_SUFFIXES = { jr: 1, sr: 1, ii: 1, iii: 1, iv: 1, phd: 1 };

function stripAffiliations(text) {
  return String(text)
    .replace(/[\*†‡¹²³]+/g, '')           // footnote / affiliation markers
    .replace(/([A-Za-z])\d+\b/g, '$1');    // trailing affiliation digits after a name
}

function splitOneName(name) {
  name = String(name).trim();
  var parts = name.split(/\s+/).filter(function (p) { return p.length; });
  var flags = [];
  if (!parts.length) return { first: '', last: '', flags: ['empty'] };
  if (parts.length === 1) return { first: '', last: parts[0], flags: ['single_token_name'] };

  var lastIdx = parts.length - 1;
  var lastBare = parts[lastIdx].toLowerCase().replace(/\./g, '');
  if (NAME_SUFFIXES[lastBare] && parts.length >= 3) {
    flags.push('verify_name_suffix');
    return {
      first: parts.slice(0, lastIdx - 1).join(' '),
      last: parts[lastIdx - 1] + ' ' + parts[lastIdx],
      flags: flags
    };
  }

  var first = parts.slice(0, -1);
  for (var i = 0; i < first.length; i++) {
    if (NAME_PARTICLES[first[i].toLowerCase().replace(/\./g, '')]) { flags.push('verify_name_particle'); break; }
  }
  return { first: first.join(' '), last: parts[parts.length - 1], flags: flags };
}

function splitAuthorBlock(text) {
  if (!text) return [];
  var t = stripAffiliations(text);
  var tokens = t.split(/\s*(?:,|;|\band\b|&|\n)\s*/i);
  var out = [];
  for (var i = 0; i < tokens.length; i++) {
    var tok = (tokens[i] || '').trim();
    if (tok) out.push(splitOneName(tok));
  }
  return out;
}


// ---------- SheetIO.gs ----------

// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 Daniel Koehler / ACADIA
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
function savePaperFile_(blob, id, originalName) {
  var folder = getOrCreateUploadsFolder_();
  var clean = String(originalName || 'paper').replace(/[\\\/:*?"<>|]+/g, '_');
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


// ---------- Code.gs ----------

// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 Daniel Koehler / ACADIA
/**
 * Code.gs — web-app entry point + client-callable endpoints.
 *
 * doGet renders the single-page UI. The browser calls processUpload (save + parse) then
 * commitSubmission (write) via google.script.run. runTests() lets an editor verify the
 * parser inside Apps Script after deployment (mirrors the Python regression suite).
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ACADIA 2026 Reference Indexer')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** Non-secret config the UI needs to preview the ID / SOURCE and populate dropdowns. */
function getConfigForClient() {
  var cfg = getConfig();
  var vols = {};
  Object.keys(cfg.VOLUMES).forEach(function (v) {
    vols[v] = { type: cfg.VOLUMES[v].type, sourcePreview: cfg.VOLUMES[v].sourceBase };
  });
  return {
    volumes: vols,
    types: cfg.TYPES,
    idPrefix: cfg.ID_PREFIX,
    idPad: cfg.ID_PAD,
    confYear: cfg.CONF_YEAR,
    maxAuthors: cfg.MAX_AUTHORS
  };
}

/** Stage A -> save the .docx to Drive and return the parsed (editable) draft. */
function processUpload(payload) {
  try {
    var cfg = getConfig();
    var meta = payload.meta || {};
    var vol = String(meta.vol || '');
    if (!cfg.VOLUMES[vol]) return { ok: false, error: 'Please choose a valid volume/track.' };
    if (meta.number == null || isNaN(parseInt(meta.number, 10))) {
      return { ok: false, error: 'Please enter your assigned Paper ID.' };
    }
    if (!payload.base64) return { ok: false, error: 'No file was received.' };

    var id = buildPaperId(cfg, vol, meta.number);
    var type = (meta.type && String(meta.type).trim()) || typeForVolume(cfg, vol);

    var blob = Utilities.newBlob(Utilities.base64Decode(payload.base64),
                 payload.mimeType || 'application/octet-stream', payload.filename || ('paper_' + id));
    var fileId = savePaperFile_(blob, id, payload.filename);   // store the upload (.docx or .pdf)

    // A PDF arrives pre-extracted from the browser (pdf.js); a .docx is unzipped + style-mapped here.
    var ex;
    if (payload.preExtracted) {
      ex = payload.preExtracted;
      ex.rawRefs = ex.rawRefs || [];
      ex.authorBlock = ex.authorBlock || '';
      ex.diagnostics = ex.diagnostics || { source: 'pdf' };
    } else {
      ex = extractDocx(blob);
      ex.diagnostics.source = 'docx';
    }
    var allAuthors = splitAuthorBlock(ex.authorBlock);
    var refs = ex.rawRefs.map(function (raw) {
      var r = parseReference(raw);
      r.review = needsReview(r.flags);   // drive the UI "needs review" highlight
      return r;
    });

    return {
      ok: true,
      meta: {
        id: id, vol: vol, number: String(meta.number), type: type,
        email: String(meta.email || '').trim(), source: sourceForVolume(cfg, vol), driveFileId: fileId
      },
      title: ex.title || '',
      keywords: ex.keywords || '',
      authors: allAuthors.slice(0, cfg.MAX_AUTHORS),
      authorOverflow: Math.max(0, allAuthors.length - cfg.MAX_AUTHORS),
      references: refs,
      diagnostics: ex.diagnostics
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/** Stage B -> validate the reviewed object and write the rows. */
function commitSubmission(reviewed) {
  try {
    var v = validateSubmission_(reviewed);
    if (!v.ok) return v;
    return writeSubmission_(reviewed);
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function validateSubmission_(r) {
  if (!r || !r.meta) return { ok: false, error: 'Missing submission data.' };
  var cfg = getConfig();
  var idRx = new RegExp('^' + cfg.ID_PREFIX.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '_v[12]_\\d{1,4}$');
  if (!idRx.test(r.meta.id || '')) {
    return { ok: false, error: 'Invalid paper ID: ' + r.meta.id };
  }
  if (cfg.TYPES.indexOf(String(r.meta.type || '').trim()) < 0) {
    return { ok: false, error: 'Invalid paper type.' };
  }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(r.meta.email || '').trim())) {
    return { ok: false, error: 'Please enter a valid contact email.' };
  }
  if (!r.title || !String(r.title).trim()) return { ok: false, error: 'Title is required.' };
  var hasLast = (r.authors || []).some(function (a) { return a && a.last && String(a.last).trim(); });
  if (!hasLast) return { ok: false, error: 'At least one author with a last name is required.' };
  return { ok: true };
}

// ---------------------------------------------------------------------------------------
// In-GAS regression test — run from the Apps Script editor (Run > runTests) after porting.
// Mirrors prototype/test_acadia.py. Unicode via \u escapes so it is encoding-independent.
// ---------------------------------------------------------------------------------------
function runTests() {
  var fails = [];
  function check(desc, cond) { if (!cond) fails.push(desc); Logger.log((cond ? 'OK  ' : 'FAIL') + ' ' + desc); }
  function P(raw) { return parseReference(raw); }

  var p;
  p = P('Fox, Michael and Miles Kemp. 2009. Interactive Architecture. New York: Princeton Architectural Press.');
  check('Fox authors', p.authors === 'Fox, Michael and Miles Kemp');
  check('Fox year', p.year === '2009');
  check('Fox title', p.title === 'Interactive Architecture');

  p = P('Brooks, Rodney A. 1990. “Elephants Don’t Play Chess.” Robotics and Autonomous Systems 6 (1): 3–15.');
  check('Brooks authors (no year leak)', p.authors === 'Brooks, Rodney A.');
  check('Brooks year', p.year === '1990');
  check('Brooks title', p.title === 'Elephants Don’t Play Chess');

  p = P('Cremers, Jan. 2011. “Energy Saving Design of Membrane Building Envelope.” In International Conference … Structural Membranes 2011, ed. E. Oñate. 148–155. Barcelona: CIMNE.');
  check('Cremers not falsely verify_year', p.flags.indexOf('verify_year') < 0);
  check('Cremers year', p.year === '2011');

  p = P('Bazjanac, Vladimir. IFC BIM-Based Methodology for Semi-Automated Building Energy Performance Simulation. Lawrence Berkeley National Laboratory, 2008.');
  check('Bazjanac authors', p.authors === 'Bazjanac, Vladimir');
  check('Bazjanac year', p.year === '2008');
  check('Bazjanac title', p.title.indexOf('IFC BIM-Based Methodology') === 0);

  p = P('Zhou, Y. W., Z. Z. Hu, J. R. Lin, et al. “A Review on 3D Spatial Data Analytics for Building Information Models.” Archives of Computational Methods in Engineering 27 (2020): 1449–1463.');
  check('Zhou year (paren)', p.year === '2020');
  check('Zhou keeps et al.', p.authors.indexOf('et al') >= 0);
  check('Zhou title quoted', p.title.indexOf('A Review on 3D Spatial Data') === 0);

  p = P('Google. Gemini 1.5 Pro in Google Docs. 2025. Accessed June 15, 2025.');
  check('Google authors', p.authors === 'Google');
  check('Google title', p.title === 'Gemini 1.5 Pro in Google Docs');
  check('Google verify_year flagged', p.flags.indexOf('verify_year') >= 0);

  p = P('Mäntylä, Martti. An Introduction to Solid Modeling. Rockville, MD: Computer Science Press, 1987.');
  check('Mantyla keeps umlaut', p.authors.indexOf('Mäntylä') >= 0);
  check('Mantyla year', p.year === '1987');

  p = P('Jones, Alice. 2014–18. “Longitudinal Study.” Design Journal 5 (2).');
  check('year range -> hyphen', p.year === '2014-18');

  var ns = splitAuthorBlock('Oliver Thomas Hamedinger and Jade Bailey');
  check('name split count', ns.length === 2);
  check('middle name -> FIRST', ns[0].first === 'Oliver Thomas' && ns[0].last === 'Hamedinger');
  var pn = splitOneName('Theo van Doesburg');
  check('particle flagged', pn.flags.indexOf('verify_name_particle') >= 0);

  var summary = (fails.length ? 'FAILED ' + fails.length + ': ' + fails.join(' | ')
                              : 'ALL PASSED');
  Logger.log(summary);
  return summary;
}
