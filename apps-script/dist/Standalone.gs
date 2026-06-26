// ===== ACADIA Reference Indexer — combined build. Edit the modular .gs files, re-run build.py. =====


// ---------- Config.gs ----------

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


// ---------- standalone page (overrides doGet; no separate HTML file needed) ----------
var STANDALONE_PAGE_HTML = "<!DOCTYPE html>\n<html>\n<head>\n  <base target=\"_top\">\n  <meta charset=\"utf-8\">\n  <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n  <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>\n  <link href=\"https://fonts.googleapis.com/css2?family=Roboto+Mono:wght@400;500;700&family=Roboto:wght@400;500;700&display=swap\" rel=\"stylesheet\">\n  <script src=\"https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js\"></script>\n  <script>window.pdfjsLib && (pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js');</script>\n  <style>\n  /* ===== ACADIA 2026 \u2014 Humanism Recoded visual identity =====\n     pure black/white \u00b7 oxblood accent (#933231) \u00b7 Roboto + Roboto Mono \u00b7 dotted-grid texture */\n  :root {\n    --ink: #0a0a0a; --paper: #ffffff; --accent: #933231; --accent-700: #6f2625;\n    --muted: #707070; --line: #e6e6e6; --line-2: #d6d6d6; --ink-soft: #2a2a2a;\n    --review: #b06a12; --review-bg: #fdf6ec; --review-line: #e0992f;\n    --ok: #0a7d52; --ok-bg: #f0fbf6; --danger: #b3261e; --danger-bg: #fdf1f0;\n    --mono: 'Roboto Mono', ui-monospace, 'SFMono-Regular', Menlo, monospace;\n    --sans: 'Roboto', system-ui, -apple-system, sans-serif;\n  }\n  * { box-sizing: border-box; }\n  html { -webkit-text-size-adjust: 100%; }\n  body {\n    font-family: var(--sans); color: var(--ink); background: var(--paper);\n    margin: 0; line-height: 1.55; font-size: 15px;\n  }\n\n  /* ---- brand top bar (echoes the conference hero) ---- */\n  .topbar {\n    background: var(--ink);\n    background-image: radial-gradient(rgba(255,255,255,.18) 1px, transparent 1.5px);\n    background-size: 15px 15px;\n    border-bottom: 3px solid var(--accent);\n    color: #fff;\n  }\n  .topbar-inner {\n    max-width: 1060px; margin: 0 auto; padding: 14px 22px;\n    display: flex; align-items: baseline; gap: 16px; flex-wrap: wrap;\n  }\n  .wordmark { font-family: var(--mono); font-weight: 700; font-size: 17px; letter-spacing: .14em; }\n  .wordmark b { color: var(--accent); font-weight: 700; }\n  .tagline { font-family: var(--mono); font-size: 11px; letter-spacing: .14em; text-transform: uppercase; color: rgba(255,255,255,.6); }\n\n  /* ---- layout ---- */\n  .wrap { max-width: 1060px; margin: 0 auto; padding: 30px 22px 90px; }\n  .lede { margin: 6px 0 26px; }\n  .eyebrow { font-family: var(--mono); font-size: 11px; letter-spacing: .16em; text-transform: uppercase; color: var(--accent); margin-bottom: 8px; }\n  h1 { font-family: var(--mono); font-size: 30px; font-weight: 700; letter-spacing: -.01em; margin: 0 0 8px; line-height: 1.08; }\n  .sub { color: var(--muted); margin: 0; max-width: 64ch; }\n  h2 {\n    font-family: var(--mono); font-size: 13px; font-weight: 700; text-transform: uppercase;\n    letter-spacing: .12em; margin: 30px 0 12px; padding-bottom: 8px; border-bottom: 1px solid var(--ink);\n  }\n  h2::before { content: \"\u258d \"; color: var(--accent); }\n\n  /* ---- cards ---- */\n  .card { border: 1px solid var(--line-2); border-top: 3px solid var(--ink); border-radius: 3px; padding: 22px; margin-bottom: 18px; background: #fff; }\n\n  /* ---- form controls ---- */\n  label { display: block; font-family: var(--mono); font-weight: 500; font-size: 11px; text-transform: uppercase; letter-spacing: .08em; margin: 0 0 6px; color: var(--ink); }\n  label .hint, .hint { font-family: var(--sans); text-transform: none; letter-spacing: 0; color: var(--muted); font-weight: 400; font-size: 12px; }\n  .hint em { font-style: normal; font-family: var(--mono); color: var(--ink-soft); }\n  input, select, textarea {\n    width: 100%; padding: 10px 11px; border: 1px solid var(--line-2); border-radius: 2px;\n    font-family: var(--sans); font-size: 14px; color: var(--ink); background: #fff;\n  }\n  input:focus, select:focus, textarea:focus { outline: none; border-color: var(--ink); box-shadow: 0 0 0 3px rgba(147,50,49,.20); }\n  textarea { resize: vertical; min-height: 40px; line-height: 1.45; }\n  .row { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }\n  .field { margin-bottom: 16px; }\n\n  /* ---- ID / SOURCE preview ---- */\n  .preview { border: 1px dashed var(--line-2); border-left: 3px solid var(--accent); border-radius: 2px; padding: 11px 13px; font-size: 13px; background: #fafafa; }\n  .preview code { font-family: var(--mono); font-size: 14px; color: var(--accent); font-weight: 700; letter-spacing: .02em; }\n  .preview .src { color: var(--muted); display: block; margin-top: 6px; word-break: break-word; font-size: 12px; }\n\n  /* ---- buttons ---- */\n  button {\n    font-family: var(--mono); font-weight: 700; font-size: 13px; text-transform: uppercase; letter-spacing: .06em;\n    border: 1px solid var(--accent); background: var(--accent); color: #fff; padding: 11px 20px; border-radius: 2px; cursor: pointer;\n    transition: background .12s, color .12s, border-color .12s;\n  }\n  button:hover { background: var(--accent-700); border-color: var(--accent-700); }\n  button:disabled { opacity: .5; cursor: default; }\n  button.ghost { background: #fff; color: var(--ink); border-color: var(--ink); }\n  button.ghost:hover { background: var(--ink); color: #fff; }\n  button.tiny { padding: 5px 11px; font-size: 11px; letter-spacing: .05em; }\n  button.danger { border-color: var(--line-2); background: #fff; color: var(--danger); }\n  button.danger:hover { background: var(--danger); border-color: var(--danger); color: #fff; }\n\n  /* ---- tables ---- */\n  table { width: 100%; border-collapse: collapse; margin-top: 10px; }\n  th { text-align: left; font-family: var(--mono); font-size: 10.5px; text-transform: uppercase; letter-spacing: .08em; color: var(--muted); padding: 7px 8px; border-bottom: 2px solid var(--ink); font-weight: 500; }\n  td { padding: 6px 7px; vertical-align: top; border-bottom: 1px solid var(--line); }\n  td.num { font-family: var(--mono); color: var(--muted); font-size: 12px; padding-top: 13px; width: 28px; }\n  .col-year input { font-family: var(--mono); }\n  tr.review td { background: var(--review-bg); }\n  tr.review td.num { border-left: 3px solid var(--review-line); color: var(--review); }\n  .flagtip { color: var(--review); font-size: 11px; margin-top: 4px; font-family: var(--mono); letter-spacing: .02em; }\n  /* resolved (edited or confirmed) flagged rows */\n  tr.resolved td { background: var(--ok-bg); }\n  tr.resolved td.num { border-left: 3px solid var(--ok); color: var(--ok); }\n  tr.resolved .flagtip, tr.resolved .confirmline { color: var(--ok); }\n  .confirmline { display: flex; align-items: center; gap: 6px; margin-top: 6px; font-family: var(--mono); font-size: 11px; color: var(--review); letter-spacing: .02em; cursor: pointer; }\n  .confirmline input { width: auto; margin: 0; accent-color: var(--accent); }\n  .confirmline.done { color: var(--ok); }\n  /* empty-field block on submit */\n  input.empty, textarea.empty { border-color: var(--danger); box-shadow: 0 0 0 3px rgba(179,38,30,.14); }\n\n  /* ---- banners ---- */\n  .banner { padding: 11px 14px; border-radius: 2px; font-size: 13px; margin-bottom: 12px; border: 1px solid; }\n  .banner.warn { background: var(--review-bg); color: var(--review); border-color: var(--review-line); }\n  .banner.err { background: var(--danger-bg); color: var(--danger); border-color: #e6b3ae; }\n  .banner.ok { background: var(--ok-bg); color: var(--ok); border-color: #a6e9c5; }\n\n  .counts { color: var(--muted); font-size: 12px; margin: 8px 2px; font-family: var(--mono); letter-spacing: .03em; }\n  .toolbar { display: flex; gap: 10px; align-items: center; margin-top: 18px; flex-wrap: wrap; }\n  .spinner { display: inline-block; width: 14px; height: 14px; border: 2px solid #fff; border-top-color: transparent; border-radius: 50%; animation: spin .7s linear infinite; vertical-align: -2px; margin-right: 7px; }\n  @keyframes spin { to { transform: rotate(360deg); } }\n  [hidden] { display: none !important; }\n  .col-year { width: 88px; } .col-rm { width: 36px; }\n  ::selection { background: var(--accent); color: #fff; }\n  @media (max-width: 720px) { .row { grid-template-columns: 1fr; } h1 { font-size: 24px; } .wrap { padding-top: 20px; } }\n</style>\n\n</head>\n<body>\n<header class=\"topbar\">\n  <div class=\"topbar-inner\">\n    <span class=\"wordmark\">ACADIA <b>2026</b></span>\n    <span class=\"tagline\">Humanism Recoded \u2014 Detroit</span>\n  </div>\n</header>\n<div class=\"wrap\">\n\n  <div class=\"lede\">\n    <div class=\"eyebrow\">// Proceedings reference indexing</div>\n    <h1>Reference Indexer</h1>\n    <p class=\"sub\">Upload your final paper (.docx). We read your references and prepare the\n      CumInCAD index for you to review and confirm \u2014 takes a couple of minutes.</p>\n  </div>\n\n  <!-- ================= STAGE A: upload ================= -->\n  <div id=\"stageA\" class=\"card\">\n    <div class=\"row\">\n      <div class=\"field\">\n        <label>Track <span class=\"hint\">(sets your volume &amp; type)</span></label>\n        <select id=\"track\"></select>\n      </div>\n      <div class=\"field\">\n        <label>Paper ID <span class=\"hint\">(your EasyChair submission number)</span></label>\n        <input id=\"number\" type=\"number\" min=\"1\" step=\"1\" placeholder=\"e.g. 42\">\n        <div class=\"hint\" style=\"margin-top:5px\">\n          Find it in EasyChair \u2192 <em>My Submissions</em>: the number in the <strong>#</strong> column next to your paper.\n        </div>\n      </div>\n    </div>\n\n    <div class=\"field\">\n      <label>Main contact email</label>\n      <input id=\"email\" type=\"email\" placeholder=\"you@university.edu\">\n    </div>\n\n    <div class=\"field\">\n      <label>Your paper <span class=\"hint\">(Word <strong>.docx</strong> recommended \u2014 most accurate; <strong>.pdf</strong> also accepted)</span></label>\n      <input id=\"file\" type=\"file\" accept=\".docx,.pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf\">\n      <div class=\"hint\" style=\"margin-top:5px\">A .docx formatted with the ACADIA template styles gives the cleanest result. A PDF works too, but references are extracted approximately \u2014 you'll review them on the next screen.</div>\n    </div>\n\n    <div class=\"preview field\">\n      Your CumInCAD ID will be <code id=\"idPreview\">\u2014</code> &nbsp;\u00b7&nbsp; type <span id=\"typePreview\">\u2014</span>\n      <span class=\"src\" id=\"srcPreview\"></span>\n    </div>\n\n    <div id=\"bannerA\"></div>\n    <div class=\"toolbar\">\n      <button id=\"uploadBtn\">Upload &amp; parse references</button>\n    </div>\n  </div>\n\n  <!-- ================= STAGE B: confirm ================= -->\n  <div id=\"stageB\" class=\"card\" hidden>\n    <div id=\"banners\"></div>\n\n    <div class=\"field\">\n      <label>Paper ID</label>\n      <div class=\"preview\"><code id=\"idB\">\u2014</code> &nbsp;\u00b7&nbsp; <span id=\"typeB\"></span></div>\n    </div>\n\n    <div class=\"field\">\n      <label>Title</label>\n      <textarea id=\"title\" rows=\"2\"></textarea>\n    </div>\n\n    <h2>Authors</h2>\n    <p class=\"hint\">Last name = the surname for the index. Fix any split below \u2014 middle names\n      belong in \u201cFirst\u201d.</p>\n    <table id=\"authorsTable\">\n      <thead><tr><th>First</th><th>Last</th><th class=\"col-rm\"></th></tr></thead>\n      <tbody></tbody>\n    </table>\n    <div class=\"toolbar\"><button class=\"ghost tiny\" id=\"addAuthor\">+ add author</button></div>\n\n    <h2>References</h2>\n    <div class=\"counts\" id=\"refCounts\"></div>\n    <table id=\"refsTable\">\n      <thead><tr><th class=\"col-rm\"></th><th>Authors</th><th class=\"col-year\">Year</th><th>Title</th><th>Source</th><th class=\"col-rm\"></th></tr></thead>\n      <tbody></tbody>\n    </table>\n    <div class=\"toolbar\"><button class=\"ghost tiny\" id=\"addRef\">+ add reference</button></div>\n\n    <div id=\"bannerB\"></div>\n    <div class=\"toolbar\">\n      <button id=\"submitBtn\">Confirm &amp; submit to index</button>\n      <button class=\"ghost\" id=\"backBtn\">Start over</button>\n    </div>\n  </div>\n\n  <div id=\"done\" hidden></div>\n\n</div>\n<script>\n(function () {\n  'use strict';\n\n  var cfg = null;     // from getConfigForClient\n  var draft = null;   // from processUpload\n  var titleNeedsConfirm = false, titleConfirm = null, titleOrig = '';\n\n  var FLAG_TEXT = {\n    verify_year: 'check year', year_not_found: 'no year found', verify_segmentation: 'check author/title split',\n    authors_empty: 'no authors found', title_missing: 'no title found', source_empty: 'no source found',\n    title_unquoted_guess: 'title was guessed', verify_author_has_url: 'URL in authors?',\n    verify_long_span: 'unusually long', verify_title_quotes: 'check quotes', year_nonstandard: 'unusual year',\n    single_token_name: 'one-word name', verify_name_particle: 'check name particle', verify_name_suffix: 'check suffix'\n  };\n\n  // ---- tiny DOM helpers ----\n  function $(id) { return document.getElementById(id); }\n  function el(tag, attrs, kids) {\n    var n = document.createElement(tag);\n    if (attrs) Object.keys(attrs).forEach(function (k) {\n      if (k === 'class') n.className = attrs[k];\n      else if (k === 'text') n.textContent = attrs[k];\n      else n.setAttribute(k, attrs[k]);\n    });\n    (kids || []).forEach(function (c) { n.appendChild(c); });\n    return n;\n  }\n  function banner(host, kind, msg) {\n    $(host).appendChild(el('div', { class: 'banner ' + kind, text: msg }));\n  }\n\n  // ---- server bridge (real Apps Script, or a mock for the preview/demo) ----\n  function serverCall(fn, arg, onOk, onErr) {\n    if (window.google && google.script && google.script.run) {        // deployed Apps Script\n      google.script.run.withSuccessHandler(onOk).withFailureHandler(function (e) {\n        onErr(e && e.message ? e.message : String(e));\n      })[fn](arg);\n      return;\n    }\n    // Local real backend (python) if one is serving /api; otherwise the static mock (demo.html).\n    fetch('/api/' + fn, { method: 'POST', headers: { 'Content-Type': 'application/json' },\n                          body: JSON.stringify(arg || {}) })\n      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })\n      .then(function (d) { onOk(d); })\n      .catch(function () { setTimeout(function () { onOk(MOCK[fn](arg)); }, 200); }); // demo fallback\n  }\n\n  var MOCK = {\n    getConfigForClient: function () {\n      return {\n        volumes: {\n          '1': { type: 'project', sourcePreview: 'ACADIA 2026: Humanism Recoded [Project Catalog \u2026]' },\n          '2': { type: 'paper', sourcePreview: 'ACADIA 2026: Humanism Recoded [Proceedings of the 46th Annual Conference \u2026]' }\n        }, types: ['paper', 'project'], idPrefix: 'acadia26', idPad: 2, maxAuthors: 8\n      };\n    },\n    processUpload: function (p) {\n      return {\n        ok: true,\n        meta: { id: 'acadia26_v' + p.meta.vol + '_' + ('0' + p.meta.number).slice(-2), vol: p.meta.vol,\n                number: String(p.meta.number), type: 'paper', email: p.meta.email, source: 'ACADIA 2026: \u2026' },\n        title: 'A Demonstration Paper on Computational Design',\n        keywords: 'demo, parsing', authorOverflow: 0,\n        authors: [{ first: 'Jane Q.', last: 'Smith', flags: [] }, { first: 'John', last: 'Doe', flags: [] }],\n        references: [\n          { authors: 'Fox, Michael and Miles Kemp', year: '2009', title: 'Interactive Architecture', source: 'New York: Princeton Architectural Press.', flags: ['title_unquoted_guess'], review: false },\n          { authors: 'Brooks, Rodney A.', year: '1990', title: 'Elephants Don\u2019t Play Chess', source: 'Robotics and Autonomous Systems 6 (1): 3\u201315.', flags: [], review: false },\n          { authors: 'Google', year: '2025', title: 'Gemini 1.5 Pro in Google Docs', source: '2025. Accessed June 15, 2025.', flags: ['verify_year'], review: true }\n        ],\n        diagnostics: { styledRefCount: 3, usedFallback: false, titleFound: true, authorFound: true }\n      };\n    },\n    commitSubmission: function (r) {\n      return { ok: true, paperId: r.meta.id, titleRowsWritten: 1, refRowsWritten: (r.references || []).length };\n    }\n  };\n\n  // ============================================================================\n  // PDF extraction (client-side, via pdf.js) \u2014 ported from prototype/dryrun_pdf.py.\n  // Produces {title, authorBlock, rawRefs} from the laid-out PDF; the SERVER then runs the\n  // same parseReference()/splitAuthorBlock() it uses for .docx. Less reliable than .docx\n  // (multi-column layout) \u2014 the review/gating step is the safety net.\n  // ============================================================================\n  var PDF_JUNK = /(computing for resil|^acadia\\b|acadia team will fill|^\\d{1,3}$|^topic\\b|et al\\.\\s*\\|)/i;\n  var PDF_HEADING = /^(references|bibliography|works cited)\\b/i;\n  var PDF_BIO = /\\b(is a|is an|is currently|is the head|holds a|received (his|her)|researches and teaches|is a university professor|earned (his|her))\\b/i;\n  var PDF_ENDS = /\\.[\u201d\u2019\"\u2019\ufffd\u00ad]?\\s*$/;\n  var PDF_AUTHOR_START = /^(?:[A-Z][A-Za-z'\u2019.\\-]+,|[A-Z][A-Za-z&'\u2019\\-]+\\.|[A-Z][A-Za-z'\u2019\\-]+ et al\\.)/;\n  var PDF_BOILER = /(leave this|do not delete|topic\\s*\\(acadia|acadia team will fill|computing for resil|^abstract\\b|^keywords\\b|^references\\b)/i;\n\n  function lineText(ln) {\n    var t = '';\n    for (var i = 0; i < ln.length; i++) {\n      if (i > 0 && (ln[i].x0 - ln[i - 1].x1) > 1) t += ' ';\n      t += ln[i].str;\n    }\n    return t.replace(/\\s+/g, ' ').trim();\n  }\n  function centralCut(items, x0, x1, W) {\n    var width = x1 - x0;\n    var sub = items.filter(function (w) { return w.x0 >= x0 - 1 && w.x1 <= x1 + 1; });\n    if (width < 170 || sub.length < 12) return null;\n    var BINS = 160, bw = width / BINS, cover = [];\n    for (var b = 0; b < BINS; b++) cover[b] = 0;\n    sub.forEach(function (w) {\n      var lo = Math.max(0, Math.floor((w.x0 - x0) / bw)), hi = Math.min(BINS - 1, Math.floor((w.x1 - x0) / bw));\n      for (var k = lo; k <= hi; k++) cover[k]++;\n    });\n    var maxc = Math.max.apply(null, cover) || 1, thr = 0.04 * maxc;\n    var loB = Math.floor(0.30 * BINS), hiB = Math.floor(0.70 * BINS);\n    var bestLen = 0, bestMid = null, k = loB;\n    while (k < hiB) {\n      if (cover[k] <= thr) { var j = k; while (j < hiB && cover[j] <= thr) j++; if (j - k > bestLen) { bestLen = j - k; bestMid = (k + j) / 2; } k = j; }\n      else k++;\n    }\n    var minBins = Math.max(2, Math.floor((0.008 * W) / bw));\n    return (bestMid !== null && bestLen >= minBins) ? (x0 + bestMid * bw) : null;\n  }\n  function columnCuts(items, W) {\n    if (items.length < 20) return [];\n    var cuts = [];\n    (function rec(x0, x1, depth) {\n      if (depth > 1) return;\n      var c = centralCut(items, x0, x1, W);\n      if (c === null) return;\n      rec(x0, c, depth + 1); cuts.push(c); rec(c, x1, depth + 1);\n    })(0, W, 0);\n    return cuts.sort(function (a, b) { return a - b; });\n  }\n  function colOf(x, cuts) { for (var i = 0; i < cuts.length; i++) if (x < cuts[i]) return i; return cuts.length; }\n  function pageLines(items, W) {\n    var cuts = columnCuts(items, W), out = [];\n    for (var ci = 0; ci <= cuts.length; ci++) {\n      var cw = items.filter(function (w) { return colOf(w.x0, cuts) === ci; });\n      cw.sort(function (a, b) { return (Math.round(a.y / 3) - Math.round(b.y / 3)) || (a.x0 - b.x0); });\n      var cur = [], cy = null, lines = [];\n      cw.forEach(function (w) {\n        if (cy === null || Math.abs(w.y - cy) <= 3) { cur.push(w); if (cy === null) cy = w.y; }\n        else { lines.push(cur); cur = [w]; cy = w.y; }\n      });\n      if (cur.length) lines.push(cur);\n      lines.forEach(function (ln) { ln.sort(function (a, b) { return a.x0 - b.x0; }); out.push(lineText(ln)); });\n    }\n    return out;\n  }\n  function itemsToLines(items) {\n    var byY = {};\n    items.forEach(function (w) { var key = Math.round(w.y / 3); (byY[key] = byY[key] || []).push(w); });\n    return Object.keys(byY).map(function (key) {\n      var ln = byY[key].sort(function (a, b) { return a.x0 - b.x0; });\n      return { y: ln[0].y, size: Math.max.apply(null, ln.map(function (w) { return w.size; })), text: lineText(ln) };\n    }).sort(function (a, b) { return a.y - b.y; });\n  }\n  function extractTitleAuthors(pageItems) {\n    var best = null;\n    for (var pi = 0; pi < Math.min(2, pageItems.length); pi++) {\n      itemsToLines(pageItems[pi]).forEach(function (L) {\n        if (L.y > 360 || L.text.length < 6 || PDF_BOILER.test(L.text)) return;\n        if (best === null || L.size > best.size + 0.5) best = { size: L.size, text: L.text, page: pi, y: L.y };\n      });\n    }\n    var authorBlock = '';\n    if (best) {\n      var below = itemsToLines(pageItems[best.page]).filter(function (L) { return L.y > best.y && L.y < best.y + 120; });\n      var auth = [];\n      for (var i = 0; i < below.length && auth.length < 3; i++) {\n        if (/^(abstract|keywords|introduction)\\b/i.test(below[i].text)) break;\n        if (below[i].text.length >= 2 && below[i].size < best.size) auth.push(below[i].text);\n      }\n      authorBlock = auth.join(', ');\n    }\n    return { title: best ? best.text : '', authorBlock: authorBlock };\n  }\n  function joinEntry(lines) {\n    var s = '';\n    for (var i = 0; i < lines.length; i++) {\n      if (i === 0) s = lines[i];\n      else if (s.charAt(s.length - 1) === '\u00ad') s = s.slice(0, -1) + lines[i];\n      else if (s.charAt(s.length - 1) === '-') s = s + lines[i];\n      else s = s + ' ' + lines[i];\n    }\n    return s.replace(/\\s+/g, ' ').trim();\n  }\n  function referencesFromLines(lines) {\n    var start = null, inline = null;\n    for (var i = 0; i < lines.length; i++) {\n      var m = lines[i].match(PDF_HEADING);\n      if (m) { start = i + 1; var rest = lines[i].slice(m[0].length).trim(); inline = rest || null; }\n    }\n    if (start === null) return [];\n    var body = (inline ? [inline] : []).concat(lines.slice(start).filter(function (l) { return l && !PDF_JUNK.test(l); }));\n    var entries = [], cur = [];\n    for (var b = 0; b < body.length; b++) {\n      var l = body[b];\n      if (PDF_BIO.test(l) && !PDF_AUTHOR_START.test(l) && !/\\b(19|20)\\d{2}\\b/.test(l)) break;\n      var isStart = PDF_AUTHOR_START.test(l) && (cur.length === 0 || PDF_ENDS.test(cur[cur.length - 1]));\n      if (isStart && cur.length) { entries.push(cur); cur = [l]; } else cur.push(l);\n    }\n    if (cur.length) entries.push(cur);\n    return entries.map(joinEntry);\n  }\n  function arrayBufferToBase64(buf) {\n    var bytes = new Uint8Array(buf), CH = 0x8000, bin = '';\n    for (var i = 0; i < bytes.length; i += CH) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));\n    return btoa(bin);\n  }\n  function extractPdfClient(buf) {\n    if (!window.pdfjsLib) return Promise.reject(new Error('PDF reader (pdf.js) did not load \u2014 check your connection, or upload the .docx.'));\n    return pdfjsLib.getDocument({ data: buf }).promise.then(function (pdf) {\n      var pageItems = [], allLines = [], seq = Promise.resolve();\n      function doPage(pno) {\n        return pdf.getPage(pno).then(function (page) {\n          var vp = page.getViewport({ scale: 1 });\n          return page.getTextContent().then(function (tc) {\n            var items = tc.items.filter(function (it) { return it.str && it.str.trim(); }).map(function (it) {\n              var x0 = it.transform[4], yb = it.transform[5];          // pdf.js: origin bottom-left\n              return { str: it.str, x0: x0, x1: x0 + (it.width || 0), y: vp.height - yb, size: it.height || 0 };\n            });\n            pageItems.push(items);                                     // for title/authors (first pages)\n            var body = items.filter(function (w) { return w.y > 40 && w.y < vp.height - 40; });\n            allLines = allLines.concat(pageLines(body, vp.width));\n          });\n        });\n      }\n      for (var p = 1; p <= pdf.numPages; p++) { (function (pp) { seq = seq.then(function () { return doPage(pp); }); })(p); }\n      return seq.then(function () {\n        var ta = extractTitleAuthors(pageItems);\n        return {\n          title: ta.title, authorBlock: ta.authorBlock, keywords: '',\n          rawRefs: referencesFromLines(allLines),\n          diagnostics: { source: 'pdf', styledRefCount: 0, usedFallback: true,\n                         titleFound: !!ta.title, titleFallback: true, authorFound: !!ta.authorBlock }\n        };\n      });\n    });\n  }\n\n  // ---- Stage A ----\n  function pad(n) {\n    var s = String(parseInt(n, 10) || '');\n    while (s && cfg && s.length < cfg.idPad) s = '0' + s;\n    return s;\n  }\n  function updatePreview() {\n    if (!cfg) return;\n    var vol = $('track').value, num = $('number').value;\n    var v = cfg.volumes[vol] || {};\n    $('idPreview').textContent = num ? (cfg.idPrefix + '_v' + vol + '_' + pad(num)) : '\u2014';\n    $('typePreview').textContent = v.type || '\u2014';\n    $('srcPreview').textContent = v.sourcePreview ? ('SOURCE: ' + v.sourcePreview) : '';\n  }\n\n  function initStageA() {\n    var sel = $('track');\n    Object.keys(cfg.volumes).forEach(function (vol) {\n      var v = cfg.volumes[vol];\n      var label = v.type.charAt(0).toUpperCase() + v.type.slice(1) + ' \u2014 vol ' + vol;\n      sel.appendChild(el('option', { value: vol, text: label }));\n    });\n    sel.value = '2'; // default: paper\n    sel.addEventListener('change', updatePreview);\n    $('number').addEventListener('input', updatePreview);\n    $('uploadBtn').addEventListener('click', doUpload);\n    updatePreview();\n  }\n\n  function doUpload() {\n    $('bannerA').innerHTML = '';\n    var file = $('file').files[0];\n    var email = $('email').value.trim();\n    var num = $('number').value;\n    if (!num) return banner('bannerA', 'err', 'Please enter your Paper ID.');\n    if (!/^[^@\\s]+@[^@\\s]+\\.[^@\\s]+$/.test(email)) return banner('bannerA', 'err', 'Please enter a valid contact email.');\n    if (!file) return banner('bannerA', 'err', 'Please choose your .docx or .pdf file.');\n    var isPdf = /\\.pdf$/i.test(file.name), isDocx = /\\.docx$/i.test(file.name);\n    if (!isPdf && !isDocx) return banner('bannerA', 'err', 'Please upload a Word .docx (recommended) or a PDF.');\n\n    var btn = $('uploadBtn');\n    btn.disabled = true;\n    btn.innerHTML = '<span class=\"spinner\"></span>Reading your paper\u2026';\n    var meta = { vol: $('track').value, number: num, email: email };\n    function done() { btn.disabled = false; btn.textContent = 'Upload & parse references'; }\n    function ok(data) { done(); if (!data || !data.ok) return banner('bannerA', 'err', (data && data.error) || 'Upload failed.'); draft = data; renderStageB(data); }\n    function fail(msg) { done(); banner('bannerA', 'err', 'Upload failed: ' + msg); }\n\n    var reader = new FileReader();\n    reader.onerror = function () { fail('could not read the file.'); };\n    if (isDocx) {\n      reader.onload = function () {\n        var res = reader.result, base64 = res.substring(res.indexOf(',') + 1);\n        serverCall('processUpload', {\n          base64: base64, filename: file.name,\n          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', meta: meta\n        }, ok, fail);\n      };\n      reader.readAsDataURL(file);\n    } else {                                          // PDF: extract in-browser, then send\n      reader.onload = function () {\n        var buf = reader.result, base64;\n        try { base64 = arrayBufferToBase64(buf); }     // BEFORE pdf.js detaches the buffer\n        catch (e) { return fail('could not read the PDF.'); }\n        extractPdfClient(buf).then(function (pre) {\n          serverCall('processUpload', {\n            preExtracted: pre, base64: base64,\n            filename: file.name, mimeType: 'application/pdf', meta: meta\n          }, ok, fail);\n        }).catch(function (e) { fail((e && e.message) || String(e)); });\n      };\n      reader.readAsArrayBuffer(file);\n    }\n  }\n\n  // ---- Stage B ----\n  function renderStageB(data) {\n    $('banners').innerHTML = '';\n    $('bannerB').innerHTML = '';\n    var d = data.diagnostics || {};\n    if (d.source === 'pdf') {\n      banner('banners', 'warn', 'You uploaded a PDF \u2014 extracting references from a PDF is approximate, so please check every row carefully. (A Word .docx gives the most accurate result.)');\n    } else if (d.usedFallback || d.styledRefCount === 0) {\n      banner('banners', 'warn', 'We could not find references in the ACADIA-Reference style, so we detected them automatically \u2014 please check the list carefully.');\n    }\n    if (d.titleFallback) banner('banners', 'warn', 'We detected your title automatically (the ACADIA-Title style wasn\u2019t applied) \u2014 please verify it.');\n    else if (!d.titleFound) banner('banners', 'warn', 'We could not detect your title \u2014 please enter it below.');\n    if (!d.authorFound) banner('banners', 'warn', 'We could not detect author names in the document \u2014 please add them below.');\n    if (data.authorOverflow > 0) banner('banners', 'warn', 'More than ' + cfg.maxAuthors + ' authors were detected; only the first ' + cfg.maxAuthors + ' are kept (CumInCAD limit).');\n\n    $('idB').textContent = data.meta.id;\n    $('typeB').textContent = 'type: ' + data.meta.type;\n    $('title').value = data.title || '';\n    titleNeedsConfirm = !!d.titleFallback;\n    setupTitleConfirm(data.title || '');\n\n    var atb = $('authorsTable').getElementsByTagName('tbody')[0];\n    atb.innerHTML = '';\n    var authors = (data.authors && data.authors.length) ? data.authors : [{ first: '', last: '' }];\n    authors.forEach(function (a) { atb.appendChild(authorRow(a)); });\n\n    var rtb = $('refsTable').getElementsByTagName('tbody')[0];\n    rtb.innerHTML = '';\n    (data.references || []).forEach(function (r) { rtb.appendChild(refRow(r)); });\n    renumber();\n    updateCounts();\n\n    $('stageA').hidden = true;\n    $('stageB').hidden = false;\n    $('addAuthor').onclick = function () { atb.appendChild(authorRow({ first: '', last: '' })); };\n    $('addRef').onclick = function () { rtb.appendChild(refRow({ authors: '', year: '', title: '', source: '', review: false })); renumber(); updateCounts(); };\n    $('submitBtn').onclick = doSubmit;\n    $('backBtn').onclick = function () { location.reload(); };\n    window.scrollTo(0, 0);\n  }\n\n  function inputCell(value, cls, multiline) {\n    var node = multiline ? el('textarea', { rows: '2' }) : el('input', { type: 'text' });\n    node.value = value || '';\n    if (cls) node.className = cls;\n    return el('td', null, [node]);\n  }\n  function removeCell(onClick) {\n    var b = el('button', { class: 'danger tiny', title: 'remove', text: '\u00d7' });\n    b.addEventListener('click', onClick);\n    return el('td', { class: 'col-rm' }, [b]);\n  }\n\n  function authorRow(a) {\n    var tr = el('tr');\n    tr.appendChild(inputCell(a.first, null, false));\n    tr.appendChild(inputCell(a.last, null, false));\n    tr.appendChild(removeCell(function () { tr.parentNode.removeChild(tr); }));\n    if (a.flags && a.flags.length) tr.className = 'review';\n    return tr;\n  }\n\n  function refRow(r) {\n    var tr = el('tr');\n    tr._flagged = !!(r.flags && r.flags.length);   // ANY flag (incl. \"title was guessed\") needs edit-or-confirm\n    tr.appendChild(el('td', { class: 'num' }));                       // # (filled by renumber)\n    var authorsTd = inputCell(r.authors, null, true);\n    if (r.flags && r.flags.length) {\n      var tips = r.flags.map(function (f) { return FLAG_TEXT[f] || f; }).join(' \u00b7 ');\n      authorsTd.appendChild(el('div', { class: 'flagtip', text: '\u26a0 ' + tips }));\n    }\n    if (tr._flagged) {\n      var line = el('label', { class: 'confirmline' });\n      var cb = el('input', { type: 'checkbox' });\n      cb.addEventListener('change', function () { refreshRow(tr); });\n      line.appendChild(cb);\n      line.appendChild(el('span', { text: 'edit a field, or tick to confirm it\u2019s correct' }));\n      tr._confirm = cb;\n      authorsTd.appendChild(line);\n    }\n    tr.appendChild(authorsTd);\n    tr.appendChild(inputCell(r.year, 'col-year', false));\n    tr.appendChild(inputCell(r.title, null, true));\n    tr.appendChild(inputCell(r.source, null, true));\n    tr.appendChild(removeCell(function () { tr.parentNode.removeChild(tr); renumber(); updateCounts(); }));\n\n    // the four editable fields, for empty-checks and edit detection\n    tr._fields = [tr.cells[1].querySelector('textarea'), tr.cells[2].querySelector('input'),\n                  tr.cells[3].querySelector('textarea'), tr.cells[4].querySelector('textarea')];\n    tr._fields.forEach(function (inp) {\n      inp._orig = inp.value;\n      inp.addEventListener('input', function () { inp.classList.remove('empty'); refreshRow(tr); });\n    });\n    refreshRow(tr);\n    return tr;\n  }\n\n  function renumber() {\n    var rows = $('refsTable').getElementsByTagName('tbody')[0].rows;\n    for (var i = 0; i < rows.length; i++) rows[i].cells[0].textContent = (i + 1);\n  }\n  function updateCounts() {\n    var rows = $('refsTable').getElementsByTagName('tbody')[0].rows;\n    var pending = 0;\n    for (var i = 0; i < rows.length; i++) if (rows[i]._flagged && !rowResolved(rows[i])) pending++;\n    $('refCounts').textContent = rows.length + ' reference' + (rows.length === 1 ? '' : 's') +\n      (pending ? ' \u00b7 ' + pending + ' still need a check (edit or tick)' : (rows.length ? ' \u00b7 all reviewed' : ''));\n  }\n\n  // ---- \"edit or confirm\" resolution for flagged rows / the auto-detected title ----\n  function rowEdited(tr) { return (tr._fields || []).some(function (i) { return i.value !== i._orig; }); }\n  function rowResolved(tr) {\n    if (!tr._flagged) return true;\n    return (tr._confirm && tr._confirm.checked) || rowEdited(tr);\n  }\n  function refreshRow(tr) {\n    if (!tr._flagged) return;\n    var ok = rowResolved(tr);\n    tr.classList.toggle('review', !ok);\n    tr.classList.toggle('resolved', ok);\n    updateCounts();\n  }\n  function setupTitleConfirm(orig) {\n    titleOrig = orig;\n    var old = document.getElementById('titleConfirmWrap'); if (old) old.remove();\n    var t = $('title'); t.classList.remove('empty');\n    if (!titleNeedsConfirm) { titleConfirm = null; return; }\n    var wrap = el('label', { class: 'confirmline', id: 'titleConfirmWrap' });\n    titleConfirm = el('input', { type: 'checkbox' });\n    wrap.appendChild(titleConfirm);\n    wrap.appendChild(el('span', { text: 'auto-detected \u2014 edit it, or tick to confirm it\u2019s right' }));\n    t.parentNode.appendChild(wrap);\n    titleConfirm.addEventListener('change', function () { wrap.classList.toggle('done', titleResolved()); });\n    t.addEventListener('input', function () { t.classList.remove('empty'); wrap.classList.toggle('done', titleResolved()); });\n  }\n  function titleResolved() {\n    if (!titleNeedsConfirm) return true;\n    return (titleConfirm && titleConfirm.checked) || ($('title').value !== titleOrig);\n  }\n\n  function doSubmit() {\n    $('bannerB').innerHTML = '';\n    document.querySelectorAll('.empty').forEach(function (e) { e.classList.remove('empty'); });\n\n    // ---- title: non-empty + (if auto-detected) edited or confirmed ----\n    var titleVal = $('title').value.trim();\n    if (!titleVal) { $('title').classList.add('empty'); return banner('bannerB', 'err', 'Please enter a title.'); }\n    if (!titleResolved()) return banner('bannerB', 'err', 'Please review the auto-detected title \u2014 edit it, or tick the confirm box.');\n\n    // ---- authors: each present row needs first AND last ----\n    var authors = [], authorIncomplete = 0;\n    var arows = $('authorsTable').getElementsByTagName('tbody')[0].rows;\n    for (var i = 0; i < arows.length; i++) {\n      var fi = arows[i].cells[0].firstChild, li = arows[i].cells[1].firstChild;\n      var f = fi.value.trim(), l = li.value.trim();\n      if (!f && !l) continue;                                  // empty row \u2192 ignore\n      if (!f) fi.classList.add('empty');\n      if (!l) li.classList.add('empty');\n      if (!f || !l) authorIncomplete++;\n      authors.push({ first: f, last: l });\n    }\n    if (!authors.length) return banner('bannerB', 'err', 'Please enter at least one author (first and last name).');\n    if (authorIncomplete) return banner('bannerB', 'err', 'Each author needs both a first and last name (highlighted).');\n\n    // ---- references: no empty fields + every flagged row resolved ----\n    var rrows = $('refsTable').getElementsByTagName('tbody')[0].rows;\n    var emptyN = 0, pendingN = 0, firstBad = null, refs = [];\n    for (var j = 0; j < rrows.length; j++) {\n      var tr = rrows[j];\n      var vals = tr._fields.map(function (inp) { return inp.value.trim(); });\n      var filled = vals.filter(function (v) { return v; }).length;\n      if (filled === 0) continue;                              // blank row \u2192 ignore\n      if (filled < 4) {\n        tr._fields.forEach(function (inp) { if (!inp.value.trim()) inp.classList.add('empty'); });\n        emptyN++; firstBad = firstBad || tr;\n      }\n      if (tr._flagged && !rowResolved(tr)) { pendingN++; firstBad = firstBad || tr; }\n      refs.push({ authors: vals[0], year: vals[1], title: vals[2], source: vals[3] });\n    }\n    if (!refs.length) return banner('bannerB', 'err', 'Please add at least one reference.');\n    if (emptyN || pendingN) {\n      var parts = [];\n      if (emptyN) parts.push(emptyN + ' reference' + (emptyN > 1 ? 's have' : ' has') + ' an empty field');\n      if (pendingN) parts.push(pendingN + ' flagged reference' + (pendingN > 1 ? 's' : '') + ' still need a check');\n      banner('bannerB', 'err', 'Can\u2019t submit yet \u2014 ' + parts.join('; ') + '. Fill every column and resolve the highlighted rows.');\n      if (firstBad) firstBad.scrollIntoView({ block: 'center', behavior: 'smooth' });\n      return;\n    }\n\n    var reviewed = { meta: draft.meta, title: titleVal, authors: authors, references: refs };\n    var btn = $('submitBtn');\n    btn.disabled = true; btn.innerHTML = '<span class=\"spinner\"></span>Submitting\u2026';\n    serverCall('commitSubmission', reviewed, function (res) {\n      btn.disabled = false; btn.textContent = 'Confirm & submit to index';\n      if (!res || !res.ok) {\n        if (res && res.error === 'duplicate_id') return banner('bannerB', 'err', 'This paper ID is already indexed. Contact the editors if you need to update it.');\n        return banner('bannerB', 'err', (res && res.error) || 'Submission failed.');\n      }\n      $('stageB').hidden = true;\n      var done = $('done'); done.hidden = false;\n      done.appendChild(el('div', { class: 'banner ok',\n        text: 'Saved \u2014 ' + res.paperId + ': 1 paper row and ' + res.refRowsWritten + ' reference rows added to the index. Thank you!' }));\n    }, function (msg) {\n      btn.disabled = false; btn.textContent = 'Confirm & submit to index';\n      banner('bannerB', 'err', 'Submission failed: ' + msg);\n    });\n  }\n\n  // ---- boot ----\n  serverCall('getConfigForClient', null, function (c) {\n    cfg = c; initStageA();\n  }, function (msg) {\n    banner('bannerA', 'err', 'Could not load configuration: ' + msg);\n  });\n})();\n</script>\n\n</body>\n</html>\n";
function doGet() {
  return HtmlService.createHtmlOutput(STANDALONE_PAGE_HTML)
    .setTitle('ACADIA 2026 Reference Indexer')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
