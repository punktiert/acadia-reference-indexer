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
