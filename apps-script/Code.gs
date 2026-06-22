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
      return { ok: false, error: 'Please enter your assigned paper number.' };
    }
    if (!payload.base64) return { ok: false, error: 'No file was received — please choose your .docx.' };

    var id = buildPaperId(cfg, vol, meta.number);
    var type = (meta.type && String(meta.type).trim()) || typeForVolume(cfg, vol);

    var bytes = Utilities.base64Decode(payload.base64);
    var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream',
                                 payload.filename || 'paper.docx');

    var fileId = saveDocx_(blob, id, payload.filename);   // capture upload before parsing

    var ex = extractDocx(blob);
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
      title: ex.title,
      keywords: ex.keywords,
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
