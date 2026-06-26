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
