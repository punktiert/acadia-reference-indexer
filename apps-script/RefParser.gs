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
