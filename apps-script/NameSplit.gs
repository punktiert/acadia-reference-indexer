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
