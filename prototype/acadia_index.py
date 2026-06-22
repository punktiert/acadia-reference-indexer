"""
ACADIA reference indexing — offline Python prototype.

This mirrors the logic that will be ported to Google Apps Script (.gs). It exists to
de-risk and regression-test the two hard parts against the real Examples corpus before
porting:

  * extract_docx(path)      -> {title, author_block, keywords, raw_refs[]}  (style-based)
  * parse_reference(raw)    -> {authors, year, title, source, flags[]}      (year-anchored)
  * split_author_block(txt) -> [{first, last, flags}]                        (display order)

Design reference: ../please-help-me-to-buzzing-conway-agent-ae9b9b0a9a8d94f07.md
"""

import re
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# --------------------------------------------------------------------------------------
# Text normalization
# --------------------------------------------------------------------------------------

def normalize_ws(s):
    """Collapse runs of whitespace to single spaces; trim. Leaves unicode (smart quotes,
    en-dashes) untouched — \\s only matches whitespace."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", s).strip()


# --------------------------------------------------------------------------------------
# DOCX extraction  (mirrors Docx.gs)
# --------------------------------------------------------------------------------------

def _p_style(p):
    pPr = p.find(f"{W}pPr")
    if pPr is None:
        return None
    pStyle = pPr.find(f"{W}pStyle")
    if pStyle is None:
        return None
    return pStyle.get(f"{W}val")


def _collect_text(el, out):
    """Concatenate every w:t descendant in document order. A single reference paragraph
    can split across several runs (italic title, formatted year); a naive grab shatters
    references, so we must walk the whole subtree."""
    for c in el:
        tag = c.tag
        if tag == f"{W}t":
            out.append(c.text or "")
        elif tag == f"{W}tab":
            out.append("\t")
        elif tag in (f"{W}br", f"{W}cr"):
            out.append(" ")
        else:
            _collect_text(c, out)


def _para_text(p):
    out = []
    _collect_text(p, out)
    return "".join(out)


def _strip_keywords_label(text):
    return re.sub(r"^\s*keywords\s*[:\-]\s*", "", text, flags=re.I).strip()


def _all_paragraphs(body):
    """Recursive descent collecting every w:p (including inside tables / text boxes)."""
    for p in body.iter(f"{W}p"):
        yield p


def extract_docx(path):
    """Extract title / authors / keywords / references from a .docx by paragraph style."""
    with zipfile.ZipFile(path) as z:
        xml_bytes = z.read("word/document.xml")
    # Force UTF-8 so smart quotes / en-dashes survive (avoids the 2024 mojibake bug).
    root = ET.fromstring(xml_bytes.decode("utf-8"))
    body = root.find(f"{W}body")

    res = {
        "title": "",
        "author_block": "",
        "keywords": "",
        "raw_refs": [],
        "style_counts": {},
        "used_fallback": False,
    }

    # Top-level paragraphs first (the common, clean case).
    top_paras = body.findall(f"{W}p")
    _map_paragraphs(top_paras, res)

    # If no styled references at top level, also scan nested paragraphs (tables/text boxes).
    if not res["raw_refs"]:
        nested = [p for p in _all_paragraphs(body)]
        tmp = {"title": res["title"], "author_block": res["author_block"],
               "keywords": res["keywords"], "raw_refs": [], "style_counts": {},
               "used_fallback": False}
        _map_paragraphs(nested, tmp)
        if tmp["raw_refs"]:
            res["raw_refs"] = tmp["raw_refs"]
            res["title"] = res["title"] or tmp["title"]
            res["author_block"] = res["author_block"] or tmp["author_block"]
            res["keywords"] = res["keywords"] or tmp["keywords"]

    # Fallback: author never applied ACADIA-Reference -> scan for a References heading.
    if not res["raw_refs"]:
        res["raw_refs"] = _fallback_references(top_paras)
        res["used_fallback"] = bool(res["raw_refs"])

    # Fallback: author didn't apply ACADIA-Title -> recover the title from the first heading.
    res["title_fallback"] = False
    if not res["title"]:
        res["title"] = _fallback_title(top_paras)
        res["title_fallback"] = bool(res["title"])

    return res


def _map_paragraphs(paras, res):
    for p in paras:
        style = _p_style(p)
        text = normalize_ws(_para_text(p))
        if not text:
            continue
        res["style_counts"][style] = res["style_counts"].get(style, 0) + 1
        if style == "ACADIA-Title":
            res["title"] = (res["title"] + " " + text).strip() if res["title"] else text
        elif style == "ACADIA-Author":
            res["author_block"] = (res["author_block"] + "\n" + text) if res["author_block"] else text
        elif style == "ACADIA-Keywords":
            res["keywords"] = _strip_keywords_label(text)
        elif style == "ACADIA-Reference":
            res["raw_refs"].append(text)


_HEADING_REF_RX = re.compile(r"^(references|bibliography|works cited)\s*$", re.I)
_TERMINATOR_RX = re.compile(
    r"^(image credits|acknowledg(e)?ments?|author bio|biograph|appendix"
    r"|figure\s*\d|fig\.?\s*\d|table\s*\d)", re.I)

# Section labels that share the title's heading style but are NOT the title.
_SECTION_WORDS = {
    "abstract", "references", "bibliography", "works cited", "keywords", "introduction",
    "background", "methodology", "methods", "results", "discussion", "conclusion",
    "conclusions", "acknowledgments", "acknowledgements",
}


def _fallback_title(paras):
    """Title fallback when ACADIA-Title isn't applied: the first heading-styled paragraph
    (authors often use ACADIA-HeaderA Not-Numbered for the title) that isn't a section label,
    else the first non-empty paragraph that isn't a section label."""
    for want_heading in (True, False):
        for p in paras:
            st = _p_style(p) or ""
            t = normalize_ws(_para_text(p))
            if not t or t.strip().lower() in _SECTION_WORDS:
                continue
            if want_heading and not st.startswith("ACADIA-HeaderANot"):
                continue
            return t
    return ""


def _fallback_references(paras):
    """Locate a References heading and collect following paragraphs until a terminator."""
    out = []
    collecting = False
    for p in paras:
        style = _p_style(p) or ""
        text = normalize_ws(_para_text(p))
        if not text:
            continue
        if not collecting:
            if _HEADING_REF_RX.match(text):
                collecting = True
            continue
        # collecting:
        if (_TERMINATOR_RX.match(text)
                or style.startswith(("ACADIA-HeaderA", "ACADIA-HeaderB", "Heading"))
                or style in ("ACADIA-FigureCaption", "ACADIA-TableCaption")):
            break
        out.append(text)
    return out


# --------------------------------------------------------------------------------------
# Reference parser  (mirrors RefParser.gs)
# --------------------------------------------------------------------------------------

# 4-digit year 1500-2099, optional range "2014-18"/"1994-2015", optional surrounding parens.
YEAR_TOKEN = re.compile(r"(\()?\b((?:1[5-9]|20)\d{2})(\s*[–\-]\s*\d{2,4})?\b(\))?")
QUOTED_SPAN = re.compile(r"[“\"]([^”\"]{2,})[”\"]")  # “...” or "..."
NA_RX = re.compile(r"^\s*(n\.?\s*a\.?|n\.?\s*d\.?|no date|forthcoming|in press)", re.I)
ACCESSED_RX = re.compile(r"accessed", re.I)
URL_RX = re.compile(r"https?://\S+", re.I)

# Tokens after which a period is NOT a sentence boundary.
ABBR = set("ed eds vol no nos pp p diss al inc ltd co corp st trans rev eg ie cf et jr "
           "sr dr prof mr mrs ms vs etc fig figs eq eqs ch chap repr".split())


def _is_guarded_period(prefix):
    """prefix = s[start:idx+1] including the period at idx. Return True if this period is an
    abbreviation/initial (not a real sentence break)."""
    # token immediately before the period (letters/dots only)
    m = re.search(r"([A-Za-z][A-Za-z\.]*)\.$", prefix)
    if not m:
        return False
    tok = m.group(1)
    base = re.sub(r"[^A-Za-z]", "", tok)
    if len(base) == 1:          # single-letter initial: "A.", "S.", "Y.", "W."
        return True
    if base.lower() in ABBR:    # "ed.", "no.", "et al.", "diss.", "Inc."
        return True
    return False


def sentence_segments(s):
    """Split into sentence segments on a real '. ' (or '." ') boundary, respecting quotes and
    abbreviations/initials. Returns list of dicts {t, a, b} with char spans into s."""
    segs = []
    in_q = False
    start = 0
    i = 0
    n = len(s)
    quotes = "“”‘’\"'"
    while i < n:
        ch = s[i]
        if ch in quotes:
            in_q = not in_q
            i += 1
            continue
        if ch == "." and not in_q:
            j = i + 1
            # period may be followed by a closing quote, then a space/end
            close = 0
            if j < n and s[j] in "”’\"'":
                close = 1
                j += 1
            if j >= n or s[j] == " ":
                end = i + 1 + close
                if not _is_guarded_period(s[start:i + 1]):
                    seg = s[start:end].strip()
                    if seg:
                        segs.append({"t": seg, "a": start, "b": end})
                    start = end
                    i = end
                    continue
        i += 1
    tail = s[start:].strip()
    if tail:
        segs.append({"t": tail, "a": start, "b": n})
    return segs


def _choose_year(matches, s):
    """Pick the publication year. Preference: parenthesized (YYYY) -> early '. YYYY.'
    (author-date) -> last bare year (notes-bibliography)."""
    if not matches:
        return None
    # 1) parenthesized year like "(2020)"
    for m in matches:
        if m.group(1) or m.group(4):
            return m
        a, b = m.start(2), m.end()
        if a - 1 >= 0 and s[a - 1] == "(" and b < len(s) and s[b:b + 1] == ")":
            return m
    # 2) an early year that forms its own sentence segment (author-date)
    EARLY = 120
    for m in matches:
        if m.start() <= EARLY:
            prev = s[max(0, m.start() - 2):m.start()]
            if prev.endswith(". ") or prev.endswith("."):
                return m
    # 3) notes-bibliography: the last bare year
    return matches[-1]


def _clean_authors(text):
    a = text.strip().strip(",").strip()
    # drop a trailing sentence period, but keep an initial's period ("Rodney A.")
    if not re.search(r"\b[A-Z]\.$", a):
        a = a.rstrip(". ").strip()
    # drop a dangling editor marker accidentally swept in
    a = re.sub(r",?\s*(ed\.|eds\.)$", "", a).strip()
    return a


def _clean_title(text):
    t = text.strip()
    t = t.strip("“”‘’\"'").strip()
    t = t.rstrip(".").strip()       # strip one trailing period for consistency
    return t


def _clean_source(text):
    return text.strip().lstrip(".").strip()


def _is_just_year(seg_text, year_m):
    core = re.sub(r"[^0-9]", "", seg_text)
    return len(seg_text) <= 8 and core[:4] == year_m.group(2)


def _is_author_date(year_m, segs, s):
    """Author-date iff the chosen year stands right after the author block as a `. YYYY.`
    token. Two shapes:
      A) the year is its own 2nd sentence segment ("Authors. YYYY. Title.")
      B) the year is the trailing token of the 1st segment, because the author block ends in
         an initial ("Brooks, Rodney A. 1990." — the "A." period is guarded, so the year did
         not split off into its own segment)."""
    if year_m is None or not segs:
        return False
    ys = year_m.start(2)
    # Case A: year forms (begins) the 2nd segment.
    if len(segs) >= 2 and segs[1]["a"] <= ys < segs[1]["b"] and _is_just_year(segs[1]["t"], year_m):
        return True
    # Case B: year is the trailing token of segment 0 (author ended in an initial).
    s0 = segs[0]
    if s0["a"] <= ys < s0["b"]:
        trailing = s0["t"][ys - s0["a"]:]
        if re.match(r"^\d{4}\b", trailing) and len(trailing) <= 8:
            return True
    return False


def parse_reference(raw):
    s = normalize_ws(raw)
    flags = []
    if not s:
        return {"authors": "", "year": "", "title": "", "source": "", "flags": ["empty"]}

    matches = list(YEAR_TOKEN.finditer(s))
    year_m = _choose_year(matches, s)
    segs = sentence_segments(s)
    quote_m = QUOTED_SPAN.search(s)

    # ---- year string ----
    if year_m is not None:
        # normalize a range separator to a plain hyphen (2024 used "2014-18", not en-dash)
        rng = re.sub(r"\s*[–—-]\s*", "-", year_m.group(3) or "")
        year = year_m.group(2) + rng
    else:
        year = ""
        if NA_RX.match(s):
            flags.append("year_nonstandard")
        else:
            flags.append("year_not_found")

    is_author_date = _is_author_date(year_m, segs, s)

    # ---- authors / title / source ----
    if is_author_date:
        authors = _clean_authors(s[:year_m.start(2)])
        after = year_m.end()
        if quote_m and quote_m.start() >= after:
            title = _clean_title(quote_m.group(0))
            source = _clean_source(s[quote_m.end():])
        else:
            # title = first segment after the year segment
            tail_segs = [seg for seg in segs if seg["a"] >= after]
            if tail_segs:
                title = _clean_title(tail_segs[0]["t"])
                source = _clean_source(s[tail_segs[0]["b"]:])
            else:
                title = _clean_title(s[after:])
                source = ""
                flags.append("source_empty")
            flags.append("title_unquoted_guess")
    else:
        if quote_m:
            authors = _clean_authors(s[:quote_m.start()])
            title = _clean_title(quote_m.group(0))
            source = _clean_source(s[quote_m.end():])
        elif len(segs) >= 2:
            authors = _clean_authors(segs[0]["t"])
            title = _clean_title(segs[1]["t"])
            source = _clean_source(s[segs[1]["b"]:])
            flags.append("title_unquoted_guess")
            if len(segs) < 3:
                flags.append("verify_segmentation")
        else:
            authors = _clean_authors(segs[0]["t"]) if segs else ""
            title = ""
            source = ""
            flags.append("verify_segmentation")

    # ---- confidence flags ----
    if not authors:
        authors = "Unknown"
        flags.append("authors_empty")
    if not title:
        flags.append("title_missing")
    if not source:
        if "source_empty" not in flags:
            flags.append("source_empty")
    if year == "":
        if "verify_year" not in flags:
            flags.append("verify_year")
    elif not is_author_date:
        # Ambiguous date: online "Accessed" sources, or the same year repeated (e.g.
        # "...Docs. 2025. Accessed June 15, 2025.") where the publication year is unclear.
        same_year = sum(1 for m in matches if m.group(2) == year_m.group(2))
        if ACCESSED_RX.search(s) and same_year >= 2:
            flags.append("verify_year")
    if URL_RX.search(authors):
        flags.append("verify_author_has_url")
    if len(title) > 300 or len(authors) > 200:
        flags.append("verify_long_span")

    # de-dup flags, keep order
    seen = set()
    flags = [f for f in flags if not (f in seen or seen.add(f))]
    return {"authors": authors, "year": year, "title": title, "source": source, "flags": flags}


# Flags that should strongly highlight a row as "needs review" in the UI. Low-severity
# info flags (e.g. an unquoted book title, which is normal) are intentionally excluded so
# the highlight stays meaningful — otherwise nearly every book would light up.
REVIEW_FLAGS = {
    "year_not_found", "verify_year", "verify_segmentation", "authors_empty",
    "title_missing", "source_empty", "verify_long_span", "verify_author_has_url",
    "empty", "verify_title_quotes",
}


def needs_review(flags):
    return any(f in REVIEW_FLAGS for f in flags)


# --------------------------------------------------------------------------------------
# Author name first/last split  (mirrors NameSplit.gs)
# --------------------------------------------------------------------------------------

PARTICLES = {"van", "von", "de", "del", "della", "der", "da", "di", "la", "le", "bin", "al", "ter", "ten", "dos", "das"}
SUFFIXES = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "phd"}


def _strip_affiliations(text):
    # remove superscript-style markers and asterisks/daggers right after names
    text = re.sub(r"[\*†‡¹²³]+", "", text)
    text = re.sub(r"(?<=[A-Za-z])\d+\b", "", text)  # trailing affiliation digits
    return text


def split_one_name(name):
    name = name.strip()
    parts = name.split()
    flags = []
    if not parts:
        return {"first": "", "last": "", "flags": ["empty"]}
    if len(parts) == 1:
        return {"first": "", "last": parts[0], "flags": ["single_token_name"]}
    # suffix handling
    last_idx = len(parts) - 1
    if parts[last_idx].lower().strip(".") in {s.strip(".") for s in SUFFIXES} and len(parts) >= 3:
        last = parts[last_idx - 1] + " " + parts[last_idx]
        first = " ".join(parts[:last_idx - 1])
        flags.append("verify_name_suffix")
        return {"first": first, "last": last, "flags": flags}
    last = parts[-1]
    first = " ".join(parts[:-1])
    if any(p.lower().strip(".") in PARTICLES for p in parts[:-1]):
        flags.append("verify_name_particle")
    return {"first": first, "last": last, "flags": flags}


def split_author_block(text):
    if not text:
        return []
    t = _strip_affiliations(text)
    tokens = re.split(r"\s*(?:,|;|\band\b|&|\n)\s*", t, flags=re.I)
    names = [tok.strip() for tok in tokens if tok and tok.strip()]
    return [split_one_name(n) for n in names]


if __name__ == "__main__":
    import json, sys
    if len(sys.argv) > 1:
        out = extract_docx(sys.argv[1])
        print(json.dumps(out, ensure_ascii=False, indent=2))
