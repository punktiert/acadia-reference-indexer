# -*- coding: utf-8 -*-
# SPDX-License-Identifier: AGPL-3.0-or-later
# Copyright (C) 2026 Daniel Koehler / ACADIA
"""
PDF extraction — reference implementation / algorithm spec for the client-side pdf.js port.

The deployed tool runs THIS algorithm in the browser (pdf.js); this PyMuPDF version is what we
iterate against the example PDFs. The hard part is the PDF, not the parser: laid-out papers have
multi-column text with no hanging indent / no blank lines, per-page gutters NOT at the midline
(pages are even 612pt vs 1224pt wide), a boilerplate cover page, running headers/footers, and
author bios after the references. We: detect columns per page by the widest whitespace gaps,
read columns left-to-right, segment entries by author-start-after-period, filter junk/bios, and
de-hyphenate — then hand each raw reference string to the SAME parser the .docx path uses.

    cd prototype && python dryrun_pdf.py            # set ACADIA_PDFS to point at a PDF folder
"""
import sys, io, os, re, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import fitz
from acadia_index import parse_reference, needs_review

FOLDER = os.environ.get(
    "ACADIA_PDFS",
    r"E:/Claude/Acadia-Indexing/Examples/Session 8_Day3_INTELLIGENCES-Models_Saldana")

JUNK = re.compile(r"(computing for resil|^acadia\b|acadia team will fill|^\d{1,3}$|^topic\b|et al\.\s*\|)", re.I)
HEADING = re.compile(r"^(references|bibliography|works cited)\b", re.I)
# author-bio section that follows references (prose, no year, no author-start)
BIO = re.compile(r"\b(is a|is an|is currently|is the head|holds a|received (his|her)|"
                 r"researches and teaches|is a university professor|earned (his|her))\b", re.I)
ENDS_SENTENCE = re.compile(r'\.[”’"”’�\xad]?\s*$')
# an entry begins with "Surname," / "Org." / "Name et al."
AUTHOR_START = re.compile(r"^(?:[A-Z][A-Za-z'’.\-]+,|[A-Z][A-Za-z&'’\-]+\.|[A-Z][A-Za-z'’\-]+ et al\.)")
BOILERPLATE = re.compile(r"(leave this|do not delete|topic\s*\(acadia|acadia team will fill|"
                         r"computing for resil|^abstract\b|^keywords\b|^references\b)", re.I)


def _central_cut(words, x0, x1, W):
    """x of the dominant central whitespace valley within [x0, x1], or None. Uses a vertical
    coverage projection (robust to a few words bridging the gutter)."""
    width = x1 - x0
    sub = [w for w in words if w[0] >= x0 - 1 and w[2] <= x1 + 1]
    if width < 170 or len(sub) < 12:
        return None
    BINS = 160
    bw = width / BINS
    cover = [0] * BINS
    for w in sub:
        lo = max(0, int((w[0] - x0) / bw)); hi = min(BINS - 1, int((w[2] - x0) / bw))
        for k in range(lo, hi + 1):
            cover[k] += 1
    thr = 0.04 * (max(cover) or 1)
    lo_b, hi_b = int(0.30 * BINS), int(0.70 * BINS)    # only the central band of this range
    best_len, best_mid, k = 0, None, lo_b
    while k < hi_b:
        if cover[k] <= thr:
            j = k
            while j < hi_b and cover[j] <= thr:
                j += 1
            if j - k > best_len:
                best_len, best_mid = j - k, (k + j) / 2.0
            k = j
        else:
            k += 1
    min_bins = max(2, int((0.008 * W) / bw))
    return (x0 + best_mid * bw) if (best_mid is not None and best_len >= min_bins) else None


def column_cuts(words, W):
    """Column boundary x's via recursive central-gutter splitting: split at the dominant central
    valley, then recurse into each half (depth ≤ 2 → up to 4 columns). Handles normal 2-column
    portrait pages and double-wide 4-column spreads; [] = single column."""
    if len(words) < 20:
        return []
    cuts = []

    def rec(x0, x1, depth):
        if depth > 1:
            return
        c = _central_cut(words, x0, x1, W)
        if c is None:
            return
        rec(x0, c, depth + 1)
        cuts.append(c)
        rec(c, x1, depth + 1)

    rec(0, W, 0)
    return sorted(cuts)


def _col_of(x, cuts):
    i = 0
    for c in cuts:
        if x < c:
            return i
        i += 1
    return i


def page_lines(page):
    """Lines of this page in reading order: column by column (left→right), each top→bottom."""
    H, W = page.rect.height, page.rect.width
    words = [w for w in page.get_text("words") if 40 < w[1] < H - 40]   # drop header/footer band
    cuts = column_cuts(words, W)
    out = []
    for ci in range(len(cuts) + 1):
        cw = [w for w in words if _col_of(w[0], cuts) == ci]
        cw.sort(key=lambda w: (round(w[1] / 3.0), w[0]))
        cur, cy, lines = [], None, []
        for w in cw:
            if cy is None or abs(w[1] - cy) <= 3:
                cur.append(w); cy = w[1] if cy is None else cy
            else:
                lines.append(cur); cur = [w]; cy = w[1]
        if cur:
            lines.append(cur)
        for ln in lines:
            ln.sort(key=lambda w: w[0])
            out.append(" ".join(w[4] for w in ln).strip())
    return out


def reading_order_lines(doc):
    out = []
    for pno in range(doc.page_count):
        out += page_lines(doc[pno])
    return out


def extract_title(doc):
    """Best-effort: the largest-font non-boilerplate text near the top of the first 2 pages."""
    best = None  # (size, text)
    for pno in range(min(2, doc.page_count)):
        for block in doc[pno].get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                txt = "".join(s["text"] for s in line["spans"]).strip()
                if len(txt) < 6 or BOILERPLATE.search(txt) or line["bbox"][1] > 360:
                    continue
                size = max((s["size"] for s in line["spans"]), default=0)
                if best is None or size > best[0] + 0.5:
                    best = (size, txt)
    return best[1] if best else ""


def join_entry(lines):
    s = ""
    for i, ln in enumerate(lines):
        if i == 0:
            s = ln
        elif s.endswith("\xad"):
            s = s[:-1] + ln                 # soft hyphen -> remove, no space
        elif s.endswith("-"):
            s = s + ln                      # compound hyphen (Semi-Automated) -> keep
        else:
            s = s + " " + ln
    return re.sub(r"\s+", " ", s).strip()


def references_from_lines(lines):
    start, inline = None, None
    for i, ln in enumerate(lines):
        m = HEADING.match(ln)
        if m:                                # take the LAST heading; may be merged with entry 1
            start = i + 1
            inline = ln[m.end():].strip() or None
    if start is None:
        return []
    body = ([inline] if inline else []) + [ln for ln in lines[start:] if ln and not JUNK.search(ln)]
    entries, cur = [], []
    for ln in body:
        if BIO.search(ln) and not AUTHOR_START.match(ln) and not re.search(r"\b(19|20)\d{2}\b", ln):
            break                            # stop at the author-bio section
        is_start = bool(AUTHOR_START.match(ln)) and (not cur or ENDS_SENTENCE.search(cur[-1]))
        if is_start and cur:
            entries.append(cur); cur = [ln]
        else:
            cur.append(ln)
    if cur:
        entries.append(cur)
    return [join_entry(e) for e in entries]


def extract_pdf(path):
    """Mirrors the client's output shape: {title, raw_refs}."""
    doc = fitz.open(path)
    res = {"title": extract_title(doc), "raw_refs": references_from_lines(reading_order_lines(doc))}
    doc.close()
    return res


def main():
    pdfs = sorted(glob.glob(os.path.join(FOLDER, "*.pdf")))
    total = review = 0
    for path in pdfs:
        ex = extract_pdf(path)
        parsed = [parse_reference(r) for r in ex["raw_refs"]]
        rev = sum(1 for p in parsed if needs_review(p["flags"]))
        total += len(parsed); review += rev
        print("\n" + "=" * 100)
        print(f"{os.path.basename(path)}  —  {len(parsed)} refs, {rev} flagged | title: {ex['title'][:58]!r}")
        print("=" * 100)
        for p in parsed[:3]:
            mark = "   <-- REVIEW" if needs_review(p["flags"]) else ""
            print(f"  {p['authors'][:58]!r} | {p['year']!r} | {p['title'][:50]!r}")
            print(f"     source: {p['source'][:80]!r}{mark}")
        if len(parsed) > 3:
            print(f"  … {len(parsed) - 3} more")
    print("\n" + "#" * 100)
    print(f"TOTAL: {total} refs across {len(pdfs)} papers; {review} flagged, {total - review} clean.")


if __name__ == "__main__":
    main()
