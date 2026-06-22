# -*- coding: utf-8 -*-
"""
Dry run: extract references from the example PDFs and run them through the parser.

NOTE: the real tool reads the author's .docx, where each reference is one clean paragraph.
Here we only have final PDFs, whose two-column layout flattens references into uniform lines
with NO paragraph breaks, NO hanging indent, and running headers injected mid-list. So this
script must re-segment entries heuristically (author-start after a sentence end) and de-wrap
columns — work the .docx path never needs. Treat segmentation glitches as a PDF artifact, not
a parser result; the 4-field split is what the parser actually does.
"""
import sys, io, os, re, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import fitz
from acadia_index import parse_reference, needs_review

FOLDER = r"E:/Claude/Acadia-Indexing/Examples/Session 8_Day3_INTELLIGENCES-Models_Saldana"

# running headers / footers / topic lines to drop if they land inside the reference list
JUNK = re.compile(r"(computing for resilience|^acadia\b|acadia team will fill|^\d{1,3}$|"
                  r"^topic\b|et al\.\s*\|)", re.I)
HEADING = re.compile(r"^(references|bibliography|works cited)\b", re.I)
# author-bio section that follows references in ACADIA papers (prose, no year, no author-start)
BIO = re.compile(r"\b(is a|is an|is currently|is the head|holds a|received (his|her)|"
                 r"researches and teaches|is a university professor|earned (his|her))\b", re.I)
ENDS_SENTENCE = re.compile(r'\.[”’"”’�\xad]?\s*$')
# an entry begins with "Surname," / "Org." / "Name et al."
AUTHOR_START = re.compile(r"^(?:[A-Z][A-Za-z'’.\-]+,|[A-Z][A-Za-z&'’\-]+\.|[A-Z][A-Za-z'’\-]+ et al\.)")


def reading_order_lines(doc):
    """All lines across pages in reading order: page by page, left column then right,
    top-to-bottom; margins dropped."""
    out = []
    for pno in range(doc.page_count):
        page = doc[pno]
        H, Wd = page.rect.height, page.rect.width
        words = [w for w in page.get_text("words") if 44 < w[1] < H - 52]
        mid = Wd / 2.0
        for col in (0, 1):
            cw = [w for w in words if (w[0] < mid) == (col == 0)]
            cw.sort(key=lambda w: (round(w[1] / 3.0), w[0]))  # group by ~line, then x
            cur, cury = [], None
            lines = []
            for w in cw:
                if cury is None or abs(w[1] - cury) <= 3:
                    cur.append(w); cury = w[1] if cury is None else cury
                else:
                    lines.append(cur); cur = [w]; cury = w[1]
            if cur:
                lines.append(cur)
            for ln in lines:
                ln.sort(key=lambda w: w[0])
                out.append(" ".join(w[4] for w in ln).strip())
    return out


def join_entry(lines):
    s = ""
    for i, ln in enumerate(lines):
        if i == 0:
            s = ln
        elif s.endswith("\xad"):
            s = s[:-1] + ln                 # soft hyphen -> remove, no space
        elif s.endswith("-"):
            s = s + ln                      # compound hyphen (Semi-Automated) -> keep, no space
        else:
            s = s + " " + ln
    return re.sub(r"\s+", " ", s).strip()


def extract_references(path):
    doc = fitz.open(path)
    lines = reading_order_lines(doc)
    doc.close()
    # find the LAST references heading (may be merged with the first entry on the same line)
    start, inline = None, None
    for i, ln in enumerate(lines):
        m = HEADING.match(ln)
        if m:
            start = i + 1
            rest = ln[m.end():].strip()
            inline = rest or None
    if start is None:
        return []
    body = ([inline] if inline else []) + [ln for ln in lines[start:] if ln and not JUNK.search(ln)]
    # segment into entries; stop when the author-bio section begins
    entries, cur = [], []
    for ln in body:
        if BIO.search(ln) and not AUTHOR_START.match(ln) and not re.search(r"\b(19|20)\d{2}\b", ln):
            break
        is_start = bool(AUTHOR_START.match(ln)) and (not cur or ENDS_SENTENCE.search(cur[-1]))
        if is_start and cur:
            entries.append(cur); cur = [ln]
        else:
            cur.append(ln)
    if cur:
        entries.append(cur)
    return [join_entry(e) for e in entries]


def main():
    pdfs = sorted(glob.glob(os.path.join(FOLDER, "*.pdf")))
    grand_total = grand_review = 0
    for path in pdfs:
        name = os.path.basename(path)
        refs = extract_references(path)
        parsed = [parse_reference(r) for r in refs]
        rev = sum(1 for p in parsed if needs_review(p["flags"]))
        grand_total += len(parsed); grand_review += rev
        print("\n" + "=" * 100)
        print(f"{name}  —  {len(parsed)} references extracted, {rev} flagged for review")
        print("=" * 100)
        for p in parsed[:4]:
            print(f"  authors : {p['authors'][:80]!r}")
            print(f"  year    : {p['year']!r}    title: {p['title'][:70]!r}")
            print(f"  source  : {p['source'][:90]!r}")
            print(f"  flags   : {p['flags']}" + ("   <-- REVIEW" if needs_review(p['flags']) else ""))
            print("  " + "-" * 60)
        if len(parsed) > 4:
            print(f"  … {len(parsed) - 4} more")
    print("\n" + "#" * 100)
    print(f"TOTAL: {grand_total} references across {len(pdfs)} papers; "
          f"{grand_review} ({100*grand_review//max(1,grand_total)}%) flagged for author review, "
          f"{grand_total - grand_review} parsed clean.")


if __name__ == "__main__":
    main()
