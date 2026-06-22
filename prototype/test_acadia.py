# -*- coding: utf-8 -*-
"""Regression tests for the prototype against the real Examples corpus."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from acadia_index import extract_docx, parse_reference, split_author_block, needs_review

# The extractor checks need the ACADIA instructions .docx. Set ACADIA_TEMPLATE to point at it;
# otherwise those checks are skipped so the suite still runs anywhere.
TEMPLATE = os.environ.get(
    "ACADIA_TEMPLATE",
    r"E:/Claude/Acadia-Indexing/Examples/ACADIA_2025_FullPaper_Instructions_styles.docx")

# Notes-bibliography corpus (year-at-end), de-wrapped from S8_145_Conway.pdf + variants.
CONWAY = [
    'Bazjanac, Vladimir. IFC BIM-Based Methodology for Semi-Automated Building Energy Performance Simulation. Lawrence Berkeley National Laboratory, 2008.',
    'Zhou, Y. W., Z. Z. Hu, J. R. Lin, et al. “A Review on 3D Spatial Data Analytics for Building Information Models.” Archives of Computational Methods in Engineering 27 (2020): 1449–1463.',
    'Douglas, David H., and Thomas K. Peucker. “Algorithms for the Reduction of the Number of Points Required to Represent a Digitized Line or Its Caricature.” Cartographica: The International Journal for Geographic Information and Geovisualization 10, no. 2 (1973): 112–122.',
    'Google. Gemini 1.5 Pro in Google Docs. 2025. Accessed June 15, 2025.',
    'Mäntylä, Martti. An Introduction to Solid Modeling. Rockville, MD: Computer Science Press, 1987.',
    'Weber, Ramon Elias. Spatial Computing for Building Performance and Design. PhD diss., Massachusetts Institute of Technology, 2024.',
]

# ALL-CAPS (Mostafavi-style) — case must not affect parsing.
ALLCAPS = [
    'SMITH, JOHN. “GENERATIVE DESIGN METHODS.” AUTOMATION IN CONSTRUCTION 18, NO. 3 (2009): 357–375.',
]


def show(label, refs):
    print("\n" + "=" * 100)
    print(label)
    print("=" * 100)
    for r in refs:
        p = parse_reference(r)
        print("\nRAW :", r[:140])
        print("  authors:", repr(p["authors"]))
        print("  year   :", repr(p["year"]))
        print("  title  :", repr(p["title"]))
        print("  source :", repr(p["source"]))
        print("  flags  :", p["flags"], "| NEEDS REVIEW" if needs_review(p["flags"]) else "")


def main():
    if not os.path.exists(TEMPLATE):
        print("ACADIA template not found at", TEMPLATE)
        print("Set ACADIA_TEMPLATE to the instructions .docx to run the full suite. Skipping.")
        return
    # ---- Extractor test ----
    ex = extract_docx(TEMPLATE)
    print("EXTRACTOR — style counts:", ex["style_counts"])
    print("EXTRACTOR — title:", repr(ex["title"][:80]))
    print("EXTRACTOR — keywords:", repr(ex["keywords"][:80]))
    print("EXTRACTOR — raw_refs:", len(ex["raw_refs"]), "| used_fallback:", ex["used_fallback"])
    assert len(ex["raw_refs"]) == 6, f"expected 6 template refs, got {len(ex['raw_refs'])}"
    assert ex["title"], "title not extracted"
    # multi-run reassembly: OpenAI/2025 ref must be one clean string containing the URL
    openai = [r for r in ex["raw_refs"] if r.startswith("OpenAI")]
    assert openai and "chat.openai.com" in openai[0], "multi-run reassembly failed"
    print("  [OK] extractor: 6 refs, title present, multi-run reference reassembled")

    show("TEMPLATE references (Chicago author-date)", ex["raw_refs"])
    show("CONWAY references (Chicago notes-bibliography)", CONWAY)
    show("ALL-CAPS reference", ALLCAPS)

    # ---- Targeted assertions ----
    checks = []

    def check(desc, cond):
        checks.append((desc, cond))
        print(("  [OK] " if cond else "  [FAIL] ") + desc)

    print("\n" + "-" * 100 + "\nASSERTIONS\n" + "-" * 100)

    p = parse_reference(ex["raw_refs"][0])  # Fox ... Interactive Architecture
    check("Fox: authors 'Fox, Michael and Miles Kemp'", p["authors"] == "Fox, Michael and Miles Kemp")
    check("Fox: year 2009", p["year"] == "2009")
    check("Fox: title 'Interactive Architecture'", p["title"] == "Interactive Architecture")
    check("Fox: source has publisher", "Princeton Architectural Press" in p["source"])

    p = parse_reference(ex["raw_refs"][2])  # Brooks (author ends in initial 'A.')
    check("Brooks: authors 'Brooks, Rodney A.' (no year leak)", p["authors"] == "Brooks, Rodney A.")
    check("Brooks: year 1990", p["year"] == "1990")
    check("Brooks: title from quotes", p["title"] == "Elephants Don’t Play Chess")

    p = parse_reference(ex["raw_refs"][3])  # Cremers (proceedings title contains a year)
    check("Cremers: NOT falsely flagged verify_year", "verify_year" not in p["flags"])
    check("Cremers: year 2011", p["year"] == "2011")

    p = parse_reference(CONWAY[0])  # Bazjanac
    check("Bazjanac: authors 'Bazjanac, Vladimir'", p["authors"] == "Bazjanac, Vladimir")
    check("Bazjanac: year 2008", p["year"] == "2008")
    check("Bazjanac: title starts 'IFC BIM-Based'", p["title"].startswith("IFC BIM-Based Methodology"))
    check("Bazjanac: source has lab+year", "Lawrence Berkeley" in p["source"])

    p = parse_reference(CONWAY[1])  # Zhou journal
    check("Zhou: year 2020 (parenthesized)", p["year"] == "2020")
    check("Zhou: authors keep 'et al.'", "et al" in p["authors"])
    check("Zhou: title is the quoted span", p["title"].startswith("A Review on 3D Spatial Data"))
    check("Zhou: source has journal+pages", "Archives of Computational Methods" in p["source"] and "1449" in p["source"])

    p = parse_reference(CONWAY[3])  # Google/Gemini double-year
    check("Google: authors 'Google'", p["authors"] == "Google")
    check("Google: title 'Gemini 1.5 Pro in Google Docs'", p["title"] == "Gemini 1.5 Pro in Google Docs")
    check("Google: verify_year flagged", "verify_year" in p["flags"])

    p = parse_reference(CONWAY[4])  # Mäntylä UTF-8
    check("Mantyla: authors keep ä", "Mäntylä" in p["authors"])
    check("Mantyla: year 1987", p["year"] == "1987")

    p = parse_reference(ALLCAPS[0])
    check("ALLCAPS: title from quotes", "GENERATIVE DESIGN METHODS" in p["title"])
    check("ALLCAPS: year 2009", p["year"] == "2009")

    # Edge tolerance (patterns seen in the 2024 data: year ranges, missing years)
    p = parse_reference('Jones, Alice. 2014–18. “Longitudinal Study.” Design Journal 5 (2).')
    print("\n  year-range ->", p)
    check("year range preserved '2014-18'", p["year"] == "2014-18")

    p = parse_reference('Anonymous. “Untitled Manifesto.” Self-published pamphlet.')
    print("  no-year ->", p)
    check("missing year -> flagged for review", p["year"] == "" and needs_review(p["flags"]))

    # ---- Name split ----
    print("\n" + "-" * 100 + "\nNAME SPLIT\n" + "-" * 100)
    names = split_author_block("Jane Q. Smith, John Doe")
    print("  'Jane Q. Smith, John Doe' ->", names)
    check("two authors parsed", len(names) == 2)
    check("Smith first='Jane Q.' last='Smith'", names[0]["first"] == "Jane Q." and names[0]["last"] == "Smith")

    names2 = split_author_block("Oliver Thomas Hamedinger and Jade Bailey")
    print("  '...Hamedinger and Jade Bailey' ->", names2)
    check("and-joined -> 2 authors", len(names2) == 2)
    check("middle name -> FIRST (2024 bug fixed)", names2[0]["first"] == "Oliver Thomas" and names2[0]["last"] == "Hamedinger")

    names3 = split_author_block("Theo van Doesburg")
    print("  'Theo van Doesburg' ->", names3)
    check("particle flagged", "verify_name_particle" in names3[0]["flags"])

    failed = [d for d, c in checks if not c]
    print("\n" + "=" * 100)
    print(f"RESULT: {len(checks) - len(failed)}/{len(checks)} checks passed")
    if failed:
        print("FAILURES:")
        for d in failed:
            print("   -", d)
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main()
