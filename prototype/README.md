# Parser prototype (offline)

Python mirror of the Apps Script parsing logic, used to de-risk and regression-test the two
hard parts against the real `Examples/` corpus before/while porting to `../apps-script/`.

- `acadia_index.py` — `extract_docx()`, `parse_reference()`, `split_author_block()`.
- `test_acadia.py` — 31 assertions over the 6 ACADIA template references (author-date),
  the Conway references (notes-bibliography), ALL-CAPS, double-year, year ranges, missing
  years, UTF-8, and name-splitting.

```bash
cd prototype
python test_acadia.py      # expects: "ALL CHECKS PASSED"
```

If you change parsing rules, change them here first, get the tests green, then mirror the
edit into `../apps-script/RefParser.gs` / `NameSplit.gs` and re-run `runTests()` in Apps
Script (and the Node parity check noted in the build history).
