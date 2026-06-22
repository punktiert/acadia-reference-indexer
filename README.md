# ACADIA Reference Indexer

A self-service tool that lets ACADIA authors index their paper references into the two
CumInCAD submission spreadsheets — turning a months-long manual editing bottleneck into a few
minutes of author self-service, with a human review step that keeps the data clean.

It lives entirely in Google: a **Google Sheet** (the data), a **Drive folder** (the uploaded
papers), and a **Google Apps Script web app** (the logic). No servers, no database, and
editors never touch code.

> **Future conference teams:** jump to **[Run it for your conference](#run-it-for-your-conference)**.

---

## What it does

1. An author opens a link, picks their **track**, enters their **EasyChair Paper ID** and
   email, and uploads their final paper **`.docx`**.
2. The app reads the **title, authors, and references** straight from the Word file — by the
   ACADIA template paragraph styles, with fallbacks when authors don't apply them — and splits
   each Chicago reference into **authors / year / title / source**.
3. The author **reviews**: anything the parser was unsure about is flagged and must be edited
   or ticked to confirm, and no field may be left empty.
4. On submit, it appends rows to two tabs of the Sheet — **`title_authors`** and
   **`references`** — in the **exact CumInCAD column order**, and saves the `.docx` to Drive.
5. Editors export the two tabs to `.xlsx` and submit to CumInCAD.

### Why not a Google Form
A Form can collect a file, but it can't read the document or show the parsed references back
for confirmation. The "parse → review → confirm" loop is the whole point, so this is an
**Apps Script web app** — Google-native, looks like a form to authors, but runs the parser.

---

## Run it for your conference

Two parts: **(A)** point the tool at your year, **(B)** deploy it (~15 min, no coding for B).

### A · Set your year
The conference particulars and look live in a few files. Edit, then rebuild the paste bundle.

| What to change | Where |
|---|---|
| Conference year, ID prefix, per-volume **SOURCE** text & types | `apps-script/Config.gs` → `DEFAULTS` *(or just the **Settings** tab after deploy)* |
| Header wordmark + tagline | `apps-script/Index.html` (`.topbar`) |
| Accent colour + theme | `apps-script/Stylesheet.html` (`--accent`, top comment) — pull the palette from that year's conference website |
| Tracks (paper / project / report …) | `apps-script/Config.gs` → `VOLUMES` + `TYPES` |

Then rebuild the single-file bundle:
```bash
cd apps-script && python build.py        # → dist/Standalone.gs
```
*(You can skip A entirely and only edit the **Settings tab** after deploy if all you need is to
change the SOURCE wording.)*

### B · Deploy (no coding)
1. Create a Google Sheet at **[sheets.new](https://sheets.new)**, name it e.g. "ACADIA 20XX Indexing".
2. **Extensions → Apps Script** → select all in `Code.gs`, paste **`apps-script/dist/Standalone.gs`**,
   **Save**. *(One file — it contains the whole app.)*
3. Back on the Sheet, **reload** → menu **ACADIA Indexer → 1. First-time setup** → approve the
   permission prompt (it's your own script: *Advanced → Go to … → Allow*). This creates the
   `title_authors`, `references`, **Settings**, and **How to publish** tabs, plus the uploads folder.
4. Fill the **Settings** tab (column B) — especially the per-volume **SOURCE** rows (ISBN, editors).
5. In the Apps Script editor: **Deploy → New deployment → Web app**, *Execute as:* **Me**,
   *Who has access:* **Anyone** → **Deploy** → copy the **Web app URL**. That's the author link.

> **Changing code later?** After editing files + `python build.py`, repaste `Standalone.gs`, then
> **Deploy → Manage deployments → ✏️ → Version: New version → Deploy** (keeps the same URL).

---

## For authors (share this)

1. In Word, make sure your references use the **ACADIA-Reference** paragraph style (and the
   ACADIA template styles for title/authors). That's what makes the import clean — though the
   tool also recovers references without it.
2. Open the link, choose your **track**, enter your **EasyChair Paper ID** and **email**,
   upload your **`.docx`**.
3. Review the parsed references: orange rows need attention — **edit or tick every flagged
   row**, fill any empty cell, and fix the author First/Last split (middle names go in *First*).
4. **Confirm & submit.** Re-submitting the same Paper ID safely replaces your earlier entry.

## For editors

- The two tabs are already in CumInCAD's exact layout → **File → Download → Microsoft Excel
  (.xlsx)** for each.
- Uploaded `.docx` files live in the Drive folder **"ACADIA Indexer uploads — ‹sheet name›"**,
  each named `‹paperID›__‹original›.docx`.
- **ACADIA Indexer → Run self-test** logs `ALL PASSED` if the parser is healthy.
- The "Google hasn't verified this app" / "created by a Google Apps Script user" notices are
  normal for a personal Apps Script web app.

---

## How the parser handles messy input

References are read by the `ACADIA-Reference` paragraph style; if authors didn't apply it, the
app scans for a "References" heading and recovers them (and recovers a title styled as a
heading). Each reference is split by **anchoring on the year**, so it handles both Chicago
variants (author-date *and* notes-bibliography), ALL-CAPS, initials, smart quotes / UTF-8,
year ranges, and missing years. Nothing is silently guessed — uncertain fields are flagged and
the author must resolve them before submitting.

## Repository layout

```
apps-script/        Google Apps Script web app (the deployable tool)
  Config.gs         year/conference config + Settings-tab reader
  Setup.gs          the "ACADIA Indexer" menu, one-click setup, Drive folder
  Docx.gs           .docx unzip + style-based extraction
  RefParser.gs      Chicago reference → authors / year / title / source
  NameSplit.gs      author display name → First / Last
  SheetIO.gs        sheet writes (locked, idempotent) + Drive save
  Code.gs           web-app entry + endpoints + runTests()
  Index/Stylesheet/JavaScript.html   the two-stage UI
  build.py → dist/  build the paste bundles (Standalone.gs = one-file deploy)
  README.md         app-level reference
prototype/          offline Python mirror that de-risks & tests the parser
  acadia_index.py   extractor + parser + name split
  test_acadia.py    regression suite (set ACADIA_TEMPLATE to run extractor checks)
  local_server.py   a REAL local backend for demos (no Google needed)
  dryrun_pdf.py     run the parser over example PDFs
```

## Local development & testing

```bash
cd prototype
python test_acadia.py           # parser regression suite
python local_server.py          # the REAL app locally at http://127.0.0.1:8091/
```
`local_server.py` serves the same UI and runs the actual parser on uploads, writing CSVs to
`prototype/output/` — handy for trying papers without deploying. Edit the modular files in
`apps-script/`, run `python build.py`, and the local server picks up the change. The deployed
Apps Script app and the Python prototype share the same parser logic (kept in lock-step).

## License

MIT — see [LICENSE](LICENSE).
