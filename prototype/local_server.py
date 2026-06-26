# SPDX-License-Identifier: AGPL-3.0-or-later
# Copyright (C) 2026 Daniel Koehler / ACADIA
"""
Local REAL backend for the indexer UI — runs the actual extractor + parser on uploads.

Serves the production UI (apps-script/dist/Index.html) and implements the same three
endpoints the deployed Apps Script app exposes, backed by acadia_index.py (the validated
parser). Uploads are saved to prototype/uploads/, and confirmed submissions are written to
prototype/output/title_authors.csv + references.csv in the exact CumInCAD column order.

This is NOT the mock — it genuinely analyses whatever .docx you upload.

    cd prototype && python local_server.py        # then open http://127.0.0.1:8091/
"""
import http.server, socketserver, json, base64, os, csv, io, tempfile

from acadia_index import extract_docx, parse_reference, split_author_block, needs_review

HERE = os.path.dirname(os.path.abspath(__file__))
UI_FILE = os.path.join(HERE, "..", "apps-script", "dist", "Index.html")
UPLOADS = os.path.join(HERE, "uploads"); os.makedirs(UPLOADS, exist_ok=True)
OUTDIR = os.path.join(HERE, "output"); os.makedirs(OUTDIR, exist_ok=True)
PORT = 8091

VOLUMES = {
    "1": {"type": "project", "sourceBase": "ACADIA 2026: Humanism Recoded [Project Catalog of the 46th Annual Conference for the Association for Computer Aided Design in Architecture (ACADIA) ISBN <ISBN>]. Detroit. 22-24 October 2026. edited by <editors>."},
    "2": {"type": "paper", "sourceBase": "ACADIA 2026: Humanism Recoded [Proceedings of the 46th Annual Conference for the Association for Computer Aided Design in Architecture (ACADIA) ISBN <ISBN>]. Detroit. 22-24 October 2026. edited by <editors>."},
}
TYPES = ["paper", "project"]
ID_PREFIX, ID_PAD, MAX_AUTHORS = "acadia26", 2, 8
SERIES = "ACADIA"

TA_HEADERS = ["ID", "SOURCE", "SERIES", "TYPE", "TITLE", "MAIN CONTACT EMAIL", "SUMMARY"] + \
    sum([[f"AUTHOR {i} FIRST", f"AUTHOR {i} LAST"] for i in range(1, 9)], [])
REF_HEADERS = ["paper", "authors", "year", "title", "source"]


def build_id(vol, number):
    return f"{ID_PREFIX}_v{vol}_{str(int(number)).zfill(ID_PAD)}"


def api_get_config(_payload):
    return {
        "volumes": {v: {"type": VOLUMES[v]["type"], "sourcePreview": VOLUMES[v]["sourceBase"]} for v in VOLUMES},
        "types": TYPES, "idPrefix": ID_PREFIX, "idPad": ID_PAD, "confYear": "2026", "maxAuthors": MAX_AUTHORS,
    }


def api_process_upload(payload):
    meta = payload.get("meta", {})
    vol = str(meta.get("vol", ""))
    if vol not in VOLUMES:
        return {"ok": False, "error": "Please choose a valid track."}
    if not str(meta.get("number", "")).strip().isdigit():
        return {"ok": False, "error": "Please enter your EasyChair Paper ID (a number)."}
    if not payload.get("base64"):
        return {"ok": False, "error": "No file received."}

    pid = build_id(vol, meta["number"])
    saved = os.path.join(UPLOADS, pid + "__" + os.path.basename(str(payload.get("filename", "paper"))))
    with open(saved, "wb") as f:
        f.write(base64.b64decode(payload["base64"]))

    pre = payload.get("preExtracted")
    if pre:                                  # PDF: the browser already extracted it (pdf.js)
        title, author_block = pre.get("title", ""), pre.get("authorBlock", "")
        raw_refs, keywords = pre.get("rawRefs", []), pre.get("keywords", "")
        diagnostics = pre.get("diagnostics", {"source": "pdf"})
    else:                                    # .docx: extract here
        ex = extract_docx(saved)
        title, author_block, raw_refs, keywords = ex["title"], ex["author_block"], ex["raw_refs"], ex["keywords"]
        diagnostics = {"source": "docx", "styledRefCount": ex["style_counts"].get("ACADIA-Reference", 0),
                       "usedFallback": ex["used_fallback"], "titleFound": bool(ex["title"]),
                       "titleFallback": ex.get("title_fallback", False), "authorFound": bool(ex["author_block"])}

    all_authors = split_author_block(author_block)
    refs = []
    for raw in raw_refs:
        p = parse_reference(raw)
        p["review"] = needs_review(p["flags"])
        refs.append(p)
    return {
        "ok": True,
        "meta": {"id": pid, "vol": vol, "number": str(meta["number"]),
                 "type": meta.get("type") or VOLUMES[vol]["type"],
                 "email": str(meta.get("email", "")).strip(), "source": VOLUMES[vol]["sourceBase"]},
        "title": title, "keywords": keywords,
        "authors": all_authors[:MAX_AUTHORS], "authorOverflow": max(0, len(all_authors) - MAX_AUTHORS),
        "references": refs, "diagnostics": diagnostics,
    }


def _rewrite_without_id(path, headers, key_col, pid, new_rows):
    rows = []
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            r = csv.reader(f)
            allrows = list(r)
        if allrows:
            rows = [row for row in allrows[1:] if not (row and row[key_col] == pid)]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)
        w.writerows(new_rows)


def api_commit(reviewed):
    meta = reviewed["meta"]; pid = meta["id"]
    title_row = [pid, meta.get("source", ""), SERIES, str(meta.get("type", "")).strip(),
                 reviewed.get("title", ""), meta.get("email", ""), ""]
    authors = reviewed.get("authors", [])
    for k in range(MAX_AUTHORS):
        a = authors[k] if k < len(authors) else {}
        title_row += [a.get("first", ""), a.get("last", "")]
    ref_rows = [[pid, r.get("authors", ""), r.get("year", ""), r.get("title", ""), r.get("source", "")]
                for r in reviewed.get("references", [])]
    _rewrite_without_id(os.path.join(OUTDIR, "title_authors.csv"), TA_HEADERS, 0, pid, [title_row])
    _rewrite_without_id(os.path.join(OUTDIR, "references.csv"), REF_HEADERS, 0, pid, ref_rows)
    return {"ok": True, "paperId": pid, "titleRowsWritten": 1, "refRowsWritten": len(ref_rows)}


ROUTES = {"getConfigForClient": api_get_config, "processUpload": api_process_upload, "commitSubmission": api_commit}


class Handler(http.server.BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json"):
        b = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype + ("; charset=utf-8" if "json" in ctype or "html" in ctype else ""))
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            with open(UI_FILE, encoding="utf-8") as f:
                self._send(200, f.read(), "text/html")
        else:
            self._send(404, "not found", "text/plain")

    def do_POST(self):
        if not self.path.startswith("/api/"):
            return self._send(404, json.dumps({"ok": False, "error": "no route"}))
        fn = self.path[len("/api/"):]
        n = int(self.headers.get("Content-Length", 0))
        payload = json.loads(self.rfile.read(n) or b"{}") if n else {}
        handler = ROUTES.get(fn)
        if not handler:
            return self._send(404, json.dumps({"ok": False, "error": "unknown endpoint " + fn}))
        try:
            self._send(200, json.dumps(handler(payload), ensure_ascii=False))
        except Exception as e:
            self._send(200, json.dumps({"ok": False, "error": str(e)}))

    def log_message(self, *a):
        pass  # quiet


if __name__ == "__main__":
    print(f"REAL indexer backend on http://127.0.0.1:{PORT}/   (uploads -> {UPLOADS}, output -> {OUTDIR})")
    socketserver.ThreadingTCPServer(("127.0.0.1", PORT), Handler).serve_forever()
