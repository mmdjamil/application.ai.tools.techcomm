# 📝 Doc Scanner — Inclusive Language & Broken Link Checker

> A lightweight Streamlit web app that scans technical documents for **non‑inclusive language** and **broken hyperlinks** — built for technical writers, editors, and documentation teams.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Streamlit](https://img.shields.io/badge/built%20with-Streamlit-FF4B4B)
![Version](https://img.shields.io/badge/version-1.3.0-brightgreen)
![License](https://img.shields.io/badge/license-MIT-green)

---

## 📖 Overview

`application.ai.tools.techcomm` is a productivity tool for technical communicators. It takes a document you upload (Word, PDF, PowerPoint, or Excel) and produces two clear reports:

1. **Inclusive Language Scan** — flags non‑inclusive terms (e.g. *blacklist*, *whitelist*, *master*, *slave*), recommends inclusive replacements, generates a rewritten sentence for each finding, lets you approve/reject each rewrite, and exports an approved corrected copy of the document with visible change markers.
2. **Broken Link Check** — extracts every `http(s)://` link in the document and verifies each one in parallel, reporting HTTP status, timeouts, and connection errors.

Both reports include the **page/section and line number** where each issue was found, and can be downloaded as CSV for further triage.

---

## 🆕 What's New in v1.3.0

- 📄 **Annotated PDF export** — accepted PDF suggestions now generate a real corrected `.pdf` with highlights and sticky notes describing each suggested replacement.
- 🟨 **Visual change markers in DOCX/XLSX** — rewritten Word documents highlight each inserted replacement in yellow, and changed Excel cells get a yellow fill plus a `Doc Scanner` comment listing the applied replacements.
- 🎞️ **Visible PPTX styling** — rewritten PowerPoint decks now style each inserted replacement so the changed word stands out during review.
- 🧷 **Safer PDF fallback** — if a PDF cannot be opened or annotated, the app still falls back to the existing plain-text export path instead of failing.

## 🆕 What's New in v1.1.0

> Released as part of [PR #2 — Add inclusive rewrite review flow and corrected text export](https://github.com/mmdjamil/application.ai.tools.techcomm/pull/2).

The inclusive language scanner evolved from **detection‑only** to a full **approval‑driven rewrite workflow**:

- ✍️ **Rewrite‑aware findings** — each finding now includes an `original_sentence` and a `suggested_sentence` (the line with the non‑inclusive term swapped for an inclusive replacement).
- 🔠 **Case‑preserving replacements** — `Blacklist → Denylist`, `BLACKLIST → DENYLIST`, `blacklist → denylist`.
- 🧩 **Smart default for multi‑option entries** — replacements like `primary / initiator` insert the first option (`primary`) into the rewritten sentence while the full string is still shown for context.
- ☑️ **Per‑finding review UI** — a new `st.data_editor` table with an **Apply?** checkbox (ticked by default) so you can accept or reject each rewrite individually.
- 📝 **Corrected copy export** — click **✅ Generate corrected copy** to preview and download `<original_basename>_corrected.txt` with only the approved replacements applied.
- 🧠 **Session‑aware scans** — scan results are cached in `st.session_state` so review interactions survive Streamlit reruns; the cache is cleared automatically when a different file is uploaded.
- 🛠️ **New helper APIs** in `scanners/inclusive_scanner.py`:
  - `pick_default_replacement(replacement)`
  - `apply_replacement_to_text(text, term, replacement)`
  - `apply_all_accepted_replacements(text, accepted)`

Existing behavior is preserved: inclusive‑scan CSV export still works (now driven by the reviewed table state), and the broken‑link flow is unchanged.

---

## ✨ Features

- 📂 **Multi‑format document parsing**
  - Microsoft Word (`.docx`) — paragraphs and tables
  - PDF (`.pdf`) — page‑by‑page, line‑by‑line
  - PowerPoint (`.pptx`) — slides, shapes, and tables
  - Excel (`.xlsx`) — every sheet and row
- 🔤 **Non‑inclusive language detection** with case‑insensitive whole‑word matching and suggested replacements.
- ✅ **Suggestion review & corrected-copy export** — approve/reject each suggested rewrite row-by-row, then download an approved corrected copy **in the original document format** with visible highlights or annotations (`.docx`, `.pptx`, `.xlsx`, `.pdf`).
- 🔗 **Concurrent link checking** (10 workers) using HEAD requests with automatic GET fallback, custom User‑Agent, and 10s timeout.
- 📊 **Summary metrics** — word count, compliance rate, total links, broken links.
- 📥 **CSV export** of both inclusive scan and link check results.
- 🚀 **Simple Streamlit UI** — drag, drop, scan.

---

## 🗂️ Project Structure

```
application.ai.tools.techcomm/
├── app.py                       # Streamlit UI entry point
├── requirements.txt             # Python dependencies
├── .streamlit/
│   └── config.toml              # Streamlit configuration
└── scanners/
    ├── __init__.py              # Public API exports
    ├── file_parser.py           # Parsers for .docx / .pdf / .pptx / .xlsx
    ├── inclusive_scanner.py     # Non-inclusive language detection
    ├── document_rewriter.py     # Format-preserving corrected-copy export
    └── link_checker.py          # URL extraction & broken-link verification
```

---

## 🚀 Getting Started

### Prerequisites

- Python **3.9 or later**
- `pip` (or `uv` / `poetry` if you prefer)

### Installation

```bash
# 1. Clone the repository
git clone https://github.com/mmdjamil/application.ai.tools.techcomm.git
cd application.ai.tools.techcomm

# 2. (Recommended) create a virtual environment
python -m venv .venv
source .venv/bin/activate          # macOS / Linux
# .venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install -r requirements.txt
```

### Run the app

```bash
streamlit run app.py
```

Then open the URL Streamlit prints (typically <http://localhost:8501>) in your browser.

---

## 🧭 How to Use

1. **Upload** a `.docx`, `.pdf`, `.pptx`, or `.xlsx` file (up to 200 MB).
2. Choose what to scan:
   - ✅ Non‑Inclusive Language Scan
   - ✅ Broken Link Check
3. Click **🚀 Start Scan**.
4. Review the results inline and click **📥 Download** to export them as CSV.

---

## 🔤 Inclusive Language Rules

The default ruleset is defined in [`scanners/inclusive_scanner.py`](scanners/inclusive_scanner.py):

| Found Term  | Suggested Replacement     |
|-------------|---------------------------|
| `blacklist` | `denylist`                |
| `whitelist` | `allowlist`               |
| `master`    | `primary / initiator`     |
| `slave`     | `secondary / target`      |

Matching is **case‑insensitive** and uses **whole‑word boundaries** (`\b`) to avoid false positives.

### ✅ Reviewing suggestions & exporting a corrected copy

1. After a scan, each finding includes **Original Sentence** and **Suggested Sentence** columns, plus an **Apply?** checkbox (default ticked).
2. Untick any suggestions you do not want to keep.
3. Click **✅ Generate corrected copy** to download the corrected document.
4. The corrected copy is returned **in the same format as the uploaded file**:
   - `.docx` → a `.docx` that opens cleanly in Word with original fonts, headings, and tables preserved, with each inserted replacement highlighted in yellow.
   - `.pptx` → a `.pptx` that opens in PowerPoint with layout and theme preserved, with each inserted replacement visibly styled for review.
   - `.xlsx` → a `.xlsx` that opens in Excel with styling and untouched formulas preserved; changed cells are filled yellow and receive a `Doc Scanner` comment listing the applied replacements.
   - `.pdf` → a corrected `.pdf` with highlights and sticky-note annotations for each accepted suggestion; if the PDF cannot be opened or annotated, the tool falls back to the plain `.txt` export.
5. For multi-option entries like `master → primary / initiator`, the rewriter applies the first option (`primary`) while preserving matched casing (`Blacklist → Denylist`, `BLACKLIST → DENYLIST`).

### Example: default replacement + case-preserving rewrite

```python
from scanners.inclusive_scanner import apply_replacement_to_text

line = "Blacklist and MASTER values in this list."
print(apply_replacement_to_text(line, "blacklist", "denylist"))
# Denylist and MASTER values in this list.

print(apply_replacement_to_text(line, "master", "primary / initiator"))
# Blacklist and PRIMARY values in this list.
```

### Add your own terms

Edit the `NON_INCLUSIVE_TERMS` dictionary in `scanners/inclusive_scanner.py`:

```python
NON_INCLUSIVE_TERMS = {
    "blacklist":  "denylist",
    "whitelist":  "allowlist",
    "master":     "primary / initiator",
    "slave":      "secondary / target",
    # Add new entries here
    "dummy":      "placeholder",
}
```

---

## 🔗 Link Checking Details

- URLs are extracted with the regex `https?://[^\s<>"\'\)\],;|\\}]+`.
- Each link is first probed with **HEAD**; if the server returns ≥ 400, the checker falls back to **GET** (streaming).
- Requests use a custom User‑Agent (`Mozilla/5.0 (DocScanner/1.0)`) with a **10‑second timeout** and follow redirects.
- A link is reported **broken** when:
  - HTTP status ≥ 400, **or**
  - the request times out, **or**
  - a connection / generic request error occurs.
- Up to **10 links are checked concurrently** via `ThreadPoolExecutor`.

---

## 🧩 Programmatic Use

You can also call the scanner modules directly without Streamlit:

```python
from scanners import parse_file, scan_for_inclusive_language, check_all_links

with open("my_doc.docx", "rb") as f:
    parsed = parse_file(f.read(), "my_doc.docx")

inclusive = scan_for_inclusive_language(parsed)
print(inclusive["total_non_inclusive_count"], "non-inclusive terms found")

links = check_all_links(parsed)
print(links["broken_link_count"], "broken links")
```

Each parsed entry has the shape:

```python
{"page": int, "line": int, "text": str}
```

---

## 📦 Dependencies

Defined in [`requirements.txt`](requirements.txt):

| Package        | Purpose                                  |
|----------------|------------------------------------------|
| `streamlit`    | Web UI                                   |
| `python-docx`  | Read `.docx` files                       |
| `pdfplumber`   | Read `.pdf` files                        |
| `pymupdf`      | Add PDF highlights and sticky notes      |
| `python-pptx`  | Read `.pptx` files                       |
| `openpyxl`     | Read `.xlsx` files                       |
| `requests`     | HTTP link checking                       |
| `pandas`       | Tabular results & CSV export             |

---

## 🛠️ Troubleshooting

- **`Unsupported file type`** — only `.docx`, `.pdf`, `.pptx`, and `.xlsx` are supported. Convert other formats first.
- **Many links reported as broken** — some sites block `HEAD`/automated requests or are behind authentication. Verify in a browser.
- **Slow scan on large PDFs** — parsing time scales with document length; the link check is parallelized but still bound by remote server response times.
- **Password‑protected or scanned (image‑only) PDFs** — text cannot be extracted; OCR is not currently supported.

---

## 📋 Changelog

### v1.3.0 — Visual markup for corrected copies
- PDFs now export as annotated `.pdf` files with highlights and sticky notes for accepted replacements when annotation is possible.
- DOCX exports highlight inserted replacement words in yellow.
- PPTX exports visibly style inserted replacement words for easier review.
- XLSX exports fill changed cells yellow and add `Doc Scanner` comments listing the applied replacements.
- The app UI now passes accepted findings into the PDF rewriter and explains whether the corrected PDF should be reviewed in a PDF reader or has fallen back to plain text.

### v1.2.0 — Format-preserving corrected-copy export
- **New `scanners/document_rewriter.py`** module with `rewrite_docx`, `rewrite_pptx`, `rewrite_xlsx`, and `rewrite_file`.
- Corrected copy is now downloaded in the **original document format** (`.docx`, `.pptx`, `.xlsx`) preserving fonts, headings, tables, slide layouts, and themes.
- PDFs fall back to the existing plain-text export; the UI now shows a caption explaining why.
- DOCX/PPTX rewriter uses a run-concatenation strategy to handle words split across multiple formatting runs.
- XLSX rewriter opens in read-write/formula-preserving mode so untouched cells and formulas are unaffected.
- `scanners/__init__.py` now exports `rewrite_file`.
- `app.py` updated: flat-text assembly removed; download button now offers the native format.
- New `tests/test_document_rewriter.py` with DOCX, PPTX, XLSX, PDF, and line-drift round-trip tests.

### v1.1.0 — Inclusive rewrite review flow and corrected text export
- Added `original_sentence` and `suggested_sentence` to every inclusive‑language finding.
- Added `pick_default_replacement`, `apply_replacement_to_text`, and `apply_all_accepted_replacements` helpers in `scanners/inclusive_scanner.py`.
- Streamlit UI: per‑finding **Apply?** review checkboxes via `st.data_editor`; scan results cached in `st.session_state` and invalidated when the uploaded file changes.
- New **✅ Generate corrected copy** action that exports `<original_basename>_corrected.txt` with only the approved rewrites applied.
- Case‑preserving rewrites (`Blacklist → Denylist`, `BLACKLIST → DENYLIST`) and first‑option insertion for multi‑option entries (e.g. `master → primary / initiator` inserts `primary`).
- README updated with the new review/export workflow and roadmap entry for original‑format round‑tripping.
- Shipped via [#2](https://github.com/mmdjamil/application.ai.tools.techcomm/pull/2).

### v1.0.0 — Initial release
- Multi‑format document parsing (`.docx`, `.pdf`, `.pptx`, `.xlsx`).
- Non‑inclusive language scan with case‑insensitive whole‑word matching and suggested replacements.
- Concurrent broken‑link checking (HEAD with GET fallback, 10 workers, 10s timeout).
- Summary metrics and CSV export for both reports.
- Streamlit drag‑and‑drop UI.

---

## 🗺️ Roadmap Ideas

- Configurable / external rule file (YAML/JSON) for inclusive terms
- OCR support for scanned PDFs
- Additional formats (Markdown, HTML, plain text)
- Severity levels and per‑rule enable/disable
- CLI mode (run without Streamlit)
- Dockerfile for one‑command deployment
- Export corrected document back into the original format (`.docx`, `.pptx`, `.xlsx`) — ✅ **Implemented in v1.2.0** (`.pdf` plain-text fallback only)

---

## 🤝 Contributing

Contributions are welcome! To propose a change:

1. Fork the repo and create a feature branch: `git checkout -b feature/my-improvement`
2. Make your changes and test locally with `streamlit run app.py`.
3. Commit with a clear message and open a Pull Request.

For larger changes, please open an issue first to discuss what you'd like to do.

---

## 📄 License

This project is released under the **MIT License**. See [`LICENSE`](LICENSE) for details (add one if not present).

---

## 🙏 Acknowledgements

Built for technical writers who care about clear, inclusive, and reliable documentation.
Powered by [Streamlit](https://streamlit.io/), [pdfplumber](https://github.com/jsvine/pdfplumber), [python-docx](https://github.com/python-openxml/python-docx), [python-pptx](https://github.com/scanny/python-pptx), and [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl).
