# 📝 Doc Scanner — Inclusive Language & Broken Link Checker

> A lightweight Streamlit web app that scans technical documents for **non‑inclusive language** and **broken hyperlinks** — built for technical writers, editors, and documentation teams.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Streamlit](https://img.shields.io/badge/built%20with-Streamlit-FF4B4B)
![License](https://img.shields.io/badge/license-MIT-green)

---

## 📖 Overview

`application.ai.tools.techcomm` is a productivity tool for technical communicators. It takes a document you upload (Word, PDF, PowerPoint, or Excel) and produces two clear reports:

1. **Inclusive Language Scan** — flags non‑inclusive terms (e.g. *blacklist*, *whitelist*, *master*, *slave*), recommends inclusive replacements, generates a rewritten sentence for each finding, and lets you export an approved corrected copy.
2. **Broken Link Check** — extracts every `http(s)://` link in the document and verifies each one in parallel, reporting HTTP status, timeouts, and connection errors.

Both reports include the **page/section and line number** where each issue was found, and can be downloaded as CSV for further triage.

---

## ✨ Features

- 📂 **Multi‑format document parsing**
  - Microsoft Word (`.docx`) — paragraphs and tables
  - PDF (`.pdf`) — page‑by‑page, line‑by‑line
  - PowerPoint (`.pptx`) — slides, shapes, and tables
  - Excel (`.xlsx`) — every sheet and row
- 🔤 **Non‑inclusive language detection** with case‑insensitive whole‑word matching and suggested replacements.
- ✅ **Suggestion review & corrected-copy export** — approve/reject each suggested rewrite row-by-row, then download an approved corrected `.txt` copy.
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
3. Click **✅ Generate corrected copy** to preview and download `<filename>_corrected.txt`.
4. For multi-option entries like `master → primary / initiator`, the rewriter applies the first option (`primary`) while preserving matched casing (`Blacklist → Denylist`, `BLACKLIST → DENYLIST`).
5. Current v1 limitation: corrected output is exported as plain `.txt` (lossless text), not back into original `.docx`/`.pdf`/`.pptx`/`.xlsx` files with formatting. Original-format round-tripping is on the roadmap.

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

## 🗺️ Roadmap Ideas

- Configurable / external rule file (YAML/JSON) for inclusive terms
- OCR support for scanned PDFs
- Additional formats (Markdown, HTML, plain text)
- Severity levels and per‑rule enable/disable
- CLI mode (run without Streamlit)
- Dockerfile for one‑command deployment
- Export corrected document back into the original format (`.docx`, `.pdf`, `.pptx`, `.xlsx`) preserving formatting

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
Powered by [Streamlit](https://streamlit.io/), [pdfplumber](https://github.com/jsvine/pdfplumber), [python-docx](https://github.com/python-openxml/python-docx), [python-pptx](https://github.com/scanny/python-pptx), and [openpyxl](https://openpyxl.readthedocs.io/).
