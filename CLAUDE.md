# CLAUDE.md — Project Context for Claude Code

## Project Overview

**pdf-toolkit** — PDF manipulation tools with a tkinter GUI, using open-source libraries only:

1. **`gui.py`** — Tkinter GUI that wraps both tools below. Each PDF is processed in its own subprocess for crash isolation.
2. **`add_toc.py`** — Automatically detects chapter headings in structured PDFs and inserts a Table of Contents page at the beginning of each file.
3. **`pdf_to_docx.py`** — High-fidelity PDF to DOCX converter that preserves text formatting, images, tables, equations, headers/footers, and page layout.

## Tech Stack

- **Python 3.12+**
- **PyMuPDF (fitz)** — PDF reading, text/image/drawing extraction, page manipulation
- **python-docx** — DOCX generation with XML-level access via lxml
- **lxml** — XML manipulation for advanced DOCX features (borders, equations)
- **Pillow** — Image validation and processing
- **tkinter** — GUI (built-in with Python)

## Architecture

### gui.py — Tkinter GUI with crash isolation

The GUI processes each PDF in a **separate subprocess** via `_run_one.py`. If a corrupt or oversized PDF causes a segfault or OOM, only that subprocess dies — the GUI stays alive and continues with the next file.

**Key features:**
- Directory picker for input PDFs
- "Add TOC" and "Convert to DOCX" buttons
- Per-file progress bar with file size display
- Stop button for cancellation between files
- Options: "H1 headings only", "Include TOC in page count"
- Stale output file handling (locked 0-byte files from prior crashes)

**Flow:**
```
GUI → for each PDF:
        subprocess: python _run_one.py <mode> <input> <output> [flags]
        parse stdout for __RESULT__:<status> tag
        log output to text widget
```

### _run_one.py — Per-file subprocess wrapper

Thin wrapper that imports `process_pdf` from either `add_toc` or `pdf_to_docx`, handles stale/locked output files, and prints a machine-readable `__RESULT__:` tag for the GUI to parse.

**CLI:** `python _run_one.py <toc|docx> <input.pdf> <output_path> [--h1-only] [--no-count-toc]`

### add_toc.py (~1000 lines)

Font-metric based heading detection pipeline:
- Body text = most frequent font size by character count
- Headings = bold text at sizes larger than body (relative per document)
- Font sizes clustered within 0.6pt into canonical buckets
- Bold detected via flags bit 4 OR font name patterns ("bold", "cmbx", "advtib")
- Multi-line heading merge: same-block (45px y-gap) + cross-block continuation
- H1/H2 selection rule: if H1 count > 3, use H1; otherwise fall back to H2
- Over-extraction guard: >50 H1s triggers narrowing to "Chapter"-prefixed entries
- Running header filtering: text on >25% of sampled pages excluded
- Output: inserts generated TOC pages + bookmark metadata via `doc.set_toc()`

**TOC page generation:**
- Long heading text wraps up to 3 lines per entry (word-wrapped); entries that still don't fit are removed
- Reduced margins (45pt left/right, 50pt top/bottom) and font sizes (9pt H1, 8pt H2) for more usable space
- Dot leaders and page numbers on the last line of each entry
- `h1_only` flag: drops all H2 sub-section entries
- `count_toc_pages` flag: controls whether displayed page numbers include the TOC pages or match original PDF numbering
- Bookmarks always navigate to the correct real page regardless of page number display setting

### pdf_to_docx.py (~1820 lines)

Per-page extraction → intermediate data structures → DOCX rendering:

```
PDF page → [Images | Tables | Text | Equations] → sort by position → render to DOCX
```

**Key data structures:** `TextRun`, `ParagraphElement`, `ImageElement`, `TableElement`, `EquationElement`

**Pipeline per page:**
1. Image extraction via `page.get_image_info(xrefs=True)` + `doc.extract_image()`
2. Table detection from vector line drawings (horizontal/vertical line intersections)
3. Text extraction with header/footer exclusion, paragraph assembly (line joining with hyphen removal)
4. Equation detection via TeX font names → batched pixmap rendering at 200 DPI
5. All elements sorted top-to-bottom, deduplicated, rendered to python-docx

**Text paragraph assembly heuristics:**
- Join lines if: same font size, small vertical gap (< 1.8x font height), no list markers
- Don't join if: font size change, large gap, sentence-ending punctuation + uppercase start, indent change
- Hyphenation: line ending with `-` + next starting lowercase → remove hyphen and join

## File Layout

```
/
├── gui.py              # Tkinter GUI (entry point)
├── _run_one.py         # Per-file subprocess wrapper for crash isolation
├── add_toc.py          # TOC generation script
├── pdf_to_docx.py      # PDF → DOCX converter
├── CLAUDE.md           # This file
├── README.md           # User-facing documentation
├── requirements.txt    # Python dependencies
├── .gitignore          # Ignores output dirs, PDFs, __pycache__
├── with_toc/           # Output: PDFs with TOC inserted (gitignored)
├── docx_output/        # Output: converted DOCX files (gitignored)
└── failed_files/       # Output: scanned/unprocessable PDFs (gitignored)
```

## Running

```bash
# Install dependencies
pip install PyMuPDF python-docx Pillow lxml

# Launch the GUI
python gui.py

# Or run scripts directly (processes all *.pdf in current directory)
python add_toc.py
python pdf_to_docx.py
```

## Key Design Decisions

- **Per-file subprocess isolation**: Each PDF is processed in its own subprocess. Corrupt or oversized PDFs that cause segfaults or OOM kills cannot crash the GUI or abort remaining files.
- **TOC text wrapping instead of truncation**: Long heading text wraps up to 3 lines with dot leaders on the last line. Entries exceeding 3 lines are removed entirely rather than being truncated with ellipsis.
- **Equations as images, not OMML**: TeX math is rendered as 200 DPI PNG images rather than converted to Office Math XML. This preserves visual fidelity without the extreme complexity of TeX→OMML conversion.
- **Table detection from vectors**: No native table markup exists in PDFs. Tables are reconstructed by finding horizontal/vertical line intersections in the drawing layer.
- **Font mapping heuristic**: PDF fonts are mapped to closest Word-compatible equivalents (Times→Times New Roman, Helvetica→Arial, CM→Times New Roman). Unknown fonts fall back to serif/sans-serif based on font flags.
- **Text sanitization**: All text is sanitized for XML 1.0 compatibility before inserting into DOCX (removes null bytes and control characters from broken PDFs).
- **Per-page error handling**: If a single page fails, the converter adds a placeholder and continues rather than aborting the entire document.

## Common Issues

- **Locked output files on Windows**: If a process crashes mid-write, Windows Search Indexer may lock the 0-byte output file. The helper script `_run_one.py` detects this and saves to a `_new` suffixed filename instead. The lock releases on reboot.
- **Type 3 fonts** (e.g., Solid State Physics PDF): Font encoding may be broken in the PDF itself. The converter reproduces whatever text PyMuPDF extracts, which may be garbled.
- **Large math-heavy PDFs**: Books with hundreds of pages of algebra produce 100+ MB DOCX files due to equation images. This is expected.
- **Multi-column layouts**: Content is extracted in reading order but rendered as single-column. Multi-column DOCX layout via XML is fragile and not implemented.
- **Very large PDFs (300MB+)**: May cause OOM in the subprocess. The GUI survives via crash isolation and reports the file as "CRASHED".
