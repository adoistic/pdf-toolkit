# CLAUDE.md — Project Context for Claude Code

## Project Overview

**pdf-toolkit** — Two Python scripts for PDF manipulation using open-source tools only:

1. **`add_toc.py`** — Automatically detects chapter headings in structured PDFs and inserts a Table of Contents page at the beginning of each file.
2. **`pdf_to_docx.py`** — High-fidelity PDF to DOCX converter that preserves text formatting, images, tables, equations, headers/footers, and page layout.

## Tech Stack

- **Python 3.12** (use `python3.12` explicitly — `python3` may point to a broken 3.14 on some systems)
- **PyMuPDF (fitz) 1.26.7** — PDF reading, text/image/drawing extraction, page manipulation
- **python-docx 0.8.11** — DOCX generation with XML-level access via lxml
- **lxml 6.0.2** — XML manipulation for advanced DOCX features (borders, equations)
- **Pillow 12.1.0** — Image validation and processing

## Architecture

### add_toc.py (951 lines)

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

### pdf_to_docx.py (1819 lines)

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
├── add_toc.py          # TOC generation script
├── pdf_to_docx.py      # PDF → DOCX converter
├── CLAUDE.md           # This file
├── README.md           # User-facing documentation
├── requirements.txt    # Python dependencies
├── with_toc/           # Output: PDFs with TOC inserted (gitignored)
└── docx_output/        # Output: converted DOCX files (gitignored)
```

## Running

```bash
# Install dependencies
pip install PyMuPDF python-docx Pillow lxml

# Generate TOC for all PDFs in the current directory
python3.12 add_toc.py

# Convert all PDFs to DOCX
python3.12 pdf_to_docx.py
```

Both scripts process all `*.pdf` files in the working directory (`/Users/siraj/Downloads/For_TOC_Make/` by default — change the `work_dir` variable to customize).

## Key Design Decisions

- **Equations as images, not OMML**: TeX math is rendered as 200 DPI PNG images rather than converted to Office Math XML. This preserves visual fidelity without the extreme complexity of TeX→OMML conversion.
- **Table detection from vectors**: No native table markup exists in PDFs. Tables are reconstructed by finding horizontal/vertical line intersections in the drawing layer.
- **Font mapping heuristic**: PDF fonts are mapped to closest Word-compatible equivalents (Times→Times New Roman, Helvetica→Arial, CM→Times New Roman). Unknown fonts fall back to serif/sans-serif based on font flags.
- **Text sanitization**: All text is sanitized for XML 1.0 compatibility before inserting into DOCX (removes null bytes and control characters from broken PDFs).
- **Per-page error handling**: If a single page fails, the converter adds a placeholder and continues rather than aborting the entire document.

## Common Issues

- **Type 3 fonts** (e.g., Solid State Physics PDF): Font encoding may be broken in the PDF itself. The converter reproduces whatever text PyMuPDF extracts, which may be garbled.
- **Large math-heavy PDFs**: Books like Chee Yap (538 pages of algebra) produce 100+ MB DOCX files due to equation images. This is expected.
- **Multi-column layouts**: Content is extracted in reading order but rendered as single-column. Multi-column DOCX layout via XML is fragile and not implemented.
