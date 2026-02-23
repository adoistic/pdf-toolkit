# pdf-toolkit

Two Python scripts for working with PDF files using open-source tools only — no Adobe Acrobat needed.

| Tool | What it does |
|------|-------------|
| **`add_toc.py`** | Detects chapter headings and inserts a Table of Contents page into each PDF |
| **`pdf_to_docx.py`** | Converts PDFs to Word documents with high fidelity — preserving fonts, images, tables, equations, and layout |

## Quick Start

```bash
# Install dependencies
pip install PyMuPDF python-docx Pillow lxml

# Place your PDFs in a folder, then:

# 1. Add Table of Contents to all PDFs
python3 add_toc.py          # Output: with_toc/ subfolder

# 2. Convert all PDFs to Word
python3 pdf_to_docx.py      # Output: docx_output/ subfolder
```

> **Note:** Both scripts process all `*.pdf` files in the current working directory. Edit the `work_dir` variable at the bottom of each script to change the input folder.

## Requirements

- Python 3.10+
- [PyMuPDF](https://pymupdf.readthedocs.io/) (fitz) — PDF reading and manipulation
- [python-docx](https://python-docx.readthedocs.io/) — Word document generation
- [Pillow](https://pillow.readthedocs.io/) — Image processing
- [lxml](https://lxml.de/) — XML manipulation (installed with python-docx)

```bash
pip install -r requirements.txt
```

---

## Tool 1: `add_toc.py` — Automatic Table of Contents

Scans each PDF for chapter headings using font metrics (size, bold weight) and inserts a formatted TOC page at the beginning.

### How it works

1. **Body text detection** — The most frequently used font size (by character count) is identified as body text
2. **Heading extraction** — Bold text at font sizes larger than body text is collected as heading candidates
3. **Multi-line merge** — Adjacent heading fragments on the same page are merged (handles line-wrapped titles)
4. **H1/H2 selection** — If the largest heading level has more than 3 entries, it's used for chapters. If 3 or fewer, it's likely a book title, so the next level down is used instead
5. **Over-extraction guard** — If more than 50 headings are detected, the script narrows to only "Chapter N"-prefixed entries
6. **TOC generation** — A formatted TOC page with dot leaders and page numbers is created and inserted before page 1
7. **Bookmarks** — PDF bookmark metadata is also set for navigation in PDF viewers

### Example output

```
$ python3 add_toc.py

Processing: Cisco ASA Handbook.pdf
  Body text: 12.0pt, Heading sizes: [13.5, 18.0]
  Found 12 H1 headings, 1 H2 heading
  TOC: 1 page
  Saved: with_toc/Cisco ASA Handbook.pdf

SUMMARY: 12/12 succeeded
```

### Features

- Handles TeX/LaTeX PDFs with Computer Modern fonts
- Detects running headers (text repeated on >25% of pages) and excludes them
- Keyword-based fallback for PDFs where font metrics alone aren't enough ("Chapter N" pattern matching)
- Deduplication of headings appearing on adjacent pages
- Hierarchy normalization to ensure valid PDF bookmark structure

---

## Tool 2: `pdf_to_docx.py` — PDF to Word Converter

Converts each PDF to a `.docx` file, preserving the original layout as closely as possible.

### What gets preserved

| Element | How it's handled |
|---------|-----------------|
| **Text** | Extracted with font name, size, bold/italic, color per span. PDF's broken lines are rejoined into proper paragraphs |
| **Fonts** | Mapped to closest Word-compatible equivalent (Times→Times New Roman, Helvetica→Arial, Courier→Courier New, etc.) |
| **Page layout** | Original page size and margins are preserved |
| **Images** | Extracted at original resolution, positioned as block or inline based on width |
| **Tables** | Detected from vector line drawings (horizontal/vertical line intersections), reconstructed with cell text and borders |
| **Equations** | TeX math (detected via CM/AMS font names) is rendered as 200 DPI images |
| **Headers/Footers** | Running text patterns are detected, excluded from body, and placed in Word header/footer zones |
| **Headings** | Detected via font-size analysis and styled with Word heading styles |
| **Hyphenation** | Line-end hyphens are removed when joining lines |
| **Colors** | RGB, grayscale, and CMYK colors are mapped to Word RGB values |

### Example output

```
$ python3 pdf_to_docx.py

Processing: Cisco ASA Handbook.pdf
  Pages: 884
  Page: 612x792pt, margins: L=72 R=71 T=72 B=85
  Body: 12.0pt, Heading sizes: [13.5, 18.0]
  Total elements rendered: 8706
  Saved: docx_output/Cisco ASA Handbook.docx (6.5 MB)

SUMMARY: 12/12 succeeded
```

### Paragraph assembly

PDF breaks text into individual lines. The converter reassembles them into paragraphs using these heuristics:

- **Join lines** when: vertical gap is small (< 1.8x font height), same font size, no list bullet or number at start
- **Don't join** when: font size changes (heading boundary), large vertical gap, line ends with `.!?` and next starts uppercase with a gap, significant indentation change
- **Hyphen handling**: if a line ends with `-` and the next starts lowercase, the hyphen is removed and words are joined

### Known limitations

- **Equations** are images, not editable Office Math — this preserves visual fidelity for complex TeX math
- **Vector diagrams** (line art, flowcharts) are not currently extracted as images
- **Multi-column layouts** are flowed into single-column output
- **Type 3 fonts** with broken encoding produce garbled text (this is a PDF-level issue)
- **Exact line breaks** differ from the original because Word re-flows text using its own line-breaking algorithm

---

## Tested On

Successfully processed 12 diverse PDFs covering:

| Type | Examples | Pages |
|------|----------|-------|
| Math/Science textbooks | Algorithmic Algebra, Noncommutative Geometry | 338–538 |
| Scientific protocols | Chromatin Protocols | 508 |
| Networking/IT handbooks | Cisco ASA, Communications Network | 784–884 |
| Security books | Hack Attacks, Hackers Beware, Hacking Firewalls | 354–792 |
| Encyclopedia | Encyclopedia of Public Health | 384 |
| Materials science | Creep in Metals and Alloys | 279 |
| Physics (Type 3 fonts) | Solid State Physics | 406 |

---

## Project Structure

```
.
├── add_toc.py           # TOC generation script (951 lines)
├── pdf_to_docx.py       # PDF → DOCX converter (1819 lines)
├── requirements.txt     # Python dependencies
├── CLAUDE.md            # Claude Code project context
├── README.md            # This file
├── with_toc/            # Output: PDFs with TOC (generated, gitignored)
└── docx_output/         # Output: DOCX files (generated, gitignored)
```

## License

MIT
