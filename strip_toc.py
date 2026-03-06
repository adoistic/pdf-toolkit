"""Strip generated TOC pages from PDFs for testing.

Detects TOC pages at the start of each PDF (containing "Table of Contents"
header text) and removes them, saving clean copies to a test directory.
"""
import fitz
import os
import sys

SRC_DIR = r"C:\Users\admin\Downloads\TOC_Issue\pdf"
DST_DIR = r"C:\Users\admin\Downloads\TOC_Issue\originals"


def count_toc_pages(doc):
    """Count how many leading pages are generated TOC pages."""
    toc_pages = 0
    for pidx in range(min(5, len(doc))):  # TOC is at most ~5 pages
        page = doc[pidx]
        text = page.get_text("text").strip()
        # Our generated TOC pages start with "Table of Contents"
        if "Table of Contents" in text and len(text) < 5000:
            toc_pages += 1
        else:
            break
    return toc_pages


def main():
    os.makedirs(DST_DIR, exist_ok=True)

    for fname in sorted(os.listdir(SRC_DIR)):
        if not fname.endswith(".pdf"):
            continue
        src = os.path.join(SRC_DIR, fname)
        dst = os.path.join(DST_DIR, fname)

        doc = fitz.open(src)
        n = count_toc_pages(doc)

        if n == 0:
            print(f"  {fname}: no TOC pages found, copying as-is")
            doc.save(dst, deflate=True, garbage=4)
        else:
            print(f"  {fname}: removing {n} TOC page(s)")
            doc.delete_pages(list(range(n)))
            # Keep bookmarks — PyMuPDF auto-adjusts page refs after delete
            doc.save(dst, deflate=True, garbage=4)

        doc.close()

    print(f"\nClean copies saved to: {DST_DIR}")


if __name__ == "__main__":
    main()
