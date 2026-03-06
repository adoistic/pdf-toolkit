#!/usr/bin/env python3
# ─────────────────────────────────────────────────────────────────────
# Copyright (c) 2025-2026 Thothica Private Limited, Delhi, India.
# All rights reserved.  Proprietary and confidential.
# Unauthorized copying or distribution is strictly prohibited.
# ─────────────────────────────────────────────────────────────────────
"""
Helper: process a single PDF in an isolated process.
Used by gui.py so that a crash on one file can't kill the whole app.

Usage:  python _run_one.py <toc|docx> <input.pdf> <output_path> [--h1-only] [--no-count-toc]
"""

import sys
import os
import tempfile
import shutil

# Ensure the toolkit modules are importable
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

mode     = sys.argv[1]                   # "toc" or "docx"
src      = sys.argv[2]                   # input PDF path
dst      = sys.argv[3]                   # output file path
h1_only        = "--h1-only" in sys.argv        # optional flag
count_toc_pages = False  # TOC page numbers always match original PDF

if mode == "toc":
    from add_toc import process_pdf
else:
    from pdf_to_docx import process_pdf


def _clean_stale(path):
    """Try to remove an existing output file (may be a 0-byte leftover from a crash)."""
    if not os.path.exists(path):
        return True
    try:
        os.remove(path)
        return True
    except PermissionError:
        return False
    except Exception:
        return False


# ── Try to clear a stale/locked output file from a previous crash ────────
if not _clean_stale(dst):
    print(f"  WARNING: Output file is locked (stale from a previous crash).")
    print(f"           Saving to a temporary name instead.")
    # Write to a sibling temp file in the same output directory
    out_dir = os.path.dirname(dst)
    base    = os.path.basename(dst)
    name, ext = os.path.splitext(base)
    dst = os.path.join(out_dir, f"{name}_new{ext}")
    # If even the fallback exists and is locked, give up
    if not _clean_stale(dst):
        print(f"  ERROR: Cannot write to output — file locked.  Skipping.")
        print("__RESULT__:failed")
        sys.exit(0)

try:
    if mode == "toc":
        result = process_pdf(src, dst, h1_only=h1_only,
                             count_toc_pages=count_toc_pages)
    else:
        result = process_pdf(src, dst)
except Exception as e:
    print(f"  FATAL: {e}")
    import traceback
    traceback.print_exc()
    result = "failed"

# Machine-readable result tag for the GUI to parse
print(f"__RESULT__:{result}")
