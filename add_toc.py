#!/usr/bin/env python3
# ─────────────────────────────────────────────────────────────────────
# Copyright (c) 2025-2026 Thothica Private Limited, Delhi, India.
# All rights reserved.  Proprietary and confidential.
# Unauthorized copying or distribution is strictly prohibited.
# ─────────────────────────────────────────────────────────────────────
"""
add_toc.py — Automatically extract chapter headings from structured PDFs
and insert a Table of Contents page at the beginning of each file.

Uses PyMuPDF (fitz) for all PDF operations.
Output is saved to a 'with_toc/' subfolder.
"""

import sys
import io

# Force UTF-8 stdout so Unicode symbols in print() don't crash on Windows cp1252
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(
        sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True,
    )

import fitz
import os
import re
import shutil
from dataclasses import dataclass
from collections import Counter, defaultdict
from pathlib import Path

# ─── Layout constants for generated TOC pages ───────────────────────────────
MARGIN_TOP = 50
MARGIN_BOTTOM = 50
MARGIN_LEFT = 45
MARGIN_RIGHT = 45
TOC_TITLE_SIZE = 14
TOC_CHAPTER_SIZE = 9
TOC_SECTION_SIZE = 8
TOC_LINE_SPACING_FACTOR = 1.6
TOC_SECTION_INDENT = 20
TOC_MAX_LINES = 3                       # max wrapped lines per entry; longer → removed
DOT_LEADER_COLOR = (0.5, 0.5, 0.5)
SEPARATOR_COLOR = (0.6, 0.6, 0.6)

# ─── Detection thresholds ───────────────────────────────────────────────────
FONT_SIZE_CLUSTER_THRESHOLD = 0.6   # pt — sizes within this are same level
PAGE_NUM_MARGIN = 60                # px from top/bottom = header/footer zone
DEDUP_PAGE_RANGE = 4                # pages within which same text = duplicate
RUNNING_HEADER_RATIO = 0.25         # text on >25% of pages = running header

# ─── Bold font-name patterns (case-insensitive substring match) ─────────────
BOLD_FONT_PATTERNS = ["bold", "cmbx", "advtib"]


# ─── Data structures ────────────────────────────────────────────────────────

@dataclass
class HeadingCandidate:
    text: str
    page_num: int           # 0-indexed into original document
    font_size: float
    is_bold: bool
    y_position: float
    level: int = 0          # assigned after clustering: 1 = H1, 2 = H2 …
    block_bbox: tuple = ()  # bounding box of the parent block


# ─── Helper functions ────────────────────────────────────────────────────────

def is_bold_font(font_name: str, flags: int) -> bool:
    """Detect bold using both font flags and font-name heuristics."""
    if flags & (1 << 4):
        return True
    name_lower = font_name.lower()
    for pat in BOLD_FONT_PATTERNS:
        if pat in name_lower:
            return True
    return False


def cluster_font_sizes(sizes, threshold=FONT_SIZE_CLUSTER_THRESHOLD):
    """
    Group font sizes within *threshold* pt into canonical buckets.
    Returns dict  {original_size: canonical_size}  where canonical = max in cluster.
    """
    if not sizes:
        return {}
    sorted_sizes = sorted(set(sizes), reverse=True)
    clusters = []
    for sz in sorted_sizes:
        placed = False
        for cluster in clusters:
            if abs(cluster[0] - sz) <= threshold:
                cluster.append(sz)
                placed = True
                break
        if not placed:
            clusters.append([sz])
    mapping = {}
    for cluster in clusters:
        canonical = max(cluster)
        for sz in cluster:
            mapping[sz] = canonical
    return mapping


def is_noise_text(text, y_pos, page_height):
    """
    Filter out page numbers, headers/footers, and non-textual debris.
    This is a CONSERVATIVE filter — only removes clearly non-heading text.
    """
    t = text.strip()
    if not t:
        return True
    # Pure numeric in header / footer zone
    stripped = t.replace(".", "").replace("-", "").replace(" ", "")
    if stripped.isdigit():
        if y_pos < PAGE_NUM_MARGIN or y_pos > page_height - PAGE_NUM_MARGIN:
            return True
        # Also filter standalone large page numbers in content area
        # that are just chapter number decorations (like "1" or "78")
        if len(stripped) <= 3 and stripped.isdigit():
            return True
    # Negative y → rendered in the page margin / running header area
    if y_pos < 0:
        return True
    # Mostly non-alphabetic and >3 chars → equation fragment
    alpha = sum(1 for c in t if c.isalpha())
    if len(t) > 3 and alpha / len(t) < 0.40:
        return True
    # Very short text with special chars → likely math/symbols
    if len(t) <= 4 and not t[0].isalpha():
        return True
    return False


def find_running_headers(doc, sample_pages=100):
    """
    Scan a sample of pages and return normalized text strings that appear
    on a large fraction of them (i.e. running headers / footers).
    """
    n = min(sample_pages, len(doc))
    text_freq = Counter()
    for pidx in range(n):
        page = doc[pidx]
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        seen_on_page = set()
        for blk in blocks:
            if "lines" not in blk:
                continue
            for line in blk["lines"]:
                lt = "".join(s["text"] for s in line["spans"]).strip()
                if lt and len(lt) > 3:
                    norm = re.sub(r"\d+", "#", lt).strip()
                    seen_on_page.add(norm)
        for t in seen_on_page:
            text_freq[t] += 1
    threshold = max(4, n * RUNNING_HEADER_RATIO)
    return {t for t, c in text_freq.items() if c >= threshold}


def looks_like_heading_text(text):
    """
    Return True if the text looks like it could plausibly be a chapter or
    section heading — starts with a letter or digit, has real words.
    Reject equation fragments, symbols-only, etc.
    """
    t = text.strip()
    if len(t) < 3:
        return False
    # Must start with a letter or digit (for numbered sections like "1.1 Foo")
    if not (t[0].isalpha() or t[0].isdigit()):
        return False
    # Must have reasonable alphabetic content
    alpha = sum(1 for c in t if c.isalpha())
    if alpha < 3:
        return False
    # Reject if it's mostly math symbols
    if alpha / len(t) < 0.40:
        return False
    # Reject text containing mathematical operators — equation fragments
    if re.search(r"[=<>{}\|\\]", t):
        return False
    # Reject math-like patterns: parenthesized expressions with no spaces
    # e.g. "E(p)/E", "f(x)", "A(n)+B"
    if re.search(r"\w\(.*\)", t) and " " not in t:
        return False
    return True


# ─── Scanned PDF Detection ───────────────────────────────────────────────────

def is_scanned_pdf(doc, sample_size=10):
    """Detect if a PDF is mostly scanned images with no extractable text.
    Samples up to sample_size pages evenly distributed through the document.
    Returns True if <20% of sampled pages have meaningful text (>50 chars).
    """
    num_pages = len(doc)
    if num_pages == 0:
        return True
    if num_pages <= sample_size:
        indices = list(range(num_pages))
    else:
        indices = [int(i * num_pages / sample_size) for i in range(sample_size)]
    pages_with_text = 0
    for idx in indices:
        page = doc[idx]
        text = page.get_text("text").strip()
        if len(text) > 50:
            pages_with_text += 1
    ratio = pages_with_text / len(indices)
    return ratio < 0.2


# ─── Core heading extraction (font-metric based) ────────────────────────────

def extract_headings_by_font(doc):
    """
    Walk every page of *doc*, collect bold text spans whose font size exceeds
    the document's body-text size.  Returns a list of HeadingCandidate or []
    if font metrics are unusable (e.g. Type-3 fonts reporting 0.1 pt).
    """
    running_headers = find_running_headers(doc)

    # ── Pass 1 — collect ALL text spans with metadata ────────────────────
    all_spans = []   # (page, size, bold, text, y, block_bbox, font_name)
    for pidx in range(len(doc)):
        page = doc[pidx]
        ph = page.rect.height
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        for blk in blocks:
            if "lines" not in blk:
                continue
            bbx = tuple(blk["bbox"])
            for line in blk["lines"]:
                for span in line["spans"]:
                    txt = span["text"].strip()
                    if not txt or len(txt) < 2:
                        continue
                    sz = round(span["size"], 1)
                    bold = is_bold_font(span["font"], span["flags"])
                    y = span["bbox"][1]
                    all_spans.append((pidx, sz, bold, txt, y, bbx, span["font"]))

    if not all_spans:
        return []

    # ── Detect broken fonts (all same tiny size) ─────────────────────────
    sizes = [s[1] for s in all_spans]
    unique = set(sizes)
    if len(unique) <= 3 and min(unique) < 1.0:
        return []           # caller should use keyword fallback

    # ── Body text = most frequent font size (considering non-bold only) ──
    # Prefer non-bold for body detection to avoid bold-heavy docs skewing it
    non_bold_sizes = [s[1] for s in all_spans if not s[2]]
    if non_bold_sizes:
        body_size = Counter(non_bold_sizes).most_common(1)[0][0]
    else:
        body_size = Counter(sizes).most_common(1)[0][0]

    # ── Cluster font sizes ───────────────────────────────────────────────
    cluster_map = cluster_font_sizes(list(unique))

    # ── Identify heading-level canonical sizes (> body) ──────────────────
    heading_sizes = sorted({
        canonical
        for _, canonical in cluster_map.items()
        if canonical > body_size + 0.5
    }, reverse=True)

    if not heading_sizes:
        return []

    # Keep top 3 heading sizes to avoid pulling in section-level bold text
    heading_sizes_set = set(heading_sizes[:3])

    # ── Collect candidates (applying noise + running-header filters) ─────
    candidates = []
    for pidx, sz, bold, txt, y, bbx, fname in all_spans:
        canonical = cluster_map.get(sz, sz)
        if canonical not in heading_sizes_set:
            continue
        # Must be bold, OR very large compared to body (for PDFs where
        # chapter titles use a non-bold display font like LegacySans-Medium)
        if not bold and canonical <= body_size * 2:
            continue
        # Apply noise filters
        ph = doc[pidx].rect.height
        if is_noise_text(txt, y, ph):
            continue
        # Check running headers
        norm = re.sub(r"\d+", "#", txt).strip()
        if norm in running_headers:
            continue
        # Check it looks like real heading text
        if not looks_like_heading_text(txt):
            continue

        candidates.append(HeadingCandidate(
            text=txt, page_num=pidx, font_size=canonical,
            is_bold=bold, y_position=y, block_bbox=bbx,
        ))

    # ── Assign heading levels (1 = largest) ──────────────────────────────
    used_sizes = sorted({c.font_size for c in candidates}, reverse=True)
    sz2lvl = {s: i + 1 for i, s in enumerate(used_sizes)}
    for c in candidates:
        c.level = sz2lvl.get(c.font_size, 99)

    return candidates


# ─── Keyword-based fallback (Type-3 / broken fonts) ─────────────────────────

CHAPTER_RE = re.compile(
    r"^(Chapter|Lecture|Part|CHAPTER|LECTURE|PART)\s+(\d+|[IVXLCDM]+)",
    re.IGNORECASE,
)


def extract_headings_by_keyword(doc):
    """
    Fallback when font metrics are unusable.  Searches for lines matching
    'Chapter N' / 'Lecture N' / 'Part N' and uses them as H1 entries.
    """
    candidates = []
    seen = {}  # chapter_id → first page

    for pidx in range(len(doc)):
        page = doc[pidx]
        text = page.get_text("text")
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        if not lines:
            continue

        for li, line in enumerate(lines[:10]):
            m = CHAPTER_RE.match(line)
            if not m:
                continue
            ctype, cnum = m.group(1), m.group(2)
            cid = f"{ctype} {cnum}"

            # ── skip if this same chapter was already seen recently ──
            if cid in seen and pidx - seen[cid] < DEDUP_PAGE_RANGE:
                continue
            seen[cid] = pidx

            # ── grab the chapter title from the next non-empty line(s) ───
            title_parts = []
            for nxt in lines[li + 1 : li + 4]:
                if nxt and not CHAPTER_RE.match(nxt):
                    # Filter lines that look like page-content (too long = body text)
                    if len(nxt) > 100:
                        break
                    title_parts.append(nxt)
                else:
                    break
            title = " ".join(title_parts).strip()
            title = re.sub(r"\s{2,}", " ", title)   # collapse whitespace
            full = f"{cid}: {title}" if title else cid

            candidates.append(HeadingCandidate(
                text=full, page_num=pidx, font_size=0,
                is_bold=True, y_position=0, level=1,
            ))
            break   # only first match per page

    return candidates


def extract_headings_from_outline(doc):
    """
    Last-resort fallback: use the PDF's existing bookmark / outline tree
    (doc.get_toc()) as heading entries.  Returns a list of HeadingCandidate
    or [] if the document has no outline.

    TOC entries are [level, title, 1-based-page].  We convert to 0-based
    page numbers and keep only level 1 and 2 entries.
    """
    toc = doc.get_toc()
    if not toc:
        return []

    candidates = []
    for level, title, page_1based in toc:
        if level > 2:
            continue
        title = title.strip()
        if not title:
            continue
        page_0 = max(0, page_1based - 1)      # convert to 0-based
        if page_0 >= len(doc):
            page_0 = len(doc) - 1
        candidates.append(HeadingCandidate(
            text=title, page_num=page_0, font_size=0,
            is_bold=True, y_position=0, level=level,
        ))
    return candidates


# ─── Post-processing ─────────────────────────────────────────────────────────

# Words that indicate a heading line is a continuation of a previous line
_CONTINUATION_STARTERS = {"and", "or", "of", "for", "in", "the", "with",
                          "to", "on", "by", "a", "an", "its", "their"}


def _is_continuation_line(text):
    """True if the text starts with a lowercase letter or common connector."""
    t = text.strip()
    if not t:
        return False
    if t[0].islower():
        return True
    first_word = t.split()[0].lower().rstrip(".,;:")
    return first_word in _CONTINUATION_STARTERS


def merge_multiline_headings(headings):
    """
    Two-pass merge for wrapped heading titles:

    Pass 1 — same block: consecutive headings on the same page, same level,
    same block, within 45px vertically → merge.

    Pass 2 — cross-block: consecutive headings on the same page, same level,
    same font size, within 50px vertically, where the second line looks like a
    continuation (starts lowercase / common connector) → merge.
    """
    if not headings:
        return []

    # ── Pass 1: same-block merge ─────────────────────────────────────────
    pass1 = []
    i = 0
    while i < len(headings):
        cur = headings[i]
        parts = [cur.text]
        j = i + 1
        while j < len(headings):
            nxt = headings[j]
            if (nxt.page_num == cur.page_num
                    and nxt.level == cur.level
                    and nxt.block_bbox == cur.block_bbox
                    and abs(nxt.y_position - headings[j - 1].y_position) < 45):
                parts.append(nxt.text)
                j += 1
            else:
                break
        pass1.append(HeadingCandidate(
            text=" ".join(parts),
            page_num=cur.page_num,
            font_size=cur.font_size,
            is_bold=cur.is_bold,
            y_position=cur.y_position,
            level=cur.level,
            block_bbox=cur.block_bbox,
        ))
        i = j

    # ── Pass 2: cross-block continuation merge ───────────────────────────
    pass2 = []
    i = 0
    while i < len(pass1):
        cur = pass1[i]
        parts = [cur.text]
        j = i + 1
        while j < len(pass1):
            nxt = pass1[j]
            if (nxt.page_num == cur.page_num
                    and nxt.level == cur.level
                    and abs(nxt.font_size - cur.font_size) < 1.0
                    and abs(nxt.y_position - pass1[j - 1].y_position) < 50
                    and _is_continuation_line(nxt.text)):
                parts.append(nxt.text)
                j += 1
            else:
                break
        pass2.append(HeadingCandidate(
            text=" ".join(parts),
            page_num=cur.page_num,
            font_size=cur.font_size,
            is_bold=cur.is_bold,
            y_position=cur.y_position,
            level=cur.level,
            block_bbox=cur.block_bbox,
        ))
        i = j

    return pass2


def merge_adjacent_headings_across_blocks(headings):
    """
    Merge headings on the SAME PAGE at DIFFERENT levels when one is clearly
    a label and the other is the title.

    Patterns handled:
      • "Chapter 2" at 13pt + "Diffusional Creep" at 17pt → "Chapter 2: Diffusional Creep"
      • "PART" + "Two" + "Putting it All Together" on same page → "Part Two: Putting it All Together"
      • "CHAPTER" + "3" + "Understanding Communication Mediums" → "Chapter 3: Understanding…"
    """
    if not headings:
        return []

    # Group headings by page to find multi-block label patterns
    by_page = defaultdict(list)
    for i, h in enumerate(headings):
        by_page[h.page_num].append((i, h))

    consumed = set()  # indices consumed by merging
    result = []

    for i, h in enumerate(headings):
        if i in consumed:
            continue

        page_group = by_page[h.page_num]

        # Pattern 1: "Chapter/Part/Lecture N" label + title on same page
        if (CHAPTER_RE.match(h.text.strip())
                and i + 1 < len(headings)
                and headings[i + 1].page_num == h.page_num
                and i + 1 not in consumed):
            nxt = headings[i + 1]
            combined = HeadingCandidate(
                text=f"{h.text.strip()}: {nxt.text.strip()}",
                page_num=h.page_num,
                font_size=max(h.font_size, nxt.font_size),
                is_bold=True,
                y_position=min(h.y_position, nxt.y_position),
                level=min(h.level, nxt.level),
                block_bbox=h.block_bbox,
            )
            result.append(combined)
            consumed.add(i)
            consumed.add(i + 1)
            continue

        # Pattern 2: bare keyword "PART"/"CHAPTER"/"SECTION" (no number)
        # followed by a number-word and then a title, all on same page
        bare = h.text.strip().upper()
        if bare in ("PART", "CHAPTER", "SECTION", "LECTURE"):
            # Collect remaining headings on same page
            remaining = [(j, hh) for j, hh in page_group
                         if j > i and j not in consumed]
            if remaining:
                # Next item is the number/name (e.g. "Two", "3")
                num_idx, num_h = remaining[0]
                label = f"{bare.title()} {num_h.text.strip()}"
                consumed.add(i)
                consumed.add(num_idx)
                # If there's also a title after the number
                if len(remaining) > 1:
                    title_idx, title_h = remaining[1]
                    label = f"{label}: {title_h.text.strip()}"
                    consumed.add(title_idx)
                    use_level = min(h.level, num_h.level, title_h.level)
                    use_size = max(h.font_size, num_h.font_size, title_h.font_size)
                else:
                    use_level = min(h.level, num_h.level)
                    use_size = max(h.font_size, num_h.font_size)
                result.append(HeadingCandidate(
                    text=label, page_num=h.page_num,
                    font_size=use_size, is_bold=True,
                    y_position=h.y_position, level=use_level,
                    block_bbox=h.block_bbox,
                ))
                continue

        result.append(h)

    return result


def deduplicate_headings(headings):
    """Remove duplicate heading text appearing within DEDUP_PAGE_RANGE pages."""
    result = []
    seen = {}   # lowered text → last page
    for h in headings:
        key = h.text.strip().lower()
        if key in seen and abs(h.page_num - seen[key]) < DEDUP_PAGE_RANGE:
            continue
        seen[key] = h.page_num
        result.append(h)
    return result


def select_toc_headings(headings, h1_only=False):
    """
    Apply the user's rule:
      • If H1 (largest heading size) has > 3 entries → keep H1 (+ H2 if present)
      • If H1 has ≤ 3 entries → it's probably a book/part title;
        promote H2 to be the primary TOC level.
    Only two levels are ever included in the final TOC.

    If *h1_only* is True, only H1-level entries are kept (H2 is dropped).

    Sanity check: if we end up with too many entries, narrow the selection.
    """
    if not headings:
        return []

    by_level = defaultdict(list)
    for h in headings:
        by_level[h.level].append(h)

    levels = sorted(by_level.keys())       # 1, 2, 3 …
    if not levels:
        return []

    h1_lvl = levels[0]
    h1_entries = by_level[h1_lvl]

    if len(h1_entries) > 3:
        # H1 is genuine chapter level → keep H1 + optional H2
        keep = {h1_lvl}
        if len(levels) > 1:
            keep.add(levels[1])
        selected = [h for h in headings if h.level in keep]
    else:
        # H1 ≤ 3 → promote H2
        if len(levels) < 2:
            return h1_entries
        h2_lvl = levels[1]
        for h in by_level[h2_lvl]:
            h.level = 1
        selected = list(by_level[h2_lvl])
        if len(levels) > 2:
            h3_lvl = levels[2]
            for h in by_level[h3_lvl]:
                h.level = 2
            selected += by_level[h3_lvl]
        selected.sort(key=lambda h: (h.page_num, h.y_position))

    # ── Sanity check: over-extraction guard ──────────────────────────────
    selected = _guard_over_extraction(selected)

    # ── Optional: drop everything except H1 ──────────────────────────────
    if h1_only:
        selected = [h for h in selected if h.level == 1]

    return selected


def _guard_over_extraction(headings):
    """
    If there are way too many entries, the detection is probably too broad
    (e.g. every bold phrase is being treated as a heading).

    Strategies:
      1. If there are "Chapter N"-prefixed entries among many H1s, keep only those
         as H1 and demote the rest.
      2. If total entries > 100 with no "Chapter" structure, keep only H1 (drop H2).
      3. If H1 alone > 80, keep only the largest font-size level.
    """
    h1s = [h for h in headings if h.level == 1]

    if len(h1s) <= 50:
        return headings  # reasonable count, no action needed

    # Strategy 1: filter to "Chapter"-prefixed entries as H1
    chapter_h1s = [h for h in h1s if CHAPTER_RE.match(h.text.strip())]
    if len(chapter_h1s) >= 4:
        # Keep chapter entries as H1, demote everything else
        chapter_pages = {h.page_num for h in chapter_h1s}
        result = []
        for h in headings:
            if h.level == 1 and h not in chapter_h1s:
                h.level = 2  # demote non-chapter H1 to H2
            result.append(h)
        return result

    # Strategy 2: too many entries total — keep only H1 (the biggest heading size)
    if len(headings) > 100:
        biggest_size = max(h.font_size for h in h1s)
        filtered = [h for h in headings
                    if h.level == 1 and abs(h.font_size - biggest_size) < 1.0]
        if len(filtered) >= 4:
            for h in filtered:
                h.level = 1
            return filtered

    # Strategy 3: if nothing else works, just keep H1
    return [h for h in headings if h.level == 1]


def normalize_hierarchy(headings):
    """
    Ensure the heading list satisfies PyMuPDF's set_toc() rules:
      1. First item must be level 1
      2. Level can only increase by 1 at a time (no jumping from 1 to 3)
    Also normalizes so levels are contiguous (1, 2) not (1, 3).
    """
    if not headings:
        return []

    # First, remap to contiguous levels
    used_levels = sorted(set(h.level for h in headings))
    remap = {old: new + 1 for new, old in enumerate(used_levels)}
    for h in headings:
        h.level = remap[h.level]

    # Ensure first entry is level 1
    if headings[0].level != 1:
        for h in headings:
            h.level = 1  # flatten if structure is broken

    # Ensure no jumps > 1
    for i in range(1, len(headings)):
        prev_level = headings[i - 1].level
        if headings[i].level > prev_level + 1:
            headings[i].level = prev_level + 1

    return headings


def supplement_with_keyword_chapters(doc, headings):
    """
    After font-based extraction, scan for CHAPTER/PART title pages that
    were missed (e.g. because their font size wasn't in the top 3).

    STRICT matching: only triggers on pages where "CHAPTER" or similar
    appears as a short, standalone line near the top — NOT embedded in
    body text paragraphs.
    """
    # Build a set of pages that already have headings (± 2 page buffer)
    existing_pages = set()
    for h in headings:
        for offset in range(-2, 3):
            existing_pages.add(h.page_num + offset)

    # Strict regex: line must be ONLY "Chapter N" or "CHAPTER N" (+ optional punctuation)
    strict_re = re.compile(
        r"^(Chapter|Lecture|CHAPTER|LECTURE)\s+(\d+|[IVXLCDM]+)\s*[.:;]?\s*$",
        re.IGNORECASE,
    )

    # First pass: count how many pages each chapter ID appears on.
    # If a chapter ID appears on >3 pages, it's a running header, not a real chapter.
    chapter_page_count = Counter()
    for pidx in range(len(doc)):
        page = doc[pidx]
        text = page.get_text("text")
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        for line in lines[:4]:
            m = strict_re.match(line)
            if m:
                cid = f"{m.group(1).title()} {m.group(2)}"
                chapter_page_count[cid] += 1
                break
            if line.upper() in ("CHAPTER", "SECTION", "LECTURE"):
                if lines.index(line) + 1 < len(lines):
                    num = lines[lines.index(line) + 1].strip()
                    if num.isdigit() or re.match(r"^[IVXLCDM]+$", num, re.IGNORECASE):
                        cid = f"Chapter {num}"
                        chapter_page_count[cid] += 1
                break

    # Filter: only keep chapter IDs that appear on ≤3 pages (real chapter title pages)
    valid_chapters = {cid for cid, cnt in chapter_page_count.items() if cnt <= 3}

    for pidx in range(len(doc)):
        if pidx in existing_pages:
            continue
        page = doc[pidx]
        text = page.get_text("text")
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        if not lines:
            continue

        # Only check very first few lines — chapter pages have the label near the top
        for li, line in enumerate(lines[:4]):
            m = strict_re.match(line)
            if not m:
                # Bare keyword on its own line (e.g. "CHAPTER" then "3" on next line)
                if line.upper() in ("CHAPTER", "SECTION", "LECTURE") and li + 1 < len(lines):
                    num_line = lines[li + 1].strip()
                    # Must be a short number/numeral
                    if len(num_line) > 10 or not (num_line.isdigit() or
                            re.match(r"^[IVXLCDM]+$", num_line, re.IGNORECASE)):
                        continue
                    # Grab title from lines after the number (must be short = heading, not body)
                    title_parts = []
                    for nxt in lines[li + 2: li + 5]:
                        if nxt and len(nxt) < 80 and not strict_re.match(nxt):
                            title_parts.append(nxt)
                        else:
                            break
                    title = " ".join(title_parts).strip()
                    cid = f"Chapter {num_line}"
                    if cid not in valid_chapters:
                        continue
                    label = cid
                    if title:
                        label = f"{label}: {title}"
                    headings.append(HeadingCandidate(
                        text=label, page_num=pidx, font_size=999,
                        is_bold=True, y_position=0, level=1,
                    ))
                    existing_pages.add(pidx)
                    break
                continue

            ctype, cnum = m.group(1), m.group(2)
            cid = f"{ctype.title()} {cnum}"
            if cid not in valid_chapters:
                continue
            # Grab title from subsequent short lines
            title_parts = []
            for nxt in lines[li + 1: li + 4]:
                if nxt and len(nxt) < 80 and not strict_re.match(nxt):
                    title_parts.append(nxt)
                else:
                    break
            title = " ".join(title_parts).strip()
            label = f"{ctype.title()} {cnum}"
            if title:
                label = f"{label}: {title}"

            headings.append(HeadingCandidate(
                text=label, page_num=pidx, font_size=999,
                is_bold=True, y_position=0, level=1,
            ))
            existing_pages.add(pidx)
            break

    headings.sort(key=lambda h: (h.page_num, h.y_position))
    return headings


# ─── TOC text wrapping ──────────────────────────────────────────────────────

def wrap_toc_entry(text, font, fontsize, full_width, last_line_width):
    """
    Word-wrap a TOC entry.
    full_width:      available width for non-last lines.
    last_line_width: available width for the last line (room for dots + page num).
    Returns a list of line strings.
    """
    words = text.split()
    if not words:
        return [text]

    lines = []
    current = words[0]

    for word in words[1:]:
        test = current + " " + word
        if font.text_length(test, fontsize=fontsize) <= full_width:
            current = test
        else:
            lines.append(current)
            current = word
    lines.append(current)

    # Ensure last line fits within last_line_width (room for dots + page number)
    while (len(lines) > 0
           and font.text_length(lines[-1], fontsize=fontsize) > last_line_width):
        last_words = lines[-1].split()
        if len(last_words) <= 1:
            break  # single word too long, can't split further
        lines[-1] = " ".join(last_words[:-1])
        lines.append(last_words[-1])

    return lines


# ─── TOC page creation ──────────────────────────────────────────────────────

def create_toc_pages(headings, page_width, page_height, count_toc_pages=True):
    """
    Build a fitz.Document containing the TOC page(s).
    Long entries wrap up to TOC_MAX_LINES lines; entries that still
    don't fit are removed entirely.
    Returns (toc_doc, filtered_headings, num_toc_pages).

    If *count_toc_pages* is True, page numbers include the TOC pages
    (original_page + num_toc_pages + 1).  If False, page numbers match
    the original PDF (original_page + 1).
    """
    toc_doc = fitz.open()

    font_regular = fitz.Font("helv")
    font_bold    = fitz.Font("hebo")
    right_edge   = page_width - MARGIN_RIGHT

    # ── Step 1: wrap every entry and drop those exceeding TOC_MAX_LINES ──
    prepared = []   # (heading, [wrapped_lines], fontsize, fontname, font_obj, indent)
    for h in headings:
        if h.level == 1:
            fs, fn, fobj, ind = TOC_CHAPTER_SIZE, "hebo", font_bold, 0
        else:
            fs, fn, fobj, ind = TOC_SECTION_SIZE, "helv", font_regular, TOC_SECTION_INDENT

        full_w = right_edge - MARGIN_LEFT - ind
        last_w = full_w - 60          # room for dot leader + page number
        wrapped = wrap_toc_entry(h.text, fobj, fs, full_w, last_w)

        if len(wrapped) > TOC_MAX_LINES:
            continue                  # entry too long — remove it
        prepared.append((h, wrapped, fs, fn, fobj, ind))

    if not prepared:
        return toc_doc, [], 0

    # ── Step 2: simulate layout to determine num_toc_pages ───────────────
    title_block = TOC_TITLE_SIZE + 6 + 12
    max_y = page_height - MARGIN_BOTTOM
    y = MARGIN_TOP + title_block      # first page includes title
    num_toc_pages = 1

    for i, (h, wrapped, fs, _fn, _fobj, _ind) in enumerate(prepared):
        extra = 4 if h.level == 1 and i > 0 else 0
        entry_h = extra + len(wrapped) * fs * TOC_LINE_SPACING_FACTOR
        if y + entry_h > max_y:
            num_toc_pages += 1
            y = MARGIN_TOP
        y += entry_h

    # ── Step 3: render ───────────────────────────────────────────────────
    entry_idx = 0
    for toc_pg in range(num_toc_pages):
        page = toc_doc.new_page(width=page_width, height=page_height)
        y = MARGIN_TOP

        if toc_pg == 0:
            y += TOC_TITLE_SIZE
            page.insert_text(
                fitz.Point(MARGIN_LEFT, y), "Table of Contents",
                fontsize=TOC_TITLE_SIZE, fontname="hebo", color=(0, 0, 0),
            )
            y += 6
            page.draw_line(
                fitz.Point(MARGIN_LEFT, y), fitz.Point(right_edge, y),
                color=SEPARATOR_COLOR, width=0.5,
            )
            y += 12

        while entry_idx < len(prepared):
            h, wrapped, fs, fn, fobj, ind = prepared[entry_idx]
            extra = 4 if h.level == 1 and entry_idx > 0 else 0
            entry_h = extra + len(wrapped) * fs * TOC_LINE_SPACING_FACTOR

            if y + entry_h > max_y:
                break

            y += extra
            page_offset = num_toc_pages if count_toc_pages else 0
            display_page = h.page_num + page_offset + 1
            pn_str = str(display_page)

            for li, line_text in enumerate(wrapped):
                is_last = (li == len(wrapped) - 1)

                y += fs
                page.insert_text(
                    fitz.Point(MARGIN_LEFT + ind, y), line_text,
                    fontsize=fs, fontname=fn, color=(0, 0, 0),
                )

                if is_last:
                    pn_w = font_regular.text_length(pn_str, fontsize=fs)
                    page.insert_text(
                        fitz.Point(right_edge - pn_w, y), pn_str,
                        fontsize=fs, fontname="helv", color=(0, 0, 0),
                    )
                    tw = fobj.text_length(line_text, fontsize=fs)
                    dot_start = MARGIN_LEFT + ind + tw + 6
                    dot_end   = right_edge - pn_w - 6
                    if dot_end > dot_start + 10:
                        dot_unit = font_regular.text_length(". ", fontsize=fs - 1)
                        if dot_unit > 0:
                            dots = ". " * int((dot_end - dot_start) / dot_unit)
                            page.insert_text(
                                fitz.Point(dot_start, y), dots,
                                fontsize=fs - 1, fontname="helv",
                                color=DOT_LEADER_COLOR,
                            )

                y += fs * (TOC_LINE_SPACING_FACTOR - 1)

            entry_idx += 1

    filtered = [e[0] for e in prepared]
    return toc_doc, filtered, num_toc_pages


# ─── Main orchestrator for one PDF ──────────────────────────────────────────

def process_pdf(src_path, output_path, h1_only=False, count_toc_pages=True):
    """Extract headings → build TOC → insert at start → save.
    Returns 'success', 'scanned', or 'skipped'.
    """
    doc = fitz.open(src_path)
    fname = os.path.basename(src_path)
    print(f"\n{'─'*60}")
    print(f"  {fname}  ({len(doc)} pages)")

    # Check for scanned/image-only PDF
    if is_scanned_pdf(doc):
        print("  ⚠  Scanned/image-only PDF (no extractable text) — skipping")
        doc.close()
        return "scanned"

    # ── 1. Extract headings ──────────────────────────────────────────────
    headings = extract_headings_by_font(doc)
    if not headings:
        print("  ⚠  Font metrics unusable — falling back to PDF outline/bookmarks")
        headings = extract_headings_from_outline(doc)
        if headings:
            print(f"  ✓  Found {len(headings)} entries from existing PDF outline")
    if not headings:
        print("  ⚠  No outline — falling back to keyword detection")
        headings = extract_headings_by_keyword(doc)
    if not headings:
        print("  ✗  No headings found.  Skipping.")
        doc.close()
        return "skipped"

    # ── 2. Post-process ──────────────────────────────────────────────────
    headings.sort(key=lambda h: (h.page_num, h.y_position))
    headings = merge_multiline_headings(headings)
    headings = merge_adjacent_headings_across_blocks(headings)
    headings = supplement_with_keyword_chapters(doc, headings)
    headings = deduplicate_headings(headings)
    headings = select_toc_headings(headings, h1_only=h1_only)
    headings.sort(key=lambda h: (h.page_num, h.y_position))
    headings = normalize_hierarchy(headings)

    if not headings:
        print("  ✗  No usable headings after filtering.  Skipping.")
        doc.close()
        return "skipped"

    # Log summary
    h1_count = sum(1 for h in headings if h.level == 1)
    h2_count = sum(1 for h in headings if h.level == 2)
    print(f"  → {len(headings)} TOC entries  (H1={h1_count}, H2={h2_count})")
    for h in headings[:8]:
        tag = f"H{h.level}"
        print(f"     [{tag}] p{h.page_num+1:>4}: {h.text[:72]}")
    if len(headings) > 8:
        print(f"     … and {len(headings) - 8} more")

    # ── 3. Create TOC pages (handles wrapping and page count) ────────────
    pw, ph = doc[0].rect.width, doc[0].rect.height
    toc_doc, headings, num_toc_pages = create_toc_pages(
        headings, pw, ph, count_toc_pages=count_toc_pages,
    )

    if not headings:
        print("  ✗  No headings fit within TOC line limit.  Skipping.")
        doc.close()
        toc_doc.close()
        return "skipped"

    # ── 4. Insert TOC at the front of the document ───────────────────────
    doc.insert_pdf(toc_doc, start_at=0)

    # ── 5. Set the bookmark / outline tree (sidebar navigation) ──────────
    toc_entries = []
    for h in headings:
        # Bookmarks always use real page position (including TOC pages)
        bk_page = h.page_num + num_toc_pages + 1   # 1-based for set_toc()
        toc_entries.append([h.level, h.text, bk_page])
    if toc_entries:
        doc.set_toc(toc_entries)

    # ── 6. Save ──────────────────────────────────────────────────────────
    doc.save(output_path, deflate=True, garbage=4)
    doc.close()
    toc_doc.close()
    print(f"  ✓  Saved → {os.path.basename(output_path)}  "
          f"(+{num_toc_pages} TOC page{'s' if num_toc_pages > 1 else ''})")
    return "success"


# ─── Entry point ─────────────────────────────────────────────────────────────

def main():
    input_dir  = Path(".")
    output_dir = input_dir / "with_toc"
    failed_dir = input_dir / "failed_files"
    output_dir.mkdir(exist_ok=True)

    pdfs = sorted(input_dir.glob("*.pdf"))
    print(f"Found {len(pdfs)} PDF files in {input_dir}\n")

    success, skipped, scanned, failed = 0, 0, 0, 0
    for p in pdfs:
        out = output_dir / p.name
        try:
            result = process_pdf(str(p), str(out))
            if result == "success":
                success += 1
            elif result == "scanned":
                scanned += 1
                failed_dir.mkdir(exist_ok=True)
                shutil.copy2(str(p), str(failed_dir / p.name))
                print(f"  → Copied to failed_files/{p.name}")
            elif result == "skipped":
                skipped += 1
            else:
                skipped += 1
        except Exception as exc:
            failed += 1
            print(f"  ✗  ERROR: {exc}")
            import traceback
            traceback.print_exc()

    print(f"\n{'═'*60}")
    summary_parts = [f"{success} succeeded"]
    if skipped:
        summary_parts.append(f"{skipped} skipped")
    if scanned:
        summary_parts.append(f"{scanned} scanned → failed_files/")
    if failed:
        summary_parts.append(f"{failed} failed")
    print(f"Done.  {' · '.join(summary_parts)}")
    print(f"Output directory: {output_dir}")


if __name__ == "__main__":
    main()
