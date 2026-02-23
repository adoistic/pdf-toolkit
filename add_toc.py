#!/usr/bin/env python3
"""
add_toc.py — Automatically extract chapter headings from structured PDFs
and insert a Table of Contents page at the beginning of each file.

Uses PyMuPDF (fitz) for all PDF operations.
Output is saved to a 'with_toc/' subfolder.
"""

import fitz
import os
import re
from dataclasses import dataclass
from collections import Counter, defaultdict
from pathlib import Path

# ─── Layout constants for generated TOC pages ───────────────────────────────
MARGIN_TOP = 72
MARGIN_BOTTOM = 72
MARGIN_LEFT = 72
MARGIN_RIGHT = 72
TOC_TITLE_SIZE = 18
TOC_CHAPTER_SIZE = 11
TOC_SECTION_SIZE = 10
TOC_LINE_SPACING_FACTOR = 1.6
TOC_SECTION_INDENT = 20
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


def select_toc_headings(headings):
    """
    Apply the user's rule:
      • If H1 (largest heading size) has > 3 entries → keep H1 (+ H2 if present)
      • If H1 has ≤ 3 entries → it's probably a book/part title;
        promote H2 to be the primary TOC level.
    Only two levels are ever included in the final TOC.

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


# ─── TOC page creation ──────────────────────────────────────────────────────

def create_toc_pages(headings, page_width, page_height, num_toc_pages):
    """
    Build a small fitz.Document containing the TOC page(s).
    Page numbers displayed are  original_page + num_toc_pages + 1  (1-based).
    """
    toc_doc = fitz.open()

    font_regular = fitz.Font("helv")
    font_bold    = fitz.Font("hebo")
    right_edge   = page_width - MARGIN_RIGHT

    heading_idx = 0
    for toc_pg in range(num_toc_pages):
        page = toc_doc.new_page(width=page_width, height=page_height)
        y = MARGIN_TOP

        if toc_pg == 0:
            y += TOC_TITLE_SIZE
            page.insert_text(
                fitz.Point(MARGIN_LEFT, y), "Table of Contents",
                fontsize=TOC_TITLE_SIZE, fontname="hebo", color=(0, 0, 0),
            )
            y += 8
            page.draw_line(
                fitz.Point(MARGIN_LEFT, y), fitz.Point(right_edge, y),
                color=SEPARATOR_COLOR, width=0.5,
            )
            y += 16

        max_y = page_height - MARGIN_BOTTOM

        while heading_idx < len(headings) and y < max_y - 15:
            h = headings[heading_idx]

            if h.level == 1:
                fontsize = TOC_CHAPTER_SIZE
                fontname = "hebo"
                fmeasure = font_bold
                indent = 0
                if heading_idx > 0:
                    y += 4
            else:
                fontsize = TOC_SECTION_SIZE
                fontname = "helv"
                fmeasure = font_regular
                indent = TOC_SECTION_INDENT

            display_page = h.page_num + num_toc_pages + 1
            pn_str = str(display_page)

            entry = h.text
            max_tw = right_edge - MARGIN_LEFT - indent - 60
            tw = fmeasure.text_length(entry, fontsize=fontsize)
            if tw > max_tw:
                while tw > max_tw and len(entry) > 10:
                    entry = entry[:-1]
                    tw = fmeasure.text_length(entry + "…", fontsize=fontsize)
                entry += "…"
                tw = fmeasure.text_length(entry, fontsize=fontsize)

            y += fontsize
            page.insert_text(
                fitz.Point(MARGIN_LEFT + indent, y), entry,
                fontsize=fontsize, fontname=fontname, color=(0, 0, 0),
            )
            pn_w = font_regular.text_length(pn_str, fontsize=fontsize)
            page.insert_text(
                fitz.Point(right_edge - pn_w, y), pn_str,
                fontsize=fontsize, fontname="helv", color=(0, 0, 0),
            )
            dot_start = MARGIN_LEFT + indent + tw + 6
            dot_end   = right_edge - pn_w - 6
            if dot_end > dot_start + 10:
                dot_unit = font_regular.text_length(". ", fontsize=fontsize - 1)
                if dot_unit > 0:
                    dots = ". " * int((dot_end - dot_start) / dot_unit)
                    page.insert_text(
                        fitz.Point(dot_start, y), dots,
                        fontsize=fontsize - 1, fontname="helv",
                        color=DOT_LEADER_COLOR,
                    )

            y += fontsize * (TOC_LINE_SPACING_FACTOR - 1)
            heading_idx += 1

    return toc_doc


# ─── Main orchestrator for one PDF ──────────────────────────────────────────

def process_pdf(src_path, output_path):
    """Extract headings → build TOC → insert at start → save."""
    doc = fitz.open(src_path)
    fname = os.path.basename(src_path)
    print(f"\n{'─'*60}")
    print(f"  {fname}  ({len(doc)} pages)")

    # ── 1. Extract headings ──────────────────────────────────────────────
    headings = extract_headings_by_font(doc)
    if not headings:
        print("  ⚠  Font metrics unusable — falling back to keyword detection")
        headings = extract_headings_by_keyword(doc)
    if not headings:
        print("  ✗  No headings found.  Skipping.")
        doc.close()
        return

    # ── 2. Post-process ──────────────────────────────────────────────────
    headings.sort(key=lambda h: (h.page_num, h.y_position))
    headings = merge_multiline_headings(headings)
    headings = merge_adjacent_headings_across_blocks(headings)
    headings = supplement_with_keyword_chapters(doc, headings)
    headings = deduplicate_headings(headings)
    headings = select_toc_headings(headings)
    headings.sort(key=lambda h: (h.page_num, h.y_position))
    headings = normalize_hierarchy(headings)

    if not headings:
        print("  ✗  No usable headings after filtering.  Skipping.")
        doc.close()
        return

    # Log summary
    h1_count = sum(1 for h in headings if h.level == 1)
    h2_count = sum(1 for h in headings if h.level == 2)
    print(f"  → {len(headings)} TOC entries  (H1={h1_count}, H2={h2_count})")
    for h in headings[:8]:
        tag = f"H{h.level}"
        print(f"     [{tag}] p{h.page_num+1:>4}: {h.text[:72]}")
    if len(headings) > 8:
        print(f"     … and {len(headings) - 8} more")

    # ── 3. Figure out how many TOC pages we need ─────────────────────────
    pw, ph = doc[0].rect.width, doc[0].rect.height
    usable = ph - MARGIN_TOP - MARGIN_BOTTOM - 50
    per_page = max(1, int(usable / (TOC_CHAPTER_SIZE * TOC_LINE_SPACING_FACTOR)))
    num_toc_pages = max(1, -(-len(headings) // per_page))  # ceil div

    # ── 4. Create TOC pages ──────────────────────────────────────────────
    toc_doc = create_toc_pages(headings, pw, ph, num_toc_pages)

    # ── 5. Insert TOC at the front of the document ───────────────────────
    doc.insert_pdf(toc_doc, start_at=0)

    # ── 6. Set the bookmark / outline tree (sidebar navigation) ──────────
    toc_entries = []
    for h in headings:
        bk_page = h.page_num + num_toc_pages + 1   # 1-based for set_toc()
        toc_entries.append([h.level, h.text, bk_page])
    if toc_entries:
        doc.set_toc(toc_entries)

    # ── 7. Save ──────────────────────────────────────────────────────────
    doc.save(output_path, deflate=True, garbage=4)
    doc.close()
    toc_doc.close()
    print(f"  ✓  Saved → {os.path.basename(output_path)}  "
          f"(+{num_toc_pages} TOC page{'s' if num_toc_pages > 1 else ''})")


# ─── Entry point ─────────────────────────────────────────────────────────────

def main():
    input_dir  = Path("/Users/siraj/Downloads/For_TOC_Make")
    output_dir = input_dir / "with_toc"
    output_dir.mkdir(exist_ok=True)

    pdfs = sorted(input_dir.glob("*.pdf"))
    print(f"Found {len(pdfs)} PDF files in {input_dir}\n")

    success, skipped, failed = 0, 0, 0
    for p in pdfs:
        out = output_dir / p.name
        try:
            process_pdf(str(p), str(out))
            if out.exists():
                success += 1
            else:
                skipped += 1
        except Exception as exc:
            failed += 1
            print(f"  ✗  ERROR: {exc}")
            import traceback
            traceback.print_exc()

    print(f"\n{'═'*60}")
    print(f"Done.  {success} succeeded · {skipped} skipped · {failed} failed")
    print(f"Output directory: {output_dir}")


if __name__ == "__main__":
    main()
