"""Temporary analysis: compare original vs with_toc TOC quality."""
import fitz
import os
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

FOLDER = r"C:\Users\admin\Downloads\Batch 1\Wrong TOC"
TOC_FOLDER = os.path.join(FOLDER, "with_toc")

# For each with_toc file, show the TOC entries and flag issues
for name in sorted(os.listdir(TOC_FOLDER)):
    if not name.endswith('.pdf'):
        continue
    path = os.path.join(TOC_FOLDER, name)
    doc = fitz.open(path)
    toc = doc.get_toc()

    issues = []
    for lvl, title, pg in toc:
        t = title.strip()
        if len(t) > 80:
            issues.append(f"LONG: {t[:60]}...")
        if len(t) < 4:
            issues.append(f"SHORT: '{t}'")
        # Check for body text leaking (sentence-like with lowercase)
        words = t.split()
        if len(words) > 12:
            issues.append(f"BODY?: {t[:60]}...")

    if issues:
        print(f"\n{name[:70]}")
        print(f"  TOC entries: {len(toc)}")
        for iss in issues[:5]:
            print(f"  * {iss}")
        if len(issues) > 5:
            print(f"  ... +{len(issues)-5} more issues")
    doc.close()
