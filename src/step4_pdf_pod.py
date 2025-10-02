#!/usr/bin/env python3
"""
step4_extract_pod.py — Find POD IDs (HU... or 39...) in invoice PDFs

What this does
--------------
- Opens a PDF (PyMuPDF / fitz)
- Tries to read Page 2 first (most invoices show POD there)
- Falls back to scanning every page
- Extracts the first POD-like ID it finds:
  - Electricity: starts with "HU"
  - Gas: starts with "39"

Why this works
--------------
- Many suppliers print a label like "Mérőpont azonosító:"
- After that label comes the POD value (letters, digits, hyphens)
- We also include a fallback that searches the whole page if the label is missing

Requirements
------------
- Python 3.10+
- Install PyMuPDF:  pip install pymupdf
  (import name is 'fitz')

How to use
----------
from step5_extract_pod import extract_pod_from_pdf
pod = extract_pod_from_pdf(r"C:\path\to\invoice.pdf")
print(pod)  # e.g., "HU-1234-5678-AB" or "3912345678ABC"
"""

from __future__ import annotations
import re
import fitz  # PyMuPDF


# ==========================
# 1) Regex building blocks
# ==========================

# The label in Hungarian often appears like:
# "Mérőpont azonosító:" / "Mérő azonosító:" / "Mero pont azonosito" (with/without accents)
# We allow common variations using character classes like [ée] and [íi].
LABEL = r"(?i)m[ée]r[őo]?(?:si)?\s*pont\s+azonos[ií]t[óo](?:ja)?\s*:?\s*"

# Allowed characters inside a POD: letters A–Z, digits 0–9, hyphens.
# Spaces are typically not used.
IDCH = r"[A-Z0-9\-]"

# Explanation of the lookaheads we use:
#  - (?= ... \d)    → ensure there is at least one digit somewhere after the start
#  - (?= ... -)     → ensure there's at least one hyphen
#  - [A-Z0-9] end   → ensure it doesn't end with a hyphen
#
# We keep 6–40 as a practical inside-length range to avoid tiny/huge garbage matches.

PATTERNS = [
    # 1) After the label, match a HU... value
    re.compile(rf"{LABEL}(HU(?={IDCH}*\d)(?={IDCH}*-){IDCH}{{6,40}}[A-Z0-9])", re.IGNORECASE),

    # 2) After the label, match a 39... value
    re.compile(rf"{LABEL}(39(?={IDCH}*\d){IDCH}{{6,40}}[A-Z0-9])", re.IGNORECASE),

    # 3) Fallback: HU... anywhere on the page
    re.compile(rf"\bHU(?={IDCH}*\d)(?={IDCH}*-){IDCH}{{6,40}}[A-Z0-9]\b", re.IGNORECASE),

    # 4) Fallback: 39... anywhere on the page
    re.compile(rf"\b39(?={IDCH}*\d){IDCH}{{6,40}}[A-Z0-9]\b", re.IGNORECASE),
]

# ==========================
# 2) Small utilities
# ==========================

def _normalize_pod(s: str) -> str:
    """
    Make the POD tidy and consistent:
    - Uppercase
    - Remove spaces
    - Collapse multiple hyphens to single
    - Trim hyphens from start/end
    """
    s = (s or "").upper()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"-{2,}", "-", s)
    s = s.strip("-")
    return s


def _find_pod_in_text(text: str) -> str | None:
    """
    Run patterns in priority order:
      1) After-label HU
      2) After-label 39
      3) Anywhere HU
      4) Anywhere 39
    Return the first normalized match, or None.
    """
    if not text:
        return None

    for pat in PATTERNS:
        matches = pat.findall(text)
        if not matches:
            continue
        pod = _normalize_pod(matches[0])
        if pod.startswith("HU") or pod.startswith("39"):
            return pod

    return None


# ==========================
# 3) Public function
# ==========================

def extract_pod_from_pdf(pdf_path: str) -> str | None:
    """
    Try to extract a POD from a PDF file.

    Strategy
    --------
    - Try Page 2 first (index 1). Many invoices place the POD there.
    - If not found, scan all pages in order.
    - Return the first valid POD (HU... or 39...), else None.
    """
    # Open with PyMuPDF
    with fitz.open(pdf_path) as doc:
        # 1) Prefer Page 2
        if doc.page_count >= 2:
            page = doc.load_page(1)
            txt = page.get_text()  # plain extracted text
            pod = _find_pod_in_text(txt)
            if pod:
                return pod

        # 2) Fallback: scan all pages
        for i in range(doc.page_count):
            page = doc.load_page(i)
            txt = page.get_text()
            pod = _find_pod_in_text(txt)
            if pod:
                return pod

    # Nothing matched
    return None


# ==========================
# 4) Quick local test
# ==========================

if __name__ == "__main__":
    # Replace with a real file path on your machine to test quickly.
    TEST_PDF = r"C:\YOURFILENAME.pdf"

    try:
        result = extract_pod_from_pdf(TEST_PDF)
        print("POD found:", result if result else "(none)")
    except Exception as e:
        print("Error while reading PDF:", e)
