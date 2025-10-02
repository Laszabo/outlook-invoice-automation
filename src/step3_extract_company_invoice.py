#!/usr/bin/env python3
"""
step3_extract_fields.py

What this does
--------------
- Looks for a greeting like:  "Tisztelt <Company>!"
- Looks for an invoice sentence like:
  "küldjük a <digits> számú elektronikus számláját"

Why regex here
--------------
- We know the wording is nearly identical across emails.
- Regex lets us capture just the changing parts:
  - <Company>      → named group "company"
  - <digits>       → first capturing group (1)

Key regex pieces (human terms)
------------------------------
- \s+        : one or more whitespace characters (space, tab, newline)
- .+?        : minimal/lazy "anything" (stop at the first match that works)
- (?P<name>) : named capture group; access with m.group("name")
- re.IGNORECASE : case-insensitive match (handles ü/á/é correctly in Python 3)
- re.DOTALL     : dot (.) also matches newlines — important for multi-line text

Usage
-----
- Pass a cleaned body string (use your step2_body_tools.normalize_body_from_mailitem first).
- Call extract_company(body) and extract_invoice(body).
"""

from __future__ import annotations
import re

# Example matches:
# "Tisztelt Jégszilánk Kft!"  → company = "Jégszilánk Kft"
# Pattern logic:
# - "Tisztelt" then at least one space
# - (?P<company>.+?)  : lazily grab the shortest text that can be the company
# - \s*!              : optional spaces then an exclamation mark
RE_GREET = re.compile(
    r"Tisztelt\s+(?P<company>.+?)\s*!",
    re.IGNORECASE | re.DOTALL,
)

# Example matches:
# "küldjük a 562003117859 számú elektronikus számláját"
# Pattern logic:
# - "küldjük" + spaces + "a" + spaces
# - (\d{9,30}) : 9 to 30 digits as the invoice number (group 1)
# - spaces + "számú elektronikus számláját"

RE_INVOICE = re.compile(
    r"küldjük\s+a\s+(\d{9,30})\s+számú\s+elektronikus\s+számláját",
    re.IGNORECASE,
)


# --- 2) Public helpers -------------------------------------------------------

def extract_company(body: str) -> str | None:
    """
    Return the company name from a "Tisztelt <Company>!" greeting.

    - Uses a named group (?P<company>...) so we can read it as m.group("company")
    - Returns None if not found
    """
    if not body:
        return None
    m = RE_GREET.search(body)
    if not m:
        return None
    return m.group("company").strip()


def extract_invoice(body: str) -> str | None:
    """
    Return the invoice number (digits only) from the standard Hungarian sentence:
    "küldjük a <digits> számú elektronikus számláját"

    - Returns None if not found
    """
    if not body:
        return None
    m = RE_INVOICE.search(body)
    if not m:
        return None
    return m.group(1).strip()


# --- 3) Quick local tests (run this file directly to try) --------------------
if __name__ == "__main__":
    # Base sample (multi-line to show DOTALL usefulness)
    sample = """Tisztelt Jégszilánk Kft!
Ezúton küldjük a 562003117859 számú elektronikus számláját...
Köszönjük!
"""

    print("Company:", extract_company(sample))  # expected: "Jégszilánk Kft"
    print("Invoice:", extract_invoice(sample))  # expected: "562003117859"

    # More samples to sanity check behavior
    samples = [
        # extra spaces before !
        "Tisztelt Valami Zrt   !\nKüldjük a 123456789 számú elektronikus számláját",
        # lowercase "tisztelt", mixed casing and newlines
        "tisztelt  Példa Bt!\n\nküldjük a 987654321 számú elektronikus számláját",
        # HTML-ish leftovers already stripped to text (line breaks remain)
        "Tisztelt  Mintacég Kft!\nKüldjük a 562003117859 számú elektronikus számláját",
        # company with punctuation
        "Tisztelt ACME-Építő Kft.! küldjük a 123456789012 számú elektronikus számláját",
    ]

    for i, s in enumerate(samples, 1):
        print(f"\n--- Sample {i} ---")
        print("Company:", extract_company(s))
        print("Invoice:", extract_invoice(s))
