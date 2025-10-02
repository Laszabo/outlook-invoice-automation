#!/usr/bin/env python3
"""
step2_body_tools.py

What this does
--------------
- Provides a function: normalize_body_from_mailitem(msg)
- It returns clean, readable text from an Outlook MailItem:
  1) Prefer plain-text body (msg.Body)
  2) If it's empty, fall back to HTML body (msg.HTMLBody) and strip tags

Why this matters
----------------
Parsing invoice numbers or company names is easier on clean text.
This keeps your extraction code simple and reliable.

Requirements
------------
- Python 3.10+
- No external packages (uses only the standard library)
"""

import re
from html import unescape


def _strip_html(html: str) -> str:
    """
    Take raw HTML and return readable plain text.

    Steps:
    - Decode HTML entities (&nbsp;, &amp;, etc.)
    - Remove <script> and <style> blocks
    - Convert <br> and </p> to newlines
    - Remove all remaining tags
    - Normalize whitespace
    """
    if not html:
        return ""

    # Decode HTML entities (e.g., &nbsp; -> space)
    txt = unescape(html)

    # Remove <script> and <style> content completely
    # (?is) = case-insensitive + dot matches newlines
    txt = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", txt)

    # Replace common HTML line boundaries with newlines
    txt = re.sub(r"(?is)<br\s*/?>", "\n", txt)
    txt = re.sub(r"(?is)</p\s*>", "\n", txt)

    # Remove all remaining tags
    txt = re.sub(r"(?is)<[^>]+>", " ", txt)

    # Collapse HTML non-breaking spaces that may remain
    txt = txt.replace("\xa0", " ")

    # Normalize spaces and compress excessive blank lines
    txt = re.sub(r"[ \t]+", " ", txt)                # collapse runs of spaces/tabs
    txt = re.sub(r"\n\s*\n\s*\n+", "\n\n", txt)      # max 1 empty line between paragraphs

    return txt.strip()


def normalize_body_from_mailitem(msg) -> str:
    """
    Return a clean text body from an Outlook MailItem.

    Logic:
    1) Try msg.Body (usually already plain text).
    2) If empty, try msg.HTMLBody and strip it with _strip_html().
    3) Normalize line endings to '\n'.

    Parameters
    ----------
    msg : Outlook MailItem (COM object)

    Returns
    -------
    str : clean text body
    """
    # 1) Plain text body if available
    body = (getattr(msg, "Body", "") or "").strip()
    if not body:
        # 2) Fallback to HTML body if plain text is empty
        html = getattr(msg, "HTMLBody", "") or ""
        body = _strip_html(html)

    # 3) Normalize line endings across Windows/Outlook variants
    body = body.replace("\r\n", "\n").replace("\r", "\n").strip()

    return body


# --- Optional: quick demo stub (safe to remove) -------------------------------
if __name__ == "__main__":
    class _FakeMsg:
        Body = ""
        HTMLBody = """
            <html>
                <head><style>p{}</style><script>var x=1;</script></head>
                <body>
                    <p>Tisztelt XY Kft!</p>
                    <p>Ezúton küldjük a <b>562111117859</b> számú elektronikus számláját.<br>
                    Üdvözlettel,<br>Számlaküldő</p>
                </body>
            </html>
        """

    cleaned = normalize_body_from_mailitem(_FakeMsg())
    print("--- Cleaned Body ---")
    print(cleaned)
