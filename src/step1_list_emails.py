#!/usr/bin/env python3
"""
step1_list_emails.py

What this does
--------------
- Opens a shared Outlook mailbox (e.g., "YourSharedInbox")
- Looks at the Inbox
- Filters emails by:
  1) Sender email address (case-insensitive)
  2) Year + Month of the ReceivedTime
- Prints a short summary: date/time and subject

Requirements
------------
- Windows + Outlook Desktop installed and configured
- You must have access to the shared mailbox
- Python 3.10+ recommended
- Install pywin32:  pip install pywin32

How to use
----------
1) Change the 4 settings below (MAILBOX_NAME, SENDER_EMAIL, YEAR, MONTH).
2) Run:  python step1_list_emails.py
"""

# === 1) SETTINGS: CHANGE THESE TO YOUR NEEDS =================================
MAILBOX_NAME = "YourSharedInbox"           # exactly as it appears in Outlook
SENDER_EMAIL = "eszamla@e2hungary.hu"   # filter by this sender (substring match)

from datetime import datetime
YEAR  = datetime.now().year             # e.g., 2025
MONTH = datetime.now().month            # e.g., 10 (October)
# If you want a fixed month, set them like:
# YEAR  = 2025
# MONTH = 9

# === 2) IMPORTS ==============================================================

try:
    import win32com.client as win32  # provided by pywin32
except Exception as e:
    raise SystemExit(
        "pywin32 is required. Install it with:\n\n"
        "    pip install pywin32\n\n"
        f"Import error details: {e}"
    )

# Outlook constants we need
OL_FOLDER_INBOX   = 6    # Inbox
OL_CLASS_MAILITEM = 43   # MailItem class ID


# === 3) OUTLOOK HELPERS ======================================================

def get_inbox_items(mailbox_name: str):
    """
    Opens the shared mailbox Inbox and returns its Items, sorted newest-first.

    If the mailbox name is wrong or you don't have permission, this will fail.
    """
    # Connect to Outlook (MAPI)
    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Tell Outlook which shared mailbox we want
    recipient = ns.CreateRecipient(mailbox_name)
    if not recipient.Resolve():  # Outlook must recognize the name
        raise SystemExit(
            f'Could not find mailbox "{mailbox_name}". '
            "Check the exact display name and your permissions."
        )

    # 6 = Inbox
    inbox = ns.GetSharedDefaultFolder(recipient, OL_FOLDER_INBOX)

    # Get all items and sort by ReceivedTime (descending=True => newest first)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    return items


def is_mail_item(item) -> bool:
    """Return True if this Outlook item is a normal MailItem."""
    return getattr(item, "Class", None) == OL_CLASS_MAILITEM


# === 4) MAIN LOGIC ===========================================================

def main():
    print(f'Listing emails from sender "{SENDER_EMAIL}" in {YEAR}-{MONTH:02d}')
    print(f'Mailbox: "{MAILBOX_NAME}"')
    print("-" * 60)

    try:
        items = get_inbox_items(MAILBOX_NAME)
    except Exception as e:
        raise SystemExit(f"Failed to open mailbox. Details:\n{e}")

    found_any = False

    # Walk through each Outlook item in the Inbox
    for it in items:
        # Skip non-email items (e.g., meeting requests)
        if not is_mail_item(it):
            continue

        # Sender check (case-insensitive substring; handles EX-style addresses too)
        sender = str(getattr(it, "SenderEmailAddress", "") or "")
        if SENDER_EMAIL.lower() not in sender.lower():
            continue

        # Date check
        when = getattr(it, "ReceivedTime", None)
        if not when:
            continue  # rare system items may lack this

        if when.year != YEAR or when.month != MONTH:
            continue

        # Print a short summary
        print("When   :", when)
        print("Subject:", getattr(it, "Subject", ""))
        print("-" * 60)
        found_any = True

    if not found_any:
        print("No emails matched your filters.")


if __name__ == "__main__":
    main()
