# 📧 Outlook Invoice Automation

Automated pipeline that reads Outlook emails, extracts company and invoice data, parses attached PDFs for POD/PRM, and routes files to the correct folders.

![Python](https://img.shields.io/badge/python-3.10+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![Status](https://img.shields.io/badge/status-active-success.svg)

## Features
- Outlook COM integration (shared mailbox)
- Company + invoice extraction from email body
- POD/PRM detection from PDF (PyMuPDF)
- Smart routing by POD prefix:
  - `HU*` → Electricity
  - `39*` → Gas
- Duplicate-safe filenames
- Exception list for manual review
- Mail state updates (mark completed)

## Architecture
```
Email Inbox → Filter (sender/month)
  → Normalize Body → Extract (Company, Invoice)
  → Parse PDF (POD/PRM)
  → Route & Rename
  → Mark Complete
```

## Quick Start
**Prereqs**
- Windows with Microsoft Outlook installed
- Python 3.10+

**Install**
```bash
git clone https://github.com/laszabo/outlook-invoice-automation.git
cd outlook-invoice-automation
pip install -r requirements.txt
```

**Configure**
Copy the example config and edit values (no secrets in Git):
```bash
copy examples\config.example.toml config.toml
# Edit config.toml (paths, mailbox, sender, year/month)
```

**Run**
```bash
python src/main.py
```

## Configuration (config.toml)
- `MAILBOX_NAME`, `SENDER_EMAIL`, `YEAR`, `MONTH`
- `OUT_ELECTRICITY`, `OUT_GAS`
- `EXCEPT_KEYWORDS` for companies to skip/flag

## Output
**Filename**: `{Company}_{POD}_{Invoice}.pdf`  
**Example**: `Halker_Kft_HU001234567890_562003117859.pdf`

## Testing
```bash
pytest -q
```

## Roadmap
- CSV audit log of processed invoices
- Multiple senders/mailboxes support
- Unit tests for regex and POD extractors
- Optional GUI
