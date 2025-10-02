# 📧 Outlook Invoice Automation

An automated invoice processing pipeline that extracts, validates, and routes PDF invoices from Outlook emails to designated folders based on utility type.

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![Status](https://img.shields.io/badge/status-active-success.svg)

## 🎯 Features

- **Automated Email Processing**: Connects to Outlook via COM interface
- **Intelligent Extraction**: Parses company names, invoice numbers, and POD identifiers
- **Smart Routing**: Automatically routes invoices based on POD prefix
  - `HU*` → Electricity invoices
  - `39*` → Gas invoices
- **Exception Handling**: Flags specific companies for manual review
- **Duplicate Detection**: Prevents filename conflicts with automatic numbering
- **State Management**: Marks processed emails as read/complete

## 🏗️ Architecture

```
Email Inbox → Filter → Extract Metadata → Parse PDF → Route to Folder → Mark Complete
```

The system follows a modular pipeline architecture with clear separation of concerns:

1. **Email Retrieval** - Connects to Outlook and filters by sender/date
2. **Body Normalization** - Cleans and standardizes email content
3. **Metadata Extraction** - Parses company name and invoice number
4. **POD Extraction** - Reads Point of Delivery identifier from PDF
5. **File Routing** - Saves files to appropriate folders with standardized naming

## 🚀 Quick Start

### Prerequisites

- Windows OS (for Outlook COM automation)
- Python 3.8+
- Microsoft Outlook installed and configured

### Installation

```bash
# Clone the repository
git clone https://github.com/laszabo/outlook-invoice-automation.git
cd outlook-invoice-automation

# Install dependencies
pip install -r requirements.txt

# Configure settings
cp examples/sample_config.py src/config.py
# Edit src/config.py with your paths and settings
```

### Usage

```bash
python src/main.py
```

## 📋 Configuration

Edit `src/config.py` to customize:

```python
# Email filtering
SENDER_EMAIL = "invoices@vendor.com"
YEAR = 2025
MONTH = 10
MAILBOX_NAME = "Shared Mailbox"

# Output directories
OUT_ELECTRICITY = r"C:\Invoices\Electricity\Incoming"
OUT_GAS = r"C:\Invoices\Gas\Incoming"

# Exception companies (require manual review)
EXCEPT_KEYWORDS = ["Company_A", "Company_B", "Company_C"]
```

## 🔧 Pipeline Stages

### 1. Email Retrieval (`step1_list_emails.py`)
Connects to Outlook mailbox and filters emails by sender and date range.

### 2. Body Normalization (`step2_body_tools.py`)
Cleans HTML tags, special characters, and encoding issues from email body.

### 3. Metadata Extraction (`step3_extract_company_invoice.py`)
Uses regex patterns to extract company name and invoice number from normalized text.

### 4. POD Extraction (`step4_pdf_pod.py`)
Parses PDF content to locate Point of Delivery identifier.

### 5. File Routing (`main.py`)
Coordinates all stages, routes files based on POD prefix, and manages email state.

## 📊 Output Format

**Filename Pattern**: `{Company}_{POD}_{Invoice}.pdf`

**Example**: `ACME_Corp_HU001234567890_INV-2025-001.pdf`

**Routing Logic**:

| POD Prefix | Utility Type | Destination Folder |
|------------|--------------|-------------------|
| HU*        | Electricity  | OUT_ELECTRICITY   |
| 39*        | Gas          | OUT_GAS           |

## 🛡️ Exception Handling

Certain companies require manual review and are automatically flagged:
- Marked as unread in Outlook
- Not processed automatically
- Configurable in `EXCEPT_KEYWORDS`

**Error Handling**:
- Missing metadata → Skip and leave email unchanged
- Exception companies → Mark unread for manual review
- PDF parsing errors → Skip attachment
- Duplicate filenames → Add numeric suffix

## 📝 Requirements

- `pywin32>=305` - Outlook COM automation
- `PyPDF2>=3.0.0` or `pdfplumber>=0.10.0` - PDF parsing
- `python-dateutil>=2.8.2` - Date handling

See `requirements.txt` for complete list.

## 🧪 Testing

```bash
# Run tests (if implemented)
python -m pytest tests/
```

## 📁 Project Structure

```
outlook-invoice-automation/
├── src/
│   ├── main.py                      # Main orchestration
│   ├── step1_list_emails.py         # Email retrieval
│   ├── step2_body_tools.py          # Body normalization
│   ├── step3_extract_company_invoice.py  # Metadata extraction
│   ├── step4_pdf_pod.py             # PDF POD extraction
│   └── config.py                    # Configuration
├── requirements.txt                 # Dependencies
├── README.md                        # This file
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## ⚠️ Disclaimer

This tool is designed for legitimate business automation purposes. Ensure you have proper authorization to access and process emails and attachments in your organization.

## 🔮 Future Enhancements

- [ ] Database logging of processed invoices
- [ ] Email notifications for exceptions
- [ ] Web dashboard for monitoring
- [ ] Support for additional utility types
- [ ] Multi-threaded processing for large volumes
- [ ] Machine learning for improved metadata extraction

## 📧 Contact

**Szabo Laszlo** - [@laszabo](https://www.linkedin.com/in/laszabo)

**Project Link**: [https://github.com/laszabo/outlook-invoice-automation](https://github.com/laszabo/outlook-invoice-automation)

---

⭐ If you find this project helpful, please consider giving it a star!
