"""
Configuration settings for invoice automation pipeline.
Copy from examples/sample_config.py and customize.
"""

# Email filtering
SENDER_EMAIL = "invoices@vendor.com"
YEAR = 2025
MONTH = 10
MAILBOX_NAME = "Shared Mailbox"

# Output directories
OUT_ELECTRICITY = r"C:\Invoices\Electricity\Incoming"
OUT_GAS = r"C:\Invoices\Gas\Incoming"

# Exception companies (require manual review)
EXCEPT_KEYWORDS = [
    "Company_A",
    "Company_B",
    "Company_C"
]
