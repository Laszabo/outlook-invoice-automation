"""
Invoice Email Processing and Routing System
===========================================

This module serves as the main orchestrator for an automated invoice processing pipeline
that extracts, validates, and routes PDF invoices from Outlook emails to designated folders
based on utility type (electricity/gas).

Architecture Overview:
---------------------
1. Email Retrieval: Connects to Outlook via COM interface to fetch emails from specific sender
2. Content Extraction: Normalizes email body and extracts company name and invoice number
3. PDF Processing: Extracts POD (Point of Delivery) identifiers from invoice PDFs
4. Smart Routing: Routes invoices to appropriate folders based on POD prefix
5. State Management: Marks emails as processed or flags exceptions for manual review

Business Logic:
--------------
- Electricity invoices (HU prefix): Routed to electricity folder
- Gas invoices (39 prefix): Routed to gas folder
- Exception companies: Marked as unread for manual processing
- Duplicate handling: Automatic numeric suffix generation

Author: Szabo Laszlo
Date: 2025
"""

import os
import re
import tempfile
import shutil
import traceback
from typing import Optional, List, Any
import win32com.client as win32

# Import pipeline modules
from step1_list_emails import get_inbox_items, SENDER_EMAIL, YEAR, MONTH, MAILBOX_NAME
from step2_body_tools import normalize_body_from_mailitem
from step3_extract_company_invoice import extract_company, extract_invoice
from step5_pdf_pod import extract_pod_from_pdf


# ============================================================================
# CONFIGURATION CONSTANTS
# ============================================================================

# Output directory paths for different utility types
OUT_ELECTRICITY = r"F:\[Company_Path]\Electricity\Invoices\Incoming"
OUT_GAS = r"F:\[Company_Path]\Gas\Commercial_Invoices\Incoming"

# Companies requiring manual review (bypass automation)
EXCEPT_KEYWORDS = ["[Company_A]", "[Company_B]", "[Company_C]", "[Company_D]", "[Company_E]"]

# Outlook message class constant
OUTLOOK_MAIL_ITEM_CLASS = 43


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def is_exception_company(name: Optional[str]) -> bool:
    """
    Check if a company name matches the exception list requiring manual processing.
    
    Args:
        name: Company name extracted from email body
        
    Returns:
        True if company is in exception list, False otherwise
        
    Example:
        >>> is_exception_company("Hanon Systems")
        True
        >>> is_exception_company("Regular Company")
        False
    """
    if not name:
        return False
    
    name_lower = name.lower()
    return any(keyword.lower() in name_lower for keyword in EXCEPT_KEYWORDS)


def safe_filename(filename: str) -> str:
    """
    Sanitize a string for use as a Windows filename.
    
    Removes illegal characters (\\/:*?"<>|) and normalizes whitespace.
    
    Args:
        filename: Raw filename string
        
    Returns:
        Sanitized filename safe for Windows filesystem
        
    Example:
        >>> safe_filename('Company/Name: "Invoice"')
        'Company Name  Invoice'
    """
    # Replace Windows-illegal characters with space
    sanitized = re.sub(r'[\\/:*?"<>|]+', " ", filename)
    
    # Collapse multiple spaces into one and trim
    sanitized = re.sub(r"\s+", " ", sanitized).strip()
    
    return sanitized


def ensure_dir(directory_path: str) -> None:
    """
    Ensure a directory exists, creating it if necessary.
    
    Args:
        directory_path: Path to directory to create/verify
        
    Note:
        Uses exist_ok=True to prevent race condition errors
    """
    if not os.path.isdir(directory_path):
        os.makedirs(directory_path, exist_ok=True)


def mark_processed(message: Any) -> None:
    """
    Mark an Outlook email as processed (read + flagged complete).
    
    Sets UnRead=False and FlagStatus=2 (completed) for visual confirmation
    in Outlook that the email has been successfully processed.
    
    Args:
        message: Outlook MailItem COM object
    """
    try:
        message.UnRead = False
        message.FlagStatus = 2  # 0=None, 1=Flagged, 2=Completed
        message.Save()
    except Exception:
        # Silently fail if Outlook state change fails (non-critical)
        pass


def mark_unread(message: Any) -> None:
    """
    Mark an Outlook email as unread for manual review.
    
    Used for exception cases that require human intervention.
    
    Args:
        message: Outlook MailItem COM object
    """
    try:
        message.UnRead = True
        message.FlagStatus = 0  # Clear any flags
        message.Save()
    except Exception:
        pass


def route_folder(pod: Optional[str]) -> Optional[str]:
    """
    Determine output folder based on POD (Point of Delivery) identifier prefix.
    
    Routing Logic:
        - HU prefix → Electricity folder (Hungarian electricity POD format)
        - 39 prefix → Gas folder (Hungarian gas POD format)
        - Unknown → Defaults to electricity folder
        
    Args:
        pod: Point of Delivery identifier extracted from PDF invoice
        
    Returns:
        Full path to target directory, or None if POD is invalid
        
    Example:
        >>> route_folder("HU001234567890")
        'F:\\HUNGARY\\Company\\ÁRAM\\...'
        >>> route_folder("39001234567890")
        'F:\\HUNGARY\\Company\\Földgáz\\...'
    """
    if not pod:
        return None
    
    pod_upper = pod.upper()
    
    if pod_upper.startswith("HU"):
        return OUT_ELECTRICITY
    elif pod_upper.startswith("39"):
        return OUT_GAS
    else:
        # Unknown prefix defaults to electricity (adjust as needed)
        return OUT_ELECTRICITY


# ============================================================================
# CORE PROCESSING LOGIC
# ============================================================================

def process_attachment(message: Any, attachment: Any) -> bool:
    """
    Process a single email attachment through the invoice extraction pipeline.
    
    Pipeline Steps:
        1. Validate attachment is PDF
        2. Extract company name and invoice number from email body
        3. Check for exception companies → mark unread if match
        4. Save attachment to temporary directory
        5. Extract POD identifier from PDF content
        6. Route to appropriate folder based on POD prefix
        7. Rename file using pattern: Company_POD_Invoice.pdf
        8. Handle duplicate filenames with numeric suffixes
        
    Args:
        message: Outlook MailItem COM object
        attachment: Outlook Attachment COM object
        
    Returns:
        True if attachment was successfully processed and saved
        False if processing failed or attachment was skipped
        
    Side Effects:
        - Creates temporary directory (cleaned up after processing)
        - Marks exception emails as unread
        - Saves PDF to target directory
    """
    # Step 1: Validate file type
    filename = str(getattr(attachment, "FileName", "") or "")
    if not filename.lower().endswith(".pdf"):
        return False

    # Step 2: Extract metadata from email body
    email_body = normalize_body_from_mailitem(message)
    company_name = extract_company(email_body)
    invoice_number = extract_invoice(email_body)

    # Step 3: Check exception list
    if is_exception_company(company_name):
        mark_unread(message)
        return False

    # Validate required metadata
    if not company_name or not invoice_number:
        return False

    # Step 4: Save attachment to temporary location
    temp_dir = tempfile.mkdtemp(prefix="tpi_")
    temp_pdf_path = os.path.join(temp_dir, filename)
    
    try:
        attachment.SaveAsFile(temp_pdf_path)
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return False

    # Step 5: Extract POD from PDF content
    pod_identifier = None
    try:
        pod_identifier = extract_pod_from_pdf(temp_pdf_path)
    except Exception:
        pass

    if not pod_identifier:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return False

    # Step 6: Determine target directory
    output_directory = route_folder(pod_identifier)
    if not output_directory:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return False
    
    ensure_dir(output_directory)

    # Step 7: Build standardized filename
    base_filename = f"{safe_filename(company_name)}_{pod_identifier}_{invoice_number}.pdf"
    output_path = os.path.join(output_directory, base_filename)

    # Step 8: Handle duplicate filenames
    final_path = output_path
    counter = 1
    while os.path.exists(final_path):
        root, ext = os.path.splitext(output_path)
        final_path = f"{root} ({counter}){ext}"
        counter += 1

    # Move file to final destination
    shutil.move(temp_pdf_path, final_path)
    shutil.rmtree(temp_dir, ignore_errors=True)
    
    return True


# ============================================================================
# MAIN ORCHESTRATION
# ============================================================================

def main() -> None:
    """
    Main orchestration function for invoice processing pipeline.
    
    Execution Flow:
        1. Connect to Outlook and retrieve emails matching filters
        2. Snapshot email collection (avoid COM collection quirks)
        3. Iterate through messages and process PDF attachments
        4. Mark successfully processed emails as read/complete
        5. Output processing statistics
        
    Filters Applied:
        - Sender: SENDER_EMAIL (configured in step1_list_emails)
        - Time: YEAR-MONTH (configured in step1_list_emails)
        - Message Class: 43 (Outlook MailItem)
        
    Output:
        Prints processing statistics to console
    """
    print(f"[{MAILBOX_NAME}] {YEAR}-{MONTH:02d} — {SENDER_EMAIL}")
    print("Downloading and routing Mirbest Csoport invoices...")

    # Snapshot email collection to avoid COM iteration issues
    items_iterator = get_inbox_items()
    messages = list(items_iterator)  # Critical: convert to static list

    # Processing counters
    total_messages = 0
    total_attachments = 0
    saved_files = 0
    skipped_emails = 0

    # Process each message
    for message in messages:
        # Verify message is a MailItem
        if getattr(message, "Class", None) != OUTLOOK_MAIL_ITEM_CLASS:
            continue

        # Validate sender and date filters
        sender = str(getattr(message, "SenderEmailAddress", "") or "")
        received_time = getattr(message, "ReceivedTime", None)
        
        if not (SENDER_EMAIL.lower() in sender.lower() and 
                received_time and 
                received_time.year == YEAR and 
                received_time.month == MONTH):
            continue

        total_messages += 1
        had_successful_save = False

        # Process attachments
        attachments = getattr(message, "Attachments", None)
        if not attachments:
            skipped_emails += 1
            continue

        # Snapshot attachment collection (COM safety)
        attachment_list = [attachments.Item(i) for i in range(1, attachments.Count + 1)]

        for attachment in attachment_list:
            total_attachments += 1
            processing_success = process_attachment(message, attachment)
            
            if processing_success:
                had_successful_save = True
                saved_files += 1

        # Update message state based on processing result
        if had_successful_save:
            mark_processed(message)
        # Note: Failed messages are left unchanged for review

    # Display processing summary
    if total_messages == 0:
        print("No emails matched sender/month filters.")
    
    print(f"\n{'='*50}")
    print(f"Processing Summary")
    print(f"{'='*50}")
    print(f"Messages scanned : {total_messages}")
    print(f"Attachments seen : {total_attachments}")
    print(f"Files saved      : {saved_files}")
    print(f"Skipped          : {skipped_emails}")
    print(f"{'='*50}")


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    import sys
    
    try:
        main()
    except Exception as error:
        print("\n[ERROR] Unhandled exception occurred:")
        print(error)
        traceback.print_exc()
    finally:
        # Interactive mode: wait for user confirmation before closing
        if len(sys.argv) == 1:
            try:
                input("\nProcessing complete. Press Enter to exit...")
            except EOFError:
                pass
