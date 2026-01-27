"""
Utils Module - CLO Trade Settlement Automation

Contains utility functions for string manipulation, file handling, and data processing.
"""

import re
from datetime import datetime
import pandas as pd


def sanitize_filename(filename: str) -> str:
    """
    Sanitizes filename by removing invalid characters.
    
    Args:
        filename: The filename to sanitize
        
    Returns:
        Sanitized filename truncated to safe length
    """
    invalid_chars = '<>:"/\\|?*\n\r\t'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename[:240]


def escape_html(text: str) -> str:
    """
    Escapes special characters in HTML to prevent XSS attacks.
    
    Args:
        text: Text to escape
        
    Returns:
        HTML-escaped text
    """
    if not text:
        return ""
    text = str(text)
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('"', '&quot;')
    text = text.replace("'", '&#x27;')
    return text


def safe_get_value(value, default=""):
    """
    Safely retrieves a value, handles None/NaN cases.
    
    Args:
        value: Value to retrieve
        default: Default value if input is None/NaN
        
    Returns:
        Cleaned string value or default
    """
    if pd.isna(value) or value is None:
        return default
    result = str(value).strip()
    return result if result else default


def clean_ticket(raw):
    """
    Cleans ticket number by removing non-alphanumeric characters.
    
    Args:
        raw: Raw ticket value
        
    Returns:
        Cleaned ticket string
    """
    if pd.isna(raw):
        return ''
    s = str(raw).replace('-', '').replace(' ', '')
    s = re.sub(r'[^A-Za-z0-9]', '', s)
    return s


def detect_prefix_type(filename):
    """
    Detects document type based on filename prefix.
    
    Args:
        filename: PDF filename
        
    Returns:
        Document type string: "Exec A&A", "TC", "FM", or ""
    """
    lower = filename.lower()
    if lower.startswith('aa_'):
        return "Exec A&A"
    if lower.startswith('confirm_'):
        return "TC"
    if lower.startswith('fundingmemo_') or lower.startswith('funding_memo_'):
        return "FM"
    if lower.startswith('payoutletter_'):
        return "Exec-POL"
    return ""


def extract_subtrade_ticket_from_filename(filename):
    """
    Extracts subtrade ticket number from filename.
    
    Expected format: {prefix}_{ticket}_{subtrade_ticket}.pdf
    Example: AA_7323626_7323626001.pdf -> 7323626001
    
    Args:
        filename: PDF filename
        
    Returns:
        Subtrade ticket string or None if not found
    """
    # Remove .pdf extension
    name_without_ext = filename.rsplit('.', 1)[0]
    
    # Split by underscore
    parts = name_without_ext.split('_')
    
    # Must have at least 3 parts: prefix, ticket, subtrade_ticket
    # Last part is subtrade_ticket
    if len(parts) >= 3:
        subtrade_ticket = parts[-1]
        return clean_ticket(subtrade_ticket)
    
    return None


def format_date(dt):
    """
    Formats datetime object to MM-DD-YYYY string.
    
    Args:
        dt: datetime object
        
    Returns:
        Formatted date string
    """
    return dt.strftime("%m-%d-%Y")


def get_clo_name(buy_sell_action, subtrade_seller, subtrade_buyer):
    """
    Determines CLO name based on buy/sell action.
    
    For SELL action, returns seller name.
    For BUY action, returns buyer name.
    
    Args:
        buy_sell_action: "BUY" or "SELL"
        subtrade_seller: Seller CLO name
        subtrade_buyer: Buyer CLO name
        
    Returns:
        Appropriate CLO name
    """
    if buy_sell_action.upper() == "SELL":
        return subtrade_seller
    elif buy_sell_action.upper() == "BUY":
        return subtrade_buyer
    return "UnknownCLO"


def build_new_filename(
    clo_name,
    document_type,
    buy_sell_action,
    credit_name,
    loan_name,
    primary_secondary_type,
    bank_name,
    date_by_file,
    ticket
):
    """
    Builds standardized filename for CLO settlement documents.
    
    Format: {CLO} {DocType}({Action}) - {Credit}({Loan}) - {Type} - {Bank} - {Date} - #{Ticket}
    
    Args:
        clo_name: CLO name
        document_type: Document type (TC, FM, Exec A&A)
        buy_sell_action: BUY or SELL
        credit_name: Credit/borrower name
        loan_name: Loan name
        primary_secondary_type: Primary or Secondary
        bank_name: Bank name
        date_by_file: Date string
        ticket: Ticket number
        
    Returns:
        Formatted filename (without extension)
    """
    return (
        f"{clo_name} {document_type}({buy_sell_action}) - "
        f"{credit_name}({loan_name}) - "
        f"{primary_secondary_type} - "
        f"{bank_name} - "
        f"{date_by_file} - "
        f"#{ticket}"
    )


def remove_ltd_suffix(clo_name: str) -> str:
    """
    Removes ", Ltd." or " Ltd" suffix from CLO name.
    
    Args:
        clo_name: CLO name with potential Ltd suffix
        
    Returns:
        CLO name without Ltd suffix
    """
    clo_name = clo_name.strip()
    if clo_name.endswith(", Ltd."):
        return clo_name[:-6]
    elif clo_name.endswith(" Ltd"):
        return clo_name[:-4]
    return clo_name
