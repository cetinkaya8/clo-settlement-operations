"""
Validators Module - CLO Trade Settlement Automation

Contains all validation functions for email addresses, filenames, and data integrity checks.
"""

import re


def validate_email_address(email_str: str) -> bool:
    """
    Validates email address format.
    
    Supports multiple email addresses separated by semicolons.
    
    Args:
        email_str: Email address or semicolon-separated list of email addresses
        
    Returns:
        True if all email addresses are valid, False otherwise
    """
    if not email_str or not email_str.strip():
        return False
    
    # Check if multiple emails exist (separated by semicolons)
    emails = [e.strip() for e in email_str.split(';') if e.strip()]
    if not emails:
        return False
    
    # Simple email format validation
    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    for email in emails:
        if not re.match(email_pattern, email):
            return False
    
    return True


def validate_filename_length(filename: str, max_length: int = 255) -> tuple[bool, str]:
    """
    Validates if filename length is within system limits.
    
    Args:
        filename: The filename to validate
        max_length: Maximum allowed filename length (default: 255 for Windows)
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    if len(filename) > max_length:
        return False, f"Filename too long ({len(filename)} characters > {max_length})"
    return True, ""


def validate_path_length(path: str, max_length: int = 260) -> tuple[bool, str]:
    """
    Validates if full path length is within system limits.
    
    Args:
        path: The full path to validate
        max_length: Maximum allowed path length (default: 260 for Windows)
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    if len(path) > max_length:
        return False, f"Path too long ({len(path)} characters > {max_length})"
    return True, ""


def validate_required_fields(field_dict: dict[str, any]) -> tuple[bool, list[str]]:
    """
    Validates that all required fields are present and non-empty.
    
    Args:
        field_dict: Dictionary of field_name -> value pairs
        
    Returns:
        Tuple of (all_valid, list_of_missing_fields)
    """
    missing_fields = []
    for field_name, value in field_dict.items():
        if not value or (isinstance(value, str) and not value.strip()):
            missing_fields.append(field_name)
    
    return len(missing_fields) == 0, missing_fields


def validate_clo_name_in_mapping(clo_name: str, mapping: dict) -> bool:
    """
    Validates if CLO name exists in the email mapping.
    
    Args:
        clo_name: CLO name to validate
        mapping: Email mapping dictionary
        
    Returns:
        True if CLO name exists in mapping, False otherwise
    """
    return clo_name in mapping


def validate_sanitization_integrity(original: str, sanitized: str, max_loss: int = 5) -> tuple[bool, str]:
    """
    Validates that sanitization didn't cause significant data loss.
    
    Args:
        original: Original string before sanitization
        sanitized: String after sanitization
        max_loss: Maximum allowed character loss
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    char_diff = len(original) - len(sanitized)
    
    if char_diff > max_loss or len(sanitized) < 20:
        return False, f"Significant data loss during sanitization ({char_diff} characters lost)"
    
    return True, ""
