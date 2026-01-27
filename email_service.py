"""
Email Service Module - CLO Trade Settlement Automation

Handles Outlook email draft creation for trade settlements.
"""

import os
from datetime import datetime
from tkinter import messagebox

from config import (
    CLO_EMAIL_MAPPING, 
    DEFAULT_CC, 
    OUTLOOK_EMAIL_ACCOUNT,
    EMAIL_SIGNATURE
)
from utils import safe_get_value, escape_html, remove_ltd_suffix
from validators import validate_email_address

# Outlook integration
try:
    import win32com.client as win32
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False


class EmailService:
    """Service for creating Outlook email drafts."""
    
    def __init__(self, folder_path: str):
        """
        Initialize email service.
        
        Args:
            folder_path: Path to folder containing attachments
        """
        self.folder_path = folder_path
        self.created_drafts = 0
        self.draft_errors = []
        self.outlook = None
        self.account = None
    
    def is_available(self) -> bool:
        """
        Check if Outlook integration is available.
        
        Returns:
            True if Outlook is available, False otherwise
        """
        return OUTLOOK_AVAILABLE
    
    def connect_to_outlook(self) -> bool:
        """
        Establishes connection to Outlook application.
        
        Returns:
            True if successful, False otherwise
        """
        if not OUTLOOK_AVAILABLE:
            messagebox.showwarning(
                "Outlook Not Available", 
                "Outlook integration is not available. Email drafts could not be created."
            )
            return False
        
        try:
            # Connect to Outlook
            try:
                self.outlook = win32.Dispatch("Outlook.Application")
            except Exception as e:
                raise Exception(f"Could not connect to Outlook application: {e}")
            
            try:
                namespace = self.outlook.GetNamespace("MAPI")
            except Exception as e:
                raise Exception(f"Could not get Outlook namespace: {e}")
            
            # Find specified email account
            self.account = None
            try:
                for acc in namespace.Accounts:
                    if hasattr(acc, 'SmtpAddress') and acc.SmtpAddress == OUTLOOK_EMAIL_ACCOUNT:
                        self.account = acc
                        print(f"Found account: {OUTLOOK_EMAIL_ACCOUNT}")
                        break
            except Exception as e:
                raise Exception(f"Could not read email accounts: {e}")
            
            if not self.account:
                raise Exception(f"Specified email account not found: {OUTLOOK_EMAIL_ACCOUNT}")
            
            return True
            
        except Exception as e:
            error_msg = f"Outlook connection error: {e}"
            messagebox.showerror("Outlook Error", error_msg)
            print(f"CRITICAL ERROR: {error_msg}")
            return False
    
    def create_email_drafts(self, files_by_ticket_clo: dict):
        """
        Creates email drafts for each (ticket, CLO) group.
        
        Args:
            files_by_ticket_clo: Dictionary with structure:
                {(ticket, clo_name): {
                    'files': [list of filenames],
                    'credit_name': str,
                    'loan_name': str,
                    'buy_sell_action': str,
                    'primary_secondary_type': str
                }}
        """
        if not self.connect_to_outlook():
            return
        
        self.created_drafts = 0
        self.draft_errors = []
        
        # Pre-validation: Check all emails for maximum attachment limit (max 3 files)
        for (ticket, clo_name), group_data in files_by_ticket_clo.items():
            files = group_data.get('files', [])
            if isinstance(files, list) and len(files) > 3:
                error_msg = (
                    f"SECURITY WARNING: Email creation cancelled!\n\n"
                    f"Ticket {ticket}, CLO {clo_name} has {len(files)} files.\n"
                    f"Maximum 3 files allowed.\n\n"
                    f"NO EMAILS WERE CREATED!"
                )
                messagebox.showerror("Security Check - Email Cancelled", error_msg)
                print(f"SECURITY ERROR: Too many attachments for Ticket {ticket}, CLO {clo_name}: {len(files)} files")
                return
        
        for (ticket, clo_name), group_data in files_by_ticket_clo.items():
            try:
                # Validate data
                files = group_data.get('files', [])
                if not files or not isinstance(files, list):
                    self.draft_errors.append((
                        f"Ticket {ticket}, CLO {clo_name}", 
                        "Associated file list is empty or invalid"
                    ))
                    continue
                
                # Safe value retrieval
                credit_name = safe_get_value(group_data.get('credit_name', ''))
                loan_name = safe_get_value(group_data.get('loan_name', ''))
                buy_sell_action = safe_get_value(group_data.get('buy_sell_action', ''), 'UNKNOWN')
                primary_secondary_type = safe_get_value(group_data.get('primary_secondary_type', ''), 'Primary')
                has_pol = group_data.get('has_pol', False)
                ticket = safe_get_value(ticket, '')
                clo_name = safe_get_value(clo_name, 'Unknown CLO')
                
                # Check required fields
                if not credit_name or not loan_name:
                    self.draft_errors.append((
                        f"Ticket {ticket}, CLO {clo_name}", 
                        f"Missing information: credit_name={credit_name}, loan_name={loan_name}"
                    ))
                    continue
                
                # Get email recipients and validate
                if clo_name not in CLO_EMAIL_MAPPING:
                    self.draft_errors.append((
                        f"Ticket {ticket}, CLO {clo_name}", 
                        f"CRITICAL ERROR: CLO name '{clo_name}' not found in email mapping list!"
                    ))
                    continue
                
                email_info = CLO_EMAIL_MAPPING.get(clo_name, {})
                recipient_to = email_info.get("to", "").strip()
                recipient_cc = email_info.get("cc", DEFAULT_CC).strip()
                
                # Validate email addresses
                if recipient_to and not validate_email_address(recipient_to):
                    self.draft_errors.append((
                        f"Ticket {ticket}, CLO {clo_name}", 
                        f"Invalid TO email address: {recipient_to}"
                    ))
                    continue
                
                if recipient_cc and not validate_email_address(recipient_cc):
                    print(f"WARNING: Invalid CC email address for Ticket {ticket}, CLO {clo_name}: {recipient_cc}")
                    recipient_cc = DEFAULT_CC
                
                if not recipient_to:
                    self.draft_errors.append((
                        f"Ticket {ticket}, CLO {clo_name}", 
                        f"TO email address not found for CLO '{clo_name}'."
                    ))
                    continue
                
                # Create email content
                subject, body = self._create_email_content(
                    clo_name, credit_name, loan_name, 
                    buy_sell_action, primary_secondary_type, ticket, has_pol
                )
                
                if not subject or not body:
                    continue
                
                # Create and configure email
                success = self._create_and_configure_email(
                    subject, body, recipient_to, recipient_cc, 
                    files, ticket, clo_name
                )
                
                if success:
                    self.created_drafts += 1
                    
            except Exception as e:
                error_msg = f"Unexpected error: {str(e)}"
                self.draft_errors.append((
                    f"Ticket {ticket if ticket else 'N/A'}, CLO {clo_name if clo_name else 'N/A'}", 
                    error_msg
                ))
                print(f"ERROR: Ticket {ticket}, CLO {clo_name} - {error_msg}")
    
    def _create_email_content(
        self, 
        clo_name: str, 
        credit_name: str, 
        loan_name: str,
        buy_sell_action: str, 
        primary_secondary_type: str, 
        ticket: str,
        has_pol: bool = False
    ) -> tuple[str, str]:
        """
        Creates email subject and body content.
        
        Returns:
            Tuple of (subject, body) or (None, None) if error
        """
        try:
            current_date = datetime.now().strftime('%m-%d-%y')
        except Exception as e:
            self.draft_errors.append((
                f"Ticket {ticket}, CLO {clo_name}", 
                f"Could not create date: {e}"
            ))
            return None, None
        
        # Determine REMIT or EXPECT
        # For normal trades: BUY=REMIT, SELL=EXPECT
        # For POL trades: BUY=EXPECT, SELL=REMIT (reversed)
        if has_pol:
            remit_or_expect = "EXPECT" if buy_sell_action.upper() == "BUY" else "REMIT"
        else:
            remit_or_expect = "REMIT" if buy_sell_action.upper() == "BUY" else "EXPECT"
        
        # Remove Ltd suffix from CLO name for body
        clo_name_for_body = remove_ltd_suffix(clo_name)
        
        # HTML escape operations
        clo_name_escaped = escape_html(clo_name)
        clo_name_body_escaped = escape_html(clo_name_for_body)
        credit_name_escaped = escape_html(credit_name)
        loan_name_escaped = escape_html(loan_name)
        buy_sell_action_escaped = escape_html(buy_sell_action)
        primary_secondary_type_escaped = escape_html(primary_secondary_type)
        ticket_escaped = escape_html(ticket)
        
        # Create subject
        try:
            subject = (
                f"{clo_name} - Trade Settlement ({buy_sell_action}) "
                f"effective ***{current_date}*** {credit_name} ({loan_name})"
            )
            if ticket:
                subject += f" #{ticket}"
            
            # Outlook subject limit
            if len(subject) > 250:
                subject = subject[:247] + "..."
        except Exception as e:
            self.draft_errors.append((
                f"Ticket {ticket}, CLO {clo_name}", 
                f"Could not create subject: {e}"
            ))
            return None, None
        
        # Create body
        try:
            if has_pol:
                body = self._build_pol_email_body(
                    clo_name_body_escaped,
                    buy_sell_action_escaped,
                    primary_secondary_type_escaped,
                    credit_name_escaped,
                    loan_name_escaped,
                    ticket_escaped,
                    remit_or_expect,
                    current_date
                )
            else:
                body = self._build_email_body(
                    clo_name_body_escaped,
                    buy_sell_action_escaped,
                    primary_secondary_type_escaped,
                    credit_name_escaped,
                    loan_name_escaped,
                    ticket_escaped,
                    remit_or_expect,
                    current_date
                )
        except Exception as e:
            self.draft_errors.append((
                f"Ticket {ticket}, CLO {clo_name}", 
                f"Could not create body: {e}"
            ))
            return None, None
        
        return subject, body
    
    def _build_email_body(
        self,
        clo_name: str,
        buy_sell_action: str,
        primary_secondary_type: str,
        credit_name: str,
        loan_name: str,
        ticket: str,
        remit_or_expect: str,
        current_date: str
    ) -> str:
        """
        Builds HTML email body with signature.
        
        Returns:
            HTML body string
        """
        body = (
            f"<strong><u>{clo_name} Ltd</u></strong>,<br><br>"
            f"Attached, please find <span style='color:blue; font-weight:bold;'>"
            f"({buy_sell_action}-{primary_secondary_type})</span> Trade settlement closing package for – "
            f"<span style='color:blue; font-weight:bold;'>{credit_name} ({loan_name})</span>"
        )
        
        if ticket:
            body += f" #{ticket}"
        
        body += (
            f"<br><br>Please <span style='color:blue; font-weight:bold;'>"
            f"***{remit_or_expect} FUNDS on {current_date}***</span><br><br>"
            "Please let me know if you have any questions.<br><br>Best Regards,"
            "<br><br>"
        )
        
        # Add signature from config
        sig = EMAIL_SIGNATURE
        body += (
            f"<span style='font-family:Arial; font-size:10pt; color:#003366;'>"
            f"<strong>{sig['name']}</strong></span><br>"
            f"<a href='mailto:{sig['email']}' style='font-family:Arial; font-size:10pt; color:#0056b3;'>"
            f"{sig['email']}</a><br>"
            f"<span style='font-family:Arial; font-size:10pt; color:#003366;'>"
            f"<strong>{sig['company']}</strong></span><br>"
            f"<span style='font-family:Arial; font-size:10pt; color:#0056b3;'>"
            f"{sig['address_line1']}<br>{sig['address_line2']}</span>"
        )
        
        return body
    
    def _build_pol_email_body(
        self,
        clo_name: str,
        buy_sell_action: str,
        primary_secondary_type: str,
        credit_name: str,
        loan_name: str,
        ticket: str,
        remit_or_expect: str,
        current_date: str
    ) -> str:
        """
        Builds HTML email body for POL (PayoutLetter) emails with signature.
        
        Returns:
            HTML body string
        """
        body = (
            f"<strong><u>{clo_name} Ltd</u></strong>,<br><br>"
            f"Attached, please find <span style='color:blue; font-weight:bold;'>"
            f"({buy_sell_action}-{primary_secondary_type})</span> Trade settlement closing package for – "
            f"<span style='color:blue; font-weight:bold;'>{credit_name} ({loan_name})</span>"
        )
        
        if ticket:
            body += f" #{ticket}"
        
        body += (
            f"<br><br>Please <span style='color:blue; font-weight:bold;'>"
            f"***{remit_or_expect} FUNDS on {current_date}***</span><br><br>"
            "<span style='color:red; font-weight:bold;'>This trade settled with POL.</span><br><br>"
            "Please let me know if you have any questions.<br><br>Best Regards,"
            "<br><br>"
        )
        
        # Add signature from config
        sig = EMAIL_SIGNATURE
        body += (
            f"<span style='font-family:Arial; font-size:10pt; color:#003366;'>"
            f"<strong>{sig['name']}</strong></span><br>"
            f"<a href='mailto:{sig['email']}' style='font-family:Arial; font-size:10pt; color:#0056b3;'>"
            f"{sig['email']}</a><br>"
            f"<span style='font-family:Arial; font-size:10pt; color:#003366;'>"
            f"<strong>{sig['company']}</strong></span><br>"
            f"<span style='font-family:Arial; font-size:10pt; color:#0056b3;'>"
            f"{sig['address_line1']}<br>{sig['address_line2']}</span>"
        )
        
        return body
    
    def _create_and_configure_email(
        self,
        subject: str,
        body: str,
        recipient_to: str,
        recipient_cc: str,
        files: list,
        ticket: str,
        clo_name: str
    ) -> bool:
        """
        Creates email item, attaches files, and saves as draft.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
        except Exception as e:
            self.draft_errors.append((
                f"Ticket {ticket}, CLO {clo_name}", 
                f"Could not create mail item: {e}"
            ))
            return False
        
        try:
            mail.SendUsingAccount = self.account
        except Exception as e:
            self.draft_errors.append((
                f"Ticket {ticket}, CLO {clo_name}", 
                f"Could not set mail account: {e}"
            ))
            return False
        
        try:
            mail.Subject = subject
            mail.BodyFormat = 2  # 2 = olFormatHTML
            mail.HTMLBody = body
        except Exception as e:
            self.draft_errors.append((
                f"Ticket {ticket}, CLO {clo_name}", 
                f"Could not set mail content: {e}"
            ))
            return False
        
        try:
            if recipient_to:
                mail.To = recipient_to
            if recipient_cc:
                mail.CC = recipient_cc
        except Exception as e:
            self.draft_errors.append((
                f"Ticket {ticket}, CLO {clo_name}", 
                f"Could not set mail recipients: {e}"
            ))
            return False
        
        # Attach files
        attached_files, missing_files = self._attach_files(mail, files)
        
        # Don't create email if no files were attached
        if not attached_files:
            self.draft_errors.append((
                f"Ticket {ticket}, CLO {clo_name}", 
                f"No files could be attached. Missing files: {', '.join(missing_files[:3])}"
            ))
            return False
        
        # Warn if some files are missing
        if missing_files:
            print(f"WARNING: Some files could not be attached for Ticket {ticket}, CLO {clo_name}: {', '.join(missing_files[:3])}")
        
        # Save as draft (DOES NOT SEND!)
        try:
            mail.Save()
            return True
        except Exception as e:
            self.draft_errors.append((
                f"Ticket {ticket}, CLO {clo_name}", 
                f"Could not save mail draft: {e}"
            ))
            return False
    
    def _attach_files(self, mail, files: list) -> tuple[list, list]:
        """
        Attaches files to email.
        
        Returns:
            Tuple of (attached_files, missing_files)
        """
        attached_files = []
        missing_files = []
        
        for file_name in files:
            if not file_name or not isinstance(file_name, str):
                missing_files.append(f"{file_name}: Invalid filename")
                continue
            
            try:
                file_path = os.path.join(self.folder_path, file_name)
            except Exception as e:
                missing_files.append(f"{file_name}: Could not create path - {e}")
                continue
            
            if not os.path.exists(file_path):
                missing_files.append(f"{file_name}: File not found")
                continue
            
            if not os.path.isfile(file_path):
                missing_files.append(f"{file_name}: Not a file")
                continue
            
            try:
                mail.Attachments.Add(file_path)
                attached_files.append(file_name)
            except Exception as e:
                missing_files.append(f"{file_name}: Could not attach - {e}")
        
        return attached_files, missing_files
