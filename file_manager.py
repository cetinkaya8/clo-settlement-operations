"""
File Manager Module - CLO Trade Settlement Automation

Handles PDF file validation, renaming, and organization into folders.
"""

import os
from datetime import datetime
from tkinter import messagebox
import pandas as pd

from config import (
    PDF_PREFIXES, 
    CLO_EMAIL_MAPPING, 
    MAX_FILENAME_LENGTH, 
    MAX_PATH_LENGTH,
    SAFE_FILENAME_LENGTH
)
from utils import (
    extract_subtrade_ticket_from_filename,
    get_clo_name,
    detect_prefix_type,
    format_date,
    build_new_filename,
    sanitize_filename
)
from validators import (
    validate_clo_name_in_mapping,
    validate_filename_length,
    validate_path_length,
    validate_sanitization_integrity
)


class FileManager:
    """Manages PDF file renaming and organization."""
    
    def __init__(self, folder_path: str):
        """
        Initialize file manager.
        
        Args:
            folder_path: Path to folder containing PDF files
        """
        self.folder_path = folder_path
        self.pdf_files = []
        self.skipped_files = []
        self.validation_errors = []
        self.rename_plan = []
        self.files_by_ticket_clo = {}
        self.organized_folders = {}
        self.organization_errors = []
    
    def scan_pdf_files(self):
        """Scans folder for PDF files."""
        all_files = os.listdir(self.folder_path)
        self.pdf_files = [f for f in all_files if f.lower().endswith('.pdf')]
    
    def validate_and_plan_renames(
        self, 
        excel_processor, 
        ticket_to_values: dict
    ) -> bool:
        """
        Phase 1: Validates all files and creates rename plan without actually renaming.
        
        Args:
            excel_processor: ExcelProcessor instance with loaded data
            ticket_to_values: Dictionary mapping ticket -> (credit_name, loan_name, bank_name)
            
        Returns:
            True if validation passed, False if critical errors found
        """
        self.validation_errors = []
        self.rename_plan = []
        self.skipped_files = []
        
        for fname in self.pdf_files:
            lower = fname.lower()
            
            # Skip files that don't have expected prefix (silently ignore)
            if not any(lower.startswith(p) for p in PDF_PREFIXES):
                continue
            
            # Extract subtrade ticket from filename
            file_subtrade_ticket = extract_subtrade_ticket_from_filename(fname)
            if not file_subtrade_ticket:
                self.skipped_files.append(fname)
                continue
            
            # Search for this subtrade ticket in Excel
            row_idx = excel_processor.get_row_by_subtrade_ticket(file_subtrade_ticket)
            if row_idx is None:
                self.skipped_files.append(fname)
                continue
            
            # Get row data
            row_data = excel_processor.get_row_data(row_idx)
            
            try:
                clo_name = get_clo_name(
                    row_data['buy_sell_action'],
                    row_data['subtrade_seller'],
                    row_data['subtrade_buyer']
                )
                
                # Critical: Check if CLO name exists in email mapping
                if not validate_clo_name_in_mapping(clo_name, CLO_EMAIL_MAPPING):
                    self.validation_errors.append((
                        fname, 
                        f"CRITICAL ERROR: CLO name '{clo_name}' not found in email mapping list!"
                    ))
                    continue
                
                ticket = row_data['ticket']
                
                # Check if ticket is active (checkbox checked in GUI)
                if ticket not in ticket_to_values:
                    self.skipped_files.append(fname)
                    continue
                
                credit_name, loan_name, bank_name, has_pol = ticket_to_values.get(ticket, ('', '', '', False))
                document_type = detect_prefix_type(fname)
                primary_secondary_type = row_data['primary_secondary_type']
                trade_date = row_data['trade_date']
                
                # Determine date based on document type
                if document_type in ("Exec A&A", "FM", "Exec-POL"):
                    date_by_file = format_date(datetime.now())
                elif document_type == "TC":
                    try:
                        date_by_file = format_date(pd.to_datetime(trade_date))
                    except:
                        self.validation_errors.append((
                            fname, 
                            f"Trade Date not found or invalid for TC file: {trade_date}"
                        ))
                        continue
                else:
                    date_by_file = format_date(datetime.now())
                
                # Build new filename
                new_name_core = build_new_filename(
                    clo_name=clo_name,
                    document_type=document_type,
                    buy_sell_action=row_data['buy_sell_action'],
                    credit_name=credit_name,
                    loan_name=loan_name,
                    primary_secondary_type=primary_secondary_type,
                    bank_name=bank_name,
                    date_by_file=date_by_file,
                    ticket=ticket
                )
                
                # Validate filename length before sanitization
                if len(new_name_core) > SAFE_FILENAME_LENGTH:
                    self.validation_errors.append((
                        fname, 
                        f"ERROR: New filename too long ({len(new_name_core)} characters). "
                        f"File cannot be properly renamed! Renaming cancelled."
                    ))
                    continue
                
                # Sanitize filename
                new_name_core_sanitized = sanitize_filename(new_name_core)
                
                # Validate filename length after sanitization
                filename_with_ext = new_name_core_sanitized + ".pdf"
                is_valid, error_msg = validate_filename_length(filename_with_ext, MAX_FILENAME_LENGTH)
                if not is_valid:
                    self.validation_errors.append((fname, f"ERROR: {error_msg}"))
                    continue
                
                # Validate full path length
                test_path = os.path.join(self.folder_path, filename_with_ext)
                is_valid, error_msg = validate_path_length(test_path, MAX_PATH_LENGTH)
                if not is_valid:
                    self.validation_errors.append((
                        fname, 
                        f"ERROR: File path too long. Folder path too long or filename doesn't fit!"
                    ))
                    continue
                
                # Validate sanitization integrity
                is_valid, error_msg = validate_sanitization_integrity(new_name_core, new_name_core_sanitized)
                if not is_valid:
                    self.validation_errors.append((fname, f"ERROR: {error_msg}"))
                    continue
                
                # Store rename plan
                self.rename_plan.append({
                    'original_name': fname,
                    'new_name_core': new_name_core_sanitized,
                    'file_subtrade_ticket': file_subtrade_ticket,
                    'ticket': ticket,
                    'clo_name': clo_name,
                    'credit_name': credit_name,
                    'loan_name': loan_name,
                    'buy_sell_action': row_data['buy_sell_action'],
                    'primary_secondary_type': primary_secondary_type,
                    'has_pol': has_pol
                })
                
            except Exception as e:
                self.validation_errors.append((fname, f"Validation error: {e}"))
                continue
        
        # If validation errors exist, show error and stop
        if self.validation_errors:
            error_msg = "CRITICAL ERROR: File naming issue detected. NO FILES WERE PROCESSED!\n\n"
            error_msg += "All operations cancelled. No files were renamed, moved, or emails sent.\n\n"
            error_msg += "Errors:\n"
            for fname, err in self.validation_errors[:5]:
                error_msg += f"- {fname}: {err}\n"
            if len(self.validation_errors) > 5:
                error_msg += f"... and {len(self.validation_errors) - 5} more errors\n"
            error_msg += "\nPlease fix the errors and try again."
            messagebox.showerror("CRITICAL ERROR - Operation Cancelled", error_msg)
            return False
        
        return True
    
    def execute_renames(self) -> dict:
        """
        Phase 2: Execute the rename plan (only called after validation passes).
        
        Returns:
            Dictionary of renamed files grouped by ticket
        """
        renamed_by_ticket = {}
        errors = []
        
        for plan in self.rename_plan:
            try:
                fname = plan['original_name']
                new_name_core = plan['new_name_core']
                file_subtrade_ticket = plan['file_subtrade_ticket']
                ticket = plan['ticket']
                clo_name = plan['clo_name']
                credit_name = plan['credit_name']
                
                # Rename file
                src = os.path.join(self.folder_path, fname)
                dst = os.path.join(self.folder_path, new_name_core + ".pdf")
                
                # Handle duplicate filenames
                suffix = 1
                final_dst = dst
                while os.path.exists(final_dst):
                    final_dst = os.path.join(self.folder_path, f"{new_name_core} ({suffix}).pdf")
                    suffix += 1
                
                os.rename(src, final_dst)
                renamed_file = os.path.basename(final_dst)
                renamed_by_ticket.setdefault(file_subtrade_ticket, []).append(renamed_file)
                
                # Group by ticket and CLO name
                group_key = (ticket, clo_name)
                if group_key not in self.files_by_ticket_clo:
                    self.files_by_ticket_clo[group_key] = {
                        'files': [],
                        'credit_name': credit_name,
                        'loan_name': plan['loan_name'],
                        'buy_sell_action': plan['buy_sell_action'],
                        'primary_secondary_type': plan['primary_secondary_type'],
                        'has_pol': plan.get('has_pol', False)
                    }
                self.files_by_ticket_clo[group_key]['files'].append(renamed_file)
                
            except Exception as e:
                errors.append((fname, f"Excel/rename error: {e}"))
                continue
        
        return renamed_by_ticket
    
    def organize_files_into_folders(self):
        """
        Organizes renamed files into folders by credit name and ticket.
        """
        self.organized_folders = {}
        self.organization_errors = []
        
        for (ticket, clo_name), group_data in self.files_by_ticket_clo.items():
            try:
                credit_name = group_data['credit_name']
                files = group_data['files']
                
                if not credit_name or not ticket:
                    continue
                
                # Create folder name: credit_name #ticket
                folder_name = f"{credit_name} #{ticket}"
                folder_name = sanitize_filename(folder_name)
                target_folder = os.path.join(self.folder_path, folder_name)
                
                # Create folder if it doesn't exist
                if not os.path.exists(target_folder):
                    os.makedirs(target_folder)
                
                # Move files
                moved_count = 0
                for file_name in files:
                    src_file = os.path.join(self.folder_path, file_name)
                    dst_file = os.path.join(target_folder, file_name)
                    
                    if os.path.exists(src_file):
                        # Handle duplicate filenames
                        if os.path.exists(dst_file) and src_file != dst_file:
                            base_name, ext = os.path.splitext(file_name)
                            suffix = 1
                            while os.path.exists(dst_file):
                                new_name = f"{base_name} ({suffix}){ext}"
                                dst_file = os.path.join(target_folder, new_name)
                                suffix += 1
                        
                        if src_file != dst_file:
                            os.rename(src_file, dst_file)
                            moved_count += 1
                
                if moved_count > 0:
                    self.organized_folders[folder_name] = moved_count
                    
            except Exception as e:
                self.organization_errors.append((f"Ticket {ticket}, Credit {credit_name}", str(e)))
