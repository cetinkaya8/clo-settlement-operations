"""
GUI Module - CLO Trade Settlement Automation

Tkinter-based graphical user interface for the application.
"""

import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar

from excel_processor import ExcelProcessor
from file_manager import FileManager
from email_service import EmailService
from config import BANK_MAPPING_FILE, CREDIT_MAPPING_FILE


class SettlementGUI:
    """Main GUI application for CLO trade settlement automation."""
    
    def __init__(self, root):
        """
        Initialize the GUI.
        
        Args:
            root: Tkinter root window
        """
        self.root = root
        self.root.title("Bulk PDF Renamer")
        self.root.geometry("800x600")
        self.root.configure(bg='#0a3d62')
        
        self.folder_var = StringVar()
        self.excel_var = StringVar()
        self.ticket_entries = {}
        self.bank_mapping = {}  # Maps Trade Seller/Buyer -> Bank Name
        self.credit_mapping = {}  # Maps Excel Credit Name -> GUI Credit Name
        
        self._load_bank_mapping()
        self._load_credit_mapping()
        self._build_ui()
    
    def _load_bank_mapping(self):
        """Loads bank name mapping from file."""
        try:
            if os.path.exists(BANK_MAPPING_FILE):
                with open(BANK_MAPPING_FILE, 'r', encoding='utf-8') as f:
                    self.bank_mapping = json.load(f)
        except Exception as e:
            print(f"Failed to load bank mapping file: {e}")
            self.bank_mapping = {}
    
    def _save_bank_mapping(self):
        """Saves bank name mapping to file."""
        try:
            with open(BANK_MAPPING_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.bank_mapping, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Failed to save bank mapping file: {e}")
    
    def _load_credit_mapping(self):
        """Loads credit name mapping from file."""
        try:
            if os.path.exists(CREDIT_MAPPING_FILE):
                with open(CREDIT_MAPPING_FILE, 'r', encoding='utf-8') as f:
                    self.credit_mapping = json.load(f)
        except Exception as e:
            print(f"Failed to load credit mapping file: {e}")
            self.credit_mapping = {}
    
    def _save_credit_mapping(self):
        """Saves credit name mapping to file."""
        try:
            with open(CREDIT_MAPPING_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.credit_mapping, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Failed to save credit mapping file: {e}")
    
    def _build_ui(self):
        """Builds the user interface components."""
        # Top frame for folder/file selection
        frame_top = tk.Frame(self.root, bg='#0a3d62', padx=10, pady=10)
        frame_top.pack(fill=tk.X)
        
        # PDF Folder selection
        tk.Label(frame_top, text="PDF Folder:", bg='#0a3d62', fg='white').pack(anchor='w')
        f_frame = tk.Frame(frame_top, bg='#0a3d62')
        f_frame.pack(fill=tk.X, pady=(0, 5))
        tk.Entry(f_frame, textvariable=self.folder_var, width=70).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(f_frame, text="Browse", command=self._browse_folder, bg='white').pack(side=tk.LEFT)
        
        # Excel File selection
        tk.Label(frame_top, text="Excel File:", bg='#0a3d62', fg='white').pack(anchor='w')
        e_frame = tk.Frame(frame_top, bg='#0a3d62')
        e_frame.pack(fill=tk.X, pady=(0, 5))
        tk.Entry(e_frame, textvariable=self.excel_var, width=70).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(e_frame, text="Browse", command=self._browse_excel, bg='white').pack(side=tk.LEFT)
        
        # Tickets frame for dynamic ticket entries
        self.tickets_frame = tk.Frame(self.root, bg='#0a3d62')
        self.tickets_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Run button
        self.run_button = tk.Button(
            self.root, 
            text="RUN", 
            command=self._on_run, 
            bg='green', 
            fg='white', 
            font=('Arial', 12, 'bold')
        )
        self.run_button.pack(pady=10)
    
    def _browse_folder(self):
        """Opens dialog to select PDF folder."""
        path = filedialog.askdirectory()
        if path:
            self.folder_var.set(path)
    
    def _browse_excel(self):
        """Opens dialog to select Excel file and populates ticket entries."""
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if path:
            self.excel_var.set(path)
            self._populate_ticket_entries(path)
    
    def _populate_ticket_entries(self, excel_path: str):
        """
        Populates the ticket entry table from Excel file.
        
        Args:
            excel_path: Path to Excel file
        """
        # Clear existing entries
        for widget in self.tickets_frame.winfo_children():
            widget.destroy()
        self.ticket_entries.clear()
        
        # Read Excel and get unique tickets
        excel_processor = ExcelProcessor(excel_path)
        if not excel_processor.read_excel():
            return
        
        unique_tickets = excel_processor.get_unique_tickets()
        
        # Create dictionaries for credit names and trade info
        ticket_credit_map = {}
        ticket_trade_info = {}
        for tkt in unique_tickets:
            credit_name = excel_processor.get_credit_name_for_ticket(tkt)
            ticket_credit_map[tkt] = credit_name
            
            trade_info = excel_processor.get_trade_info_for_ticket(tkt)
            ticket_trade_info[tkt] = trade_info
        
        # Create table headers
        tk.Label(
            self.tickets_frame, text="Active", bg='#0a3d62', fg='white', 
            font=('Arial', 9, 'bold')
        ).grid(row=0, column=0, padx=5, pady=5)
        
        tk.Label(
            self.tickets_frame, text="Ticket", bg='#0a3d62', fg='white', 
            font=('Arial', 9, 'bold')
        ).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(
            self.tickets_frame, text="Credit Name", bg='#0a3d62', fg='white', 
            font=('Arial', 9, 'bold')
        ).grid(row=0, column=2, padx=5, pady=5)
        
        tk.Label(
            self.tickets_frame, text="Loan Name", bg='#0a3d62', fg='white', 
            font=('Arial', 9, 'bold')
        ).grid(row=0, column=3, padx=5, pady=5)
        
        tk.Label(
            self.tickets_frame, text="Bank Name", bg='#0a3d62', fg='white', 
            font=('Arial', 9, 'bold')
        ).grid(row=0, column=4, padx=5, pady=5)
        
        tk.Label(
            self.tickets_frame, text="POL", bg='#0a3d62', fg='white', 
            font=('Arial', 9, 'bold')
        ).grid(row=0, column=5, padx=5, pady=5)
        
        # Create row for each ticket
        for i, tkt in enumerate(unique_tickets, start=1):
            is_active = tk.BooleanVar(value=True)  # Default: active
            
            # Format: "Ticket : Credit Name" if credit name exists, otherwise just "Ticket"
            credit_name = ticket_credit_map.get(tkt, '')
            display_text = f"{tkt} : {credit_name}" if credit_name else tkt
            
            ticket_label = tk.Label(self.tickets_frame, text=display_text, bg='#0a3d62', fg='white')
            ticket_label.grid(row=i, column=1, padx=5, pady=3)
            
            credit_name = tk.StringVar()
            loan_name = tk.StringVar()
            bank_name = tk.StringVar()
            
            # Auto-fill Credit Name from mapping if available
            excel_credit_name = ticket_credit_map.get(tkt, '')
            if excel_credit_name and excel_credit_name in self.credit_mapping:
                credit_name.set(self.credit_mapping[excel_credit_name])
            
            # Auto-fill Bank Name from mapping if available
            trade_info = ticket_trade_info.get(tkt, {})
            if trade_info:
                trade_action = trade_info.get('trade_action', '')
                if trade_action == 'BUY':
                    mapping_key = trade_info.get('trade_seller', '')
                elif trade_action == 'SELL':
                    mapping_key = trade_info.get('trade_buyer', '')
                else:
                    mapping_key = ''
                
                if mapping_key and mapping_key in self.bank_mapping:
                    bank_name.set(self.bank_mapping[mapping_key])
            
            credit_entry = tk.Entry(self.tickets_frame, textvariable=credit_name, width=15)
            credit_entry.grid(row=i, column=2, padx=5, pady=3)
            
            loan_entry = tk.Entry(self.tickets_frame, textvariable=loan_name, width=15)
            loan_entry.grid(row=i, column=3, padx=5, pady=3)
            
            bank_entry = tk.Entry(self.tickets_frame, textvariable=bank_name, width=15)
            bank_entry.grid(row=i, column=4, padx=5, pady=3)
            
            # POL checkbox
            has_pol = tk.BooleanVar(value=False)  # Default: No POL
            pol_checkbox = tk.Checkbutton(
                self.tickets_frame,
                variable=has_pol,
                bg='#0a3d62',
                activebackground='#0a3d62'
            )
            pol_checkbox.grid(row=i, column=5, padx=5, pady=3)
            
            # Store widgets for this row
            row_widgets = [ticket_label, credit_entry, loan_entry, bank_entry, pol_checkbox]
            
            # Create checkbox
            checkbox = tk.Checkbutton(
                self.tickets_frame,
                variable=is_active,
                bg='#0a3d62',
                activebackground='#0a3d62',
                command=lambda t=tkt, v=is_active, w=row_widgets: self._toggle_ticket_row(t, v, w)
            )
            checkbox.grid(row=i, column=0, padx=5, pady=3)
            
            # Store in ticket_entries
            self.ticket_entries[tkt] = (credit_name, loan_name, bank_name, is_active, has_pol)
    
    def _toggle_ticket_row(self, tkt, is_active_var, widgets):
        """
        Toggles visual state of ticket row based on checkbox.
        
        Args:
            tkt: Ticket number
            is_active_var: BooleanVar for active state
            widgets: List of widgets in the row
        """
        if is_active_var.get():
            # Active state - normal colors
            for widget in widgets:
                if isinstance(widget, tk.Label):
                    widget.config(fg='white', bg='#0a3d62')
                elif isinstance(widget, tk.Entry):
                    widget.config(state='normal', bg='white')
        else:
            # Inactive state - grayed out
            for widget in widgets:
                if isinstance(widget, tk.Label):
                    widget.config(fg='#808080', bg='#0a3d62')
                elif isinstance(widget, tk.Entry):
                    widget.config(state='disabled', bg='#d3d3d3')
    
    def _on_run(self):
        """Executes the main processing workflow."""
        folder_path = self.folder_var.get().strip()
        excel_path = self.excel_var.get().strip()
        
        # Validate inputs
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("Folder Error", "Please select a valid PDF folder.")
            return
        
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Excel Error", "Please select a valid Excel file.")
            return
        
        # Extract only active tickets
        filled_ticket_entries = {}
        for tkt, (credit_name_var, loan_name_var, bank_name_var, is_active_var, has_pol_var) in self.ticket_entries.items():
            if is_active_var.get():  # Only include active tickets
                filled_ticket_entries[tkt] = (
                    credit_name_var.get().strip(),
                    loan_name_var.get().strip(),
                    bank_name_var.get().strip(),
                    has_pol_var.get()
                )
        
        # Disable button during processing
        self.run_button.config(state=tk.DISABLED)
        
        try:
            self._process_workflow(folder_path, excel_path, filled_ticket_entries)
            self._update_bank_mapping(excel_path, filled_ticket_entries)
            self._update_credit_mapping(excel_path, filled_ticket_entries)
        finally:
            self.run_button.config(state=tk.NORMAL)
    
    def _process_workflow(self, folder_path: str, excel_path: str, ticket_entries: dict):
        """
        Executes the complete workflow: Excel read, validation, rename, email, organize.
        
        Args:
            folder_path: Path to PDF folder
            excel_path: Path to Excel file
            ticket_entries: Dictionary of active tickets with their details
        """
        # Step 1: Read Excel
        excel_processor = ExcelProcessor(excel_path)
        if not excel_processor.read_excel():
            return
        
        # Step 2: Validate tickets
        for tkt, entries in ticket_entries.items():
            credit_name, loan_name, bank_name, has_pol = entries
            if not credit_name or not loan_name or not bank_name:
                messagebox.showerror(
                    "Input Error", 
                    f"For active ticket {tkt}, credit_name, loan_name, and bank_name cannot be empty!"
                )
                return
        
        # Step 3: File validation and renaming
        file_manager = FileManager(folder_path)
        file_manager.scan_pdf_files()
        
        # Phase 1: Validate all files (no renaming yet)
        if not file_manager.validate_and_plan_renames(excel_processor, ticket_entries):
            return  # Validation failed, errors already shown
        
        # Phase 2: Execute renames (validation passed)
        renamed_by_ticket = file_manager.execute_renames()
        
        # Step 4: Create email drafts
        email_service = EmailService(folder_path)
        if email_service.is_available():
            email_service.create_email_drafts(file_manager.files_by_ticket_clo)
        
        # Step 5: Organize files into folders
        file_manager.organize_files_into_folders()
        
        # Step 6: Show summary
        self._show_summary(
            folder_path,
            excel_path,
            len(renamed_by_ticket),
            email_service.created_drafts,
            len(file_manager.organized_folders),
            len(file_manager.skipped_files),
            len(email_service.draft_errors),
            len(file_manager.organization_errors)
        )
        
        # Print detailed logs to console
        print("=== RENAMED BY TICKET ===")
        print(renamed_by_ticket)
        print("\n=== FILES BY TICKET & CLO ===")
        print(file_manager.files_by_ticket_clo)
        
        if file_manager.skipped_files:
            print("\n=== SKIPPED FILES ===")
            print(file_manager.skipped_files)
        
        if email_service.draft_errors:
            print("\n=== DRAFT ERRORS ===")
            print(email_service.draft_errors)
        
        if file_manager.organized_folders:
            print("\n=== ORGANIZED FOLDERS ===")
            print(file_manager.organized_folders)
        
        if file_manager.organization_errors:
            print("\n=== ORGANIZATION ERRORS ===")
            print(file_manager.organization_errors)
    
    def _update_bank_mapping(self, excel_path: str, ticket_entries: dict):
        """
        Updates bank name mapping based on current ticket entries.
        
        Args:
            excel_path: Path to Excel file
            ticket_entries: Dictionary of tickets with their details
        """
        excel_processor = ExcelProcessor(excel_path)
        if not excel_processor.read_excel():
            return
        
        mapping_updated = False
        for tkt, (credit_name, loan_name, bank_name, has_pol) in ticket_entries.items():
            if not bank_name:  # Skip empty bank names
                continue
            
            trade_info = excel_processor.get_trade_info_for_ticket(tkt)
            if not trade_info:
                continue
            
            trade_action = trade_info.get('trade_action', '')
            
            # Determine mapping key based on trade action
            if trade_action == 'BUY':
                mapping_key = trade_info.get('trade_seller', '')
            elif trade_action == 'SELL':
                mapping_key = trade_info.get('trade_buyer', '')
            else:
                continue
            
            if mapping_key:
                # Update mapping if not exists or different
                if mapping_key not in self.bank_mapping or self.bank_mapping[mapping_key] != bank_name:
                    self.bank_mapping[mapping_key] = bank_name
                    mapping_updated = True
                    print(f"Mapping updated: {mapping_key} -> {bank_name}")
        
        if mapping_updated:
            self._save_bank_mapping()
    
    def _update_credit_mapping(self, excel_path: str, ticket_entries: dict):
        """
        Updates credit name mapping based on current ticket entries.
        
        Args:
            excel_path: Path to Excel file
            ticket_entries: Dictionary of tickets with their details
        """
        excel_processor = ExcelProcessor(excel_path)
        if not excel_processor.read_excel():
            return
        
        mapping_updated = False
        for tkt, (credit_name, loan_name, bank_name, has_pol) in ticket_entries.items():
            if not credit_name:  # Skip empty credit names
                continue
            
            # Get credit name from Excel
            excel_credit_name = excel_processor.get_credit_name_for_ticket(tkt)
            if not excel_credit_name:
                continue
            
            # Update mapping if not exists or different
            if excel_credit_name not in self.credit_mapping or self.credit_mapping[excel_credit_name] != credit_name:
                self.credit_mapping[excel_credit_name] = credit_name
                mapping_updated = True
                print(f"Credit mapping updated: {excel_credit_name} -> {credit_name}")
        
        if mapping_updated:
            self._save_credit_mapping()
    
    def _show_summary(
        self, 
        folder_path: str, 
        excel_path: str,
        renamed_count: int,
        drafts_count: int,
        organized_count: int,
        skipped_count: int,
        draft_errors_count: int,
        org_errors_count: int
    ):
        """
        Shows summary message box with processing results.
        
        Args:
            folder_path: Processed folder path
            excel_path: Excel file path
            renamed_count: Number of renamed tickets
            drafts_count: Number of email drafts created
            organized_count: Number of organized folders
            skipped_count: Number of skipped files
            draft_errors_count: Number of draft errors
            org_errors_count: Number of organization errors
        """
        summary_lines = [
            f"Processed folder: {folder_path}",
            f"Excel: {excel_path}",
            f"Renamed tickets: {renamed_count}",
            f"Created email drafts: {drafts_count}",
            f"Organized folders: {organized_count}"
        ]
        
        if skipped_count:
            summary_lines.append(f"Skipped files: {skipped_count}")
        
        if draft_errors_count:
            summary_lines.append(f"Draft errors: {draft_errors_count} (see console)")
        
        if org_errors_count:
            summary_lines.append(f"Organization errors: {org_errors_count} (see console)")
        
        messagebox.showinfo("Operation Complete", "\n".join(summary_lines))
    
    def run(self):
        """Starts the GUI main loop."""
        self.root.mainloop()
