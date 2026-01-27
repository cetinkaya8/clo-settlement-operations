"""
Excel Processor Module - CLO Trade Settlement Automation

Handles reading and processing Excel files containing trade data.
"""

import pandas as pd
from tkinter import messagebox
from config import EXCEL_COLUMN_NAMES
from utils import clean_ticket


class ExcelProcessor:
    """Processes Excel files containing trade settlement data."""
    
    def __init__(self, excel_path: str):
        """
        Initialize Excel processor.
        
        Args:
            excel_path: Path to Excel file
        """
        self.excel_path = excel_path
        self.df = None
        self.subtrade_ticket_lookup = {}
    
    def read_excel(self) -> bool:
        """
        Reads Excel file and validates columns.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            self.df = pd.read_excel(self.excel_path, dtype=str, header=0).fillna('')
        except Exception as e:
            messagebox.showerror("Excel Read Error", f"Excel file could not be read:\n{e}")
            return False
        
        # Check Excel column names
        missing_cols = [
            col_name for col_name in EXCEL_COLUMN_NAMES.values() 
            if col_name not in self.df.columns
        ]
        
        if missing_cols:
            messagebox.showerror(
                "Excel Column Error",
                "Some expected columns were not found in Excel file.\n\n"
                f"Expected columns: {list(EXCEL_COLUMN_NAMES.values())}\n"
                f"Missing columns: {missing_cols}\n\n"
                f"Existing columns: {list(self.df.columns)}"
            )
            return False
        
        self._build_subtrade_lookup()
        return True
    
    def _build_subtrade_lookup(self):
        """
        Builds O(1) lookup dictionary for subtrade tickets.
        
        Key: cleaned subtrade_ticket
        Value: row index
        """
        self.subtrade_ticket_lookup = {}
        for idx in range(len(self.df)):
            subtrade_ticket_raw = self.df.at[idx, EXCEL_COLUMN_NAMES['subtrade_ticket']]
            cleaned_key = clean_ticket(subtrade_ticket_raw)
            if cleaned_key:
                self.subtrade_ticket_lookup[cleaned_key] = idx
    
    def get_row_by_subtrade_ticket(self, subtrade_ticket: str):
        """
        Retrieves row index for a given subtrade ticket.
        
        Args:
            subtrade_ticket: Cleaned subtrade ticket number
            
        Returns:
            Row index or None if not found
        """
        return self.subtrade_ticket_lookup.get(subtrade_ticket)
    
    def get_row_data(self, row_idx: int) -> dict:
        """
        Retrieves all relevant data from a specific row.
        
        Args:
            row_idx: Row index
            
        Returns:
            Dictionary with trade data
        """
        if row_idx is None or row_idx >= len(self.df):
            return {}
        
        return {
            'buy_sell_action': str(self.df.at[row_idx, EXCEL_COLUMN_NAMES['buy_sell_action']]).upper(),
            'subtrade_seller': self.df.at[row_idx, EXCEL_COLUMN_NAMES['subtrade_seller']],
            'subtrade_buyer': self.df.at[row_idx, EXCEL_COLUMN_NAMES['subtrade_buyer']],
            'primary_secondary_type': self.df.at[row_idx, EXCEL_COLUMN_NAMES['primary_secondary_type']],
            'ticket': self.df.at[row_idx, EXCEL_COLUMN_NAMES['ticket']],
            'trade_date': self.df.at[row_idx, EXCEL_COLUMN_NAMES['trade_date']],
            'subtrade_ticket': self.df.at[row_idx, EXCEL_COLUMN_NAMES['subtrade_ticket']]
        }
    
    def get_unique_tickets(self) -> list:
        """
        Gets list of unique ticket numbers from Excel.
        
        Returns:
            List of unique tickets
        """
        if self.df is None:
            return []
        
        ticket_col_name = EXCEL_COLUMN_NAMES.get('ticket', '')
        if not ticket_col_name or ticket_col_name not in self.df.columns:
            return []
        
        return self.df[ticket_col_name].unique().tolist()
    
    def get_credit_name_for_ticket(self, ticket: str) -> str:
        """
        Gets the first Credit Name associated with a ticket.
        
        Args:
            ticket: Ticket number
            
        Returns:
            Credit Name string or empty string if not found
        """
        if self.df is None:
            return ''
        
        ticket_col_name = EXCEL_COLUMN_NAMES.get('ticket', '')
        credit_col_name = EXCEL_COLUMN_NAMES.get('credit_name', '')
        
        if not ticket_col_name or not credit_col_name:
            return ''
        
        if ticket_col_name not in self.df.columns or credit_col_name not in self.df.columns:
            return ''
        
        # Find rows matching this ticket
        matching_rows = self.df[self.df[ticket_col_name] == ticket]
        
        if matching_rows.empty:
            return ''
        
        # Return the first Credit Name found
        credit_name = str(matching_rows.iloc[0][credit_col_name]).strip()
        return credit_name if credit_name != 'nan' and credit_name else ''
    
    def get_trade_info_for_ticket(self, ticket: str) -> dict:
        """
        Gets Trade Action and Seller/Buyer information for a ticket.
        
        Args:
            ticket: Ticket number
            
        Returns:
            Dictionary with 'trade_action', 'trade_seller', 'trade_buyer'
        """
        if self.df is None:
            return {}
        
        ticket_col_name = EXCEL_COLUMN_NAMES.get('ticket', '')
        action_col_name = EXCEL_COLUMN_NAMES.get('buy_sell_action', '')
        seller_col_name = EXCEL_COLUMN_NAMES.get('subtrade_seller', '')
        buyer_col_name = EXCEL_COLUMN_NAMES.get('subtrade_buyer', '')
        
        if not all([ticket_col_name, action_col_name, seller_col_name, buyer_col_name]):
            return {}
        
        required_cols = [ticket_col_name, action_col_name, seller_col_name, buyer_col_name]
        if not all(col in self.df.columns for col in required_cols):
            return {}
        
        # Find rows matching this ticket
        matching_rows = self.df[self.df[ticket_col_name] == ticket]
        
        if matching_rows.empty:
            return {}
        
        # Return the first match
        first_row = matching_rows.iloc[0]
        return {
            'trade_action': str(first_row[action_col_name]).strip().upper(),
            'trade_seller': str(first_row[seller_col_name]).strip(),
            'trade_buyer': str(first_row[buyer_col_name]).strip()
        }
