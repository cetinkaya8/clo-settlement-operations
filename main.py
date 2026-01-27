"""
Main Entry Point - CLO Trade Settlement Automation

This application automates the CLO trade settlement documentation workflow by:
- Reading trade data from Excel files
- Matching and renaming PDF documents according to standardized conventions
- Creating Outlook email drafts with proper recipients and attachments
- Organizing processed files into structured folders

Version: 2.0 (Modularized)
"""

import tkinter as tk
from gui import SettlementGUI


def main():
    """Main entry point for the application."""
    root = tk.Tk()
    app = SettlementGUI(root)
    app.run()


if __name__ == "__main__":
    main()
