# CLO Trade Settlement Automation

A professional, modular Python application for automating CLO trade settlement documentation workflows.

## 📁 Project Structure

```
Settlement Code/
│
├── main.py                    # Entry point - Run this file
├── config.py                  # Configuration & constants
├── validators.py              # Validation functions
├── utils.py                   # Utility functions
├── excel_processor.py         # Excel file processing
├── file_manager.py           # PDF file management
├── email_service.py          # Outlook email automation
├── gui.py                    # GUI interface
│
└── code_loan_operations.py   # Original monolithic file (backup)
```

## 🎯 Features

- **Excel Integration**: Reads trade data from Excel files with validation
- **PDF Renaming**: Automatically renames PDFs using standardized CLO naming conventions
- **Email Automation**: Creates Outlook draft emails with proper recipients and attachments
- **File Organization**: Organizes processed files into structured folders
- **Validation System**: Comprehensive validation before any file operations
- **User-Friendly GUI**: Tkinter-based interface with ticket management

## 🚀 Getting Started

### Prerequisites

- Python 3.7+
- Required packages:
  ```bash
  pip install pandas openpyxl pywin32
  ```

### Running the Application

Simply run the main entry point:

```bash
python main.py
```

## 📦 Module Descriptions

### `config.py`
- Excel column name mappings
- CLO email recipient mappings
- System constants and limits
- Email signature configuration

### `validators.py`
- Email address validation
- Filename and path length validation
- Data integrity checks
- CLO name mapping validation

### `utils.py`
- String sanitization and HTML escaping
- Date formatting
- Ticket number cleaning
- Filename building
- Prefix type detection

### `excel_processor.py`
- Excel file reading and validation
- Subtrade ticket lookup (O(1) performance)
- Row data extraction
- Unique ticket retrieval

### `file_manager.py`
- PDF file scanning and validation
- Two-phase rename process (validate then execute)
- File organization into folders
- Comprehensive error handling

### `email_service.py`
- Outlook integration
- Email draft creation (no automatic sending)
- Attachment management
- HTML email body generation

### `gui.py`
- Tkinter-based user interface
- Ticket table with checkbox activation
- Folder and file selection
- Workflow orchestration

### `main.py`
- Application entry point
- GUI initialization

## 🔒 Safety Features

1. **Two-Phase Processing**: 
   - Phase 1: Validate all files without making changes
   - Phase 2: Execute operations only if validation passes

2. **Comprehensive Validation**:
   - Filename length checks (Windows limits)
   - Path length validation
   - CLO name verification
   - Email address validation
   - Data integrity checks

3. **No Automatic Email Sending**:
   - All emails are saved as DRAFTS only
   - User must manually review and send

4. **Error Handling**:
   - Detailed error messages
   - Console logging for debugging
   - Graceful failure recovery

## 💡 Usage Tips

1. **Select Folders**: Choose the folder containing your PDF files
2. **Load Excel**: Select the Excel file with trade data
3. **Configure Tickets**: 
   - Check/uncheck tickets to activate/deactivate
   - Fill in Credit Name, Loan Name, and Bank Name for active tickets
4. **Run**: Click the RUN button to process

## 🔧 Configuration

To modify email recipients or add new CLOs, edit the `CLO_EMAIL_MAPPING` dictionary in `config.py`.

To change Excel column names, update `EXCEL_COLUMN_NAMES` in `config.py`.

## 📝 Notes

- All file operations are validated before execution
- Console output provides detailed processing logs
- Skipped files are logged but don't stop processing
- Original file is preserved as `code_loan_operations.py`

## 👤 Author

[-]

## 📄 Version

2.0 (Modularized Architecture)
