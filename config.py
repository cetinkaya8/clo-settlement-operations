"""
Configuration Module - CLO Trade Settlement Automation

Contains all configuration constants, email mappings, and Excel column definitions.
"""

# Excel column name mappings
EXCEL_COLUMN_NAMES = {
    'ticket': 'Ticket #',
    'subtrade_ticket': 'Subtrade Ticket #',
    'trade_date': 'Trade Date',
    'primary_secondary_type': 'Trade Type',
    'buy_sell_action': 'Trade Action',
    'subtrade_seller': 'Subtrade Seller',
    'subtrade_buyer': 'Subtrade Buyer',
    'credit_name': 'Credit Name'
}

# Default CC email address for all CLO communications
DEFAULT_CC = "user1@example.com; user2@example.com; user3@example.com"

# CLO Name to Email Recipient Mapping
CLO_EMAIL_MAPPING = {
    "Company CLO XVI, Ltd.": {
        "to": "cloxvi@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO XV, Ltd.": {
        "to": "cloxv@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO XR, Ltd.": {
        "to": "cloxr@trustee.example.com; collateral@trustee.example.com",
        "cc": "user1@example.com; user2@example.com; user3@example.com; contact1@trustee.example.com; contact2@trustee.example.com"
    },
    "Company CLO XIV, Ltd.": {
        "to": "cloxiv@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO XIII, Ltd.": {
        "to": "cloxiii@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO XII, Ltd.": {
        "to": "cloxii@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO XI, Ltd.": {
        "to": "cloxi@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO IX, Ltd.": {
        "to": "cloix@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO VIII, Ltd.": {
        "to": "cloviii@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO VII, Ltd.": {
        "to": "clovii@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO IV, Ltd.": {
        "to": "cloiv@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO IV-R, Ltd.": {
        "to": "cloivr@trustee.example.com; collateral@trustee.example.com",
        "cc": "user1@example.com; user2@example.com; user3@example.com; contact1@trustee.example.com; contact2@trustee.example.com"
    },
    "Company CLO III, Ltd.": {
        "to": "cloiii@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
    "Company CLO IX-R, Ltd.": {
        "to": "cloixr@trustee.example.com; collateral@trustee.example.com",
        "cc": "user1@example.com; user2@example.com; user3@example.com; contact1@trustee.example.com; contact2@trustee.example.com"
    },
    "Company CLO XVII, Ltd.": {
        "to": "cloxvii@trustee.example.com; collateral@trustee.example.com",
        "cc": DEFAULT_CC
    },
}

# Expected PDF file prefixes
PDF_PREFIXES = ('confirm_', 'aa_', 'fundingmemo_', 'funding_memo_', 'payoutletter_')

# Outlook configuration
OUTLOOK_EMAIL_ACCOUNT = "user@example.com"

# File system limits
MAX_FILENAME_LENGTH = 255
MAX_PATH_LENGTH = 260
SAFE_FILENAME_LENGTH = 250  # Leave room for .pdf extension

# Bank Name mapping file path
BANK_MAPPING_FILE = 'bank_name_mapping.json'

# Credit Name mapping file path
CREDIT_MAPPING_FILE = 'credit_name_mapping.json'

# Email signature configuration
EMAIL_SIGNATURE = {
    'name': 'Your Name',
    'email': 'user@example.com',
    'company': 'Your Company Name, Inc.',
    'address_line1': 'Your Address Line 1',
    'address_line2': 'Your City, State ZIP'
}
