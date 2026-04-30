"""
Configuration file for the Daily Outbound Report pipeline.
All hardcoded values (servers, paths, search criteria) are managed here.

Copy this file to config.py and fill in your environment-specific values.
"""
from pathlib import Path
from utils.date_utils import today_str, recent_dates_str

# ============================================================
# Paths
# ============================================================
DESKTOP_PATH  = Path.home() / 'Desktop'
_BASE_DATA    = DESKTOP_PATH / 'PROJECT' / 'Outbound' / 'Data'
_BASE_OUTPUT  = DESKTOP_PATH / 'PROJECT' / 'Outbound' / 'Output'
DATA_DIR      = _BASE_DATA   / today_str('f')
OUTPUT_DIR    = _BASE_OUTPUT / today_str('f')
TEMPLATE_PATH = _BASE_DATA   / 'template.xlsx'

# Network drive .env — holds POP3/SMTP credentials
ENV_BASE = Path(r'\\<network-drive-ip>\<share>\<path-to-env-folder>')
ENV_PATH = ENV_BASE / '.env'

MASTER_FILE_PREFIX = 'template'

# ============================================================
# POP3 (Inbound) Server
# ============================================================
POP3_SERVER   = 'pop3.example.com'
POP3_PORT     = 995
POP3_USER_ENV = 'POP3_USER'
POP3_PASS_ENV = 'POP3_PASS'

# Max number of emails to scan (newest first)
EMAIL_SEARCH_COUNT = 1000

# ============================================================
# SMTP (Outbound) Server
# ============================================================
SMTP_SERVER   = 'smtp.example.com'
SMTP_PORT     = 25
SMTP_USER_ENV = 'SMTP_USER'
SMTP_PASS_ENV = 'SMTP_PASS'

# ============================================================
# Email Targets — which attachments to download
# ============================================================
EMAIL_TARGETS = [
    {
        "name": "Provider C Report",
        "sender": "sender@provider-c.com",
        "contains": ["Outbound Report - Provider C"],
    },
    {
        "name": "Provider A Daily Exception List",
        "sender": None,
        "contains": ["Daily Exception List - Provider A"],
    },
    {
        "name": "Provider B Daily Exception List",
        "contains": ["Daily Exception List - Provider B"],
    },
    {
        "name": "Provider E Daily Manifest",
        "contains": ["Daily Manifest Report - Provider E"],
    },
    {
        "name": "Provider D Daily Report",
        "contains": ["Daily Outbound Report - Provider D"],
    }
]

EXPECTED_ATTACHMENT_COUNT = len(EMAIL_TARGETS)

# Max number of daily folders to keep in Data/ and Output/
MAX_DAILY_FOLDERS = 5

# ============================================================
# Failure Alert Email
# ============================================================
NOTIFICATION_FROM = 'sender@example.com'
NOTIFICATION_TO   = ['recipient@example.com']
NOTIFICATION_CC   = []
FAILURE_SUBJECT   = 'Daily Report - Files missing'

# ============================================================
# Report Email
# ============================================================
REPORT_SUBJECT = 'Consolidated Manifest & Exception Report – {}'  # .format(date)
REPORT_TO = ['recipient@example.com']
REPORT_CC = []
