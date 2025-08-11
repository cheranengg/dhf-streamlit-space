# src/__init__.py
"""
DHF Automation – Streamlit UI Helper Modules
--------------------------------------------
Contains:
- api_client:   Backend API caller
- drive_utils:  Google Drive (pydrive2) helpers
- ui_helpers:   DataFrame formatting, guardrails, Excel conversion

Usage:
    from src.api_client import call_backend
    from src.drive_utils import init_drive, drive_upload_bytes
    from src.ui_helpers import normalize_requirements, fill_tbd, basic_guardrails_df
"""

__version__ = "1.0.0"
