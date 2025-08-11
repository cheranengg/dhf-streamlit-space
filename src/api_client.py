# src/api_client.py
import os
from typing import Dict, Any

import requests

try:
    import streamlit as st
    _HAS_ST = True
except Exception:
    _HAS_ST = False

def _get_secret(key: str, default: str = "") -> str:
    if _HAS_ST:
        try:
            return st.secrets.get(key, default)
        except Exception:
            pass
    return os.environ.get(key, default)

BACKEND_URL = _get_secret("BACKEND_URL", "http://localhost:8080").rstrip("/")
BACKEND_TOKEN = _get_secret("BACKEND_TOKEN", "dev-token")
DEFAULT_TIMEOUT = int(_get_secret("BACKEND_TIMEOUT_SEC", "1800"))  # 30 min

def call_backend(endpoint: str, payload: Dict[str, Any], timeout_sec: int = DEFAULT_TIMEOUT) -> Dict[str, Any]:
    """
    POST JSON to the FastAPI backend with bearer auth.
    Raises RuntimeError on non-2xx.
    """
    url = f"{BACKEND_URL}{endpoint}"
    headers = {
        "Authorization": f"Bearer {BACKEND_TOKEN}",
        "Content-Type": "application/json",
    }
    r = requests.post(url, headers=headers, json=payload, timeout=timeout_sec)
    if r.status_code >= 400:
        raise RuntimeError(f"Backend error {r.status_code} at {endpoint}: {r.text[:1000]}")
    return r.json()
