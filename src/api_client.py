# src/api_client.py
import os
import requests
from typing import Dict, Any

BACKEND_URL = os.environ.get("BACKEND_URL", "http://localhost:8080").rstrip("/")
BACKEND_TOKEN = os.environ.get("BACKEND_TOKEN", "dev-token")

DEFAULT_TIMEOUT = int(os.environ.get("BACKEND_TIMEOUT_SEC", "1800"))  # 30 min

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
        snippet = r.text[:1000]
        raise RuntimeError(f"Backend error {r.status_code} at {endpoint}: {snippet}")
    return r.json()
