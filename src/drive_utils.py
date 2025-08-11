# src/drive_utils.py
import os
import io
import json
from typing import Optional

try:
    from pydrive2.auth import GoogleAuth
    from pydrive2.drive import GoogleDrive
    _HAS_DRIVE = True
except Exception:
    GoogleDrive = None  # type: ignore
    _HAS_DRIVE = False


def init_drive() -> Optional["GoogleDrive"]:
    """
    Initialize Google Drive client using a service account.
    Expects env var SERVICE_ACCOUNT_JSON to contain the full JSON string.
    Returns GoogleDrive instance or None if not configured.
    """
    if not _HAS_DRIVE:
        return None

    svc_json = os.environ.get("SERVICE_ACCOUNT_JSON")
    if not svc_json:
        return None

    os.makedirs(".secrets", exist_ok=True)
    svc_path = os.path.join(".secrets", "service_account.json")
    with open(svc_path, "w", encoding="utf-8") as f:
        f.write(svc_json)

    gauth = GoogleAuth()
    try:
        # Preferred pydrive2 helper for service accounts
        gauth.LoadServiceAccountCredentials(svc_path)
    except Exception:
        # Fallback path
        gauth.settings.update({
            "client_config_backend": "service",
            "service_config": {
                "client_json_file_path": svc_path,
                "client_user_email": json.loads(svc_json).get("client_email", ""),
            },
            "oauth_scope": ["https://www.googleapis.com/auth/drive"],
        })
        gauth.ServiceAuth()

    return GoogleDrive(gauth)


def drive_upload_bytes(drive: "GoogleDrive", folder_id: str, filename: str, data: bytes) -> str:
    """
    Upload raw bytes as a file to Google Drive (into the given folder).
    Returns the created file ID.
    """
    file = drive.CreateFile({"title": filename, "parents": [{"id": folder_id}]})
    file.content = io.BytesIO(data)
    file.Upload()
    return file["id"]
