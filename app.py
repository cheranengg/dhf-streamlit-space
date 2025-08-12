# app.py — DHF Streamlit UI (Preview + Single ZIP Download)
# ---------------------------------------------------------
# - Input: Product Requirements (Excel upload OR pasted CSV)
# - Calls FastAPI backend (HA → DVP → Trace Matrix) with bearer token
# - Previews only (no editing)
# - Exports: Hazard_Analysis.xlsx, Design_Verification_Protocol.xlsx, Trace_Matrix.xlsx
# - Packs all 3 into one ZIP for a single-click download
# - (Optional) Uploads to Google Drive if SERVICE_ACCOUNT_JSON + DRIVE_FOLDER_ID are set

import os
import io
import json
import zipfile
import typing as t
import pandas as pd
import streamlit as st
import requests

# ---------------- Constants & Secrets ----------------
TBD = "TBD - Human / SME input"
BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8080").rstrip("/")
BACKEND_TOKEN = st.secrets.get("BACKEND_TOKEN", "dev-token")
DEFAULT_DRIVE_FOLDER_ID = st.secrets.get("DRIVE_FOLDER_ID", "")
OUTPUT_DIR = st.secrets.get("OUTPUT_DIR", os.path.abspath("./streamlit_outputs"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --------------- Google Drive (pydrive2) ---------------
_HAS_DRIVE = True
try:
    from pydrive2.auth import GoogleAuth
    from pydrive2.drive import GoogleDrive
except Exception:
    _HAS_DRIVE = False


def init_drive() -> t.Optional["GoogleDrive"]:
    if not _HAS_DRIVE:
        return None
    svc_json = st.secrets.get("SERVICE_ACCOUNT_JSON")
    if not svc_json:
        return None
    os.makedirs(".secrets", exist_ok=True)
    svc_path = os.path.join(".secrets", "service_account.json")
    with open(svc_path, "w", encoding="utf-8") as f:
        f.write(svc_json)
    gauth = GoogleAuth()
    try:
        gauth.LoadServiceAccountCredentials(svc_path)
    except Exception:
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
    file = drive.CreateFile({"title": filename, "parents": [{"id": folder_id}]})
    file.content = io.BytesIO(data)
    file.Upload()
    return file["id"]


# ---------------- Helpers ----------------
def normalize_requirements(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for c in df.columns:
        lc = c.strip().lower()
        if lc in {"requirement id", "req id", "requirement_id"}:
            rename_map[c] = "Requirement ID"
        elif lc in {"verification id", "verification_id", "verif id"}:
            rename_map[c] = "Verification ID"
        elif lc in {"requirement", "requirements", "requirement text", "requirement_desc", "requirement description"}:
            rename_map[c] = "Requirements"
    df = df.rename(columns=rename_map)
    for col in ["Requirement ID", "Verification ID", "Requirements"]:
        if col not in df.columns:
            df[col] = None
    return df[["Requirement ID", "Verification ID", "Requirements"]].copy()


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.read()


def fill_tbd(df: pd.DataFrame, cols: t.List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = TBD
        out[c] = out[c].fillna(TBD)
        out.loc[out[c].astype(str).str.strip().eq(""), c] = TBD
    return out


def call_backend(endpoint: str, payload: dict) -> dict:
    url = f"{BACKEND_URL}{endpoint}"
    headers = {"Authorization": f"Bearer {BACKEND_TOKEN}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json=payload, timeout=1200)
    if r.status_code >= 400:
        raise RuntimeError(f"Backend error {r.status_code}: {r.text[:500]}")
    return r.json()


# ---------------- UI ----------------
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")
st.title("🧩 DHF Automation – Infusion Pump")
st.caption("Requirements → Hazard Analysis → DVP → Trace Matrix | Guardrails + Preview | Export to Excel (+ optional Google Drive)")

st.markdown("**Provide Product Requirements** (choose one):")
colA, colB = st.columns(2)
with colA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with colB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,The pump shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=140, placeholder=sample)

run_btn = st.button("▶️ Generate DHF Package", type="primary")

if run_btn:
    # -------- Parse Requirements --------
    with st.spinner("Parsing requirements..."):
        if uploaded is not None:
            try:
                req_df = pd.read_excel(uploaded)
            except Exception as e:
                st.error(f"Failed to read Excel: {e}")
                st.stop()
        elif pasted.strip():
            try:
                req_df = pd.read_csv(io.StringIO(pasted))
            except Exception as e:
                st.error(f"Failed to parse pasted CSV: {e}")
                st.stop()
        else:
            st.error("Please upload an Excel file or paste CSV text.")
            st.stop()
        req_df = normalize_requirements(req_df)

    st.success(f"Loaded {len(req_df)} requirements.")
    st.dataframe(req_df.head(20), use_container_width=True)

    # -------- Backend: Hazard Analysis --------
    with st.spinner("Running Hazard Analysis (backend)..."):
        ha_payload = {
            "requirements": [
                {
                    "Requirement ID": r["Requirement ID"],
                    "Verification ID": r.get("Verification ID"),
                    "Requirements": r.get("Requirements"),
                }
                for _, r in req_df.iterrows()
            ]
        }
        ha_resp = call_backend("/hazard-analysis", ha_payload)
        ha_rows = ha_resp.get("ha", [])
        ha_df = pd.DataFrame(ha_rows)

    ha_df = fill_tbd(
        ha_df,
        [
            "requirement_id", "risk_id", "risk_to_health", "hazard", "hazardous_situation",
            "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
        ],
    )
    st.subheader("Hazard Analysis (preview)")
    st.dataframe(ha_df.head(20), use_container_width=True)

    # -------- Backend: DVP --------
    with st.spinner("Generating Design Verification Protocol (backend)..."):
        dvp_payload = {
            "requirements": ha_payload["requirements"],
            "ha": ha_rows,
        }
        dvp_resp = call_backend("/dvp", dvp_payload)
        dvp_rows = dvp_resp.get("dvp", [])
        dvp_df = pd.DataFrame(dvp_rows)

    dvp_df = fill_tbd(dvp_df, ["verification_id", "verification_method", "acceptance_criteria", "sample_size", "test_procedure"])
    st.subheader("Design Verification Protocol (preview)")
    st.dataframe(dvp_df.head(20), use_container_width=True)

    # -------- Backend: Trace Matrix --------
    with st.spinner("Building Trace Matrix (backend)..."):
        tm_payload = {
            "requirements": ha_payload["requirements"],
            "ha": ha_rows,
            "dvp": dvp_rows,
        }
        tm_resp = call_backend("/trace-matrix", tm_payload)
        tm_rows = tm_resp.get("trace_matrix", [])
        tm_df = pd.DataFrame(tm_rows)

    tm_required_cols = [
        "verification_id", "requirement_id", "requirements",
        "risk_ids", "risks_to_health", "ha_risk_controls", "verification_method", "acceptance_criteria",
    ]
    tm_df = fill_tbd(tm_df, tm_required_cols)

    st.subheader("Trace Matrix (preview)")
    st.dataframe(tm_df.head(20), use_container_width=True)

    # -------- Exports (Create 3 Excel files) --------
    st.subheader("Export")

    ha_export_cols = [
        "requirement_id","risk_id","risk_to_health","hazard","hazardous_situation",
        "harm","sequence_of_events","severity_of_harm","p0","p1","poh","risk_index","risk_control",
    ]
    dvp_export_cols = ["verification_id","verification_method","acceptance_criteria","sample_size","test_procedure","requirements","requirement_id"]
    tm_export_cols = [
        "verification_id","requirement_id","requirements",
        "risk_ids","risks_to_health","ha_risk_controls","verification_method","acceptance_criteria",
    ]

    ha_bytes = df_to_excel_bytes(ha_df[ha_export_cols])
    dvp_bytes = df_to_excel_bytes(dvp_df[dvp_export_cols])
    tm_bytes  = df_to_excel_bytes(tm_df[tm_export_cols])

    # Save locally (optional)
    with open(os.path.join(OUTPUT_DIR, "Hazard_Analysis.xlsx"), "wb") as f:
        f.write(ha_bytes)
    with open(os.path.join(OUTPUT_DIR, "Design_Verification_Protocol.xlsx"), "wb") as f:
        f.write(dvp_bytes)
    with open(os.path.join(OUTPUT_DIR, "Trace_Matrix.xlsx"), "wb") as f:
        f.write(tm_bytes)

    # Build a single ZIP for one-click download
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Hazard_Analysis.xlsx", ha_bytes)
        zf.writestr("Design_Verification_Protocol.xlsx", dvp_bytes)
        zf.writestr("Trace_Matrix.xlsx", tm_bytes)
    zip_buf.seek(0)
    zip_bytes = zip_buf.read()

    clicked = st.download_button(
        "⬇️ Download DHF package (.zip)",
        data=zip_bytes,
        file_name="DHF_Package.zip",
        mime="application/zip",
        type="primary",
    )
    if clicked:
        st.success("Hazard Analysis, Design Verification Protocol, Trace Matrix documents downloaded successfully.")

    # -------- Optional: Upload to Google Drive --------
    if DEFAULT_DRIVE_FOLDER_ID:
        with st.spinner("Uploading to Google Drive..."):
            drive = init_drive()
            if drive:
                try:
                    ha_id = drive_upload_bytes(drive, DEFAULT_DRIVE_FOLDER_ID, "Hazard_Analysis.xlsx", ha_bytes)
                    dvp_id = drive_upload_bytes(drive, DEFAULT_DRIVE_FOLDER_ID, "Design_Verification_Protocol.xlsx", dvp_bytes)
                    tm_id  = drive_upload_bytes(drive, DEFAULT_DRIVE_FOLDER_ID, "Trace_Matrix.xlsx", tm_bytes)
                    st.info({"Hazard_Analysis.xlsx": ha_id,
                             "Design_Verification_Protocol.xlsx": dvp_id,
                             "Trace_Matrix.xlsx": tm_id})
                except Exception as e:
                    st.warning(f"Drive upload failed: {e}")
            else:
                st.info("Drive not initialized (missing SERVICE_ACCOUNT_JSON secret?). Files saved locally only.")

st.markdown("---")
st.markdown(
    f"""
**Files & Hosting**
- Outputs saved locally to: `{OUTPUT_DIR}`.
- Backend: `{BACKEND_URL}` (bearer-auth).
- Guardrails fill any missing fields with `{TBD}`.
"""
)
