# app.py — DHF Streamlit UI (Preview + ZIP Download)

import os
import io
import json
import zipfile
import typing as t
from collections import Counter

import pandas as pd
import streamlit as st
import requests

# =========================
# Config & constants
# =========================
TBD = "TBD - Human / SME input"

BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8080")
BACKEND_TOKEN = st.secrets.get("BACKEND_TOKEN", "dev-token")

# Row caps (preview + Excel export)
HA_MAX_ROWS = int(os.getenv("HA_MAX_ROWS", "50"))
DVP_MAX_ROWS = int(os.getenv("DVP_MAX_ROWS", "50"))
TM_MAX_ROWS  = int(os.getenv("TM_MAX_ROWS", "50"))

# How many requirements we send to backend
REQ_MAX = int(os.getenv("REQ_MAX", "50"))

# Optional Google Drive upload
DEFAULT_DRIVE_FOLDER_ID = st.secrets.get("DRIVE_FOLDER_ID", "")
OUTPUT_DIR = st.secrets.get("OUTPUT_DIR", os.path.abspath("./streamlit_outputs"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# =========================
# Style (green primary button, compact layout)
# =========================
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")
st.markdown("""
<style>
/* Make primary buttons green */
.stButton > button[kind="primary"] {
  background-color: #16a34a !important;  /* tailwind-green-600 */
  color: white !important;
  border: 1px solid #15803d !important;
}
.stButton > button[kind="primary"]:hover {
  background-color: #15803d !important;  /* green-700 */
}
/* Header spacing */
.block-container { padding-top: 1.2rem; }
</style>
""", unsafe_allow_html=True)

# =========================
# Image helpers
# =========================
ASSETS_DIR = "streamlit_assets"
INFUSION1_PATH = os.path.join(ASSETS_DIR, "Infusion1.jpg")  # bedside pump (first)
INFUSION_PATH  = os.path.join(ASSETS_DIR, "Infusion.jpg")   # lab image (second)

def try_image(path: str) -> t.Optional[bytes]:
    try:
        with open(path, "rb") as f:
            return f.read()
    except Exception:
        return None

# =========================
# Optional Google Drive (pydrive2)
# =========================
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
            'client_config_backend': 'service',
            'service_config': {
                'client_json_file_path': svc_path,
                'client_user_email': json.loads(svc_json).get('client_email', ''),
            },
            'oauth_scope': ['https://www.googleapis.com/auth/drive']
        })
        gauth.ServiceAuth()
    return GoogleDrive(gauth)

def drive_upload_bytes(drive: "GoogleDrive", folder_id: str, filename: str, data: bytes) -> str:
    file = drive.CreateFile({"title": filename, "parents": [{"id": folder_id}]})
    file.content = io.BytesIO(data)
    file.Upload()
    return file["id"]

# =========================
# Helpers
# =========================
def normalize_requirements(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for c in df.columns:
        lc = str(c).strip().lower()
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

def call_backend(endpoint: str, payload: dict) -> dict:
    url = f"{BACKEND_URL.rstrip('/')}{endpoint}"
    headers = {"Authorization": f"Bearer {BACKEND_TOKEN}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json=payload, timeout=1200)
    if r.status_code >= 400:
        raise RuntimeError(f"Backend error {r.status_code}: {r.text[:500]}")
    return r.json()

def head_cap(df: pd.DataFrame, n: int) -> pd.DataFrame:
    return df.head(n).copy() if isinstance(df, pd.DataFrame) and not df.empty else df

def fill_tbd(df: pd.DataFrame, cols: t.List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = TBD
        out[c] = out[c].fillna(TBD)
        out.loc[out[c].astype(str).str.strip().eq(""), c] = TBD
        out.loc[out[c].astype(str).str.upper().eq("NA"), c] = TBD
    return out

# =========================
# Excel formatting helpers (openpyxl)
# =========================
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

GREEN = "D1FAE5"  # soft green header
HEADER_FONT = Font(bold=True)
WRAP_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)

def df_to_excel_bytes_with_format(
    df: pd.DataFrame,
    col_width_default: int = 15,
    col_width_overrides: t.Optional[dict] = None,
    freeze_cell: str = "D2",
    row_height: int = 30
) -> bytes:
    wb = Workbook()
    ws: Worksheet = wb.active

    # Write DataFrame
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        ws.append(row)

    # Header style
    max_col = ws.max_column
    for c in range(1, max_col + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")

    # Column widths
    widths = {i: col_width_default for i in range(1, max_col + 1)}
    if col_width_overrides:
        for col_name, width in col_width_overrides.items():
            if col_name in df.columns:
                idx = df.columns.get_loc(col_name) + 1
                widths[idx] = width
    for i, w in widths.items():
        ws.column_dimensions[chr(64 + i)].width = w

    # Row heights + alignment
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = row_height
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).alignment = WRAP_ALIGN

    # Freeze panes
    ws.freeze_panes = freeze_cell

    # Save
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# =========================
# Metrics
# =========================
def compute_metrics(ha_df: pd.DataFrame, dvp_df: pd.DataFrame, tm_df: pd.DataFrame) -> dict:
    metrics = {}

    # HA
    ha_count = len(ha_df)
    sev_counts = Counter(ha_df.get("severity_of_harm", []))
    top_sev = max(sev_counts.items(), key=lambda x: x[1])[0] if sev_counts else "–"
    rth_unique = ha_df.get("risk_to_health", pd.Series(dtype=str)).nunique()
    metrics["HA"] = {
        "rows": ha_count,
        "unique_risk_to_health": int(rth_unique),
        "most_common_severity": str(top_sev)
    }

    # DVP
    non_tbd_method = (dvp_df.get("verification_method", "").astype(str).str.contains("TBD", case=False, na=False) == False).sum()
    non_tbd_accept = (dvp_df.get("acceptance_criteria", "").astype(str).str.contains("TBD", case=False, na=False) == False).sum()
    dvp_count = max(len(dvp_df), 1)
    metrics["DVP"] = {
        "rows": len(dvp_df),
        "method_filled_pct": round(100 * non_tbd_method / dvp_count, 1),
        "acceptance_filled_pct": round(100 * non_tbd_accept / dvp_count, 1),
    }

    # TM
    tm_count = len(tm_df)
    tm_cov_vid = (tm_df.get("Verification ID", "").astype(str).str.upper() != "TBD - HUMAN / SME INPUT").sum()
    tm_cov_risk = (tm_df.get("Risk ID", "").astype(str).str.upper() != "TBD - HUMAN / SME INPUT").sum()
    metrics["TM"] = {
        "rows": tm_count,
        "verification_linked_pct": round(100 * tm_cov_vid / max(tm_count, 1), 1),
        "risk_linked_pct": round(100 * tm_cov_risk / max(tm_count, 1), 1),
    }
    return metrics

def render_metrics(metrics: dict):
    st.subheader("Evaluation Metrics")
    ha = metrics.get("HA", {})
    dvp = metrics.get("DVP", {})
    tm  = metrics.get("TM", {})
    st.markdown(f"""
- **Hazard Analysis**: {ha.get('rows','–')} rows · {ha.get('unique_risk_to_health','–')} risk types · most common severity **{ha.get('most_common_severity','–')}**
- **Design Verification Protocol**: methods filled **{dvp.get('method_filled_pct','–')}%**, acceptance criteria filled **{dvp.get('acceptance_filled_pct','–')}%**
- **Trace Matrix**: verification linked **{tm.get('verification_linked_pct','–')}%**, risk linked **{tm.get('risk_linked_pct','–')}%**
""")
    st.caption("Notes: These simple coverage indicators help spot obvious gaps. Review high-severity risks and any TBD fields with an SME before sign-off.")

# =========================
# Page Header (with images on the right)
# =========================
left, right = st.columns([3, 2], vertical_alignment="center")
with left:
    st.title("🧩 DHF Automation – Infusion Pump")
    st.caption("Requirements → Hazard Analysis → DVP → TM")
with right:
    img1 = try_image(INFUSION1_PATH)  # bedside pump (first)
    img2 = try_image(INFUSION_PATH)   # lab image (second)
    r1, r2 = st.columns(2)
    if img1:
        r1.image(img1, caption="", use_column_width=True)
    if img2:
        r2.image(img2, caption="", use_column_width=True)

# =========================
# Inputs
# =========================
st.markdown("**Provide Product Requirements** (choose one):")
colA, colB = st.columns(2)
with colA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with colB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=140, placeholder=sample)

run_btn = st.button("▶️ Generate DHF Packages", type="primary")

# =========================
# Main flow
# =========================
if run_btn:
    # Parse Requirements
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

    total_reqs = len(req_df)
    st.success(f"Loaded {total_reqs} requirements.")
    st.dataframe(head_cap(req_df, 20), use_container_width=True)

    # Limit what we send to backend
    req_df_limited = req_df.head(REQ_MAX).copy()
    if total_reqs > REQ_MAX:
        st.info(f"Sending only the first {REQ_MAX} requirements to backend (set REQ_MAX env var to change).")

    # ---- HA
    with st.spinner("Running Hazard Analysis (backend)..."):
        ha_payload = {
            "requirements": [
                {
                    "Requirement ID": r["Requirement ID"],
                    "Verification ID": r.get("Verification ID"),
                    "Requirements": r.get("Requirements"),
                }
                for _, r in req_df_limited.iterrows()
            ]
        }
        ha_resp = call_backend("/hazard-analysis", ha_payload)
        ha_rows = ha_resp.get("ha", [])
        ha_df = pd.DataFrame(ha_rows)

    ha_df = fill_tbd(
        ha_df,
        [
            "risk_id", "risk_to_health", "hazard", "hazardous_situation",
            "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
        ],
    )
    st.subheader(f"Hazard Analysis (preview, first {HA_MAX_ROWS})")
    st.dataframe(head_cap(ha_df, HA_MAX_ROWS), use_container_width=True)

    # ---- DVP
    with st.spinner("Generating Design Verification Protocol (backend)..."):
        dvp_payload = {
            "requirements": ha_payload["requirements"],
            "ha": ha_rows,
        }
    try:
        dvp_resp = call_backend("/dvp", dvp_payload)
        dvp_rows = dvp_resp.get("dvp", [])
    except Exception as e:
        st.error(f"DVP backend error: {e}")
        st.stop()
    dvp_df = pd.DataFrame(dvp_rows)
    dvp_df = fill_tbd(dvp_df, ["verification_id", "verification_method", "acceptance_criteria", "sample_size", "test_procedure"])
    st.subheader(f"Design Verification Protocol (preview, first {DVP_MAX_ROWS})")
    st.dataframe(head_cap(dvp_df, DVP_MAX_ROWS), use_container_width=True)

    # ---- TM
    with st.spinner("Building Trace Matrix (backend)..."):
        tm_payload = {
            "requirements": ha_payload["requirements"],
            "ha": ha_rows,
            "dvp": dvp_rows,
        }
        tm_resp = call_backend("/tm", tm_payload)
        tm_rows = tm_resp.get("tm", [])
        tm_df = pd.DataFrame(tm_rows)

    TM_ORDER = [
        "Requirement ID",
        "Requirements",
        "Requirement (Yes/No)",
        "Risk ID",
        "Risk to Health",
        "HA Risk Control",
        "Verification ID",
        "Verification Method",
    ]
    tm_df = tm_df.reindex(columns=TM_ORDER)
    tm_df = fill_tbd(tm_df, TM_ORDER)

    st.subheader(f"Trace Matrix (preview, first {TM_MAX_ROWS})")
    st.dataframe(head_cap(tm_df, TM_MAX_ROWS), use_container_width=True)

    # ---- Metrics (instead of guardrails)
    metrics = compute_metrics(ha_df, dvp_df, tm_df)
    render_metrics(metrics)

    # ---- Exports with formatting
    ha_export_cols = [
        "risk_id", "risk_to_health", "hazard", "hazardous_situation",
        "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
    ]
    dvp_export_cols = ["verification_id", "verification_method", "acceptance_criteria", "sample_size", "test_procedure"]
    tm_export_cols  = TM_ORDER[:]

    ha_bytes = df_to_excel_bytes_with_format(
        head_cap(ha_df[ha_export_cols], HA_MAX_ROWS),
        col_width_default=15,
        col_width_overrides={"risk_control": 100},
        freeze_cell="D2",  # freeze row 2, col 4
        row_height=30
    )
    dvp_bytes = df_to_excel_bytes_with_format(
        head_cap(dvp_df[dvp_export_cols], DVP_MAX_ROWS),
        col_width_default=15,
        col_width_overrides={"acceptance_criteria": 100, "test_procedure": 100},
        freeze_cell="D2",
        row_height=30
    )
    tm_bytes = df_to_excel_bytes_with_format(
        head_cap(tm_df[tm_export_cols], TM_MAX_ROWS),
        col_width_default=15,
        col_width_overrides={},
        freeze_cell="D2",
        row_height=30
    )

    # Save locally
    with open(os.path.join(OUTPUT_DIR, "Hazard_Analysis.xlsx"), "wb") as f:
        f.write(ha_bytes)
    with open(os.path.join(OUTPUT_DIR, "Design_Verification_Protocol.xlsx"), "wb") as f:
        f.write(dvp_bytes)
    with open(os.path.join(OUTPUT_DIR, "Trace_Matrix.xlsx"), "wb") as f:
        f.write(tm_bytes)

    # One-click ZIP Download
    st.subheader("Download DHF Package")
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Hazard_Analysis.xlsx", ha_bytes)
        zf.writestr("Design_Verification_Protocol.xlsx", dvp_bytes)
        zf.writestr("Trace_Matrix.xlsx", tm_bytes)
    zip_buf.seek(0)

    clicked = st.download_button(
        "⬇️ Download DHF package (3 Excel files, ZIP)",
        data=zip_buf.getvalue(),
        file_name="DHF_Package.zip",
        mime="application/zip",
        type="primary"
    )
    if clicked:
        st.success("DHF documents downloaded successfully — Hazard Analysis, Design Verification Protocol, Trace Matrix")
        st.info("Note: Please involve Medical Device SME reviews before final approval.")
