# app.py — DHF Streamlit UI (Preview + ZIP Download)
# --------------------------------------------------

import os
import io
import json
import zipfile
import typing as t

import numpy as np
import pandas as pd
import requests
import streamlit as st

# Excel styling
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------- Constants & Secrets ----------------
TBD = "TBD - Human / SME input"
TBD_CANON = {"", "na", "n/a", "tbd", "none", "null", "nan", "tbd - human / sme input"}

BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8080")
BACKEND_TOKEN = st.secrets.get("BACKEND_TOKEN", "dev-token")

# Row caps (preview + Excel export)
HA_MAX_ROWS = int(os.getenv("HA_MAX_ROWS", "50"))
DVP_MAX_ROWS = int(os.getenv("DVP_MAX_ROWS", "50"))
TM_MAX_ROWS  = int(os.getenv("TM_MAX_ROWS", "50"))

# NEW: cap how many requirements we send to backend
REQ_MAX = int(os.getenv("REQ_MAX", "50"))

# Optional Google Drive upload
DEFAULT_DRIVE_FOLDER_ID = st.secrets.get("DRIVE_FOLDER_ID", "")
OUTPUT_DIR = st.secrets.get("OUTPUT_DIR", os.path.abspath("./streamlit_outputs"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- Optional Google Drive (pydrive2) ----------------
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


# ---------------- Helpers ----------------
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


def df_to_excel_bytes_with_style(
    df: pd.DataFrame,
    long_cols: t.Iterable[str] = (),
    default_col_width: int = 15,
    long_col_width: int = 100,
    freeze_pane_addr: str = "D2",
) -> bytes:
    """Write df to Excel with formatting."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
        ws = w.book.active

        # Header style
        header_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
        header_font = Font(bold=True)
        wrap_align = Alignment(wrap_text=True, vertical="center", horizontal="left")

        # Column widths
        long_idx = {str(c): i + 1 for i, c in enumerate(df.columns) if str(c) in set(long_cols)}
        for i, c in enumerate(df.columns, start=1):
            col_letter = get_column_letter(i)
            if str(c) in long_idx:
                ws.column_dimensions[col_letter].width = long_col_width
            else:
                ws.column_dimensions[col_letter].width = default_col_width

        # Row height + wrap + header style
        for r in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in r:
                cell.alignment = wrap_align
                if cell.row == 1:
                    cell.fill = header_fill
                    cell.font = header_font
            ws.row_dimensions[cell.row].height = 30

        # Freeze panes
        try:
            ws.freeze_panes = ws[freeze_pane_addr]
        except Exception:
            pass

    buf.seek(0)
    return buf.read()


def basic_guardrails_df(df: pd.DataFrame, required_cols: t.List[str]) -> pd.DataFrame:
    issues = []
    for i, row in df.iterrows():
        for c in required_cols:
            val = str(row.get(c, "")).strip()
            if not val or val in {"NA", TBD}:
                issues.append((i, c, "Missing or TBD"))
    return pd.DataFrame(issues, columns=["row_index", "column", "issue"]) if issues else pd.DataFrame(columns=["row_index", "column", "issue"])


def fill_tbd(df: pd.DataFrame, cols: t.List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = TBD
        out[c] = out[c].fillna(TBD)
        out.loc[out[c].astype(str).str.strip().eq(""), c] = TBD
        out.loc[out[c].astype(str).str.upper().eq("NA"), c] = TBD
    return out


def call_backend(endpoint: str, payload: dict) -> dict:
    url = f"{BACKEND_URL.rstrip('/')}{endpoint}"
    headers = {"Authorization": f"Bearer {BACKEND_TOKEN}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json=payload, timeout=1200)
    if r.status_code >= 400:
        raise RuntimeError(f"Backend error {r.status_code}: {r.text[:500]}")
    return r.json()


def head_cap(df: pd.DataFrame, n: int) -> pd.DataFrame:
    return df.head(n).copy() if isinstance(df, pd.DataFrame) and not df.empty else df


# ---------------- Metrics helpers ----------------
def pct(n: int, d: int) -> float:
    return 0.0 if d <= 0 else round(100.0 * n / d, 1)

def _not_tbd_scalar(x) -> bool:
    s = str(x if x is not None else "").strip()
    return bool(s) and s.lower() not in TBD_CANON

def ha_metrics(ha_df: pd.DataFrame) -> dict:
    if ha_df is None or ha_df.empty:
        return {}
    keys = [
        "risk_to_health","hazard","hazardous_situation","harm",
        "sequence_of_events","severity_of_harm","p0","p1","poh","risk_index","risk_control"
    ]
    exist_cols = [k for k in keys if k in ha_df.columns]
    filled = int(
        ha_df[exist_cols].applymap(_not_tbd_scalar).values.sum()
    )
    completeness = pct(filled, len(ha_df) * len(exist_cols))
    uniq = ha_df[["risk_to_health","hazard","hazardous_situation","harm"]].drop_duplicates().shape[0] if all(k in ha_df.columns for k in ["risk_to_health","hazard","hazardous_situation","harm"]) else 0
    diversity = pct(uniq, len(ha_df))
    return {
        "Completeness": completeness,   # (Threshold shown in UI)
        "Scenario Diversity": diversity # (Threshold shown in UI)
    }

def _has_numbers_units(text: str) -> bool:
    if not isinstance(text, str):
        return False
    return bool(
        pd.Series([text]).str.contains(r"[±]?\d+(?:\.\d+)?\s?(mA|µA|A|V|kV|ms|s|min|h|mL/h|mL|L|kPa|Pa|%|dB|Ω|°C|g|kg|N|cycles|cm|mm|µL|m)", case=False, regex=True).iloc[0]
    )

def dvp_metrics(dvp_df: pd.DataFrame) -> dict:
    if dvp_df is None or dvp_df.empty:
        return {}
    # accept various casings
    cols = {c.lower(): c for c in dvp_df.columns}
    vm = cols.get("verification method") or cols.get("verification_method")
    tp = cols.get("test procedure") or cols.get("test_procedure")
    ac = cols.get("acceptance criteria") or cols.get("acceptance_criteria")
    ss = cols.get("sample size") or cols.get("sample_size")

    filled_cols = [c for c in [vm, tp, ac, ss] if c]
    filled = int(dvp_df[filled_cols].applymap(_not_tbd_scalar).values.sum())
    completeness = pct(filled, len(dvp_df) * len(filled_cols)) if filled_cols else 0.0

    measurable = int(dvp_df[tp].apply(_has_numbers_units).sum()) if tp else 0
    meas_pct = pct(measurable, len(dvp_df))
    return {
        "DVP Field Completeness": completeness, # threshold in UI
        "Measurable Test Steps": meas_pct       # threshold in UI
    }

def _token_overlap(a: str, b: str) -> float:
    def toks(x):
        return {t.lower() for t in (x or "").split() if t.isalpha() and len(t) >= 4}
    aa, bb = toks(a), toks(b)
    if not aa or not bb:
        return 0.0
    inter = len(aa & bb)
    denom = max(1, min(len(aa), len(bb)))
    return inter / denom

def tm_metrics(tm_df: pd.DataFrame) -> dict:
    if tm_df is None or tm_df.empty:
        return {}
    # names from our TM schema
    req = "Requirements"
    rc  = "HA Risk Control"
    ok = tm_df[[req, rc]].apply(
        lambda r: _token_overlap(str(r[req]), str(r[rc])) >= 0.2, axis=1
    ).sum()
    mapping = pct(int(ok), len(tm_df))
    return {"Req↔RiskControl Mapping": mapping}  # threshold shown in UI

def render_metrics(ha_df, dvp_df, tm_df):
    ha = ha_metrics(ha_df)
    dvp = dvp_metrics(dvp_df)
    tm  = tm_metrics(tm_df)

    st.subheader("Evaluation Metrics")
    st.caption("Objective: these metrics quickly summarize coverage, consistency, and requirement↔control mapping quality across the generated documents.")

    # display with thresholds
    T = {
        "HA": {
            "Completeness": 80,
            "Scenario Diversity": 40,
        },
        "DVP": {
            "DVP Field Completeness": 80,
            "Measurable Test Steps": 60,
        },
        "TM": {
            "Req↔RiskControl Mapping": 60,
        }
    }

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(f"HA • Completeness (Threshold: {T['HA']['Completeness']}%)", f"{ha.get('Completeness',0)}%")
        st.metric(f"HA • Scenario Diversity (Threshold: {T['HA']['Scenario Diversity']}%)", f"{ha.get('Scenario Diversity',0)}%")
    with col2:
        st.metric(f"DVP • Field Completeness (Threshold: {T['DVP']['DVP Field Completeness']}%)", f"{dvp.get('DVP Field Completeness',0)}%")
        st.metric(f"DVP • Measurable Steps (Threshold: {T['DVP']['Measurable Test Steps']}%)", f"{dvp.get('Measurable Test Steps',0)}%")
    with col3:
        st.metric(f"TM • Req↔RiskControl Mapping (Threshold: {T['TM']['Req↔RiskControl Mapping']}%)", f"{tm.get('Req↔RiskControl Mapping',0)}%")

    st.caption("Please involve Medical Device SMEs for final review of these documents before approval.")


# ---------------- UI ----------------
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")

# Title (single line) + images on the right
top_cols = st.columns([0.70, 0.30])
with top_cols[0]:
    st.markdown(
        "<h1 style='margin-top:0; white-space:nowrap;'>DHF Automation – Infusion Pump</h1>",
        unsafe_allow_html=True
    )
    st.markdown("Requirements → Hazard Analysis → DVP → TM")

with top_cols[1]:
    c1, c2 = st.columns(2)
    with c1:
        st.image("streamlit_assets/Infusion1.jpg", use_container_width=True)
    with c2:
        st.image("streamlit_assets/Infusion.jpg", use_container_width=True)

st.divider()

st.markdown("**Provide Product Requirements** (choose one):")
colA, colB = st.columns(2)
with colA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with colB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=140, placeholder=sample)

# One button only (kept near the top)
run_btn = st.button("▶️ Generate DHF Packages", type="primary", use_container_width=False)

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

    total_reqs = len(req_df)
    st.success(f"Loaded {total_reqs} requirements.")
    st.dataframe(head_cap(req_df, 20), use_container_width=True)

    # -------- Limit how many requirements are SENT to backend --------
    req_df_limited = req_df.head(REQ_MAX).copy()
    if total_reqs > REQ_MAX:
        st.info(f"Sending only the first {REQ_MAX} requirements to backend (set REQ_MAX env var to change).")

    # -------- Call Backend: HA --------
    with st.spinner("Running Hazard Analysis (backend)..."):
        ha_payload = {
            "requirements": [
                {
                    "Requirement ID": r["Requirement ID"],
                    "Verification ID": r.get("Verification ID"),
                    "Requirements": r.get("Requirements"),
                }
                for _, r in req_df_limited.iterrows()  # <— limited
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

    # -------- Call Backend: DVP --------
    with st.spinner("Generating Design Verification Protocol (backend)..."):
        dvp_payload = {
            "requirements": ha_payload["requirements"],  # already limited
            "ha": ha_rows,
        }
    try:
        dvp_resp = call_backend("/dvp", dvp_payload)
        dvp_rows = dvp_resp.get("dvp", [])
    except Exception as e:
        st.error(f"DVP backend error: {e}")
        st.stop()
    dvp_df = pd.DataFrame(dvp_rows)

    # Standardize names to match export schema
    rename_map = {
        "verification_id": "Verification ID",
        "requirement_id": "Requirement ID",
        "requirements": "Requirements",
        "verification_method": "Verification Method",
        "sample_size": "Sample size",
        "test_procedure": "Test Procedure",
        "acceptance_criteria": "Acceptance criteria",
    }
    dvp_df = dvp_df.rename(columns=rename_map)
    dvp_order = [
        "Verification ID", "Requirement ID", "Requirements",
        "Verification Method", "Sample size", "Test Procedure", "Acceptance criteria"
    ]
    dvp_df = dvp_df.reindex(columns=dvp_order)
    dvp_df = fill_tbd(dvp_df, dvp_order)

    st.subheader(f"Design Verification Protocol (preview, first {DVP_MAX_ROWS})")
    st.dataframe(head_cap(dvp_df, DVP_MAX_ROWS), use_container_width=True)

    # -------- Call Backend: TM --------
    with st.spinner("Building Trace Matrix (backend)..."):
        tm_payload = {
            "requirements": ha_payload["requirements"],  # already limited
            "ha": ha_rows,
            "dvp": dvp_rows,
        }
        tm_resp = call_backend("/tm", tm_payload)  # backend returns {"ok":true,"tm":[...]}
        tm_rows = tm_resp.get("tm", [])
        tm_df = pd.DataFrame(tm_rows)

    # Enforce exact TM column order
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

    # -------- Evaluation Metrics (with thresholds) --------
    render_metrics(ha_df, dvp_df, tm_df)

    # -------- Prepare styled Excel exports (capped for preview) --------
    ha_export_cols = [
        "risk_id", "risk_to_health", "hazard", "hazardous_situation",
        "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
    ]
    # Use only intersecting columns to avoid KeyErrors
    ha_export_cols = [c for c in ha_export_cols if c in ha_df.columns]
    ha_bytes = df_to_excel_bytes_with_style(
        head_cap(ha_df[ha_export_cols], HA_MAX_ROWS),
        long_cols=["risk_control"],
        long_col_width=100,
        default_col_width=15,
        freeze_pane_addr="D2"
    )

    dvp_bytes = df_to_excel_bytes_with_style(
        head_cap(dvp_df, DVP_MAX_ROWS),
        long_cols=["Acceptance criteria", "Test Procedure"],
        long_col_width=100,
        default_col_width=15,
        freeze_pane_addr="D2"
    )

    tm_bytes  = df_to_excel_bytes_with_style(
        head_cap(tm_df, TM_MAX_ROWS),
        long_cols=["Requirements", "HA Risk Control"],
        long_col_width=100,
        default_col_width=15,
        freeze_pane_addr="D2"
    )

    # Save locally
    with open(os.path.join(OUTPUT_DIR, "Hazard_Analysis.xlsx"), "wb") as f:
        f.write(ha_bytes)
    with open(os.path.join(OUTPUT_DIR, "Design_Verification_Protocol.xlsx"), "wb") as f:
        f.write(dvp_bytes)
    with open(os.path.join(OUTPUT_DIR, "Trace_Matrix.xlsx"), "wb") as f:
        f.write(tm_bytes)

    # -------- One-click ZIP Download --------
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
        st.success("DHF documents downloaded successfully - Hazard Analysis, Design Verification Protocol, Trace Matrix")
        st.info("Note: Please involve Medical Device - SME reviews, before final approval.")

    # -------- Optional Google Drive upload --------
    if DEFAULT_DRIVE_FOLDER_ID:
        with st.spinner("Uploading files to Google Drive..."):
            drive = init_drive()
            if drive:
                try:
                    ha_id = drive_upload_bytes(drive, DEFAULT_DRIVE_FOLDER_ID, "Hazard_Analysis.xlsx", ha_bytes)
                    dvp_id = drive_upload_bytes(drive, DEFAULT_DRIVE_FOLDER_ID, "Design_Verification_Protocol.xlsx", dvp_bytes)
                    tm_id = drive_upload_bytes(drive, DEFAULT_DRIVE_FOLDER_ID, "Trace_Matrix.xlsx", tm_bytes)
                    st.info("Uploaded to Google Drive (file IDs shown below).")
                    st.json({
                        "Hazard_Analysis.xlsx": ha_id,
                        "Design_Verification_Protocol.xlsx": dvp_id,
                        "Trace_Matrix.xlsx": tm_id,
                    })
                except Exception as e:
                    st.warning(f"Drive upload failed: {e}")
            else:
                st.info("Drive not initialized (missing SERVICE_ACCOUNT_JSON secret). Files saved locally only.")
