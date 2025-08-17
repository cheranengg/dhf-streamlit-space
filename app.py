# app.py — DHF Streamlit UI (Preview + ZIP Download)
# --------------------------------------------------

import os
import io
import json
import zipfile
import typing as t
from pathlib import Path

import pandas as pd
import streamlit as st
import requests

# ---------------- Constants & Secrets ----------------
TBD = "TBD - Human / SME input"

BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8080")
BACKEND_TOKEN = st.secrets.get("BACKEND_TOKEN", "dev-token")

# Row caps (preview + Excel export)
HA_MAX_ROWS = int(os.getenv("HA_MAX_ROWS", "50"))
DVP_MAX_ROWS = int(os.getenv("DVP_MAX_ROWS", "50"))
TM_MAX_ROWS  = int(os.getenv("TM_MAX_ROWS", "50"))

# Cap how many requirements we send to backend
REQ_MAX = int(os.getenv("REQ_MAX", "50"))

# Optional Google Drive upload (kept but hidden in UI)
DEFAULT_DRIVE_FOLDER_ID = st.secrets.get("DRIVE_FOLDER_ID", "")
OUTPUT_DIR = st.secrets.get("OUTPUT_DIR", os.path.abspath("./streamlit_outputs"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ------------ UI polish: make top area denser & button visible -------------
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")
st.markdown("""
<style>
/* tighten default paddings to keep button visible above the fold */
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
[data-testid="stVerticalBlock"] div:has(> .stDownloadButton) { margin-top: 0.5rem; }
.kpi-note { color:#198754; font-weight:700; font-size: 1.05rem; }
.kpi-value { font-size: 1.9rem; font-weight: 800; }
.kpi-label { font-size: 1rem; font-weight: 600; color: #444; }
.kpi-ok { color: #198754; }      /* green  */
.kpi-bad { color: #d00000; }     /* red    */
.heading-tight h1 { line-height: 1.2; }
</style>
""", unsafe_allow_html=True)

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


def call_backend(endpoint: str, payload: dict) -> dict:
    url = f"{BACKEND_URL.rstrip('/')}{endpoint}"
    headers = {"Authorization": f"Bearer {BACKEND_TOKEN}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json=payload, timeout=1200)
    if r.status_code >= 400:
        raise RuntimeError(f"Backend error {r.status_code}: {r.text[:500]}")
    return r.json()


def head_cap(df: pd.DataFrame, n: int) -> pd.DataFrame:
    return df.head(n).copy() if isinstance(df, pd.DataFrame) and not df.empty else df


# ---------- Excel styling helpers (openpyxl) ----------
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows

HEADER_FILL = PatternFill("solid", fgColor="A9D18E")  # soft green
HEADER_FONT = Font(bold=True, color="000000")

def _style_ws_common(ws, freeze_cell="D2", default_col_width=15, long_cols=None, row_height=30):
    """Apply common table styles."""
    long_cols = long_cols or {}
    # header row (row 1): bold + fill
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # widths
    for col in range(1, ws.max_column + 1):
        letter = ws.cell(1, col).column_letter
        ws.column_dimensions[letter].width = default_col_width
    for col_name, width in long_cols.items():
        # find index for the column name
        for j in range(1, ws.max_column + 1):
            if str(ws.cell(1, j).value).strip() == col_name:
                ws.column_dimensions[ws.cell(1, j).column_letter].width = width
                break

    # rows
    for i in range(1, ws.max_row + 1):
        ws.row_dimensions[i].height = row_height
        for j in range(1, ws.max_column + 1):
            ws.cell(i, j).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.freeze_panes = freeze_cell

def df_to_excel_bytes_styled(df: pd.DataFrame, long_cols: dict, freeze_cell="D2", default_col_width=15) -> bytes:
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    _style_ws_common(ws, freeze_cell=freeze_cell, default_col_width=default_col_width, long_cols=long_cols)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ------------- KPI utilities (colored) -------------
TBD_CANON = {None, "", "NA", "TBD - Human / SME input", "TBD", "N/A"}

def pct(n: int, d: int) -> float:
    return 0.0 if d <= 0 else round(100.0 * n / d, 1)

def _not_tbd_val(x) -> bool:
    s = str(x) if (x is not None) else ""
    s = s.strip()
    return bool(s) and (s not in TBD_CANON)

def render_metric(label: str, value: float, threshold: float):
    color_class = "kpi-ok" if value >= threshold else "kpi-bad"
    return f"""
    <div style="text-align:center; margin: 6px 10px 22px 10px;">
        <div class="kpi-label">{label} <span style="color:#888;">(Threshold: {threshold:.0f}%)</span></div>
        <div class="kpi-value {color_class}">{value:.1f}%</div>
    </div>
    """

# ---------- KPI computation ----------
# Thresholds (you can tweak)
HA_COMPLETENESS_T = 80
HA_DIVERSITY_T    = 70
DVP_COMPLETENESS_T = 80
DVP_SEV_SAMPLE_T   = 70
TM_MAPPING_T       = 80

def ha_metrics(ha_df: pd.DataFrame) -> dict:
    if ha_df is None or ha_df.empty:
        return {"completeness": 0.0, "diversity": 0.0}

    keys = ["risk_to_health", "hazard", "hazardous_situation", "harm",
            "sequence_of_events", "severity_of_harm", "p0", "p1", "poh",
            "risk_index", "risk_control"]
    filled = 0
    for _, row in ha_df.iterrows():
        for k in keys:
            filled += 1 if _not_tbd_val(row.get(k)) else 0
    completeness = pct(filled, len(ha_df) * len(keys))

    # diversity across selected HA fields
    uniq = ha_df[["risk_to_health", "hazard", "hazardous_situation", "harm"]].dropna().astype(str).apply(lambda r: " | ".join(r.values), axis=1).nunique()
    diversity = pct(uniq, len(ha_df))
    return {"completeness": completeness, "diversity": diversity}

def dvp_metrics(dvp_df: pd.DataFrame, ha_df: pd.DataFrame) -> dict:
    if dvp_df is None or dvp_df.empty:
        return {"completeness": 0.0, "sev_sample": 0.0}

    fields = ["verification_id", "verification_method", "acceptance_criteria", "sample_size", "test_procedure"]
    filled = 0
    for _, row in dvp_df.iterrows():
        for k in fields:
            filled += 1 if _not_tbd_val(row.get(k)) else 0
    completeness = pct(filled, len(dvp_df) * len(fields))

    # severity ↔ sample size alignment
    sev_by_req = {}
    for _, r in ha_df.iterrows():
        rid = str(r.get("requirement_id") or r.get("Requirement ID") or "").strip()
        sev = str(r.get("severity_of_harm") or "").strip()
        if rid and sev.isdigit():
            sev_by_req[rid] = int(sev)

    aligned = 0
    considered = 0
    for _, row in dvp_df.iterrows():
        rid = str(row.get("requirement_id") or row.get("Requirement ID") or "").strip()
        ssz = str(row.get("sample_size") or "").strip()
        if rid and rid in sev_by_req and ssz.isdigit():
            considered += 1
            sev = sev_by_req[rid]  # 1..5
            size = int(ssz)
            # expected sizes (simple rule: 1→10, 2→20, 3→30, 4→40, 5→50), ±20% tolerance
            expected = {1:10, 2:20, 3:30, 4:40, 5:50}.get(sev, 30)
            lower = int(expected * 0.8)
            upper = int(expected * 1.2)
            if lower <= size <= upper:
                aligned += 1
    sev_sample = pct(aligned, considered if considered > 0 else 1)
    return {"completeness": completeness, "sev_sample": sev_sample}

def tm_metrics(tm_df: pd.DataFrame) -> dict:
    if tm_df is None or tm_df.empty:
        return {"mapping": 0.0}

    # completeness for TM main fields
    fields = ["Requirement ID", "Requirements", "Risk ID", "Risk to Health", "HA Risk Control",
              "Verification ID", "Verification Method", "Acceptance Criteria"]
    filled = 0
    for _, row in tm_df.iterrows():
        for k in fields:
            filled += 1 if _not_tbd_val(row.get(k)) else 0
    completeness = pct(filled, len(tm_df) * len(fields))

    # mapping quality: does HA risk control content appear lexically similar to requirement text?
    good = 0
    total = 0
    for _, row in tm_df.iterrows():
        req = str(row.get("Requirements") or "").lower()
        rc  = str(row.get("HA Risk Control") or "").lower()
        if not req or not rc or rc in TBD_CANON:
            continue
        total += 1
        # cheap lexical overlap (set-based)
        req_words = set([w for w in req.replace(",", " ").replace("/", " ").split() if len(w) > 3])
        rc_words  = set([w for w in rc.replace(",", " ").replace("/", " ").split() if len(w) > 3])
        overlap = len(req_words & rc_words)
        union   = max(1, len(req_words | rc_words))
        jaccard = overlap / union
        if jaccard >= 0.20:  # modest bar because of paraphrase
            good += 1
    mapping = pct(good, total if total > 0 else 1)
    return {"completeness": completeness, "mapping": mapping}


# ---------------- Title & Images ----------------
# Find images in /app/assets or /streamlit_assets
def _img_path(name: str) -> t.Optional[str]:
    for p in [Path("app/assets")/name, Path("streamlit_assets")/name, Path("assets")/name]:
        if p.exists():
            return str(p)
    return None

col_hd1, col_hd2 = st.columns([0.7, 0.3])
with col_hd1:
    st.markdown('<div class="heading-tight">', unsafe_allow_html=True)
    st.title("DHF Automation – Infusion Pump")
    st.caption("Requirements → Hazard Analysis → DVP → TM")
    st.markdown('</div>', unsafe_allow_html=True)
with col_hd2:
    c1, c2 = st.columns(2)
    imgA = _img_path("Infusion1.jpg")
    imgB = _img_path("Infusion.jpg")
    if imgA:
        c1.image(imgA)
    if imgB:
        c2.image(imgB)

st.markdown("---")

# ---------------- Input Section ----------------
st.markdown("**Provide Product Requirements** (choose one):")
colA, colB = st.columns(2)
with colA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with colB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=140, placeholder=sample)

# Single green primary button
run_btn = st.button("▶️ Generate DHF Packages", type="primary")

# area for immediate success text after download
download_msg_slot = st.empty()

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
                for _, r in req_df_limited.iterrows()
            ]
        }
        ha_resp = call_backend("/hazard-analysis", ha_payload)
        ha_rows = ha_resp.get("ha", [])
        ha_df = pd.DataFrame(ha_rows)

    st.subheader(f"Hazard Analysis (preview, first {HA_MAX_ROWS})")
    st.dataframe(head_cap(ha_df, HA_MAX_ROWS), use_container_width=True)

    # -------- Call Backend: DVP --------
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

    st.subheader(f"Design Verification Protocol (preview, first {DVP_MAX_ROWS})")
    st.dataframe(head_cap(dvp_df, DVP_MAX_ROWS), use_container_width=True)

    # -------- Call Backend: TM --------
    with st.spinner("Building Trace Matrix (backend)..."):
        tm_payload = {
            "requirements": ha_payload["requirements"],
            "ha": ha_rows,
            "dvp": dvp_rows,
        }
        tm_resp = call_backend("/tm", tm_payload)
        tm_rows = tm_resp.get("tm", [])
        tm_df = pd.DataFrame(tm_rows)

    # enforce TM column order for preview/export
    TM_ORDER = [
        "Requirement ID",
        "Requirements",
        "Requirement (Yes/No)",
        "Risk ID",
        "Risk to Health",
        "HA Risk Control",
        "Verification ID",
        "Verification Method",
        "Acceptance Criteria",
    ]
    tm_df = tm_df.reindex(columns=TM_ORDER)

    st.subheader(f"Trace Matrix (preview, first {TM_MAX_ROWS})")
    st.dataframe(head_cap(tm_df, TM_MAX_ROWS), use_container_width=True)

    # -------- Evaluation Metrics (colored + thresholds) --------
    st.subheader("Evaluation Metrics")
    st.markdown(
        "Objective: these metrics quickly summarize **coverage**, **consistency**, and "
        "**requirement↔control mapping** quality across the generated documents."
    )

    ha_m = ha_metrics(ha_df)
    dvp_m = dvp_metrics(dvp_df, ha_df)
    tm_m  = tm_metrics(tm_df)

    row1 = st.columns(3)
    with row1[0]:
        st.markdown(render_metric("HA • Completeness", ha_m["completeness"], HA_COMPLETENESS_T), unsafe_allow_html=True)
    with row1[1]:
        st.markdown(render_metric("DVP • Field Completeness", dvp_m["completeness"], DVP_COMPLETENESS_T), unsafe_allow_html=True)
    with row1[2]:
        st.markdown(render_metric("TM • Req↔RiskControl Mapping", tm_m["mapping"], TM_MAPPING_T), unsafe_allow_html=True)

    row2 = st.columns(3)
    with row2[0]:
        st.markdown(render_metric("HA • Scenario Diversity", ha_m["diversity"], HA_DIVERSITY_T), unsafe_allow_html=True)
    with row2[1]:
        st.markdown(render_metric("DVP • Severity↔Sample Alignment", dvp_m["sev_sample"], DVP_SEV_SAMPLE_T), unsafe_allow_html=True)
    with row2[2]:
        # also show TM completeness on row2/right for balance
        st.markdown(render_metric("TM • Field Completeness", tm_m["completeness"], DVP_COMPLETENESS_T), unsafe_allow_html=True)

    st.markdown(
        "<p class='kpi-note'>Note: Please involve Medical Device SMEs for final review of these documents before approval.</p>",
        unsafe_allow_html=True
    )

    # -------- Prepare styled Excel exports (capped for preview) --------
    # Column widths per your spec:
    #   • DVP: all 15 except Acceptance Criteria & Test Procedure & Requirements → 100
    #   • HA : all 15 except Risk control → 100, Hazardous situation & Sequence of events → 40
    #   • TM : all 15 except Requirements & HA Risk Control → 100
    ha_export_cols = [
        "requirement_id", "risk_id", "risk_to_health", "hazard", "hazardous_situation",
        "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
    ]
    # map missing columns gracefully
    ha_export = head_cap(ha_df[ [c for c in ha_export_cols if c in ha_df.columns] ], HA_MAX_ROWS)
    ha_bytes = df_to_excel_bytes_styled(
        ha_export,
        long_cols={
            "risk_control": 100,
            "hazardous_situation": 40,
            "sequence_of_events": 40,
        },
        freeze_cell="D2",
        default_col_width=15
    )

    dvp_export_cols = ["verification_id", "requirement_id", "requirements", "verification_method",
                       "sample_size", "test_procedure", "acceptance_criteria"]
    dvp_export = head_cap(dvp_df[ [c for c in dvp_export_cols if c in dvp_df.columns] ], DVP_MAX_ROWS)
    dvp_bytes = df_to_excel_bytes_styled(
        dvp_export,
        long_cols={
            "requirements": 100,
            "test_procedure": 100,
            "acceptance_criteria": 100
        },
        freeze_cell="D2",
        default_col_width=15
    )

    tm_export_cols = TM_ORDER[:]  # same order as preview
    tm_export = head_cap(tm_df[tm_export_cols], TM_MAX_ROWS)
    tm_bytes  = df_to_excel_bytes_styled(
        tm_export,
        long_cols={
            "Requirements": 100,
            "HA Risk Control": 100
        },
        freeze_cell="D2",
        default_col_width=15
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
        download_msg_slot.markdown(
            "<div style='margin-top:6px; font-weight:800;'>"
            "DHF documents downloaded successfully - Hazard Analysis, Design Verification Protocol, Trace Matrix"
            "</div>",
            unsafe_allow_html=True
        )

    # -------- Optional Google Drive upload (silent UI) --------
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
