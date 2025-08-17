# app.py — DHF Streamlit UI (Preview + ZIP Download)
# --------------------------------------------------

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
TBD_CANON = {TBD, "TBD", "NA", "N/A", "None", "null", ""}

BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8080")
BACKEND_TOKEN = st.secrets.get("BACKEND_TOKEN", "dev-token")

# Row caps (preview + Excel export)
HA_MAX_ROWS = int(os.getenv("HA_MAX_ROWS", "50"))
DVP_MAX_ROWS = int(os.getenv("DVP_MAX_ROWS", "50"))
TM_MAX_ROWS  = int(os.getenv("TM_MAX_ROWS", "50"))

# NEW: cap how many requirements we send to backend
REQ_MAX = int(os.getenv("REQ_MAX", "50"))

# Optional Google Drive upload (left wired but unused)
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


def df_to_excel_bytes(df: pd.DataFrame, formatter: t.Callable[[pd.ExcelWriter], None]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Sheet1")
        formatter(w)  # apply widths/styles/freeze panes
    buf.seek(0)
    return buf.read()


def call_backend(endpoint: str, payload: dict) -> dict:
    url = f"{BACKEND_URL.rstrip('/')}{endpoint}"
    headers = {"Authorization": f"Bearer {BACKEND_TOKEN}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json=payload, timeout=1200)
    if r.status_code >= 400:
        raise RuntimeError(f"Backend error {r.status_code}: {r.text[:500]}")
    return r.json()


def head_cap(df: pd.DataFrame, n: int) -> pd.DataFrame:
    return df.head(n).copy() if isinstance(df, pd.DataFrame) and not df.empty else df


def _not_tbd(x) -> bool:
    s = str("" if x is None else x).strip()
    return bool(s) and s not in TBD_CANON


def _pct(num: int, den: int) -> float:
    if den <= 0:
        return 0.0
    return round(100.0 * max(0, num) / den, 1)


# ---------------- Metrics (HA / DVP / TM) ----------------
# Thresholds for green/red
TH = {
    "ha_field": 0.80,
    "ha_div": 0.70,
    "dvp_field": 0.80,
    "dvp_align": 0.70,
    "tm_field": 0.80,
    "tm_cov": 0.80,  # NEW: TM Requirement Coverage
}

def ha_metrics(ha_df: pd.DataFrame) -> dict:
    if ha_df is None or ha_df.empty:
        return {"field": 0.0, "div": 0.0}

    keys = [
        "risk_id", "risk_to_health", "hazard", "hazardous_situation",
        "harm", "sequence_of_events", "severity_of_harm",
        "p0", "p1", "poh", "risk_index", "risk_control",
    ]
    # Field completeness = fraction of populated cells across selected columns
    filled = 0
    for k in keys:
        if k in ha_df.columns:
            filled += int(ha_df[k].apply(_not_tbd).sum())
        else:
            # column missing → contributes 0 populated cells
            filled += 0
    field = _pct(filled, len(ha_df) * len(keys))

    # Scenario diversity = unique tuples across hazard/hazardous_situation/harm/sequence_of_events
    cols = [c for c in ["hazard", "hazardous_situation", "harm", "sequence_of_events"] if c in ha_df.columns]
    if cols:
        uniq = len(ha_df[cols].drop_duplicates())
    else:
        uniq = 0
    div = _pct(uniq, len(ha_df))
    return {"field": field, "div": div}


def dvp_metrics(dvp_df: pd.DataFrame) -> dict:
    if dvp_df is None or dvp_df.empty:
        return {"field": 0.0, "align": 0.0}

    keys = ["verification_id", "verification_method", "acceptance_criteria", "sample_size", "test_procedure"]
    filled = 0
    for k in keys:
        if k in dvp_df.columns:
            filled += int(dvp_df[k].apply(_not_tbd).sum())
    field = _pct(filled, len(dvp_df) * len(keys))

    # alignment: sample size present & numeric; proxy for HA severity alignment
    good = 0
    if "sample_size" in dvp_df.columns:
        for v in dvp_df["sample_size"]:
            s = str(v or "").strip()
            if s.isdigit() and int(s) > 0:
                good += 1
    align = _pct(good, len(dvp_df))
    return {"field": field, "align": align}


def tm_metrics(tm_df: pd.DataFrame) -> dict:
    if tm_df is None or tm_df.empty:
        return {"field": 0.0, "coverage": 0.0}

    # Field completeness over important columns
    keys = ["Requirement ID", "Requirements", "Requirement (Yes/No)",
            "Risk ID", "Risk to Health", "HA Risk Control",
            "Verification ID", "Verification Method"]
    filled = 0
    for k in keys:
        if k in tm_df.columns:
            filled += int(tm_df[k].apply(_not_tbd).sum())
    field = _pct(filled, len(tm_df) * len(keys))

    # NEW: Requirement Coverage = % of Requirement rows that have ANY HA Risk Control
    req_rows = tm_df[tm_df.get("Requirement (Yes/No)", "Requirement").eq("Requirement")] if "Requirement (Yes/No)" in tm_df.columns else tm_df
    have_control = 0
    if not req_rows.empty and "HA Risk Control" in req_rows.columns:
        have_control = int(req_rows["HA Risk Control"].apply(_not_tbd).sum())
    coverage = _pct(have_control, len(req_rows)) if len(req_rows) else 0.0

    return {"field": field, "coverage": coverage}


def kpi(value: float, threshold: float) -> str:
    color = "#1f7a4f" if value >= threshold * 100 else "#c62828"
    return f"<span style='color:{color}; font-weight:700;'>{value:.1f}%</span>"


def render_metrics(ha_df: pd.DataFrame, dvp_df: pd.DataFrame, tm_df: pd.DataFrame):
    ha = ha_metrics(ha_df)
    dvp = dvp_metrics(dvp_df)
    tm  = tm_metrics(tm_df)

    st.subheader("Evaluation Metrics")
    st.caption(
        "Objective: these metrics quickly summarize **coverage**, **consistency**, "
        "and **requirement↔control mapping** quality across the generated documents."
    )

    # Row 1: HA Field • DVP Field • TM Field
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            f"**HA • Field Completeness (Threshold: {int(TH['ha_field']*100)}%)**\n\n"
            f"{kpi(ha['field'], TH['ha_field'])}",
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"**DVP • Field Completeness (Threshold: {int(TH['dvp_field']*100)}%)**\n\n"
            f"{kpi(dvp['field'], TH['dvp_field'])}",
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"**TM • Field Completeness (Threshold: {int(TH['tm_field']*100)}%)**\n\n"
            f"{kpi(tm['field'], TH['tm_field'])}",
            unsafe_allow_html=True,
        )

    # Row 2: HA Diversity • DVP Alignment • TM Requirement Coverage (NEW)
    c4, c5, c6 = st.columns(3)
    with c4:
        st.markdown(
            f"**HA • Scenario Diversity (Threshold: {int(TH['ha_div']*100)}%)**\n\n"
            f"{kpi(ha['div'], TH['ha_div'])}",
            unsafe_allow_html=True,
        )
    with c5:
        st.markdown(
            f"**DVP • Severity↔Sample Alignment (Threshold: {int(TH['dvp_align']*100)}%)**\n\n"
            f"{kpi(dvp['align'], TH['dvp_align'])}",
            unsafe_allow_html=True,
        )
    with c6:
        st.markdown(
            f"**TM • Requirement Coverage (Threshold: {int(TH['tm_cov']*100)}%)**\n\n"
            f"{kpi(tm['coverage'], TH['tm_cov'])}",
            unsafe_allow_html=True,
        )

    st.markdown(
        "<div style='margin-top:12px; font-weight:700; color:#1f4bbb;'>"
        "Note: Please involve Medical Device SMEs for final review of these documents before approval."
        "</div>",
        unsafe_allow_html=True,
    )


# ---------------- Excel styling helpers ----------------
def style_common(ws):
    from openpyxl.styles import Alignment, Font, PatternFill
    # Header row
    for cell in ws[1]:
        cell.font = Font(bold=True, color="000000")
        cell.fill = PatternFill("solid", fgColor="99FF99")  # light green
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    # Body
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    # Row height
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 30
    # Freeze panes: below header, 4th column
    ws.freeze_panes = "D2"


def format_ha(writer: pd.ExcelWriter):
    ws = writer.book["Sheet1"]
    style_common(ws)
    # Column widths
    widths = {}
    for idx, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
        header = str(col[0].value)
        # base 15
        w = 15
        if header in ("risk_control", "risk control", "Risk Control", "HA Risk Control"):
            w = 100
        if header in ("hazardous_situation", "hazardous situation", "Hazardous Situation"):
            w = 40
        if header in ("sequence_of_events", "sequence of events", "Sequence of Events"):
            w = 40
        widths[idx] = w
    for i, w in widths.items():
        ws.column_dimensions[chr(64 + i)].width = w


def format_dvp(writer: pd.ExcelWriter):
    ws = writer.book["Sheet1"]
    style_common(ws)
    # Column widths
    widths = {}
    for idx, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
        header = str(col[0].value)
        w = 15
        if header.lower().strip() in {"acceptance criteria", "acceptance_criteria"}:
            w = 100
        if header.lower().strip() in {"test procedure", "test_procedure"}:
            w = 100
        if header.lower().strip() in {"requirements", "requirement"}:
            w = 100
        widths[idx] = w
    for i, w in widths.items():
        ws.column_dimensions[chr(64 + i)].width = w


def format_tm(writer: pd.ExcelWriter):
    ws = writer.book["Sheet1"]
    style_common(ws)
    widths = {}
    for idx, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
        header = str(col[0].value)
        w = 15
        if header in ("Requirements",):
            w = 100
        if header in ("HA Risk Control",):
            w = 100
        widths[idx] = w
    for i, w in widths.items():
        ws.column_dimensions[chr(64 + i)].width = w


# ---------------- UI ----------------
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")

# Title row with puzzle icon + purple title + two images to the right
col_logo, col_title, col_img1, col_img2 = st.columns([0.12, 1.2, 0.9, 0.9])

with col_logo:
    st.image("streamlit_assets/puzzle.png", width=54)
with col_title:
    st.markdown(
        "<div style='font-size:44px; font-weight:800; color:#6f42c1;'>"
        "DHF Automation – Infusion Pump</div>",
        unsafe_allow_html=True,
    )
    st.caption("Requirements → Hazard Analysis → DVP → TM")
with col_img1:
    st.image("streamlit_assets/Infusion.jpg", use_container_width=True)
with col_img2:
    st.image("streamlit_assets/Infusion1.jpg", use_container_width=True)

st.markdown("**Provide Product Requirements** (choose one):")
colA, colB = st.columns(2)
with colA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with colB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=140, placeholder=sample)

# Top, always-visible primary button
run_btn = st.button("🟩 Generate DHF Packages", type="primary")

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

    # Fill empties with TBD
    def fill_tbd(df: pd.DataFrame, cols: t.List[str]) -> pd.DataFrame:
        out = df.copy()
        for c in cols:
            if c not in out.columns:
                out[c] = TBD
            out[c] = out[c].fillna(TBD)
            out.loc[out[c].astype(str).str.strip().eq(""), c] = TBD
            out.loc[out[c].astype(str).str.upper().eq("NA"), c] = TBD
        return out

    ha_df = fill_tbd(
        ha_df,
        [
            "requirement_id", "risk_id", "risk_to_health", "hazard", "hazardous_situation",
            "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
        ],
    )
    # Remove the 'requirement_id' column from preview per earlier preference? (kept for exports)
    st.subheader(f"Hazard Analysis (preview, first {HA_MAX_ROWS})")
    st.dataframe(head_cap(ha_df.drop(columns=[], errors="ignore"), HA_MAX_ROWS), use_container_width=True)

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
    dvp_df = fill_tbd(dvp_df, ["verification_id", "requirement_id", "requirements",
                               "verification_method", "acceptance_criteria", "sample_size", "test_procedure"])
    st.subheader(f"Design Verification Protocol (preview, first {DVP_MAX_ROWS})")
    st.dataframe(head_cap(dvp_df, DVP_MAX_ROWS), use_container_width=True)

    # -------- Call Backend: TM --------
    with st.spinner("Building Trace Matrix (backend)..."):
        tm_payload = {
            "requirements": ha_payload["requirements"],
            "ha": ha_rows,
            "dvp": dvp_rows,
        }
        tm_resp = call_backend("/tm", tm_payload)  # backend returns {"ok":true,"tm":[...]}
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
        "Acceptance Criteria",
    ]
    tm_df = tm_df.reindex(columns=TM_ORDER)
    tm_df = fill_tbd(tm_df, TM_ORDER)

    st.subheader(f"Trace Matrix (preview, first {TM_MAX_ROWS})")
    st.dataframe(head_cap(tm_df, TM_MAX_ROWS), use_container_width=True)

    # -------- Evaluation Metrics (with thresholds) --------
    render_metrics(ha_df, dvp_df, tm_df)

    # -------- Prepare styled Excel exports (capped for preview) --------
    ha_export_cols = [
        "requirement_id", "risk_id", "risk_to_health", "hazard", "hazardous_situation",
        "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
    ]
    dvp_export_cols = ["verification_id", "requirement_id", "requirements",
                       "verification_method", "sample_size", "test_procedure", "acceptance_criteria"]
    tm_export_cols = TM_ORDER[:]  # same order as preview

    ha_bytes = df_to_excel_bytes(head_cap(ha_df[ha_export_cols], HA_MAX_ROWS), format_ha)
    dvp_bytes = df_to_excel_bytes(head_cap(dvp_df[dvp_export_cols], DVP_MAX_ROWS), format_dvp)
    tm_bytes  = df_to_excel_bytes(head_cap(tm_df[tm_export_cols], TM_MAX_ROWS), format_tm)

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
        st.markdown(
            "<div style='margin-top:8px; font-weight:800;'>"
            "DHF documents downloaded successfully — Hazard Analysis, Design Verification Protocol, Trace Matrix"
            "</div>",
            unsafe_allow_html=True,
        )

st.markdown("---")
st.caption(
    "Preview and exports are capped to: "
    f"HA={HA_MAX_ROWS}, DVP={DVP_MAX_ROWS}, TM={TM_MAX_ROWS}. "
    f"Sending only the first {REQ_MAX} requirements to backend."
)
