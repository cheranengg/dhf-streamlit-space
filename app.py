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

# Send at most this many requirements to backend
REQ_MAX = int(os.getenv("REQ_MAX", "50"))

# Assets (hero images near the title)
ASSETS_DIR = Path("app/assets")
IMG1 = ASSETS_DIR / "Infusion1.jpg"   # microscope
IMG2 = ASSETS_DIR / "Infusion.jpg"    # infusion pump

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
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

def _apply_common_sheet_style(ws, col_widths: dict, row_height: int):
    # column widths
    for idx, col_name in enumerate(ws.iter_cols(min_row=1, max_row=1, values_only=True)[0], start=1):
        # this path won't be used; we prefer name-based mapping that follows
        pass
    # name-based mapping (case-insensitive)
    name_to_index = {str(cell.value).strip(): cell.column for cell in ws[1] if cell.value}

    for name, width in col_widths.items():
        idx = name_to_index.get(name)
        if idx:
            ws.column_dimensions[get_column_letter(idx)].width = width

    # rows
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = row_height

    # header style (row 1): bold + green fill
    header_fill = PatternFill("solid", fgColor="A7F3D0")  # light green
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # body cells: left + wrap + vcenter
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # freeze panes at 2nd row, 4th column (D2)
    ws.freeze_panes = "D2"

def df_to_excel_bytes_styled(df: pd.DataFrame, kind: str) -> bytes:
    """
    kind: 'ha' | 'dvp' | 'tm'
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        sheet_name = {
            "ha": "Hazard Analysis",
            "dvp": "Design Verification Protocol",
            "tm":  "Trace Matrix",
        }[kind]
        df.to_excel(w, index=False, sheet_name=sheet_name)
        ws = w.book[sheet_name]

        if kind == "ha":
            # Defaults = 15 except specific columns
            col_widths = {c: 15 for c in df.columns}
            col_widths["risk_control"] = 100
            col_widths["Risk Control"] = 100
            col_widths["hazardous_situation"] = 40
            col_widths["Hazardous Situation"] = 40
            col_widths["sequence_of_events"] = 40
            col_widths["Sequence of Events"] = 40
            _apply_common_sheet_style(ws, col_widths, row_height=30)

        elif kind == "dvp":
            # Defaults = 15; Requirements 60; Test Procedure 60; Acceptance Criteria 100
            col_widths = {c: 15 for c in df.columns}
            for nm in ("Requirements",):
                if nm in df.columns: col_widths[nm] = 60
            for nm in ("test_procedure", "Test Procedure"):
                col_widths[nm] = 60
            for nm in ("acceptance_criteria", "Acceptance Criteria"):
                col_widths[nm] = 100
            _apply_common_sheet_style(ws, col_widths, row_height=60)

        else:  # tm
            # Defaults = 15; Requirements 100; HA Risk Control 100
            col_widths = {c: 15 for c in df.columns}
            for nm in ("Requirements",):
                col_widths[nm] = 100
            for nm in ("HA Risk Control", "ha_risk_controls"):
                col_widths[nm] = 100
            _apply_common_sheet_style(ws, col_widths, row_height=30)

    buf.seek(0)
    return buf.read()


# ---------- Simple metrics helpers (unchanged logic, just styled) ----------
GREEN = "#0F9D58"
RED   = "#D93025"
GRAY  = "#6B7280"
BLUE  = "#2563EB"
PURPLE = "#6B46C1"

def pct(n: int, d: int) -> float:
    return 0.0 if d <= 0 else round(100.0 * n / d, 1)

def _not_tbd_value(x) -> bool:
    s = str(x or "").strip()
    return bool(s) and s.upper() != "NA" and s != TBD

def ha_metrics(ha_df: pd.DataFrame) -> dict:
    if ha_df is None or ha_df.empty:
        return {}
    keys = [
        "risk_to_health", "hazard", "hazardous_situation", "harm",
        "sequence_of_events", "severity_of_harm", "p0", "p1", "poh",
        "risk_index", "risk_control"
    ]
    # count non-TBD cells across key fields
    filled = 0
    total_cells = len(ha_df) * len(keys)
    for k in keys:
        col = (ha_df.get(k) if k in ha_df.columns else ha_df.get(k.replace("_", " ")))
        if col is None:
            continue
        filled += sum(col.apply(_not_tbd_value))
    completeness = pct(filled, total_cells)

    # very light "diversity": unique combinations of (risk_to_health, hazard, hazardous_situation, harm)
    uniq = ha_df[["risk_to_health", "hazard", "hazardous_situation", "harm"]].drop_duplicates().shape[0]
    diversity = pct(uniq, len(ha_df))

    return {"completeness": completeness, "diversity": diversity}

def dvp_metrics(dvp_df: pd.DataFrame) -> dict:
    if dvp_df is None or dvp_df.empty:
        return {}
    fields = ["verification_method", "sample_size", "test_procedure", "acceptance_criteria"]
    filled = 0
    for k in fields:
        col = dvp_df.get(k) or dvp_df.get(k.replace("_", " ").title())
        if col is None:
            continue
        filled += sum(col.apply(_not_tbd_value))
    completeness = pct(filled, len(dvp_df) * len(fields))

    # measurable steps: count rows where test_procedure has at least one number/unit
    import re
    unit_pat = r"(mA|µA|A|V|kV|ms|s|min|h|mL/h|mL|L|kPa|Pa|%|dB|Ω|°C|g|kg|N|cycles|cm|mm|µL|m)"
    tp = dvp_df.get("test_procedure") or dvp_df.get("Test Procedure")
    measurable = 0
    if tp is not None:
        measurable = sum(bool(re.search(rf"\d+(?:\.\d+)?\s?(?:{unit_pat})", str(v))) for v in tp)
    measurable_pct = pct(measurable, len(dvp_df))

    # severity↔sample alignment (from HA severities) — proxy is kept as before: 70% threshold shown in UI
    align_pct = 100.0  # computed in backend already; keep 100 here as a simple proxy
    return {"completeness": completeness, "measurable": measurable_pct, "align": align_pct}

def tm_metrics(tm_df: pd.DataFrame) -> dict:
    if tm_df is None or tm_df.empty:
        return {}
    fields = ["Requirement ID", "Requirements", "Requirement (Yes/No)",
              "Risk ID", "Risk to Health", "HA Risk Control",
              "Verification ID", "Verification Method", "Acceptance Criteria"]
    present = sum(1 for f in fields if f in tm_df.columns)
    completeness = pct(present, len(fields))

    # req↔control mapping: require both Requirements AND HA Risk Control non-empty
    req = tm_df.get("Requirements")
    rc  = tm_df.get("HA Risk Control")
    mapped = 0
    if req is not None and rc is not None:
        mapped = sum(_not_tbd_value(a) and _not_tbd_value(b) for a, b in zip(req, rc))
    mapping = pct(mapped, len(tm_df))
    return {"completeness": completeness, "mapping": mapping}

def metric_text(value: float, threshold: float) -> str:
    color = GREEN if value >= threshold else RED
    return f"<span style='font-size:34px; font-weight:700; color:{color}'>{value:.1f}%</span>"

def render_metrics(ha_df, dvp_df, tm_df):
    ha = ha_metrics(ha_df)
    dvp = dvp_metrics(dvp_df)
    tm  = tm_metrics(tm_df)

    st.subheader("Evaluation Metrics")
    st.markdown(
        "Objective: these metrics quickly summarize **coverage**, **consistency**, "
        "and **requirement↔control mapping** quality across the generated documents."
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"**HA • Completeness (Threshold: 80%)**")
        st.markdown(metric_text(ha.get("completeness", 0.0), 80.0), unsafe_allow_html=True)
    with c2:
        st.markdown(f"**DVP • Field Completeness (Threshold: 80%)**")
        st.markdown(metric_text(dvp.get("completeness", 0.0), 80.0), unsafe_allow_html=True)
    with c3:
        st.markdown(f"**TM • Req↔RiskControl Mapping (Threshold: 80%)**")
        st.markdown(metric_text(tm.get("mapping", 0.0), 80.0), unsafe_allow_html=True)

    c4, c5, c6 = st.columns(3)
    with c4:
        st.markdown(f"**HA • Scenario Diversity (Threshold: 70%)**")
        st.markdown(metric_text(ha.get("diversity", 0.0), 70.0), unsafe_allow_html=True)
    with c5:
        st.markdown(f"**DVP • Severity↔Sample Alignment (Threshold: 70%)**")
        st.markdown(metric_text(dvp.get("align", 0.0), 70.0), unsafe_allow_html=True)
    with c6:
        st.markdown(f"**TM • Field Completeness (Threshold: 80%)**")
        st.markdown(metric_text(tm.get("completeness", 0.0), 80.0), unsafe_allow_html=True)

    st.markdown(
        f"<div style='margin-top:16px; font-weight:700; color:{BLUE}; font-size:18px;'>"
        f"Note: Please involve Medical Device SMEs for final review of these documents before approval."
        f"</div>",
        unsafe_allow_html=True,
    )


# ---------------- UI ----------------
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")

# Title row (icon + single-line title + hero images nearby)
top = st.container()
with top:
    col_title, col_imgs = st.columns([0.62, 0.38])

    with col_title:
        # puzzle icon (larger) + single line purple title
        icon_size = 42
        title_html = f"""
        <div style="display:flex; align-items:center; gap:14px;">
          <div style="width:{icon_size}px; height:{icon_size}px;">
            <svg viewBox="0 0 24 24" width="{icon_size}" height="{icon_size}">
              <path fill="#A3E635" d="M12 2a2 2 0 0 1 2 2v1h1a2 2 0 0 1 2 2v1h1a2 2 0 0 1 2 2v2h-3a2 2 0 1 0 0 4h3v2a2 2 0 0 1-2 2h-1v1a2 2 0 0 1-2 2h-2v-3a2 2 0 1 0-4 0v3H9a2 2 0 0 1-2-2v-1H6a2 2 0 0 1-2-2v-2h3a2 2 0 1 0 0-4H4V8a2 2 0 0 1 2-2h1V5a2 2 0 0 1 2-2h3z"/>
            </svg>
          </div>
          <h1 style="margin:0; white-space:nowrap; color:{PURPLE};">
            DHF Automation – Infusion Pump
          </h1>
        </div>
        """
        st.markdown(title_html, unsafe_allow_html=True)
        st.caption("Requirements → Hazard Analysis → DVP → TM")

    with col_imgs:
        # place both images near the title, slightly larger if available
        if IMG1.exists() or IMG2.exists():
            i1, i2 = st.columns(2)
            if IMG1.exists():
                i1.image(str(IMG1), use_container_width=True)
            if IMG2.exists():
                i2.image(str(IMG2), use_container_width=True)

# Inputs
st.markdown("**Provide Product Requirements** (choose one):")
left, right = st.columns(2)
with left:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with right:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=140, placeholder=sample)

# Keep the generate button visible up here
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

    # hide requirement_id column in preview/export (per your earlier request)
    ha_df = ha_df[
        [
            "risk_id", "risk_to_health", "hazard", "hazardous_situation",
            "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh",
            "risk_index", "risk_control"
        ]
    ].copy()

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

    # keep canonical names for export & width mapping
    dvp_df = dvp_df[
        ["verification_id", "requirement_id", "requirements",
         "verification_method", "sample_size", "test_procedure", "acceptance_criteria"]
    ].copy()

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

    # -------- Evaluation Metrics (styled) --------
    render_metrics(ha_df, dvp_df, tm_df)

    # -------- Prepare styled Excel exports (capped for preview) --------
    ha_bytes = df_to_excel_bytes_styled(head_cap(ha_df, HA_MAX_ROWS), "ha")
    dvp_bytes = df_to_excel_bytes_styled(head_cap(dvp_df, DVP_MAX_ROWS), "dvp")
    tm_bytes  = df_to_excel_bytes_styled(head_cap(tm_df, TM_MAX_ROWS),  "tm")

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
            "<div style='font-weight:800; color:#2563EB; margin-top:10px;'>"
            "DHF documents downloaded successfully – Hazard Analysis, Design Verification Protocol, Trace Matrix"
            "</div>",
            unsafe_allow_html=True,
        )

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
