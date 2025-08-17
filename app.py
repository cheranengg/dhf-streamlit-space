# app.py — DHF Streamlit UI (Preview + ZIP Download, puzzle icon header)
# ---------------------------------------------------------------------

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

BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8080")
BACKEND_TOKEN = st.secrets.get("BACKEND_TOKEN", "dev-token")

# Row caps (preview + Excel export)
HA_MAX_ROWS = int(os.getenv("HA_MAX_ROWS", "50"))
DVP_MAX_ROWS = int(os.getenv("DVP_MAX_ROWS", "50"))
TM_MAX_ROWS  = int(os.getenv("TM_MAX_ROWS", "50"))

# NEW: cap how many requirements we send to backend
REQ_MAX = int(os.getenv("REQ_MAX", "50"))

# Optional Google Drive upload (left intact, but unused unless you set secrets)
DEFAULT_DRIVE_FOLDER_ID = st.secrets.get("DRIVE_FOLDER_ID", "")
OUTPUT_DIR = st.secrets.get("OUTPUT_DIR", os.path.abspath("./streamlit_outputs"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ------------ UI constants (header + images) ----------
TITLE_PURPLE = "#6b42cc"
PUZZLE_ICON_PATH = "streamlit_assets/puzzle.png"
IMG1_PATH = "streamlit_assets/Infusion1.jpg"
IMG2_PATH = "streamlit_assets/Infusion.jpg"

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


def fill_tbd(df: pd.DataFrame, cols: t.List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = TBD
        out[c] = out[c].fillna(TBD)
        out.loc[out[c].astype(str).str.strip().eq(""), c] = TBD
        out.loc[out[c].astype(str).str.upper().eq("NA"), c] = TBD
    return out


# ---------------- Excel styling helpers ----------------
def _style_sheet(ws, col_widths: dict, row_height: int = 30):
    from openpyxl.styles import Alignment, PatternFill, Font
    # Header row: bold, green background
    header_fill = PatternFill(start_color="B7E1CD", end_color="B7E1CD", fill_type="solid")
    header_font = Font(bold=True)
    align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Freeze pane at 2nd row, 4th column
    ws.freeze_panes = ws.cell(row=2, column=4)

    # Column widths
    for col_name, width in col_widths.items():
        if col_name in ws[1]:
            # If someone passes "A" etc. we handle by header match; simplest: iterate columns
            pass
    # Set widths by header text
    header_map = {cell.value: cell.column_letter for cell in ws[1] if cell.value}
    for name, width in col_widths.items():
        col_letter = header_map.get(name)
        if col_letter:
            ws.column_dimensions[col_letter].width = width

    # Row heights + alignment
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=False):
        ws.row_dimensions[row[0].row].height = row_height
        for cell in row:
            cell.alignment = align

    # Header styling
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font


def df_to_excel_bytes_styled(df: pd.DataFrame, col_widths: dict, row_height: int = 30) -> bytes:
    from openpyxl import Workbook
    buf = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Write header
    ws.append(list(df.columns))
    # Write data
    for _, r in df.iterrows():
        ws.append(list(r.values))

    _style_sheet(ws, col_widths=col_widths, row_height=row_height)
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------- Metrics (unchanged structure, KPI coloring) ----------------
TBD_CANON = {TBD, "TBD", "NA", "N/A", ""}

def _pct(num: int, den: int) -> float:
    if den <= 0:
        return 0.0
    return round(100.0 * num / den, 1)

def _not_tbd_val(x) -> bool:
    s = str(x or "").strip()
    return bool(s) and s not in TBD_CANON

def ha_metrics(ha_df: pd.DataFrame) -> dict:
    if ha_df is None or ha_df.empty:
        return {"completeness": 0.0, "diversity": 0.0}

    keys = ["risk_to_health", "hazard", "hazardous_situation", "harm",
            "sequence_of_events", "severity_of_harm", "p0", "p1", "poh",
            "risk_index", "risk_control"]
    filled = 0
    for k in keys:
        filled += int(ha_df[k].apply(_not_tbd_val).sum()) if k in ha_df.columns else 0
    completeness = _pct(filled, len(ha_df) * len(keys))

    # Scenario diversity – distinct tuples across three fields
    uniq = ha_df[["risk_to_health", "hazard", "hazardous_situation"]].drop_duplicates().shape[0]
    diversity = _pct(uniq, len(ha_df))

    return {"completeness": completeness, "diversity": diversity}

def dvp_metrics(dvp_df: pd.DataFrame) -> dict:
    if dvp_df is None or dvp_df.empty:
        return {"field_complete": 0.0, "sev_sample": 0.0}

    keys = ["verification_method", "sample_size", "test_procedure", "acceptance_criteria"]
    filled = 0
    for k in keys:
        filled += int(dvp_df[k].apply(_not_tbd_val).sum()) if k in dvp_df.columns else 0
    field_complete = _pct(filled, len(dvp_df) * len(keys))

    # Severity ↔ sample size alignment — simplistic: numeric sample present => count it
    def _ok_sample(x: str) -> bool:
        try:
            return int(str(x).strip()) > 0
        except Exception:
            return False
    sev_sample = _pct(int(dvp_df["sample_size"].apply(_ok_sample).sum()), len(dvp_df))

    return {"field_complete": field_complete, "sev_sample": sev_sample}

def tm_metrics(tm_df: pd.DataFrame) -> dict:
    if tm_df is None or tm_df.empty:
        return {"map_pct": 0.0, "field_complete": 0.0}

    # Mapping % : requirement text shares ≥1 meaningful token with HA Risk Control
    req_col = "Requirements"
    rc_col  = "HA Risk Control"
    def _tokens(s: str) -> set:
        if not isinstance(s, str):
            return set()
        toks = [t.lower() for t in s.replace("/", " ").replace(",", " ").split()]
        return {t for t in toks if len(t) >= 4 and t not in {"shall", "system", "device", "patient"}}

    mapped = 0
    for _, r in tm_df.iterrows():
        req = _tokens(str(r.get(req_col, "")))
        rc  = _tokens(str(r.get(rc_col, "")))
        if req and rc and (req & rc):
            mapped += 1
    map_pct = _pct(mapped, len(tm_df))

    # Field completeness for TM core columns
    keys = ["Requirement ID", "Requirements", "Risk ID", "Risk to Health",
            "HA Risk Control", "Verification ID", "Verification Method", "Acceptance Criteria"]
    filled = 0
    for k in keys:
        filled += int(tm_df[k].apply(_not_tbd_val).sum()) if k in tm_df.columns else 0
    field_complete = _pct(filled, len(tm_df) * len(keys))

    return {"map_pct": map_pct, "field_complete": field_complete}

def _metric_line(label: str, value: float, threshold: float) -> None:
    color = "#198754" if value >= threshold else "#C1121F"  # green or red
    st.markdown(
        f"""
        <div style="font-size:22px;margin-top:8px;">
          <span style="opacity:0.75">{label} (Threshold: {int(threshold)}%)</span><br/>
          <span style="font-size:44px;font-weight:800;color:{color};">{value:.1f}%</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

def render_metrics(ha_df: pd.DataFrame, dvp_df: pd.DataFrame, tm_df: pd.DataFrame):
    ha = ha_metrics(ha_df)
    dvp = dvp_metrics(dvp_df)
    tm  = tm_metrics(tm_df)

    st.subheader("Evaluation Metrics")
    st.markdown(
        "Objective: these metrics quickly summarize **coverage**, **consistency**, and "
        "**requirement↔control mapping** quality across the generated documents."
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        _metric_line("HA • Completeness", ha["completeness"], 80)
    with c2:
        _metric_line("DVP • Field Completeness", dvp["field_complete"], 80)
    with c3:
        _metric_line("TM • Req↔RiskControl Mapping", tm["map_pct"], 80)

    c4, c5, c6 = st.columns(3)
    with c4:
        _metric_line("HA • Scenario Diversity", ha["diversity"], 70)
    with c5:
        _metric_line("DVP • Severity↔Sample Alignment", dvp["sev_sample"], 70)
    with c6:
        _metric_line("TM • Field Completeness", tm["field_complete"], 80)

    st.markdown(
        '<div style="margin-top:12px;font-weight:800;color:#0b63d1;">'
        'Note: Please involve Medical Device SMEs for final review of these documents before approval.'
        '</div>',
        unsafe_allow_html=True,
    )


# ---------------- Page setup ----------------
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")

# ---- Header (puzzle icon + single-line purple title + images on right) ----
st.markdown(
    f"""
    <style>
      .title-wrap {{
        display: grid;
        grid-template-columns: auto 1fr auto auto;
        align-items: center;
        column-gap: 16px;
        margin: 8px 0 0 0;
      }}
      .title-text {{
        font-size: 56px;
        font-weight: 800;
        color: {TITLE_PURPLE};
        line-height: 1.0;
        margin: 0;
        white-space: nowrap;
      }}
      .title-icon img {{
        height: 56px;
        width: auto;
        display: block;
      }}
      @media (max-width: 1100px) {{
        .title-text {{ font-size: 48px; }}
        .title-icon img {{ height: 48px; }}
      }}
      /* Green primary button */
      .stButton > button {{
        background: #22c55e !important;
        color: white !important;
        border: 0 !important;
        padding: 0.6rem 1.0rem !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
      }}
    </style>
    """,
    unsafe_allow_html=True,
)

left, mid, img1_col, img2_col = st.columns([0.07, 0.63, 0.15, 0.15])

with left:
    st.markdown('<div class="title-icon">', unsafe_allow_html=True)
    try:
        st.image(PUZZLE_ICON_PATH, use_container_width=False)
    except Exception:
        st.write("")  # ignore if asset missing
    st.markdown("</div>", unsafe_allow_html=True)

with mid:
    st.markdown('<div class="title-wrap">', unsafe_allow_html=True)
    st.markdown('<h1 class="title-text">DHF Automation – Infusion Pump</h1>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with img1_col:
    try:
        st.image(IMG1_PATH, use_container_width=True)
    except Exception:
        pass
with img2_col:
    try:
        st.image(IMG2_PATH, use_container_width=True)
    except Exception:
        pass

st.caption("Requirements → Hazard Analysis → DVP → TM")

# ---------------- Inputs ----------------
st.markdown("**Provide Product Requirements** (choose one):")
colA, colB = st.columns(2)
with colA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with colB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=140, placeholder=sample)

# Single, always-visible generate button (top)
generate_clicked = st.button("Generate DHF Packages", use_container_width=False)

if generate_clicked:
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

    ha_df = fill_tbd(
        ha_df,
        [
            "requirement_id", "risk_id", "risk_to_health", "hazard", "hazardous_situation",
            "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
        ],
    )
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
    dvp_df = fill_tbd(dvp_df, ["verification_id", "requirement_id", "requirements",
                               "verification_method", "sample_size", "test_procedure", "acceptance_criteria"])
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

    # Exact TM column order & names for export
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

    # -------- Evaluation Metrics --------
    render_metrics(ha_df, dvp_df, tm_df)

    # -------- Prepare styled Excel exports (capped for preview) --------
    ha_export_cols = [
        "requirement_id", "risk_id", "risk_to_health", "hazard", "hazardous_situation",
        "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
    ]
    dvp_export_cols = ["verification_id", "requirement_id", "requirements",
                       "verification_method", "sample_size", "test_procedure", "acceptance_criteria"]
    tm_export_cols = TM_ORDER[:]  # same order as preview

    ha_bytes = df_to_excel_bytes_styled(
        head_cap(ha_df[ha_export_cols], HA_MAX_ROWS),
        col_widths={
            # All 15 except:
            "hazardous_situation": 40,
            "sequence_of_events": 40,
            "risk_control": 100,
        },
        row_height=30,
    )

    dvp_bytes = df_to_excel_bytes_styled(
        head_cap(dvp_df[dvp_export_cols], DVP_MAX_ROWS),
        col_widths={
            # All 15, except requirements & test_procedure at 60
            "requirements": 60,
            "test_procedure": 60,
        },
        row_height=60,  # requested row height for DVP
    )

    tm_bytes = df_to_excel_bytes_styled(
        head_cap(tm_df[tm_export_cols], TM_MAX_ROWS),
        col_widths={
            # All 15, except:
            "Requirements": 100,
            "HA Risk Control": 100,
        },
        row_height=30,
    )

    # Save locally (optional)
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

    if st.download_button(
        "⬇️ Download DHF package (3 Excel files, ZIP)",
        data=zip_buf.getvalue(),
        file_name="DHF_Package.zip",
        mime="application/zip",
        type="primary"
    ):
        st.markdown(
            '<div style="font-weight:800;margin-top:6px;">'
            'DHF documents downloaded successfully - Hazard Analysis, Design Verification Protocol, Trace Matrix'
            '</div>',
            unsafe_allow_html=True,
        )
