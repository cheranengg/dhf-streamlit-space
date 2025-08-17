# app.py — DHF Streamlit UI
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
TBD_CANON = {TBD, "NA", "N/A", "NONE", "NULL", ""}

BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8080")
BACKEND_TOKEN = st.secrets.get("BACKEND_TOKEN", "dev-token")

# Row caps (preview + Excel export)
HA_MAX_ROWS = int(os.getenv("HA_MAX_ROWS", "50"))
DVP_MAX_ROWS = int(os.getenv("DVP_MAX_ROWS", "50"))
TM_MAX_ROWS  = int(os.getenv("TM_MAX_ROWS", "50"))

# how many requirements to send to backend
REQ_MAX = int(os.getenv("REQ_MAX", "50"))

OUTPUT_DIR = os.path.abspath("./streamlit_outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ----------------- Page setup -----------------
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")
left, right = st.columns([0.63, 0.37])

with left:
    st.title("🧩 DHF Automation – Infusion Pump")
    st.caption("Requirements → Hazard Analysis → DVP → TM")

with right:
    # show images side-by-side under the title, medium size, no deprecation warnings
    c1, c2 = st.columns(2, gap="small")
    assets_dir = Path("app/assets")
    img1 = str(assets_dir / "Infusion1.jpg")  # microscope
    img2 = str(assets_dir / "Infusion.jpg")   # infusion pump
    with c1:
        if Path(img1).exists():
            st.image(img1, use_container_width=True)
    with c2:
        if Path(img2).exists():
            st.image(img2, use_container_width=True)

st.markdown("")

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

def fill_tbd(df: pd.DataFrame, cols: t.List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = TBD
        out[c] = out[c].fillna(TBD)
        out.loc[out[c].astype(str).str.strip().eq(""), c] = TBD
        out.loc[out[c].astype(str).str.upper().isin({"NA","N/A"}), c] = TBD
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

# ---------- Excel styling ----------
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows

HEADER_FILL = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")  # green
HEADER_FONT = Font(bold=True)
ALIGN_LEFT_CENTER = Alignment(horizontal="left", vertical="center", wrap_text=True)

def _apply_common_sheet_style(ws, col_widths: dict, row_height: int = 30, freeze_cell: str = "D2"):
    # header style
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_LEFT_CENTER
    # column widths
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width
    # row heights
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = row_height
    # wrap + left align all cells
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = ALIGN_LEFT_CENTER
    # freeze panes
    ws.freeze_panes = freeze_cell

def df_to_excel_bytes_styled(df: pd.DataFrame, col_width_map: t.Dict[str, int], freeze_cell: str = "D2") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    # Build address map for widths
    # Map displayed column letters to desired widths by header name
    from openpyxl.utils import get_column_letter
    width_by_letter = {}
    for idx, col in enumerate(df.columns, start=1):
        letter = get_column_letter(idx)
        width_by_letter[letter] = col_width_map.get(col, 15)
    _apply_common_sheet_style(ws, width_by_letter, row_height=30, freeze_cell=freeze_cell)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ---------- Metrics ----------
def _pct(n: int, d: int) -> float:
    return 0.0 if d <= 0 else round(100.0 * n / max(d, 1), 1)

def _not_tbd_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip()
    return (~s.isin(TBD_CANON)) & s.ne("")

def ha_metrics(ha_df: pd.DataFrame) -> dict:
    if ha_df is None or ha_df.empty:
        return {}
    keys = ["risk_to_health","hazard","hazardous_situation","harm",
            "sequence_of_events","severity_of_harm","p0","p1","poh","risk_index","risk_control"]
    total_cells = len(ha_df) * len(keys)
    filled = 0
    for k in keys:
        if k in ha_df.columns:
            filled += int(_not_tbd_series(ha_df[k]).sum())
    completeness = _pct(filled, total_cells)
    diversity = _pct(
        len(
            ha_df[["risk_to_health","hazard","hazardous_situation","harm"]]
            .astype(str).drop_duplicates()
        ),
        len(ha_df)
    )
    harm_specificity = _pct(
        int(ha_df["harm"].astype(str).str.len().ge(8).sum()) if "harm" in ha_df.columns else 0,
        len(ha_df)
    )
    return {
        "Completeness": completeness,    # Threshold 90%
        "Scenario Diversity": diversity,  # Threshold 60%
        "Harm Specificity": harm_specificity,  # Threshold 70%
    }

def dvp_metrics(dvp_df: pd.DataFrame) -> dict:
    if dvp_df is None or dvp_df.empty:
        return {}
    has_meas = dvp_df.get("test_procedure", pd.Series(dtype=str)).astype(str).str.contains(r"\d").sum()
    proc_quality = _pct(int(has_meas), len(dvp_df))
    method_coverage = _pct(
        int(_not_tbd_series(dvp_df.get("verification_method", pd.Series(dtype=str))).sum()),
        len(dvp_df)
    )
    has_accept = _pct(
        int(_not_tbd_series(dvp_df.get("acceptance_criteria", pd.Series(dtype=str))).sum()),
        len(dvp_df)
    )
    return {
        "Procedure Quality": proc_quality,        # Threshold 70%
        "Method Coverage": method_coverage,       # Threshold 90%
        "Acceptance Coverage": has_accept,        # Threshold 95%
    }

def tm_metrics(tm_df: pd.DataFrame) -> dict:
    if tm_df is None or tm_df.empty:
        return {}
    # risk control ↔ requirement mapping %
    rc = tm_df.get("HA Risk Control", tm_df.get("ha_risk_controls", pd.Series(dtype=str))).astype(str)
    req = tm_df.get("Requirements", tm_df.get("requirements", pd.Series(dtype=str))).astype(str)
    mapped = (rc.str.len().ge(6) & req.str.len().ge(6)).sum()
    mapping = _pct(int(mapped), len(tm_df))
    # coverage (Verification ID present)
    coverage = _pct(int(_not_tbd_series(tm_df.get("Verification ID", tm_df.get("verification_id", pd.Series(dtype=str)))).sum()), len(tm_df))
    return {
        "Risk-Control Mapping": mapping,  # Threshold 80%
        "Verification Coverage": coverage # Threshold 95%
    }

def render_metrics(ha_df, dvp_df, tm_df):
    ha = ha_metrics(ha_df)
    dvp = dvp_metrics(dvp_df)
    tm  = tm_metrics(tm_df)

    st.subheader("Evaluation Metrics")
    st.caption("Objective: quick, data-driven checks to highlight completeness and consistency of the generated DHF artifacts.")
    st.caption("Please involve Medical Device SMEs for final review of these documents before approval.")

    def line(label, val, thr):
        st.write(f"- **{label}**: {val:.1f}%  _(Threshold: {thr}%)_")

    colA, colB, colC = st.columns(3)
    with colA:
        st.markdown("**Hazard Analysis**")
        if ha: 
            line("Completeness", ha["Completeness"], 90)
            line("Scenario Diversity", ha["Scenario Diversity"], 60)
            line("Harm Specificity", ha["Harm Specificity"], 70)
        else:
            st.write("_No rows_")

    with colB:
        st.markdown("**Design Verification Protocol**")
        if dvp:
            line("Procedure Quality", dvp["Procedure Quality"], 70)
            line("Method Coverage", dvp["Method Coverage"], 90)
            line("Acceptance Coverage", dvp["Acceptance Coverage"], 95)
        else:
            st.write("_No rows_")

    with colC:
        st.markdown("**Trace Matrix**")
        if tm:
            line("Risk-Control Mapping", tm["Risk-Control Mapping"], 80)
            line("Verification Coverage", tm["Verification Coverage"], 95)
        else:
            st.write("_No rows_")

# ---------------- Inputs (top, single button) ----------------
st.markdown("**Provide Product Requirements** (choose one):")
topA, topB = st.columns([0.55, 0.45])

with topA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with topB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=120, placeholder=sample)

run_btn = st.button("▶️ Generate DHF Packages", type="primary")

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
        try:
            ha_resp = call_backend("/hazard-analysis", ha_payload)
        except Exception as e:
            st.error(f"HA backend error: {e}")
            st.stop()
        ha_rows = ha_resp.get("ha", [])
        ha_df = pd.DataFrame(ha_rows)

    ha_df = fill_tbd(
        ha_df,
        [
            "risk_id", "risk_to_health", "hazard", "hazardous_situation",
            "harm", "sequence_of_events", "severity_of_harm",
            "p0", "p1", "poh", "risk_index", "risk_control",
        ],
    )
    st.subheader(f"Hazard Analysis (preview, first {HA_MAX_ROWS})")
    st.dataframe(head_cap(ha_df, HA_MAX_ROWS), use_container_width=True)

    # -------- Call Backend: DVP --------
    with st.spinner("Generating Design Verification Protocol (backend)..."):
        dvp_payload = {
            "requirements": ha_payload["requirements"],  # limited
            "ha": ha_rows,
        }
        try:
            dvp_resp = call_backend("/dvp", dvp_payload)
            dvp_rows = dvp_resp.get("dvp", [])
        except Exception as e:
            st.error(f"DVP backend error: {e}")
            st.stop()
    dvp_df = pd.DataFrame(dvp_rows)
    dvp_df = fill_tbd(dvp_df, [
        "verification_id","requirement_id","requirements",
        "verification_method","sample_size","test_procedure","acceptance_criteria"
    ])
    st.subheader(f"Design Verification Protocol (preview, first {DVP_MAX_ROWS})")
    st.dataframe(head_cap(dvp_df, DVP_MAX_ROWS), use_container_width=True)

    # -------- Call Backend: TM --------
    with st.spinner("Building Trace Matrix (backend)..."):
        tm_payload = {
            "requirements": ha_payload["requirements"],
            "ha": ha_rows,
            "dvp": dvp_rows,
        }
        try:
            tm_resp = call_backend("/tm", tm_payload)  # backend returns {"ok":true,"tm":[...]}
            tm_rows = tm_resp.get("tm", [])
        except Exception as e:
            st.error(f"TM backend error: {e}")
            st.stop()

    tm_df = pd.DataFrame(tm_rows)

    # Exact TM column order & names for preview/export
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
    # If backend uses snake_case, rename to Title Case
    tm_df = tm_df.rename(columns={
        "requirement_id": "Requirement ID",
        "requirements": "Requirements",
        "Requirement (Yes/No)": "Requirement (Yes/No)",
        "risk_ids": "Risk ID",
        "risks_to_health": "Risk to Health",
        "ha_risk_controls": "HA Risk Control",
        "verification_id": "Verification ID",
        "verification_method": "Verification Method",
    })
    tm_df = tm_df.reindex(columns=TM_ORDER)
    tm_df = fill_tbd(tm_df, TM_ORDER)

    st.subheader(f"Trace Matrix (preview, first {TM_MAX_ROWS})")
    st.dataframe(head_cap(tm_df, TM_MAX_ROWS), use_container_width=True)

    # -------- Evaluation Metrics (with thresholds) --------
    render_metrics(ha_df, dvp_df, tm_df)

    # -------- Prepare styled Excel exports (capped for preview) --------
    # HA export
    ha_export_cols = [
        "risk_id", "risk_to_health", "hazard", "hazardous_situation",
        "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
    ]
    ha_x = head_cap(ha_df[ha_export_cols], HA_MAX_ROWS).copy()
    # Column widths: 15 each except Risk Control = 100
    ha_widths = {c: 15 for c in ha_x.columns}
    if "risk_control" in ha_x.columns:
        ha_widths["risk_control"] = 100
    ha_bytes = df_to_excel_bytes_styled(ha_x, ha_widths, freeze_cell="D2")

    # DVP export (order + titles like your screenshot)
    dvp_x = head_cap(dvp_df, DVP_MAX_ROWS).copy().rename(columns={
        "verification_id": "Verification ID",
        "requirement_id": "Requirement ID",
        "requirements": "Requirements",
        "verification_method": "Verification Method",
        "sample_size": "Sample size",
        "test_procedure": "Test Procedure",
        "acceptance_criteria": "Acceptance criteria",
    })
    dvp_order = [
        "Verification ID","Requirement ID","Requirements",
        "Verification Method","Sample size","Test Procedure","Acceptance criteria"
    ]
    dvp_x = dvp_x.reindex(columns=dvp_order)
    # Widths: 15 for all except Acceptance criteria & Test Procedure = 100
    dvp_widths = {c: 15 for c in dvp_order}
    dvp_widths["Acceptance criteria"] = 100
    dvp_widths["Test Procedure"] = 100
    dvp_bytes = df_to_excel_bytes_styled(dvp_x, dvp_widths, freeze_cell="D2")

    # TM export (order already set). Set Requirements & HA Risk Control = 100
    tm_x = head_cap(tm_df, TM_MAX_ROWS).copy()
    tm_widths = {c: 15 for c in tm_x.columns}
    if "Requirements" in tm_x.columns:
        tm_widths["Requirements"] = 100
    if "HA Risk Control" in tm_x.columns:
        tm_widths["HA Risk Control"] = 100
    tm_bytes  = df_to_excel_bytes_styled(tm_x, tm_widths, freeze_cell="D2")

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

    clicked = st.download_button(
        "⬇️ Download DHF package (3 Excel files, ZIP)",
        data=zip_buf.getvalue(),
        file_name="DHF_Package.zip",
        mime="application/zip",
        type="primary"
    )
    if clicked:
        st.success("DHF documents downloaded successfully — Hazard Analysis, Design Verification Protocol, Trace Matrix.")
        st.info("Note: Please involve Medical Device SMEs for final review of these documents before approval.")
