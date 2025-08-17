# app.py — DHF Streamlit UI (Preview + ZIP Download) with metric thresholds

import os
import io
import re
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

# How many requirements to send to backend
REQ_MAX = int(os.getenv("REQ_MAX", "50"))

# Assets (images)
ASSET_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "streamlit_assets"))
INFUSION_IMG  = os.path.join(ASSET_DIR, "Infusion.jpg")
INFUSION_IMG1 = os.path.join(ASSET_DIR, "Infusion1.jpg")

# Output dir (optional local save area)
OUTPUT_DIR = st.secrets.get("OUTPUT_DIR", os.path.abspath("./streamlit_outputs"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- Thresholds ----------------
# Tweak these anytime; they render as "(Threshold: xx%)" and drive the delta coloring.
THRESHOLDS = {
    "HA": {
        "Completeness": 80,
        "Diversity": 50,
        "Severity coverage": 80,
        "Specific controls": 60,
    },
    "DVP": {
        "Measurable procedures": 70,
        "Acceptance measurability": 70,
        "Method assigned": 90,
    },
    "TM": {
        "Linkage rate": 90,
        "Column coverage": 85,
        "Req ↔ Risk Control mapping": 60,
    },
}

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


# ---------------- Styled Excel writer ----------------
def styled_excel_bytes(df: pd.DataFrame, col_widths: dict, freeze_cell: str = "D2",
                       wrap=True, header_fill="00C6F6D5",  # soft green
                       all_cols_default_width=15, row_height=30,
                       align="left", valign="center") -> bytes:
    """
    One-sheet Excel file with formatting:
      - column widths (override via col_widths)
      - row height
      - wrap text
      - header bold + fill
      - freeze panes at `freeze_cell`
    """
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Write header + data
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append(list(row.values))

    # Header style
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor=header_fill)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill

    # Freeze
    ws.freeze_panes = freeze_cell  # "D2" => freeze top row + first 3 cols

    # Column widths + alignment + wrapping
    alignment = Alignment(horizontal=align, vertical=valign, wrap_text=bool(wrap))
    for col_idx, col_name in enumerate(df.columns, start=1):
        width = col_widths.get(col_name, all_cols_default_width)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
        for r in range(1, len(df) + 2):  # include header row
            cell = ws.cell(row=r, column=col_idx)
            cell.alignment = alignment

    # Row heights
    for r in range(1, len(df) + 2):
        ws.row_dimensions[r].height = row_height

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------- Metric helpers ----------------
TBD_CANON = {TBD, "TBD", "NA", "N/A", "", None}

_UNIT_RE = re.compile(
    r"[±]?\d+(?:\.\d+)?\s?(mA|µA|A|V|kV|ms|s|min|h|mL/h|mL|L|kPa|Pa|%|dB|Ω|°C|g|kg|N|cycles|cm|mm|µL|m)\b",
    re.IGNORECASE,
)

STOP = set("""
a an the and or of to in for with on at from by is are was were be being been as that this these those shall will must should may can could within across per via using than under over
system device pump flow alarm accuracy test inspect verify measure calibration detection air occlusion leakage dielectric insulation earth volume label software usability emissions immunity
""".split())

def pct(n, d):
    return 0 if d <= 0 else round(100.0 * n / d, 1)

def _not_tbd(x) -> bool:
    s = str(x or "").strip()
    return bool(s) and s not in TBD_CANON

def _has_numbers_units(text: str) -> bool:
    return bool(_UNIT_RE.search(str(text or "")))

def _is_specific(text: str) -> bool:
    t = str(text or "").strip()
    if len(t) < 30:
        return False
    if re.search(r"\bIEC|ISO|UL|FDA|62304|60601|80369|11135|61000\b", t):
        return True
    return bool(re.search(r"\b(test|verify|calibrat|inspect|measure|trigger|alarm)\b", t, re.I))

def _tokens(s: str) -> set:
    toks = re.findall(r"[A-Za-z]+", str(s or "").lower())
    return {t for t in toks if t not in STOP and len(t) > 2}

def ha_metrics(ha_df: pd.DataFrame) -> dict:
    if ha_df is None or ha_df.empty:
        return {}
    keys = ["risk_to_health", "hazard", "hazardous_situation", "harm", "risk_control",
            "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index"]
    filled = sum(_not_tbd(ha_df.get(k, "")) for k in keys for _ in ha_df.index)
    completeness = pct(filled, len(ha_df) * len(keys))
    uniq = ha_df[["risk_to_health", "hazard", "hazardous_situation", "harm"]].drop_duplicates().shape[0]
    diversity = pct(uniq, len(ha_df))
    sev_ok = ha_df["severity_of_harm"].astype(str).str.fullmatch(r"[1-5]").sum()
    sev_cov = pct(sev_ok, len(ha_df))
    ctrl_spec = ha_df["risk_control"].apply(_is_specific).sum()
    ctrl_score = pct(ctrl_spec, len(ha_df))
    return {
        "Completeness": completeness,
        "Diversity": diversity,
        "Severity coverage": sev_cov,
        "Specific controls": ctrl_score,
    }

def dvp_metrics(dvp_df: pd.DataFrame) -> dict:
    if dvp_df is None or dvp_df.empty:
        return {}
    proc_score = pct(dvp_df["test_procedure"].apply(_has_numbers_units).sum(), len(dvp_df))
    ac_score   = pct(dvp_df["acceptance_criteria"].apply(_has_numbers_units).sum(), len(dvp_df))
    method_ok  = pct(dvp_df["verification_method"].apply(_not_tbd).sum(), len(dvp_df))
    return {
        "Measurable procedures": proc_score,
        "Acceptance measurability": ac_score,
        "Method assigned": method_ok,
    }

def _overlap_ratio(a: str, b: str) -> float:
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return 0.0
    inter = len(ta & tb)
    uni = len(ta | tb)
    return inter / uni if uni else 0.0

def tm_metrics(tm_df: pd.DataFrame) -> dict:
    if tm_df is None or tm_df.empty:
        return {}
    # Linkage rate (Verification ID present)
    link_score = pct(tm_df["Verification ID"].apply(_not_tbd).sum(), len(tm_df))
    # Column coverage
    tm_cols = ["Requirement ID","Requirements","Requirement (Yes/No)","Risk ID",
               "Risk to Health","HA Risk Control","Verification ID","Verification Method"]
    cov = sum(_not_tbd(tm_df.get(c, "")) for c in tm_cols for _ in tm_df.index)
    coverage = pct(cov, len(tm_df) * len(tm_cols))
    # Mapping % between Requirements and HA Risk Control (token overlap > 0.12)
    overlaps = 0
    total_considered = 0
    for _, row in tm_df.iterrows():
        req = str(row.get("Requirements", "") or "")
        rc  = str(row.get("HA Risk Control", "") or "")
        if _not_tbd(req) and _not_tbd(rc):
            total_considered += 1
            if _overlap_ratio(req, rc) >= 0.12:
                overlaps += 1
    mapping = pct(overlaps, total_considered if total_considered else len(tm_df))
    return {
        "Linkage rate": link_score,
        "Column coverage": coverage,
        "Req ↔ Risk Control mapping": mapping,
    }

def _metric_block(col, title, metrics: dict, thresholds: dict):
    col.markdown(f"**{title}**")
    for k, v in metrics.items():
        thr = thresholds.get(k, 0)
        delta_val = round(v - thr, 1)
        label = f"{k}  (Threshold: {thr}%)"
        col.metric(label, f"{v}%", delta=f"{delta_val}%")


def render_metrics(ha_df, dvp_df, tm_df):
    ha = ha_metrics(ha_df)
    dvp = dvp_metrics(dvp_df)
    tm  = tm_metrics(tm_df)
    st.subheader("Evaluation Metrics")

    col1, col2, col3 = st.columns(3)
    _metric_block(col1, "Hazard Analysis", ha, THRESHOLDS["HA"])
    _metric_block(col2, "Design Verification Protocol", dvp, THRESHOLDS["DVP"])
    _metric_block(col3, "Trace Matrix", tm, THRESHOLDS["TM"])

    st.caption(
        "Notes: Scores are auto-computed from generated tables (completeness, measurability, linkage, and mapping). "
        "They do not substitute clinical correctness — please involve Medical Device SMEs for review."
    )


# ---------------- UI ----------------
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")
st.title("🧩 DHF Automation – Infusion Pump")
st.caption("Requirements → Hazard Analysis → DVP → TM")

# Images on the right side
right = st.columns([2, 1])[1]
with right:
    c1, c2 = st.columns(2, gap="medium")
    if os.path.exists(INFUSION_IMG1):
        c1.image(INFUSION_IMG1, use_container_width=True)
    if os.path.exists(INFUSION_IMG):
        c2.image(INFUSION_IMG, use_container_width=True)

st.markdown("**Provide Product Requirements** (choose one):")
colA, colB = st.columns(2)
with colA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with colB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=140, placeholder=sample)

run_btn = st.button("✅ Generate DHF Packages", type="primary")

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

    ha_df = fill_tbd(
        ha_df,
        [
            "risk_id", "risk_to_health", "hazard", "hazardous_situation",
            "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh",
            "risk_index", "risk_control",
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
    dvp_df = fill_tbd(dvp_df, ["verification_id", "verification_method", "acceptance_criteria", "sample_size", "test_procedure"])

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

    # Exact TM column order & names (as in your Excel)
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
        "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh",
        "risk_index", "risk_control",
    ]
    dvp_export_cols = ["verification_id", "verification_method", "acceptance_criteria", "sample_size", "test_procedure"]
    tm_export_cols = TM_ORDER[:]

    ha_x = head_cap(ha_df[ha_export_cols], HA_MAX_ROWS)
    dvp_x = head_cap(dvp_df[dvp_export_cols], DVP_MAX_ROWS)
    tm_x  = head_cap(tm_df[tm_export_cols], TM_MAX_ROWS)

    # Styling per your spec
    ha_bytes = styled_excel_bytes(
        ha_x,
        col_widths={**{c:15 for c in ha_export_cols}, "risk_control":100},
        freeze_cell="D2"
    )
    dvp_bytes = styled_excel_bytes(
        dvp_x,
        col_widths={**{c:15 for c in dvp_export_cols}, "acceptance_criteria":100, "test_procedure":100},
        freeze_cell="D2"
    )
    tm_bytes = styled_excel_bytes(
        tm_x,
        col_widths={c:15 for c in tm_export_cols},
        freeze_cell="D2"
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
