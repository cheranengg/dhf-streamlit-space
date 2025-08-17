# app.py — DHF Streamlit UI (compact top layout + dual Generate buttons + thresholds)
import os, io, re, json, zipfile, typing as t
import pandas as pd
import streamlit as st
import requests

# ---------------- Constants & Secrets ----------------
TBD = "TBD - Human / SME input"

BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8080")
BACKEND_TOKEN = st.secrets.get("BACKEND_TOKEN", "dev-token")

HA_MAX_ROWS = int(os.getenv("HA_MAX_ROWS", "50"))
DVP_MAX_ROWS = int(os.getenv("DVP_MAX_ROWS", "50"))
TM_MAX_ROWS  = int(os.getenv("TM_MAX_ROWS", "50"))
REQ_MAX      = int(os.getenv("REQ_MAX", "50"))

ASSET_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "streamlit_assets"))
INFUSION_IMG  = os.path.join(ASSET_DIR, "Infusion.jpg")
INFUSION_IMG1 = os.path.join(ASSET_DIR, "Infusion1.jpg")

OUTPUT_DIR = st.secrets.get("OUTPUT_DIR", os.path.abspath("./streamlit_outputs"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

THRESHOLDS = {
    "HA": {"Completeness": 80, "Diversity": 50, "Severity coverage": 80, "Specific controls": 60},
    "DVP": {"Measurable procedures": 70, "Acceptance measurability": 70, "Method assigned": 90},
    "TM": {"Linkage rate": 90, "Column coverage": 85, "Req ↔ Risk Control mapping": 60},
}

# ---------------- Page + CSS (compact top) ----------------
st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")
st.markdown("""
<style>
/* tighten top padding, reduce gaps */
.main .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
/* compact text inputs */
.css-1q7k1v2, .stTextArea textarea {min-height: 8rem;}
/* small caption spacing */
.smallcap {margin-top:-0.5rem; color:#6b7280;}
/* reduce image margin */
.compact-img {margin-top: .4rem; margin-bottom: .6rem;}
/* toolbar row */
.toolbar {display:flex; gap:1rem; align-items:center; justify-content:space-between;}
h1 {margin-bottom:.2rem;}
</style>
""", unsafe_allow_html=True)

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
        # show friendly info when backend is down / wrong URL / cold-start
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

# ---- Styled Excel ----
def styled_excel_bytes(df: pd.DataFrame, col_widths: dict, freeze_cell: str = "D2",
                       wrap=True, header_fill="00C6F6D5",
                       all_cols_default_width=15, row_height=30,
                       align="left", valign="center") -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append(list(row.values))
    header_font = Font(bold=True); header_fill = PatternFill("solid", fgColor=header_fill)
    for cell in ws[1]: cell.font = header_font; cell.fill = header_fill
    ws.freeze_panes = freeze_cell
    alignment = Alignment(horizontal=align, vertical=valign, wrap_text=bool(wrap))
    for col_idx, col_name in enumerate(df.columns, start=1):
        width = col_widths.get(col_name, all_cols_default_width)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
        for r in range(1, len(df) + 2):
            ws.cell(row=r, column=col_idx).alignment = alignment
    for r in range(1, len(df) + 2): ws.row_dimensions[r].height = row_height
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()

# ---- Metrics helpers (same as previous answer, abbreviated only where trivial) ----
TBD_CANON = {TBD, "TBD", "NA", "N/A", "", None}
_UNIT_RE = re.compile(r"[±]?\d+(?:\.\d+)?\s?(mA|µA|A|V|kV|ms|s|min|h|mL/h|mL|L|kPa|Pa|%|dB|Ω|°C|g|kg|N|cycles|cm|mm|µL|m)\b", re.I)
STOP = set("a an the and or of to in for with on at from by is are was were be being been as that this these those shall will must should may can could within across per via using than under over system device pump flow alarm accuracy test inspect verify measure calibration detection air occlusion leakage dielectric insulation earth volume label software usability emissions immunity".split())

def pct(n, d): return 0 if d <= 0 else round(100.0 * n / d, 1)
def _not_tbd(x) -> bool: s = str(x if x is not None else "").strip(); return bool(s) and s not in TBD_CANON
def _has_numbers_units(text: str) -> bool: return bool(_UNIT_RE.search(str(text or "")))
def _is_specific(text: str) -> bool:
    t = str(text or "").strip()
    if len(t) < 30: return False
    if re.search(r"\b(IEC|ISO|UL|FDA|62304|60601|80369|11135|61000)\b", t): return True
    return bool(re.search(r"\b(test|verify|calibrat|inspect|measure|trigger|alarm)\b", t, re.I))
def _tokens(s: str) -> set:
    toks = re.findall(r"[A-Za-z]+", str(s or "").lower())
    return {t for t in toks if t not in STOP and len(t) > 2}

def ha_metrics(ha_df: pd.DataFrame) -> dict:
    if ha_df is None or ha_df.empty: return {}
    keys = ["risk_to_health","hazard","hazardous_situation","harm","risk_control","sequence_of_events","severity_of_harm","p0","p1","poh","risk_index"]
    filled = 0; total_slots = len(ha_df) * len(keys)
    for k in keys:
        if k in ha_df.columns: filled += ha_df[k].apply(_not_tbd).sum()
    completeness = pct(filled, total_slots)
    uniq = ha_df[["risk_to_health","hazard","hazardous_situation","harm"]].drop_duplicates().shape[0]
    diversity = pct(uniq, len(ha_df))
    sev_ok = ha_df["severity_of_harm"].astype(str).str.fullmatch(r"[1-5]").sum() if "severity_of_harm" in ha_df.columns else 0
    sev_cov = pct(sev_ok, len(ha_df))
    ctrl_spec = ha_df["risk_control"].apply(_is_specific).sum() if "risk_control" in ha_df.columns else 0
    ctrl_score = pct(ctrl_spec, len(ha_df))
    return {"Completeness": completeness, "Diversity": diversity, "Severity coverage": sev_cov, "Specific controls": ctrl_score}

def dvp_metrics(dvp_df: pd.DataFrame) -> dict:
    if dvp_df is None or dvp_df.empty: return {}
    proc_score = pct(dvp_df["test_procedure"].apply(_has_numbers_units).sum(), len(dvp_df)) if "test_procedure" in dvp_df else 0
    ac_score   = pct(dvp_df["acceptance_criteria"].apply(_has_numbers_units).sum(), len(dvp_df)) if "acceptance_criteria" in dvp_df else 0
    method_ok  = pct(dvp_df["verification_method"].apply(_not_tbd).sum(), len(dvp_df)) if "verification_method" in dvp_df else 0
    return {"Measurable procedures": proc_score, "Acceptance measurability": ac_score, "Method assigned": method_ok}

def _overlap_ratio(a: str, b: str) -> float:
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb: return 0.0
    inter = len(ta & tb); uni = len(ta | tb)
    return inter / uni if uni else 0.0

def tm_metrics(tm_df: pd.DataFrame) -> dict:
    if tm_df is None or tm_df.empty: return {}
    link_score = pct(tm_df.get("Verification ID", pd.Series(dtype=str)).apply(_not_tbd).sum(), len(tm_df))
    tm_cols = ["Requirement ID","Requirements","Requirement (Yes/No)","Risk ID","Risk to Health","HA Risk Control","Verification ID","Verification Method"]
    cov = 0
    for c in tm_cols:
        if c in tm_df.columns: cov += tm_df[c].apply(_not_tbd).sum()
    coverage = pct(cov, len(tm_df)*len(tm_cols))
    overlaps = 0; total_considered = 0
    for _, row in tm_df.iterrows():
        req = str(row.get("Requirements","") or ""); rc = str(row.get("HA Risk Control","") or "")
        if _not_tbd(req) and _not_tbd(rc):
            total_considered += 1
            if _overlap_ratio(req, rc) >= 0.12: overlaps += 1
    mapping = pct(overlaps, total_considered if total_considered else len(tm_df))
    return {"Linkage rate": link_score, "Column coverage": coverage, "Req ↔ Risk Control mapping": mapping}

def _metric_block(col, title, metrics: dict, thresholds: dict):
    col.markdown(f"**{title}**")
    for k, v in metrics.items():
        thr = thresholds.get(k, 0); delta_val = round(v - thr, 1)
        col.metric(f"{k}  (Threshold: {thr}%)", f"{v}%", delta=f"{delta_val}%")

def render_metrics(ha_df, dvp_df, tm_df):
    ha = ha_metrics(ha_df); dvp = dvp_metrics(dvp_df); tm  = tm_metrics(tm_df)
    st.subheader("Evaluation Metrics")
    col1, col2, col3 = st.columns(3)
    _metric_block(col1, "Hazard Analysis", ha, THRESHOLDS["HA"])
    _metric_block(col2, "Design Verification Protocol", dvp, THRESHOLDS["DVP"])
    _metric_block(col3, "Trace Matrix", THRESHOLDS["TM"] and tm, THRESHOLDS["TM"])
    st.caption("Notes: Scores are computed from generated tables (completeness, measurability, linkage, mapping) and support SME review.")

# ---------------- Header / Toolbar ----------------
st.markdown('<div class="toolbar">', unsafe_allow_html=True)
st.title("🧩 DHF Automation – Infusion Pump")
st.caption("Requirements → Hazard Analysis → DVP → TM")
# Top toolbar Generate button (mirrors the main one)
top_run = st.button("✅ Generate DHF Packages", type="primary", key="top")
st.markdown('</div>', unsafe_allow_html=True)

# Images under the title (compact)
if os.path.exists(INFUSION_IMG1) or os.path.exists(INFUSION_IMG):
    c1, c2, _ = st.columns([1,1,3], gap="small")
    if os.path.exists(INFUSION_IMG1):
        c1.image(INFUSION_IMG1, use_container_width=True, output_format="JPEG")
    if os.path.exists(INFUSION_IMG):
        c2.image(INFUSION_IMG, use_container_width=True, output_format="JPEG")

# ---------------- Inputs (kept near top) ----------------
st.markdown("**Provide Product Requirements** (choose one):")
colA, colB = st.columns(2)
with colA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with colB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=120, placeholder=sample)

# Main button (below inputs). Either top or main triggers the same run.
main_run = st.button("✅ Generate DHF Packages", type="primary", key="main")
run_btn = top_run or main_run

# ---------------- Pipeline ----------------
if run_btn:
    with st.spinner("Parsing requirements..."):
        if uploaded is not None:
            try:
                req_df = pd.read_excel(uploaded)
            except Exception as e:
                st.error(f"Failed to read Excel: {e}"); st.stop()
        elif pasted.strip():
            try:
                req_df = pd.read_csv(io.StringIO(pasted))
            except Exception as e:
                st.error(f"Failed to parse pasted CSV: {e}"); st.stop()
        else:
            st.error("Please upload an Excel file or paste CSV text."); st.stop()
        req_df = normalize_requirements(req_df)

    total_reqs = len(req_df)
    st.success(f"Loaded {total_reqs} requirements.")
    st.dataframe(head_cap(req_df, 20), use_container_width=True)

    req_df_limited = req_df.head(REQ_MAX).copy()
    if total_reqs > REQ_MAX:
        st.info(f"Sending only the first {REQ_MAX} requirements to backend (set REQ_MAX env var to change).")

    # ---- HA ----
    with st.spinner("Running Hazard Analysis (backend)..."):
        ha_payload = {"requirements": [
            {"Requirement ID": r["Requirement ID"], "Verification ID": r.get("Verification ID"), "Requirements": r.get("Requirements")}
            for _, r in req_df_limited.iterrows()
        ]}
        try:
            ha_resp = call_backend("/hazard-analysis", ha_payload)
        except RuntimeError as e:
            st.error(f"{e}\n\nTip: verify BACKEND_URL, bearer token, and that /health returns ok.")
            st.stop()
        ha_rows = ha_resp.get("ha", [])
        ha_df = pd.DataFrame(ha_rows)

    ha_df = fill_tbd(ha_df, [
        "risk_id","risk_to_health","hazard","hazardous_situation","harm",
        "sequence_of_events","severity_of_harm","p0","p1","poh","risk_index","risk_control",
    ])
    st.subheader(f"Hazard Analysis (preview, first {HA_MAX_ROWS})")
    st.dataframe(head_cap(ha_df, HA_MAX_ROWS), use_container_width=True)

    # ---- DVP ----
    with st.spinner("Generating Design Verification Protocol (backend)..."):
        dvp_payload = {"requirements": ha_payload["requirements"], "ha": ha_rows}
    try:
        dvp_resp = call_backend("/dvp", dvp_payload)
    except RuntimeError as e:
        st.error(f"{e}\n\nTip: verify the DVP endpoint is available."); st.stop()
    dvp_rows = dvp_resp.get("dvp", [])
    dvp_df = pd.DataFrame(dvp_rows)
    dvp_df = fill_tbd(dvp_df, ["verification_id","verification_method","acceptance_criteria","sample_size","test_procedure"])
    st.subheader(f"Design Verification Protocol (preview, first {DVP_MAX_ROWS})")
    st.dataframe(head_cap(dvp_df, DVP_MAX_ROWS), use_container_width=True)

    # ---- TM ----
    with st.spinner("Building Trace Matrix (backend)..."):
        tm_payload = {"requirements": ha_payload["requirements"], "ha": ha_rows, "dvp": dvp_rows}
        try:
            tm_resp = call_backend("/tm", tm_payload)
        except RuntimeError as e:
            st.error(f"{e}\n\nTip: verify the TM endpoint is available."); st.stop()
        tm_rows = tm_resp.get("tm", [])
        tm_df = pd.DataFrame(tm_rows)

    TM_ORDER = [
        "Requirement ID","Requirements","Requirement (Yes/No)","Risk ID",
        "Risk to Health","HA Risk Control","Verification ID","Verification Method",
    ]
    tm_df = tm_df.reindex(columns=TM_ORDER)
    tm_df = fill_tbd(tm_df, TM_ORDER)

    st.subheader(f"Trace Matrix (preview, first {TM_MAX_ROWS})")
    st.dataframe(head_cap(tm_df, TM_MAX_ROWS), use_container_width=True)

    # ---- Metrics ----
    render_metrics(ha_df, dvp_df, tm_df)

    # ---- Styled Excel ----
    ha_cols = ["risk_id","risk_to_health","hazard","hazardous_situation","harm","sequence_of_events","severity_of_harm","p0","p1","poh","risk_index","risk_control"]
    dvp_cols = ["verification_id","verification_method","acceptance_criteria","sample_size","test_procedure"]
    tm_cols  = TM_ORDER[:]

    ha_x = head_cap(ha_df[ha_cols], HA_MAX_ROWS)
    dvp_x = head_cap(dvp_df[dvp_cols], DVP_MAX_ROWS)
    tm_x  = head_cap(tm_df[tm_cols], TM_MAX_ROWS)

    ha_bytes = styled_excel_bytes(ha_x, {**{c:15 for c in ha_cols}, "risk_control":100}, freeze_cell="D2")
    dvp_bytes = styled_excel_bytes(dvp_x, {**{c:15 for c in dvp_cols}, "acceptance_criteria":100, "test_procedure":100}, freeze_cell="D2")
    tm_bytes  = styled_excel_bytes(tm_x, {c:15 for c in tm_cols}, freeze_cell="D2")

    with open(os.path.join(OUTPUT_DIR, "Hazard_Analysis.xlsx"), "wb") as f: f.write(ha_bytes)
    with open(os.path.join(OUTPUT_DIR, "Design_Verification_Protocol.xlsx"), "wb") as f: f.write(dvp_bytes)
    with open(os.path.join(OUTPUT_DIR, "Trace_Matrix.xlsx"), "wb") as f: f.write(tm_bytes)

    # ---- ZIP download ----
    st.subheader("Download DHF Package")
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Hazard_Analysis.xlsx", ha_bytes)
        zf.writestr("Design_Verification_Protocol.xlsx", dvp_bytes)
        zf.writestr("Trace_Matrix.xlsx", tm_bytes)
    zip_buf.seek(0)
    if st.download_button("⬇️ Download DHF package (3 Excel files, ZIP)",
                          data=zip_buf.getvalue(), file_name="DHF_Package.zip",
                          mime="application/zip", type="primary"):
        st.success("DHF documents downloaded successfully - Hazard Analysis, Design Verification Protocol, Trace Matrix")
        st.info("Note: Please involve Medical Device - SME reviews, before final approval.")
