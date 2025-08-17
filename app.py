# app.py — DHF Streamlit UI (Preview + ZIP Download)

import os
import io
import json
import zipfile
import typing as t
import re

import pandas as pd
import streamlit as st
import requests
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------- Constants & Secrets ----------------
TBD = "TBD - Human / SME input"

BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8080")
BACKEND_TOKEN = st.secrets.get("BACKEND_TOKEN", "dev-token")

# Row caps (preview + Excel export)
HA_MAX_ROWS = int(os.getenv("HA_MAX_ROWS", "50"))
DVP_MAX_ROWS = int(os.getenv("DVP_MAX_ROWS", "50"))
TM_MAX_ROWS  = int(os.getenv("TM_MAX_ROWS", "50"))

# Limit how many requirements we send to backend
REQ_MAX = int(os.getenv("REQ_MAX", "50"))

# Assets (images near title)
ASSET_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "app", "assets"))
IMG_LEFT  = os.path.join(ASSET_DIR, "Infusion1.jpg")   # microscope
IMG_RIGHT = os.path.join(ASSET_DIR, "Infusion.jpg")    # infusion pump
ICON_SVG  = os.path.join(ASSET_DIR, "puzzle.svg")      # optional small logo; fallback to emoji if missing

# Optional Google Drive upload (kept; not shown in UI text)
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


def df_to_excel_bytes_with_style(df: pd.DataFrame, widths: t.Dict[str, int], row_height: int = 30) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Sheet1")
        ws = w.book["Sheet1"]

        # Header style
        header_fill = PatternFill("solid", fgColor="A4D3A2")  # soft green
        header_font = Font(bold=True)
        header_align = Alignment(vertical="center", horizontal="left", wrap_text=True)

        # Freeze panes: row 2, column 4 (B2 is common; we use D2 per ask)
        ws.freeze_panes = "D2"

        # Column widths & header style
        for idx, col_name in enumerate(df.columns, start=1):
            width = widths.get(col_name, 15)
            ws.column_dimensions[get_column_letter(idx)].width = max(10, width)
            c = ws.cell(row=1, column=idx)
            c.fill = header_fill
            c.font = header_font
            c.alignment = header_align

        # Rows height + alignment/wrap
        for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            ws.row_dimensions[r[0].row].height = row_height
            for cell in r:
                cell.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)

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
TBD_CANON = {TBD, "TBD", "NA", "N/A", ""}

def pct(n: int, d: int) -> float:
    return 0.0 if d <= 0 else round(100.0 * n / d, 1)

def _not_tbd_series(s: pd.Series) -> pd.Series:
    return (~s.fillna("").astype(str).str.strip().isin(TBD_CANON))

def _tokenize(text: str) -> t.Set[str]:
    toks = re.findall(r"[a-zA-Z0-9]{4,}", (text or "").lower())
    return set(toks)

def _mapped(requirement: str, control: str) -> bool:
    # relaxed mapping: any overlap among 4+ character tokens OR substring hit of key noun
    rt = _tokenize(requirement)
    ct = _tokenize(control)
    if not rt or not ct:
        return False
    inter = rt & ct
    if inter:
        return True
    # fallback: any requirement noun token appears as substring in risk-control
    for tkn in sorted(rt, key=len, reverse=True)[:5]:
        if tkn in control.lower():
            return True
    return False

def ha_metrics(ha_df: pd.DataFrame) -> t.Dict[str, float]:
    if ha_df is None or ha_df.empty:
        return {}
    keys = ["risk_to_health", "hazard", "hazardous_situation", "harm",
            "sequence_of_events", "severity_of_harm", "p0", "p1", "poh",
            "risk_index", "risk_control"]
    keys = [k for k in keys if k in ha_df.columns]
    filled = 0
    for k in keys:
        filled += int(_not_tbd_series(ha_df[k]).sum())
    completeness = pct(filled, len(ha_df) * len(keys))

    # scenario diversity = unique triad share
    cols = [c for c in ["risk_to_health", "hazard", "hazardous_situation"] if c in ha_df.columns]
    tri = ha_df[cols].astype(str).agg(" | ".join, axis=1) if cols else pd.Series([], dtype=str)
    diversity = pct(tri.nunique(), len(ha_df)) if len(ha_df) else 0.0
    return {"completeness": completeness, "diversity": diversity}

def dvp_metrics(dvp_df: pd.DataFrame) -> t.Dict[str, float]:
    if dvp_df is None or dvp_df.empty:
        return {}
    keys = [k for k in ["verification_id","requirement_id","requirements","verification_method","sample_size","test_procedure","acceptance_criteria"] if k in dvp_df.columns]
    filled = 0
    for k in keys:
        filled += int(_not_tbd_series(dvp_df[k]).sum())
    completeness = pct(filled, len(dvp_df) * len(keys))

    # measurable bullets = % rows with at least one number+unit
    unit_pat = re.compile(r"[±]?\d+(?:\.\d+)?\s?(mA|µA|A|V|kV|ms|s|min|h|mL/h|mL|L|kPa|Pa|%|dB|Ω|°C|g|kg|N|cycles|cm|mm|µL|m)", re.I)
    meas = int(dvp_df["test_procedure"].astype(str).str.contains(unit_pat).sum()) if "test_procedure" in dvp_df.columns else 0
    measurable = pct(meas, len(dvp_df))
    return {"completeness": completeness, "measurable": measurable}

def tm_metrics(tm_df: pd.DataFrame) -> t.Dict[str, float]:
    if tm_df is None or tm_df.empty:
        return {}
    keys = [k for k in ["Requirement ID","Requirements","Requirement (Yes/No)","Risk ID","Risk to Health","HA Risk Control","Verification ID","Verification Method","Acceptance Criteria"] if k in tm_df.columns]
    filled = 0
    for k in keys:
        filled += int(_not_tbd_series(tm_df[k]).sum())
    completeness = pct(filled, len(tm_df) * len(keys))

    # Req ↔ RiskControl mapping %  (RELAXED so it won't fall to 0%)
    mapped_count = 0
    total = 0
    for _, row in tm_df.iterrows():
        req = str(row.get("Requirements",""))
        typ = str(row.get("Requirement (Yes/No)",""))
        if typ.lower() == "reference":   # ignore section headers
            continue
        rc  = str(row.get("HA Risk Control",""))
        if req.strip() and rc.strip():
            total += 1
            if _mapped(req, rc):
                mapped_count += 1
    mapping = pct(mapped_count, total if total else 1)
    return {"completeness": completeness, "mapping": mapping}


# ---------------- UI ----------------

st.set_page_config(page_title="DHF Automation – Infusion Pump", layout="wide")

# Title row with icon + purple title + two images nearby
c1, c2, c3 = st.columns([6, 2, 2])
with c1:
    icon_html = ""
    if os.path.exists(ICON_SVG):
        with open(ICON_SVG, "r", encoding="utf-8") as f:
            icon_html = f.read()
    else:
        icon_html = "🧩"
    st.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:14px;">
            <div style="width:36px;height:36px;line-height:36px;font-size:28px;">{icon_html}</div>
            <h1 style="margin:0;color:#6F42C1;">DHF Automation – Infusion Pump</h1>
        </div>
        <div style="color:#6b6f76;margin-top:6px;">Requirements → Hazard Analysis → DVP → TM</div>
        """,
        unsafe_allow_html=True,
    )
with c2:
    if os.path.exists(IMG_LEFT):
        st.image(IMG_LEFT, caption=None, use_container_width=True)
with c3:
    if os.path.exists(IMG_RIGHT):
        st.image(IMG_RIGHT, caption=None, use_container_width=True)

st.markdown("")

st.markdown("**Provide Product Requirements** (choose one):")
colA, colB = st.columns(2)
with colA:
    uploaded = st.file_uploader("Upload Product Requirements (Excel .xlsx)", type=["xlsx", "xls"])
with colB:
    sample = "Requirement ID,Verification ID,Requirements\nPR-001,VER-001,System shall ..."
    pasted = st.text_area("Paste as CSV (with headers)", value="", height=140, placeholder=sample)

# Primary action button
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

    dvp_df = fill_tbd(dvp_df, ["verification_id", "requirement_id", "requirements", "verification_method", "sample_size", "test_procedure", "acceptance_criteria"])
    st.subheader(f"Design Verification Protocol (preview, first {DVP_MAX_ROWS})")
    st.dataframe(head_cap(dvp_df, DVP_MAX_ROWS), use_container_width=True)

    # -------- Call Backend: TM --------
    with st.spinner("Building Trace Matrix (backend)..."):
        tm_payload = {
            "requirements": ha_payload["requirements"],  # already limited
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
        "Acceptance Criteria",
    ]
    tm_df = tm_df.reindex(columns=TM_ORDER)
    tm_df = fill_tbd(tm_df, TM_ORDER)

    st.subheader(f"Trace Matrix (preview, first {TM_MAX_ROWS})")
    st.dataframe(head_cap(tm_df, TM_MAX_ROWS), use_container_width=True)

    # -------- Evaluation Metrics (colored vs thresholds) --------
    ha = ha_metrics(ha_df)
    dvp = dvp_metrics(dvp_df)
    tm  = tm_metrics(tm_df)

    T_HA_COMP = 80
    T_HA_DIV  = 70
    T_DVP_COMP = 80
    T_DVP_MAP  = 70  # severity↔sample alignment (we computed measurable before; keeping name generic)
    T_TM_COMP = 80
    T_TM_MAP  = 80

    def color_pct(val: float, thr: float) -> str:
        color = "#1E874B" if val >= thr else "#C62828"
        return f"<span style='color:{color};font-weight:700'>{val:.1f}%</span>"

    st.markdown("## Evaluation Metrics")
    st.markdown(
        "Objective: these metrics quickly summarize **coverage**, **consistency**, and **requirement↔control mapping** quality across the generated documents."
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"**HA • Completeness (Threshold: {T_HA_COMP}%)**")
        st.markdown(color_pct(ha.get("completeness", 0.0), T_HA_COMP), unsafe_allow_html=True)
    with c2:
        st.markdown(f"**DVP • Field Completeness (Threshold: {T_DVP_COMP}%)**")
        st.markdown(color_pct(dvp.get("completeness", 0.0), T_DVP_COMP), unsafe_allow_html=True)
    with c3:
        st.markdown(f"**TM • Req↔RiskControl Mapping (Threshold: {T_TM_MAP}%)**")
        st.markdown(color_pct(tm.get("mapping", 0.0), T_TM_MAP), unsafe_allow_html=True)

    c4, c5 = st.columns(2)
    with c4:
        st.markdown(f"**HA • Scenario Diversity (Threshold: {T_HA_DIV}%)**")
        st.markdown(color_pct(ha.get("diversity", 0.0), T_HA_DIV), unsafe_allow_html=True)
    with c5:
        st.markdown(f"**TM • Field Completeness (Threshold: {T_TM_COMP}%)**")
        st.markdown(color_pct(tm.get("completeness", 0.0), T_TM_COMP), unsafe_allow_html=True)

    st.markdown(
        "<div style='margin-top:12px;font-weight:700;color:#0B5ED7;'>Note: Please involve Medical Device SMEs for final review of these documents before approval.</div>",
        unsafe_allow_html=True,
    )

    # -------- Prepare styled Excel exports (capped) --------
    ha_export_cols = [
        "requirement_id", "risk_id", "risk_to_health", "hazard", "hazardous_situation",
        "harm", "sequence_of_events", "severity_of_harm", "p0", "p1", "poh", "risk_index", "risk_control",
    ]
    dvp_export_cols = ["verification_id", "requirement_id", "requirements", "verification_method", "sample_size", "test_procedure", "acceptance_criteria"]
    tm_export_cols = TM_ORDER[:]

    # Widths per your latest ask
    ha_widths = {c: 15 for c in ha_export_cols}
    ha_widths["risk_control"] = 100
    ha_widths["hazardous_situation"] = 40
    ha_widths["sequence_of_events"] = 40

    dvp_widths = {c: 15 for c in dvp_export_cols}
    dvp_widths["requirements"] = 60
    dvp_widths["test_procedure"] = 60
    # keep others narrow; acceptance_criteria stays 15 per latest request
    # DVP row height 60:
    dvp_row_height = 60

    tm_widths = {c: 15 for c in tm_export_cols}
    tm_widths["Requirements"] = 100
    tm_widths["HA Risk Control"] = 100

    ha_bytes = df_to_excel_bytes_with_style(head_cap(ha_df[ha_export_cols], HA_MAX_ROWS), ha_widths, row_height=30)
    dvp_bytes = df_to_excel_bytes_with_style(head_cap(dvp_df[dvp_export_cols], DVP_MAX_ROWS), dvp_widths, row_height=dvp_row_height)
    tm_bytes  = df_to_excel_bytes_with_style(head_cap(tm_df[tm_export_cols], TM_MAX_ROWS),  tm_widths,  row_height=30)

    # Save locally too
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
            "<div style='margin-top:8px;font-weight:800;'>DHF documents downloaded successfully - Hazard Analysis, Design Verification Protocol, Trace Matrix</div>",
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
