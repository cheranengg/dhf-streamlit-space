# src/ui_helpers.py
import io
from typing import List
import pandas as pd

TBD = "TBD - Human / SME input"

def normalize_requirements(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize common column headings to: Requirement ID, Verification ID, Requirements
    and ensure they exist (fill missing with None). Keep only these three columns.
    """
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


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """
    Convert a DataFrame to Excel bytes using openpyxl.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.read()


def basic_guardrails_df(df: pd.DataFrame, required_cols: List[str]) -> pd.DataFrame:
    """
    Very simple guardrail pass: mark rows/columns that are missing or TBD.
    """
    issues = []
    for i, row in df.iterrows():
        for c in required_cols:
            val = str(row.get(c, "")).strip()
            if not val or val in {"NA", TBD}:
                issues.append((i, c, "Missing or TBD"))
    return pd.DataFrame(issues, columns=["row_index", "column", "issue"]) if issues else pd.DataFrame(columns=["row_index", "column", "issue"])


def fill_tbd(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """
    Ensure a set of columns exist and fill null/empty values with the TBD token.
    """
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = TBD
        out[c] = out[c].fillna(TBD)
        out.loc[out[c].astype(str).str.strip() == "", c] = TBD
    return out
