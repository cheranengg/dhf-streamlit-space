import os, json, requests, pandas as pd, streamlit as st

BACKEND_URL = os.getenv("BACKEND_URL").rstrip("/")
TOKEN = os.getenv("BACKEND_TOKEN")

st.set_page_config(page_title="DHF Automation", layout="wide")
st.title("DHF Automation – Infusion Pump")

uploaded = st.file_uploader("Upload Product Requirements (CSV/XLSX)", type=["csv","xlsx"])
if uploaded:
    if uploaded.name.endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)

    st.write("Preview:", df.head())

    headers = {"Authorization": f"Bearer {TOKEN}"}
    req_payload = {"requirements": df.to_dict(orient="records")}

    # HA
    with st.spinner("Generating Hazard Analysis..."):
        ha = requests.post(f"{BACKEND_URL}/hazard-analysis", json=req_payload, headers=headers, timeout=300).json().get("ha",[])
    st.subheader("Hazard Analysis")
    st.dataframe(pd.DataFrame(ha))

    # DVP
    with st.spinner("Generating DVP..."):
        dvp = requests.post(f"{BACKEND_URL}/dvp", json={"requirements": req_payload["requirements"], "ha": ha}, headers=headers, timeout=600).json().get("dvp",[])
    st.subheader("Design Verification Protocol")
    st.dataframe(pd.DataFrame(dvp))

    # TM
    with st.spinner("Generating Trace Matrix..."):
        tm = requests.post(f"{BACKEND_URL}/trace-matrix", json={"requirements": req_payload["requirements"], "ha": ha, "dvp": dvp}, headers=headers, timeout=600).json().get("trace_matrix",[])
    st.subheader("Traceability Matrix")
    st.dataframe(pd.DataFrame(tm))

    # Downloads
    st.download_button("Download HA (CSV)", pd.DataFrame(ha).to_csv(index=False), "hazard_analysis.csv", "text/csv")
    st.download_button("Download DVP (CSV)", pd.DataFrame(dvp).to_csv(index=False), "dvp.csv", "text/csv")
    st.download_button("Download TM (CSV)", pd.DataFrame(tm).to_csv(index=False), "trace_matrix.csv", "text/csv")
