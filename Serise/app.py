import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import io

# =========================
# CONFIG: GitHub Raw URLs
# =========================
master_raw_url = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/4fdbf29cdca8c7127216f11f550ade65a4e65b2e/Serise/MasterSeriesHistory.xlsx"
rules_raw_url = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/4fdbf29cdca8c7127216f11f550ade65a4e65b2e/Serise/SampleSeriesRules.xlsx"

# =========================
# Helpers
# =========================
@st.cache_data
def load_from_github(raw_url):
    resp = requests.get(raw_url)
    if resp.ok:
        return pd.read_excel(BytesIO(resp.content), engine="openpyxl")
    else:
        st.error(f"Could not fetch file from {raw_url}")
        return None

def df_to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# =========================
# App UI
# =========================
st.title("📊 Series Matching Tool")
st.markdown("""
This tool compares **RequestedSeries** between your comparison file and a master file,
calculates `MajorID` usage percentages, applies business rules, and highlights the most relevant matches.
---
""")

# Sidebar: Settings & Downloads
st.sidebar.header("⚙️ Settings")
top_n = st.sidebar.number_input("Number of top series to show", min_value=1, value=1)

st.sidebar.subheader("📥 Download Templates")
master_df_default = load_from_github(master_raw_url)
rules_df_default = load_from_github(rules_raw_url)

if master_df_default is not None:
    st.sidebar.download_button(
        "Download Master Template",
        data=df_to_excel_bytes(master_df_default),
        file_name="MasterSeriesHistory.xlsx"
    )

if rules_df_default is not None:
    st.sidebar.download_button(
        "Download Rules Template",
        data=df_to_excel_bytes(rules_df_default),
        file_name="SampleSeriesRules.xlsx"
    )

# =========================
# Upload Section
# =========================
st.header("📂 File Uploads")

master_file = st.file_uploader("Upload Master File", type=["xlsx"])
if master_file:
    master_df = pd.read_excel(master_file, engine="openpyxl")
else:
    st.info("No Master file uploaded — loading from GitHub.")
    master_df = master_df_default

rules_file = st.file_uploader("Upload Rules File (optional)", type=["xlsx"])
if rules_file:
    rules_df = pd.read_excel(rules_file, engine="openpyxl")
else:
    st.info("No Rules file uploaded — loading from GitHub.")
    rules_df = rules_df_default

comparison_file = st.file_uploader("Upload Comparison File", type=["xlsx"])
if comparison_file:
    comparison_df = pd.read_excel(comparison_file, engine="openpyxl")
else:
    comparison_df = None

# =========================
# Processing
# =========================
if comparison_df is not None and master_df is not None:
    st.subheader("🔄 Processing Results")
    # TODO: Replace this with your actual matching function
    # For now, just display sample head
    result_df = comparison_df.head(top_n)
    
    st.write("Sample output (replace with your matching logic):")
    st.dataframe(result_df)

    st.download_button(
        "Download Results",
        data=df_to_excel_bytes(result_df),
        file_name="series_output.xlsx"
    )
