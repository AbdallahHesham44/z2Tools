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

# Template files
template_master_url = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/1c93e405525d5480fd43c46e15c3a1b12872d1ee/Serise/TempleteMasterSeriesHistory.xlsx"
template_input_url = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/1c93e405525d5480fd43c46e15c3a1b12872d1ee/Serise/TempleteInput_series.xlsx"
template_rules_url = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/1c93e405525d5480fd43c46e15c3a1b12872d1ee/Serise/TempleteSampleSeriesRules.xlsx"

# =========================
# Helpers
# =========================
@st.cache_data
def load_from_github(raw_url):
    resp = requests.get(raw_url)
    if resp.ok:
        return pd.read_excel(BytesIO(resp.content))
    else:
        st.error(f"Could not fetch file from {raw_url}")
        return None

def df_to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
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

st.sidebar.subheader("📥 Download Files")

# Download main files
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

# Download additional templates
template_master_df = load_from_github(template_master_url)
template_input_df = load_from_github(template_input_url)
template_rules_df = load_from_github(template_rules_url)

if template_master_df is not None:
    st.sidebar.download_button(
        "Download Template: MasterSeriesHistory",
        data=df_to_excel_bytes(template_master_df),
        file_name="TempleteMasterSeriesHistory.xlsx"
    )

if template_input_df is not None:
    st.sidebar.download_button(
        "Download Template: Input_series",
        data=df_to_excel_bytes(template_input_df),
        file_name="TempleteInput_series.xlsx"
    )

if template_rules_df is not None:
    st.sidebar.download_button(
        "Download Template: SampleSeriesRules",
        data=df_to_excel_bytes(template_rules_df),
        file_name="TempleteSampleSeriesRules.xlsx"
    )

# =========================
# Upload Section
# =========================
st.header("📂 File Uploads")

master_file = st.file_uploader("Upload Master File", type=["xlsx"])
if master_file:
    master_df = pd.read_excel(master_file)
else:
    st.info("No Master file uploaded — loading from GitHub.")
    master_df = master_df_default

rules_file = st.file_uploader("Upload Rules File (optional)", type=["xlsx"])
if rules_file:
    rules_df = pd.read_excel(rules_file)
else:
    st.info("No Rules file uploaded — loading from GitHub.")
    rules_df = rules_df_default

comparison_file = st.file_uploader("Upload Comparison File", type=["xlsx"])
if comparison_file:
    comparison_df = pd.read_excel(comparison_file)
else:
    comparison_df = None

# =========================
# Processing
# =========================
if comparison_df is not None and master_df is not None:
    st.subheader("🔄 Processing Results")
    # TODO: Replace this with your actual matching function
    result_df = comparison_df.head(top_n)
    
    st.write("Sample output (replace with your matching logic):")
    st.dataframe(result_df)

    st.download_button(
        "Download Results",
        data=df_to_excel_bytes(result_df),
        file_name="series_output.xlsx"
    )
