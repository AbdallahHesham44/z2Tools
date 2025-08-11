import streamlit as st
from utils import load_from_github, match_series
import pandas as pd
from io import BytesIO

# GitHub raw URLs for Excel files
MASTER_FILE_URL = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/main/Serise/MasterSeriesHistory.xlsx"
RULES_FILE_URL = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/main/Serise/SampleSeriesRules.xlsx"

# Template files
TEMPLATE_MASTER_URL = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/main/Serise/TempleteMasterSeriesHistory.xlsx"
TEMPLATE_INPUT_URL = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/main/Serise/TempleteInput_series.xlsx"
TEMPLATE_RULES_URL = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/main/Serise/TempleteSampleSeriesRules.xlsx"

st.set_page_config(page_title="Series Matching Tool", layout="wide")
st.title("📊 Series Matching Tool")

# Sidebar - Download template files
st.sidebar.header("📥 Download Templates")
st.sidebar.download_button(
    label="Download Master Template",
    data=load_from_github(TEMPLATE_MASTER_URL).to_excel(index=False),
    file_name="TempleteMasterSeriesHistory.xlsx"
)
st.sidebar.download_button(
    label="Download Input Template",
    data=load_from_github(TEMPLATE_INPUT_URL).to_excel(index=False),
    file_name="TempleteInput_series.xlsx"
)
st.sidebar.download_button(
    label="Download Rules Template",
    data=load_from_github(TEMPLATE_RULES_URL).to_excel(index=False),
    file_name="TempleteSampleSeriesRules.xlsx"
)

# File uploaders
st.header("📤 Upload Files")
comparison_file = st.file_uploader("Upload Comparison File (Excel)", type=["xlsx"])
rules_file = st.file_uploader("Upload Rules File (Excel)", type=["xlsx"])

if st.button("Run Matching") and comparison_file and rules_file:
    comparison_df = pd.read_excel(comparison_file, engine="openpyxl")
    master_df = load_from_github(MASTER_FILE_URL)
    rules_df = pd.read_excel(rules_file, engine="openpyxl")

    st.subheader("🔄 Processing...")
    result_df = match_series(comparison_df, master_df, rules_df, top_n=5)

    st.success("✅ Matching complete!")
    st.dataframe(result_df)

    # Download results
    output = BytesIO()
    result_df.to_excel(output, index=False)
    st.download_button(
        label="💾 Download Results",
        data=output.getvalue(),
        file_name="SeriesMatchingResults.xlsx"
    )
elif st.button("Run Matching") and (not comparison_file or not rules_file):
    st.error("⚠ Please upload both comparison and rules files before running.")
