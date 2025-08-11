import streamlit as st
import pandas as pd
import requests
from io import BytesIO

@st.cache_data
def load_from_github(raw_url):
    resp = requests.get(raw_url)
    return pd.read_excel(BytesIO(resp.content), engine="openpyxl") if resp.ok else None

st.title("Series Matching App")

# Sidebar controls
top_n = st.sidebar.number_input("Number of top series to show", min_value=1, value=1)

st.sidebar.download_button("Download Master Template", data=load_from_github(master_raw_url).to_excel(index=False), file_name="MasterSeriesHistory.xlsx")
st.sidebar.download_button("Download Rules Template", data=load_from_github(rules_raw_url).to_excel(index=False), file_name="SampleSeriesRules.xlsx")

master_file = st.file_uploader("Upload Master File", type=["xlsx", "csv"])
rules_file = st.file_uploader("Upload Rules File", type=["xlsx", "csv"], help="Optional")

if not master_file:
    st.info("Master file not uploaded—loading default from GitHub.")
    master_df = load_from_github(master_raw_url)
else:
    master_df = pd.read_excel(master_file, engine="openpyxl")

if rules_file:
    rules_df = pd.read_excel(rules_file, engine="openpyxl")
else:
    st.info("Rules file not uploaded—loading default from GitHub.")
    rules_df = load_from_github(rules_raw_url)

comparison_file = st.file_uploader("Upload Comparison File", type=["xlsx", "csv"])

if comparison_file and master_df is not None:
    # Run your pipeline with the uploaded/master versions
    result_df = compare_requested_series(..., top_n=int(top_n))
    st.write(result_df)
    st.download_button("Download Result", result_df.to_excel(index=False), file_name="series_output.xlsx")

