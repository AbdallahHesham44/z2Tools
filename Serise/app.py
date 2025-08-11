import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# =========================
# URLs for template files
# =========================
TEMPLATE_MASTER_URL = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/1c93e405525d5480fd43c46e15c3a1b12872d1ee/Serise/TempleteMasterSeriesHistory.xlsx"
TEMPLATE_INPUT_URL = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/1c93e405525d5480fd43c46e15c3a1b12872d1ee/Serise/TempleteInput_series.xlsx"
TEMPLATE_RULES_URL = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/1c93e405525d5480fd43c46e15c3a1b12872d1ee/Serise/TempleteSampleSeriesRules.xlsx"

# =========================
# Utility functions
# =========================
def load_from_github(url):
    """Load Excel file from a GitHub raw URL."""
    resp = requests.get(url)
    resp.raise_for_status()
    return pd.read_excel(BytesIO(resp.content), engine="openpyxl")

def df_to_excel_bytes(df):
    """Convert DataFrame to Excel bytes for download."""
    output = BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    return output.getvalue()

def match_series(comparison_df, master_df, rules_df, top_n):
    """Match requested series with master data applying rules."""
    results = []
    unique_requests = comparison_df["RequestedSeries"].dropna().unique()
    total = len(unique_requests)

    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, requested in enumerate(unique_requests, start=1):
        status_text.text(f"Processing {i}/{total}: {requested}")
        progress_bar.progress(i / total)

        matches = master_df[master_df["SeriesName"].str.contains(str(requested), case=False, na=False)].copy()
        if not matches.empty:
            matches["UsagePercent"] = matches["UsageCount"] / matches["UsageCount"].sum() * 100
            matches["RequestedSeries"] = requested
            matches = matches.sort_values("UsagePercent", ascending=False).head(top_n)
            results.append(matches)
        else:
            results.append(pd.DataFrame([{"RequestedSeries": requested, "SeriesName": None, "UsagePercent": 0}]))

    result_df = pd.concat(results, ignore_index=True)

    if "MinUsagePercent" in rules_df.columns:
        min_threshold = rules_df["MinUsagePercent"].max()
        result_df = result_df[result_df["UsagePercent"] >= min_threshold]

    progress_bar.empty()
    status_text.empty()
    return result_df

# =========================
# Streamlit App UI
# =========================
st.set_page_config(page_title="Series Matcher", layout="wide")
st.title("📊 Series Matching Tool")

# Sidebar downloads
st.sidebar.header("📥 Download Templates")
st.sidebar.download_button(
    label="Download Master Template",
    data=df_to_excel_bytes(load_from_github(TEMPLATE_MASTER_URL)),
    file_name="TempleteMasterSeriesHistory.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.sidebar.download_button(
    label="Download Input Template",
    data=df_to_excel_bytes(load_from_github(TEMPLATE_INPUT_URL)),
    file_name="TempleteInput_series.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.sidebar.download_button(
    label="Download Rules Template",
    data=df_to_excel_bytes(load_from_github(TEMPLATE_RULES_URL)),
    file_name="TempleteSampleSeriesRules.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# File uploads
st.subheader("📤 Upload Files")
comparison_file = st.file_uploader("Upload Comparison File (Input Series)", type=["xlsx"])
master_file = st.file_uploader("Upload Master File", type=["xlsx"])
rules_file = st.file_uploader("Upload Rules File", type=["xlsx"])
top_n = st.number_input("Top N Matches", min_value=1, max_value=20, value=5)

# Process matching
if st.button("🔍 Run Matching"):
    if comparison_file and master_file and rules_file:
        comparison_df = pd.read_excel(comparison_file)
        master_df = pd.read_excel(master_file)
        rules_df = pd.read_excel(rules_file)

        result_df = match_series(comparison_df, master_df, rules_df, top_n)
        st.success("✅ Matching completed!")

        st.dataframe(result_df)

        st.download_button(
            label="📥 Download Results",
            data=df_to_excel_bytes(result_df),
            file_name="MatchedSeriesResults.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload all three files before running the matching process.")
