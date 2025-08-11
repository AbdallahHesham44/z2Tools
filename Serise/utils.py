import pandas as pd
import requests
from io import BytesIO
import streamlit as st

def load_from_github(url):
    """Load an Excel file from a GitHub raw link."""
    resp = requests.get(url)
    resp.raise_for_status()
    return pd.read_excel(BytesIO(resp.content), engine="openpyxl")

def match_series(comparison_df, master_df, rules_df, top_n):
    """Match requested series with master series using rules."""
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
