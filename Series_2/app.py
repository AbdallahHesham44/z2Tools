import streamlit as st
import pandas as pd
from io import BytesIO
from series_processing import compare_requested_series
from github_utils import load_excel_from_github, save_excel_to_github, modify_master_or_rules

REPO_NAME = "AbdallahHesham44/z2Tools"
MASTER_PATH = "Serise/MasterSeriesHistory.xlsx"
RULES_PATH = "Serise/SampleSeriesRules.xlsx"

st.title("📊 Series Processing Tool")

process_choice = st.radio("Choose Process", ["Upload & Process", "Update/Delete GitHub Files"])

if process_choice == "Upload & Process":
    uploaded_file = st.file_uploader("Upload Input_series.xlsx", type=["xlsx"])
    if uploaded_file:
        # Load master & rules from GitHub
        token = st.text_input("GitHub Token (for private repos)", type="password")
        if token:
            master_df = load_excel_from_github(REPO_NAME, MASTER_PATH, token)
            rules_df = load_excel_from_github(REPO_NAME, RULES_PATH, token)

            output_df = compare_requested_series(uploaded_file, BytesIO(master_df.to_excel(index=False)), BytesIO(rules_df.to_excel(index=False)))

            # Download processed file
            towrite = BytesIO()
            output_df.to_excel(towrite, index=False)
            towrite.seek(0)
            st.download_button("📥 Download Processed Excel", data=towrite, file_name="Processed_Output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif process_choice == "Update/Delete GitHub Files":
    action = st.radio("Action", ["append", "delete"])
    uploaded_changes = st.file_uploader("Upload file with changes (xlsx)", type=["xlsx"])
    token = st.text_input("GitHub Token", type="password")

    if uploaded_changes and token:
        changes_df = pd.read_excel(uploaded_changes)

        # Load files from GitHub
        master_df = load_excel_from_github(REPO_NAME, MASTER_PATH, token)
        rules_df = load_excel_from_github(REPO_NAME, RULES_PATH, token)

        # Modify files
        master_df = modify_master_or_rules(master_df, changes_df, action)
        rules_df = modify_master_or_rules(rules_df, changes_df, action)

        # Save back to GitHub
        save_excel_to_github(REPO_NAME, MASTER_PATH, master_df, f"{action} MasterSeriesHistory", token)
        save_excel_to_github(REPO_NAME, RULES_PATH, rules_df, f"{action} SeriesRules", token)

        st.success(f"✅ Successfully {action}ed data in GitHub files")
