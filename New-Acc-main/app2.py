# app.py
import streamlit as st
import pandas as pd
import os
import tempfile
from io import BytesIO

# =========================
# PROJECT IMPORTS
# =========================
from mapping_utils import read_mapping_file, save_mapping_to_disk
from github_utils import download_mapping_file_from_github
from preprocessor import preprocess_input_file

from pipeline_wrappers import (
    run_fixed_pipeline_with_prefix_support,
    run_detailed_analysis_with_prefix_support
)

from result_combiner import combine_results

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="ACC Unified Pipeline",
    page_icon="⚙️",
    layout="wide"
)

st.title("⚙️ ACC Unified Pipeline")
st.markdown("Upload your input file and run the complete processing pipeline.")

# =========================
# ENV DEFAULTS
# =========================
os.environ.setdefault("GITHUB_TOKEN", "")
os.environ.setdefault("GITHUB_OWNER", "Hima9791")
os.environ.setdefault("GITHUB_REPO", "map")
os.environ.setdefault("GITHUB_FILE_PATH", "mapping.xlsx")


# =========================
# HELPERS
# =========================
def ensure_mapping(mapping_path, uploaded_mapping=None):
    """
    Ensure mapping.xlsx exists.
    Priority:
    1. Uploaded mapping file
    2. Existing local file
    3. Download from GitHub
    """

    if uploaded_mapping is not None:
        with open(mapping_path, "wb") as f:
            f.write(uploaded_mapping.getbuffer())
        return True, "Uploaded mapping file saved."

    if os.path.exists(mapping_path):
        return True, "Existing mapping file found."

    try:
        df_map = download_mapping_file_from_github()
        save_mapping_to_disk(df_map, mapping_path)
        return True, "Mapping downloaded from GitHub."
    except Exception as e:
        return False, str(e)


def convert_csv_to_excel(uploaded_file, workdir):
    """
    Convert CSV to XLSX if needed.
    """
    filename = uploaded_file.name

    if filename.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)

        excel_path = os.path.join(
            workdir,
            filename.rsplit(".", 1)[0] + "_converted.xlsx"
        )

        df.to_excel(excel_path, index=False, engine="openpyxl")
        return excel_path

    else:
        input_path = os.path.join(workdir, filename)

        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        return input_path


def dataframe_download_button(df, filename, label):
    """
    Create download button for dataframe.
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    st.download_button(
        label=label,
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================
# SIDEBAR
# =========================
st.sidebar.header("⚙️ Settings")

use_github = st.sidebar.checkbox(
    "Use GitHub mapping if missing",
    value=True
)

# =========================
# FILE UPLOADS
# =========================
uploaded_input = st.file_uploader(
    "📂 Upload Input File",
    type=["xlsx", "csv"]
)

uploaded_mapping = st.file_uploader(
    "🗂️ Upload Mapping File (Optional)",
    type=["xlsx"]
)

# =========================
# MAIN PROCESS
# =========================
if uploaded_input is not None:

    if st.button("🚀 Run Pipeline", type="primary"):

        with tempfile.TemporaryDirectory() as workdir:

            try:
                # =========================
                # STEP 1 - PREP INPUT
                # =========================
                st.info("Preparing input file...")

                input_path = convert_csv_to_excel(
                    uploaded_input,
                    workdir
                )

                mapping_path = os.path.join(workdir, "mapping.xlsx")

                ok, msg = ensure_mapping(
                    mapping_path,
                    uploaded_mapping
                )

                if not ok:
                    st.error(f"Mapping error: {msg}")
                    st.stop()

                st.success(msg)

                # =========================
                # STEP 2 - PREPROCESS
                # =========================
                st.info("Running preprocessing...")

                df_pre = preprocess_input_file(input_path)

                pre_xlsx = os.path.join(
                    workdir,
                    "preprocessed.xlsx"
                )

                df_pre.to_excel(
                    pre_xlsx,
                    index=False,
                    engine="openpyxl"
                )

                st.success("Preprocessing completed.")

                # =========================
                # STEP 3 - FIXED PIPELINE
                # =========================
                st.info("Running fixed pipeline...")

                with open(pre_xlsx, "rb") as fh:
                    file_bytes = fh.read()

                df_fixed = run_fixed_pipeline_with_prefix_support(
                    file_bytes,
                    mapping_path
                )

                fixed_xlsx = os.path.join(
                    workdir,
                    "fixed_output.xlsx"
                )

                df_fixed.to_excel(
                    fixed_xlsx,
                    index=False,
                    engine="openpyxl"
                )

                st.success("Fixed pipeline completed.")

                # =========================
                # STEP 4 - DETAILED ANALYSIS
                # =========================
                st.info("Running detailed analysis...")

                detailed_xlsx = os.path.join(
                    workdir,
                    "detailed_analysis.xlsx"
                )

                run_detailed_analysis_with_prefix_support(
                    input_df=df_pre,
                    mapping_file=mapping_path,
                    output_file=detailed_xlsx
                )

                st.success("Detailed analysis completed.")

                # =========================
                # STEP 5 - COMBINE RESULTS
                # =========================
                st.info("Combining results...")

                final_xlsx = os.path.join(
                    workdir,
                    "final_combined.xlsx"
                )

                combine_results(
                    processed_df=df_fixed,
                    analysis_file=detailed_xlsx,
                    output_file=final_xlsx
                )

                st.success("Final report created successfully.")

                # =========================
                # DISPLAY RESULTS
                # =========================
                st.header("✅ Pipeline Completed")

                tab1, tab2, tab3, tab4 = st.tabs([
                    "Preprocessed",
                    "Fixed Output",
                    "Detailed Analysis",
                    "Final Report"
                ])

                # -------------------------
                # PREPROCESSED
                # -------------------------
                with tab1:
                    st.dataframe(df_pre.head(100))

                    dataframe_download_button(
                        df_pre,
                        "preprocessed.xlsx",
                        "⬇️ Download Preprocessed"
                    )

                # -------------------------
                # FIXED OUTPUT
                # -------------------------
                with tab2:
                    st.dataframe(df_fixed.head(100))

                    dataframe_download_button(
                        df_fixed,
                        "fixed_output.xlsx",
                        "⬇️ Download Fixed Output"
                    )

                # -------------------------
                # DETAILED ANALYSIS
                # -------------------------
                with tab3:
                    detailed_df = pd.read_excel(detailed_xlsx)

                    st.dataframe(detailed_df.head(100))

                    dataframe_download_button(
                        detailed_df,
                        "detailed_analysis.xlsx",
                        "⬇️ Download Detailed Analysis"
                                        )
                    # =========================
# STEP 5 - COMBINE RESULTS
# =========================
st.info("Combining results...")

final_xlsx = os.path.join(
    workdir,
    "final_combined.xlsx"
)

combined_ok = combine_results(
    processed_df=df_fixed,
    analysis_file=detailed_xlsx,
    output_file=final_xlsx
)

# Debug
st.write("Detailed file exists:", os.path.exists(detailed_xlsx))
st.write("Fixed DF shape:", df_fixed.shape)

# Check final file
if not os.path.exists(final_xlsx):
    st.error("Final combined file was not created.")

    st.write("Possible reasons:")
    st.write("- combine_results failed")
    st.write("- detailed analysis file missing")
    st.write("- invalid dataframe")
    st.write("- exception inside combine_results")

else:
    st.success("Final report created successfully.")

# =========================
# DISPLAY RESULTS
# =========================
st.header("✅ Pipeline Completed")

tab1, tab2, tab3, tab4 = st.tabs([
    "Preprocessed",
    "Fixed Output",
    "Detailed Analysis",
    "Final Report"
])

# -------------------------
# PREPROCESSED
# -------------------------
with tab1:
    st.dataframe(df_pre.head(100))

    dataframe_download_button(
        df_pre,
        "preprocessed.xlsx",
        "⬇️ Download Preprocessed"
    )

# -------------------------
# FIXED OUTPUT
# -------------------------
with tab2:
    st.dataframe(df_fixed.head(100))

    dataframe_download_button(
        df_fixed,
        "fixed_output.xlsx",
        "⬇️ Download Fixed Output"
    )

# -------------------------
# DETAILED ANALYSIS
# -------------------------
with tab3:

    if os.path.exists(detailed_xlsx):

        detailed_df = pd.read_excel(detailed_xlsx)

        st.dataframe(detailed_df.head(100))

        dataframe_download_button(
            detailed_df,
            "detailed_analysis.xlsx",
            "⬇️ Download Detailed Analysis"
        )

    else:
        st.warning("Detailed analysis file not found.")

# -------------------------
# FINAL REPORT
# -------------------------
with tab4:

    if os.path.exists(final_xlsx):

        final_df = pd.read_excel(final_xlsx)

        st.dataframe(final_df.head(100))

        dataframe_download_button(
            final_df,
            "final_combined.xlsx",
            "⬇️ Download Final Report"
        )

    else:
        st.warning("Final report file not found.")
