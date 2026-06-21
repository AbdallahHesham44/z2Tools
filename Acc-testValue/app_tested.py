import streamlit as st
import pandas as pd
import tempfile
import os

from mapping_utils import save_mapping_to_disk
from github_utils import download_mapping_file_from_github
from preprocessor import preprocess_input_file
from pipeline_wrappers import (
    run_fixed_pipeline_with_prefix_support,
    run_detailed_analysis_with_prefix_support
)
from result_combiner import combine_results


st.set_page_config(
    page_title="ACC Unified Pipeline",
    layout="wide"
)

st.title("🚀 ACC Unified Pipeline")

# =====================================================
# Uploads
# =====================================================

input_file = st.file_uploader(
    "Upload Input File",
    type=["xlsx", "xls", "csv"]
)

mapping_file = st.file_uploader(
    "Upload Mapping File (Optional)",
    type=["xlsx"]
)

use_github = st.checkbox(
    "Download mapping.xlsx from GitHub if not uploaded",
    value=True
)

# =====================================================
# Run
# =====================================================

if st.button("▶ Run Pipeline"):

    if input_file is None:
        st.error("Please upload input file")
        st.stop()

    progress = st.progress(0)
    status = st.empty()

    with tempfile.TemporaryDirectory() as wd:

        # ==========================================
        # Save input
        # ==========================================

        input_path = os.path.join(
            wd,
            input_file.name
        )

        with open(input_path, "wb") as f:
            f.write(input_file.getbuffer())

        # ==========================================
        # CSV -> XLSX
        # ==========================================

        status.info("Converting file if needed...")

        if input_path.lower().endswith(".csv"):

            df_csv = pd.read_csv(input_path)

            converted_path = os.path.join(
                wd,
                "converted.xlsx"
            )

            df_csv.to_excel(
                converted_path,
                index=False,
                engine="openpyxl"
            )

            input_path = converted_path

        progress.progress(10)

        # ==========================================
        # Mapping
        # ==========================================

        mapping_path = os.path.join(
            wd,
            "mapping.xlsx"
        )

        status.info("Preparing mapping file...")

        if mapping_file:

            with open(mapping_path, "wb") as f:
                f.write(mapping_file.getbuffer())

        elif use_github:

            df_map = download_mapping_file_from_github()
            save_mapping_to_disk(df_map, mapping_path)

        else:

            st.error(
                "Please upload mapping file or enable GitHub download"
            )
            st.stop()

        progress.progress(20)

        # ==========================================
        # Preprocess
        # ==========================================

        status.info("Preprocessing...")

        df_pre = preprocess_input_file(input_path)

        preprocessed_path = os.path.join(
            wd,
            "preprocessed.xlsx"
        )

        df_pre.to_excel(
            preprocessed_path,
            index=False,
            engine="openpyxl"
        )

        progress.progress(40)

        # ==========================================
        # Fixed Pipeline
        # ==========================================

        status.info("Running Fixed Pipeline...")

        with open(preprocessed_path, "rb") as fh:

            df_fixed = run_fixed_pipeline_with_prefix_support(
                fh.read(),
                mapping_path
            )

        fixed_path = os.path.join(
            wd,
            "fixed_output.xlsx"
        )

        df_fixed.to_excel(
            fixed_path,
            index=False,
            engine="openpyxl"
        )

        progress.progress(60)

        # ==========================================
        # Detailed Analysis
        # ==========================================

        status.info("Running Detailed Analysis...")

        detailed_path = os.path.join(
            wd,
            "detailed_analysis.xlsx"
        )

        run_detailed_analysis_with_prefix_support(
            input_df=df_pre,
            mapping_file=mapping_path,
            output_file=detailed_path
        )

        progress.progress(80)

        # ==========================================
        # Combine
        # ==========================================

        status.info("Combining Results...")

        final_path = os.path.join(
            wd,
            "final_combined.xlsx"
        )

        combine_results(
            processed_df=df_fixed,
            analysis_file=detailed_path,
            output_file=final_path
        )

        progress.progress(100)

        status.success("Pipeline Completed Successfully")

        # ==========================================
        # Downloads
        # ==========================================

        col1, col2, col3, col4 = st.columns(4)

        with open(preprocessed_path, "rb") as f:
            col1.download_button(
                "📥 Preprocessed",
                f,
                file_name="preprocessed.xlsx"
            )

        with open(fixed_path, "rb") as f:
            col2.download_button(
                "📥 Fixed Output",
                f,
                file_name="fixed_output.xlsx"
            )

        with open(detailed_path, "rb") as f:
            col3.download_button(
                "📥 Detailed Analysis",
                f,
                file_name="detailed_analysis.xlsx"
            )

        with open(final_path, "rb") as f:
            col4.download_button(
                "📥 Final Report",
                f,
                file_name="final_combined.xlsx"
            )
