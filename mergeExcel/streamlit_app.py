import streamlit as st
import pandas as pd
import zipfile
import tempfile
import os
from io import BytesIO

def merge_excel_files_from_zip(zip_file):
    merged_df = pd.DataFrame()

    with tempfile.TemporaryDirectory() as temp_dir:
        zip_path = os.path.join(temp_dir, "uploaded.zip")
        with open(zip_path, "wb") as f:
            f.write(zip_file.getbuffer())

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        for filename in os.listdir(temp_dir):
            if filename.endswith((".xlsx", ".xls")):
                file_path = os.path.join(temp_dir, filename)
                try:
                    df = pd.read_excel(file_path)
                    merged_df = pd.concat([merged_df, df], ignore_index=True)
                except Exception as e:
                    st.warning(f"❌ Failed to read {filename}: {e}")
    
    return merged_df


st.set_page_config(page_title="Excel Merger", page_icon="📊")
st.title("📁 Merge Excel Files from Folder")

st.markdown("🔄 Upload a `.zip` file containing multiple Excel files with the **same header**, and this tool will merge them into one Excel file.")

uploaded_zip = st.file_uploader("📦 Upload ZIP Folder of Excel Files", type=["zip"])

if uploaded_zip is not None:
    with st.spinner("Processing and merging Excel files..."):
        merged_df = merge_excel_files_from_zip(uploaded_zip)

        if not merged_df.empty:
            st.success("✅ Files merged successfully!")

            # Display preview
            st.dataframe(merged_df.head())

            # Create downloadable Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='MergedData')
            output.seek(0)

            st.download_button(
                label="⬇️ Download Merged Excel",
                data=output,
                file_name="merged_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("⚠️ No Excel files were found or could be processed.")
