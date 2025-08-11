import streamlit as st
import pandas as pd
import zipfile
import rarfile
import tempfile
import os
from io import BytesIO

def extract_archive(file_buffer, temp_dir, file_type):
    archive_path = os.path.join(temp_dir, f"uploaded.{file_type}")
    
    with open(archive_path, "wb") as f:
        f.write(file_buffer.getbuffer())
    
    extracted_files = []

    if file_type == "zip":
        with zipfile.ZipFile(archive_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            extracted_files = zip_ref.namelist()

    elif file_type == "rar":
        with rarfile.RarFile(archive_path, 'r') as rar_ref:
            rar_ref.extractall(temp_dir)
            extracted_files = rar_ref.namelist()

    return extracted_files

def merge_excel_files(temp_dir):
    merged_df = pd.DataFrame()

    for filename in os.listdir(temp_dir):
        if filename.endswith((".xlsx", ".xls")):
            file_path = os.path.join(temp_dir, filename)
            try:
                df = pd.read_excel(file_path)
                merged_df = pd.concat([merged_df, df], ignore_index=True)
            except Exception as e:
                st.warning(f"❌ Failed to read {filename}: {e}")
    
    return merged_df


# --- Streamlit UI ---
st.set_page_config(page_title="Excel Merger", page_icon="📊")
st.title("📁 Merge Excel Files from Archive")

st.markdown("""
🔄 Upload a `.zip` or `.rar` file containing multiple Excel files with the **same headers**.
This tool will extract and merge them into one Excel file.
""")

uploaded_archive = st.file_uploader("📦 Upload ZIP or RAR Folder of Excel Files", type=["zip", "rar"])

if uploaded_archive is not None:
    file_type = uploaded_archive.name.split(".")[-1].lower()

    with st.spinner(f"Extracting and merging files from .{file_type}..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                extract_archive(uploaded_archive, temp_dir, file_type)
                merged_df = merge_excel_files(temp_dir)
            except Exception as e:
                st.error(f"❌ Failed to process archive: {e}")
                st.stop()

            if not merged_df.empty:
                st.success("✅ Files merged successfully!")
                st.dataframe(merged_df.head())

                # Prepare downloadable Excel
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
