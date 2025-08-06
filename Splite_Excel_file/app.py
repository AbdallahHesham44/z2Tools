import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
from io import BytesIO

def split_excel_file(df, rows_per_file, base_name, progress_callback):
    output_files = []
    total_rows = len(df)
    num_files = (total_rows + rows_per_file - 1) // rows_per_file

    for i in range(num_files):
        start = i * rows_per_file
        end = min(start + rows_per_file, total_rows)
        chunk = df.iloc[start:end]

        temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        chunk.to_excel(temp_output.name, index=False)
        output_files.append((f"{base_name}_part_{i+1}.xlsx", temp_output.name))

        progress_callback((i + 1) / num_files)

    return output_files

def compress_files_to_zip(file_list):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for arcname, filepath in file_list:
            zipf.write(filepath, arcname)
    zip_buffer.seek(0)
    return zip_buffer

# --- Streamlit UI ---
st.title("📊 Excel Splitter & Downloader")

uploaded_file = st.file_uploader("Upload a large Excel file (.xlsx)", type=["xlsx"])

rows_per_file = st.number_input("Number of rows per output file", min_value=1, value=5000, step=100)

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    base_name = os.path.splitext(uploaded_file.name)[0]

    if st.button("Split & Download"):
        st.info(f"Splitting `{uploaded_file.name}` into parts of {rows_per_file} rows...")

        progress_bar = st.progress(0)

        # Split Excel and get list of temporary file paths
        output_files = split_excel_file(df, rows_per_file, base_name, progress_callback=progress_bar.progress)

        # Compress all parts into a single zip file
        zip_buffer = compress_files_to_zip(output_files)

        # Clean up temporary files
        for _, path in output_files:
            os.remove(path)

        st.success("✅ Done! Download your split files below.")
        st.download_button(
            label="📥 Download ZIP",
            data=zip_buffer,
            file_name=f"{base_name}_split_files.zip",
            mime="application/zip"
        )
