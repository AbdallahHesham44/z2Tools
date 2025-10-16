import streamlit as st
import pandas as pd
import zipfile
import rarfile
import tempfile
import os
from io import BytesIO

st.set_page_config(page_title="Excel Deduplicator", layout="wide")

st.title("🧹 Excel Deduplicator Tool")
st.write("Upload one or more Excel or compressed files (ZIP/RAR). The app will remove duplicates from each Excel file and return a ZIP file containing cleaned versions.")

# Function to extract files
def extract_files(uploaded_file, temp_dir):
    extracted_files = []
    file_name = uploaded_file.name.lower()
    file_path = os.path.join(temp_dir, uploaded_file.name)

    with open(file_path, "wb") as f:
        f.write(uploaded_file.read())

    if file_name.endswith(".zip"):
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
            extracted_files.extend([
                os.path.join(temp_dir, f) for f in zip_ref.namelist() if f.endswith((".xlsx", ".xls"))
            ])
    elif file_name.endswith(".rar"):
        with rarfile.RarFile(file_path, "r") as rar_ref:
            rar_ref.extractall(temp_dir)
            extracted_files.extend([
                os.path.join(temp_dir, f) for f in rar_ref.namelist() if f.endswith((".xlsx", ".xls"))
            ])
    elif file_name.endswith((".xlsx", ".xls")):
        extracted_files.append(file_path)

    return extracted_files


# Function to remove duplicates and save new version
def remove_duplicates(excel_path, output_dir):
    try:
        df = pd.read_excel(excel_path)
        before = len(df)
        df = df.drop_duplicates()
        after = len(df)
        cleaned_name = os.path.splitext(os.path.basename(excel_path))[0] + "_cleaned.xlsx"
        output_path = os.path.join(output_dir, cleaned_name)
        df.to_excel(output_path, index=False)
        return cleaned_name, before, after
    except Exception as e:
        return f"Error_{os.path.basename(excel_path)}", 0, 0


uploaded_files = st.file_uploader(
    "📂 Upload Excel/ZIP/RAR files",
    type=["xlsx", "xls", "zip", "rar"],
    accept_multiple_files=True
)

if uploaded_files:
    with tempfile.TemporaryDirectory() as temp_dir, tempfile.TemporaryDirectory() as output_dir:
        all_files = []
        summary = []

        for uploaded in uploaded_files:
            extracted = extract_files(uploaded, temp_dir)
            all_files.extend(extracted)

        st.info(f"✅ Found {len(all_files)} Excel files inside uploads.")

        progress = st.progress(0)
        for i, fpath in enumerate(all_files):
            name, before, after = remove_duplicates(fpath, output_dir)
            summary.append((name, before, after))
            progress.progress((i + 1) / len(all_files))

        # Create output ZIP
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for fname in os.listdir(output_dir):
                fpath = os.path.join(output_dir, fname)
                zipf.write(fpath, arcname=fname)

        zip_buffer.seek(0)

        st.success("🎉 Cleaning completed!")
        st.dataframe(pd.DataFrame(summary, columns=["File", "Rows Before", "Rows After"]))

        st.download_button(
            label="⬇️ Download Cleaned Files (ZIP)",
            data=zip_buffer,
            file_name="cleaned_excel_files.zip",
            mime="application/zip"
        )
