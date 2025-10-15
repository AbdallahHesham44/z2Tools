import streamlit as st
import pandas as pd
import zipfile
import rarfile
import tempfile
import os
from io import BytesIO
import math

EXCEL_ROW_LIMIT = 1_048_576  # Excel sheet row limit

# ----------------------------------------------------------------
# Extract ZIP or RAR archive
# ----------------------------------------------------------------
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

# ----------------------------------------------------------------
# Merge all Excel files by sheet names
# ----------------------------------------------------------------
def merge_excel_files_by_sheets(temp_dir):
    sheet_data = {}  # {sheet_name: DataFrame}

    for filename in os.listdir(temp_dir):
        if filename.endswith((".xlsx", ".xls")):
            file_path = os.path.join(temp_dir, filename)
            try:
                # Read all sheets
                xls = pd.read_excel(file_path, sheet_name=None)
                for sheet_name, df in xls.items():
                    if sheet_name not in sheet_data:
                        sheet_data[sheet_name] = pd.DataFrame()
                    sheet_data[sheet_name] = pd.concat(
                        [sheet_data[sheet_name], df], ignore_index=True
                    )
            except Exception as e:
                st.warning(f"❌ Failed to read {filename}: {e}")
    
    return sheet_data

# ----------------------------------------------------------------
# Save merged sheets to Excel files with row limit split
# ----------------------------------------------------------------
def save_to_excel_with_row_limit(sheet_data):
    output_files = []
    
    for sheet_name, df in sheet_data.items():
        if len(df) > EXCEL_ROW_LIMIT - 1:
            num_parts = math.ceil(len(df) / (EXCEL_ROW_LIMIT - 1))
            for i in range(num_parts):
                start = i * (EXCEL_ROW_LIMIT - 1)
                end = (i + 1) * (EXCEL_ROW_LIMIT - 1)
                part_df = df.iloc[start:end]
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    part_df.to_excel(writer, index=False, sheet_name=f"{sheet_name}_part{i+1}")
                output.seek(0)
                filename = f"{sheet_name}_part{i+1}.xlsx"
                output_files.append((filename, output))
        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name)
            output.seek(0)
            output_files.append((f"{sheet_name}.xlsx", output))
    
    return output_files

# ----------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------
st.set_page_config(page_title="Excel Merger", page_icon="📊")
st.title("📁 Merge Excel Files from Archive (Multi-Sheet Support)")

st.markdown("""
🔄 Upload a `.zip` or `.rar` file containing multiple Excel files with the **same sheet names**.  
This tool will merge them **sheet-by-sheet** into final Excel files.  
If any sheet exceeds **1,048,576 rows**, it will be split into multiple `.xlsx` files.
""")

uploaded_archive = st.file_uploader("📦 Upload ZIP or RAR Folder of Excel Files", type=["zip", "rar"])

if uploaded_archive is not None:
    file_type = uploaded_archive.name.split(".")[-1].lower()

    with st.spinner(f"Extracting and merging files from .{file_type}..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                extract_archive(uploaded_archive, temp_dir, file_type)
                sheet_data = merge_excel_files_by_sheets(temp_dir)
            except Exception as e:
                st.error(f"❌ Failed to process archive: {e}")
                st.stop()

            if sheet_data:
                st.success("✅ Files merged successfully!")
                for sheet_name, df in sheet_data.items():
                    st.write(f"**Sheet:** {sheet_name} — {len(df)} rows")
                    st.dataframe(df.head())

                output_files = save_to_excel_with_row_limit(sheet_data)
                
                for filename, file_buffer in output_files:
                    st.download_button(
                        label=f"⬇️ Download {filename}",
                        data=file_buffer,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("⚠️ No Excel files were found or could be processed.")
