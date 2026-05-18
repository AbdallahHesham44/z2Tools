# Streamlit App — Filter Multiple Excel Files From ZIP (Memory Optimized)

```python
import streamlit as st
import pandas as pd
import zipfile
import io
import os
from tempfile import TemporaryDirectory

st.set_page_config(page_title="ZIP Excel Filter Tool", layout="wide")

st.title("📦 ZIP Excel Filter Tool")
st.write("Upload a ZIP file containing Excel files and filter rows with custom criteria.")

# =========================================================
# HELPERS
# =========================================================

def read_excel_columns(excel_file):
    """Read only headers to save memory"""
    try:
        df = pd.read_excel(excel_file, nrows=5)
        return df.columns.tolist()
    except Exception:
        return []


def process_excel(
    excel_bytes,
    selected_columns,
    filter_column,
    filter_value,
    required_equal_columns,
    allowed_datadefinition_values,
):
    """
    Process one excel file with low memory usage
    """

    try:
        # Read only selected columns
        usecols = list(set(selected_columns + [filter_column] + required_equal_columns + ["DataDefinition"]))

        df = pd.read_excel(
            io.BytesIO(excel_bytes),
            usecols=lambda x: x in usecols,
            dtype=str,
        )

        # Fill NaN
        df = df.fillna("")

        # =========================================
        # FILTER MAIN COLUMN
        # =========================================
        if filter_column and filter_value != "":
            df = df[df[filter_column].astype(str) == str(filter_value)]

        # =========================================
        # CHECK EQUAL COLUMNS
        # Example:
        # PartCount == CountRows
        # =========================================
        if len(required_equal_columns) >= 2:
            base_col = required_equal_columns[0]

            for col in required_equal_columns[1:]:
                df = df[
                    df[base_col].astype(str)
                    == df[col].astype(str)
                ]

        # =========================================
        # FILTER DataDefinition VALUES
        # =========================================
        if allowed_datadefinition_values:
            if "DataDefinition" in df.columns:
                df = df[
                    df["DataDefinition"].isin(allowed_datadefinition_values)
                ]

        # =========================================
        # KEEP ONLY SELECTED COLUMNS
        # =========================================
        final_columns = [
            c for c in selected_columns if c in df.columns
        ]

        df = df[final_columns]

        return df

    except Exception as e:
        return pd.DataFrame({"ERROR": [str(e)]})


# =========================================================
# UPLOAD ZIP
# =========================================================

uploaded_zip = st.file_uploader(
    "Upload ZIP File",
    type=["zip"]
)

if uploaded_zip:

    with TemporaryDirectory() as tmpdir:

        zip_path = os.path.join(tmpdir, "uploaded.zip")

        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            excel_files = [
                f for f in zip_ref.namelist()
                if f.endswith('.xlsx') or f.endswith('.xls')
            ]

            if not excel_files:
                st.error("No Excel files found inside ZIP")
                st.stop()

            st.success(f"Found {len(excel_files)} Excel files")

            # =========================================
            # READ FIRST FILE ONLY FOR COLUMNS
            # =========================================
            first_excel = excel_files[0]

            with zip_ref.open(first_excel) as f:
                columns = read_excel_columns(f)

            if not columns:
                st.error("Could not read columns")
                st.stop()

            # =========================================
            # UI
            # =========================================
            st.subheader("⚙️ Filter Settings")

            col1, col2 = st.columns(2)

            with col1:
                filter_column = st.selectbox(
                    "Filter Column",
                    columns
                )

                filter_value = st.text_input(
                    "Filter Value",
                    value="10"
                )

            with col2:
                selected_columns = st.multiselect(
                    "Columns To Keep",
                    columns,
                    default=columns[:10]
                )

            st.subheader("🔄 Equal Columns Check")

            required_equal_columns = st.multiselect(
                "Columns That Must Have Same Value",
                columns,
                default=[
                    c for c in ["PartCount", "CountRows"]
                    if c in columns
                ]
            )

            st.subheader("📘 DataDefinition Filter")

            datadef_text = st.text_area(
                "Allowed DataDefinition Values (one per line)",
                value=""
            )

            allowed_datadefinition_values = [
                x.strip()
                for x in datadef_text.splitlines()
                if x.strip()
            ]

            # =========================================
            # PROCESS
            # =========================================
            if st.button("🚀 Process Files"):

                output_buffer = io.BytesIO()

                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:

                    progress = st.progress(0)
                    status = st.empty()

                    total_files = len(excel_files)

                    for idx, excel_name in enumerate(excel_files):

                        status.write(f"Processing: {excel_name}")

                        with zip_ref.open(excel_name) as f:
                            excel_bytes = f.read()

                        result_df = process_excel(
                            excel_bytes=excel_bytes,
                            selected_columns=selected_columns,
                            filter_column=filter_column,
                            filter_value=filter_value,
                            required_equal_columns=required_equal_columns,
                            allowed_datadefinition_values=allowed_datadefinition_values,
                        )

                        # Limit sheet name length
                        safe_sheet_name = os.path.basename(excel_name)[:31]

                        result_df.to_excel(
                            writer,
                            sheet_name=safe_sheet_name,
                            index=False
                        )

                        progress.progress((idx + 1) / total_files)

                output_buffer.seek(0)

                st.success("Processing Complete ✅")

                st.download_button(
                    label="⬇️ Download Result Excel",
                    data=output_buffer,
                    file_name="Filtered_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )



# =========================================================
# MEMORY OPTIMIZATION NOTES
# =========================================================

st.markdown(r"""
---

## 🚀 Memory Optimization Used

This app is optimized for low RAM environments (1GB):

- Reads only selected columns
- Processes files one-by-one
- Uses temporary directory
- Does not keep all files in memory
- Uses `dtype=str` to reduce mixed type issues
- Reads only first file headers initially

---

## ✅ Example Filters

### Example 1

Filter:
- `Is_Split = 10`
- `PartCount == CountRows`

### Example 2

Allowed DataDefinition:

    Resistor
    Capacitor
    Voltage

---

## ▶️ Run

    pip install streamlit pandas openpyxl
    streamlit run app.py

""")
