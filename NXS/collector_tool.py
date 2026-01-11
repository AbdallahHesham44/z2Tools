import streamlit as st
import pandas as pd
import zipfile
import io

# ---------------------------
# CONFIG
# ---------------------------
IGNORE_SHEETS = ["Status", "Recipe","Intro"]

st.set_page_config(page_title="Zip Excel Merger", layout="wide")
st.title("📊 Merge Excel Files from ZIP (Multi-Sheet)")

uploaded_zip = st.file_uploader(
    "Upload ZIP file containing Excel files",
    type=["zip"]
)

if uploaded_zip:
    sheet_data = {}  # {sheet_name: [DataFrames]}

    with zipfile.ZipFile(uploaded_zip) as z:
        excel_files = [f for f in z.namelist() if f.lower().endswith((".xlsx", ".xls"))]

        if not excel_files:
            st.error("No Excel files found inside the ZIP.")
        else:
            st.success(f"Found {len(excel_files)} Excel files")

            for excel_file in excel_files:
                with z.open(excel_file) as f:
                    try:
                        xls = pd.ExcelFile(f)

                        for sheet in xls.sheet_names:
                            if sheet in IGNORE_SHEETS:
                                continue

                            df = pd.read_excel(xls, sheet_name=sheet)

                            if df.empty:
                                continue

                            sheet_data.setdefault(sheet, []).append(df)

                    except Exception as e:
                        st.warning(f"Failed to read {excel_file}: {e}")

    if sheet_data:
        output_buffer = io.BytesIO()

        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            for sheet_name, dfs in sheet_data.items():
                merged_df = pd.concat(dfs, ignore_index=True)
                merged_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        st.success("Excel file created successfully!")

        st.download_button(
            label="⬇️ Download Merged Excel",
            data=output_buffer.getvalue(),
            file_name="Merged_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("📑 Output Summary")
        for sheet_name, dfs in sheet_data.items():
            st.write(f"- **{sheet_name}** → {sum(len(df) for df in dfs)} rows")

    else:
        st.warning("No valid sheets found after filtering.")
