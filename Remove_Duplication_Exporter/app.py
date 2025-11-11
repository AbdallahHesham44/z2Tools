import streamlit as st
import pandas as pd

st.title("Excel Parts Filter App")

st.write("Upload two files: one reference file (master list) and one file to filter.")

# Upload files
file1 = st.file_uploader("Upload reference Excel file (file 1)", type=["xlsx"])
file2 = st.file_uploader("Upload file to be filtered (file 2)", type=["xlsx"])

if file1 and file2:

    try:
        # Read Excel files
        df_ref = pd.read_excel(file1, dtype=str)
        df_data = pd.read_excel(file2, dtype=str)

        st.write("✅ Files uploaded successfully")

        # Detect common part number column
        part_col = None
        possible_cols = ["PartNumber", "part", "PN", "PartNumberC", "PartNumberX"]

        for col in df_ref.columns:
            if col in possible_cols:
                part_col = col
                break

        if not part_col:
            st.error("❌ No part number column found in reference file")
        else:
            st.write(f"✅ Using part number column: **{part_col}**")

            # Filter logic: keep rows where part exists in reference file
            filtered_df = df_data[df_data[part_col].isin(df_ref[part_col])]

            st.write("### ✅ Filtered Result")
            st.dataframe(filtered_df)

            # Download button
            @st.cache_data
            def convert_to_excel(df):
                return df.to_excel(index=False, engine='xlsxwriter')

            if st.button("Download filtered file"):
                excel_file = filtered_df.to_excel("filtered_output.xlsx", index=False)
                st.success("✅ File saved as filtered_output.xlsx")

            # Alternative direct download
            st.download_button(
                label="⬇️ Download Filtered Excel",
                data=filtered_df.to_excel(index=False),
                file_name="filtered_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing files: {e}")
