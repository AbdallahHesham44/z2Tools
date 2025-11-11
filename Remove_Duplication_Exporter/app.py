import streamlit as st
import pandas as pd

st.title("Excel Parts Filter Based on Matching Columns")

# Upload files
file1 = st.file_uploader("Upload reference Excel file (File 1)", type=["xlsx"])
file2 = st.file_uploader("Upload file to be filtered (File 2)", type=["xlsx"])

if file1 and file2:
    try:
        # Load data
        df1 = pd.read_excel(file1, dtype=str)
        df2 = pd.read_excel(file2, dtype=str)

        st.write("✅ Files loaded successfully")

        # Select columns
        st.subheader("Step 1: Select matching columns in each file")

        col1 = st.selectbox(
            "Select Part Number Column from File 1",
            df1.columns,
            key="col1"
        )

        col2 = st.selectbox(
            "Select Part Number Column from File 2",
            df2.columns,
            key="col2"
        )

        # Filter data
        st.subheader("Step 2: Filter File 2 based on File 1 column values")

        filtered_df = df2[df2[col2].isin(df1[col1])]

        st.write("### ✅ Filtered Result:")
        st.dataframe(filtered_df)

        # Download button
        st.download_button(
            label="⬇️ Download Filtered Excel File",
            data=filtered_df.to_excel(index=False),
            file_name="filtered_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
