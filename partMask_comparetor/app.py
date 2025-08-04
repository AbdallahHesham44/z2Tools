def get_diff_characters(part, masked):
    min_len = min(len(part), len(masked))
    for i in range(min_len):
        if part[i] != masked[i]:
            return part[i]
    if len(part) > len(masked):
        return part[min_len]  # first extra character in PartNumber
    return 'no_diff'


import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Part Number Diff Tool", layout="wide")
st.title("🔍 Part Number First Difference Extractor")

# Upload Excel file
uploaded_file = st.file_uploader("📤 Upload Excel file with PartNumber and MaskedText", type=["xlsx"])

if uploaded_file:
    try:
        # Load Excel file
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ File loaded with {len(df)} rows.")

        # Clean and preprocess
        for col in ['PartNumber', 'MaskedText']:
            if col in df.columns:
                df[col] = df[col].fillna('').astype(str).str.strip()

        if 'MaskedText' in df.columns:
            df['MaskedText'] = df['MaskedText'].str.rstrip('-')

        # Function to get only the first differing character
        def get_diff_characters(part, masked):
            min_len = min(len(part), len(masked))
            for i in range(min_len):
                if part[i] != masked[i]:
                    return part[i]
            if len(part) > len(masked):
                return part[min_len]
            return 'no_diff'

        # Compute differences
        df['length'] = df.apply(
            lambda row: 'lengthIssue' if len(row['MaskedText']) > len(row['PartNumber']) else 'lengthApprove',
            axis=1
        )
        df['diff_char'] = df.apply(
            lambda row: get_diff_characters(row['PartNumber'], row['MaskedText']),
            axis=1
        )

        # Show preview
        st.subheader("📋 Preview of First Differences")
        st.dataframe(df[['PartNumber', 'MaskedText', 'length', 'diff_char']].head())

        # Download processed file
        to_download = io.BytesIO()
        df.to_excel(to_download, index=False)
        to_download.seek(0)
        st.download_button("📥 Download File with Differences", to_download, file_name="first_diff_output.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to start.")
