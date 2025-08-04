import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Part Number Diff Tool", layout="wide")
st.title("🔍 Part Number Difference Extractor")

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

        # Function to get only differing characters
        def get_diff_characters(part, masked):
            diff = ''
            min_len = min(len(part), len(masked))
            for i in range(min_len):
                if part[i] != masked[i]:
                    diff += part[i]
            if len(part) > len(masked):
                diff += part[min_len:]
            return diff if diff else 'no_diff'

        # Compute differences
        diff_chars = []
        lengths = []

        for index, row in df.iterrows():
            part_number = row['PartNumber']
            masked_text = row['MaskedText']

            lengths.append('lengthIssue' if len(masked_text) > len(part_number) else 'lengthApprove')
            diff_chars.append(get_diff_characters(part_number, masked_text))

        df['length'] = lengths
        df['diff_char'] = diff_chars

        # Show preview
        st.subheader("📋 Preview of Differences")
        st.dataframe(df[['PartNumber', 'MaskedText', 'length', 'diff_char']].head())

        # Download processed file
        to_download = io.BytesIO()
        df.to_excel(to_download, index=False)
        to_download.seek(0)
        st.download_button("📥 Download File with Differences", to_download, file_name="diff_output.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to start.")
