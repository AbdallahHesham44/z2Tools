def get_consecutive_diff(part, masked):
    diff = ''
    i = 0
    in_diff = False

    while i < len(part) and i < len(masked):
        if part[i] != masked[i]:
            if not in_diff:
                in_diff = True
            diff += part[i]
        else:
            if in_diff:
                break  # stop collecting when characters match again
        i += 1

    # ✅ If the masked is a perfect prefix, and part has extras:
    if not in_diff and len(part) > len(masked):
        diff = part[len(masked):]

    # ✅ If already in diff and part still has different trailing chars
    elif in_diff and i < len(part):
        while i < len(part) and (i >= len(masked) or part[i] != masked[i]):
            diff += part[i]
            i += 1

    return diff if diff else 'no_diff'

import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Part Number Difference Tool", layout="wide")
st.title("🔍 Part Number Difference Extractor (Full Character Support)")

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
        df['MaskedText'] = df['MaskedText'].str.rstrip('-')

        # Function to get first sequence of mismatching characters
        def get_consecutive_diff(part, masked):
            diff = ''
            i = 0
            in_diff = False

            while i < len(part) and i < len(masked):
                if part[i] != masked[i]:
                    if not in_diff:
                        in_diff = True
                    diff += part[i]
                else:
                    if in_diff:
                        break
                i += 1

            if in_diff and i < len(part):
                while i < len(part) and (i >= len(masked) or part[i] != masked[i]):
                    diff += part[i]
                    i += 1

            return diff if diff else 'no_diff'

        # Apply difference logic
        df['length'] = df.apply(
            lambda row: 'lengthIssue' if len(row['MaskedText']) > len(row['PartNumber']) else 'lengthApprove',
            axis=1
        )

        df['diff_char'] = df.apply(
            lambda row: get_consecutive_diff(row['PartNumber'], row['MaskedText']),
            axis=1
        )

        # Show results
        st.subheader("📋 Preview of Differences")
        st.dataframe(df[['PartNumber', 'MaskedText', 'length', 'diff_char']].head(20))

        # Export
        to_download = io.BytesIO()
        df.to_excel(to_download, index=False)
        to_download.seek(0)
        st.download_button("📥 Download Diff Output", to_download, file_name="consecutive_diff_output.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to start.")
