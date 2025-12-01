import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Part Number Masking Tool", layout="wide")
st.title("🔧 Part Number Masking Tool")

# Upload Excel file
uploaded_file = st.file_uploader("📤 Upload Excel file with PartNumber and MaskedText", type=["xlsx"])

if uploaded_file:
    try:
        # Load Excel file into DataFrame
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ File loaded with {len(df)} rows.")

        # Clean and preprocess
        for col in ['PartNumber', 'MaskedText']:
            if col in df.columns:
                df[col] = df[col].fillna('').astype(str).str.strip()

        # --- Function to get diff_char ---
        def extract_diff_char(part, masked):
            # if lengths don't match, handle carefully
            diff = []
            for p, m in zip(part, masked):
                if p != m and m != '_':  # treat "_" in masked as wildcard
                    diff.append(m)
            # if masked is longer
            if len(masked) > len(part):
                diff.extend(list(masked[len(part):]))
            return "".join(diff) if diff else 'no_diff'

        # Apply diff_char logic
        df['diff_char'] = df.apply(lambda row: extract_diff_char(row['PartNumber'], row['MaskedText']), axis=1)

        # Length status
        df['length'] = df.apply(lambda row: 'lengthIssue' if len(row['MaskedText']) > len(row['PartNumber']) else 'lengthApprove', axis=1)

        # Status
        df['status'] = df.apply(lambda row: 'Match' if row['diff_char'] == 'no_diff' else 'NotMatch', axis=1)

        # Show preview
        st.subheader("📋 Preview of Processed Data")
        st.dataframe(df[['PartNumber', 'MaskedText', 'diff_char', 'length', 'status']].head())

        # Download processed file
        to_download = io.BytesIO()
        df.to_excel(to_download, index=False)
        to_download.seek(0)
        st.download_button("📥 Download Processed File", to_download, file_name="masked_output.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to start.")
