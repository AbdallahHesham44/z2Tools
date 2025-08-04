import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Part Number Difference Tool", layout="wide")
st.title("🔍 Exact Character-by-Character Difference Tool")

uploaded_file = st.file_uploader("📤 Upload Excel file with PartNumber and MaskedText", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ File loaded with {len(df)} rows.")

        # Clean data
        for col in ['PartNumber', 'MaskedText']:
            if col in df.columns:
                df[col] = df[col].fillna('').astype(str).str.strip()
        df['MaskedText'] = df['MaskedText'].str.rstrip('-')

        # Step 1: Character-level diff function
        def get_diff_chars(part, masked):
            diff = ''
            max_len = max(len(part), len(masked))
            for i in range(max_len):
                p_char = part[i] if i < len(part) else ''
                m_char = masked[i] if i < len(masked) else ''
                if p_char != m_char:
                    diff += p_char
            return diff if diff else 'no_diff'

        # Step 2: Generate diff_char and length flag
        df['length'] = df.apply(
            lambda row: 'lengthIssue' if len(row['MaskedText']) > len(row['PartNumber']) else 'lengthApprove',
            axis=1
        )
        df['diff_char'] = df.apply(
            lambda row: get_diff_chars(row['PartNumber'], row['MaskedText']),
            axis=1
        )

        # Step 3: Optional masked_code reconstruction from known suffix patterns
        df['masked_code'] = ''
        suffix_list = df.loc[df['diff_char'] != 'no_diff', 'diff_char'].dropna().unique().tolist()
        suffix_list = sorted(suffix_list, key=len, reverse=True)  # longest first

        for suffix_item in suffix_list:
            if suffix_item:
                mask = (df['diff_char'] == 'no_diff') & (df['PartNumber'].str.endswith(suffix_item, na=False))
                if mask.any():
                    df.loc[mask, 'masked_code'] = df.loc[mask, 'PartNumber'].str[:-len(suffix_item)]
                    df.loc[mask, 'diff_char'] = suffix_item  # update diff_char with matched suffix

        # Show preview
        st.subheader("📋 Differences Found")
        st.dataframe(df[['PartNumber', 'MaskedText', 'length', 'diff_char', 'masked_code']].head(20))

        # Download results
        to_download = io.BytesIO()
        df.to_excel(to_download, index=False)
        to_download.seek(0)
        st.download_button("📥 Download Results", to_download, file_name="part_diff_output.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to begin.")
