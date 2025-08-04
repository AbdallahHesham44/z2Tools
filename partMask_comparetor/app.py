import streamlit as st
import pandas as pd
import io
import re

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

        if 'MaskedText' in df.columns:
            df['MaskedText'] = df['MaskedText'].str.rstrip('-')

        # Function to extract suffix using tokenized comparison
        def get_suffix_by_token_comparison(part, masked):
            part_tokens = re.split(r'[\W_]+', part)
            masked_tokens = re.split(r'[\W_]+', masked)

            common_tokens = []
            for p_tok, m_tok in zip(part_tokens, masked_tokens):
                if p_tok == m_tok:
                    common_tokens.append(p_tok)
                else:
                    break

            # Build common prefix string
            if common_tokens:
                # Determine original delimiter (default to '-')
                delimiter = '-' if '-' in part else ''
                common_prefix = delimiter.join(common_tokens)
                if part.startswith(common_prefix):
                    suffix = part[len(common_prefix):]
                    return suffix.lstrip('-')  # clean extra dash if any
            if part.startswith(masked):
                return part[len(masked):]
            elif masked in part:
                pos = part.find(masked)
                return part[pos + len(masked):]
            return 'no_diff'

        # Step 1: Enhanced suffix extraction
        suffix_values = []
        lengths = []

        for index, row in df.iterrows():
            part_number = row['PartNumber']
            masked_text = row['MaskedText']

            if len(masked_text) > len(part_number):
                lengths.append('lengthIssue')
            else:
                lengths.append('lengthApprove')

            diff_value = get_suffix_by_token_comparison(part_number, masked_text)
            suffix_values.append(diff_value if diff_value else 'no_diff')

        df['length'] = lengths
        df['suffix_value'] = suffix_values

        # Step 2: Generate masked_code
        df['masked_code'] = ''
        suffix_list = df.loc[df['suffix_value'] != 'no_diff', 'suffix_value'].unique().tolist()
        suffix_list.sort(key=len, reverse=True)

        for suffix_item in suffix_list:
            if pd.notna(suffix_item) and suffix_item != '':
                mask = (df['suffix_value'] == 'no_diff') & (df['PartNumber'].str.endswith(suffix_item, na=False))
                if mask.any():
                    df.loc[mask, 'masked_code'] = df.loc[mask, 'PartNumber'].str[:-len(suffix_item)]
                    df.loc[mask, 'suffix_value'] = suffix_item

        # Show preview
        st.subheader("📋 Preview of Processed Data")
        st.dataframe(df[['PartNumber', 'MaskedText', 'length', 'suffix_value', 'masked_code']].head())

        # Download processed file
        to_download = io.BytesIO()
        df.to_excel(to_download, index=False)
        to_download.seek(0)
        st.download_button("📥 Download Processed File", to_download, file_name="masked_output.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to start.")
