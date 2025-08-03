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

        if 'MaskedText' in df.columns:
            df['MaskedText'] = df['MaskedText'].str.rstrip('-')

        # Step 1: Suffix extraction
        for index, row in df.iterrows():
            part_number = row['PartNumber']
            masked_text = row['MaskedText']

            if len(masked_text) > len(part_number):
                df.at[index, 'length'] = 'lengthIssue'
            else:
                df.at[index, 'length'] = 'lengthApprove'

            if part_number.startswith(masked_text):
                diff_value = part_number[len(masked_text):]
            elif masked_text in part_number:
                pos = part_number.find(masked_text)
                diff_value = part_number[pos + len(masked_text):]
            else:
                diff_value = 'no_diff'

            df.at[index, 'suffix_value'] = diff_value if diff_value else 'no_diff'

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
        st.dataframe(df[['PartNumber', 'MaskedText', 'suffix_value', 'masked_code']].head())

        # Download processed file
        to_download = io.BytesIO()
        df.to_excel(to_download, index=False)
        to_download.seek(0)
        st.download_button("📥 Download Processed File", to_download, file_name="masked_output.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to start.")
