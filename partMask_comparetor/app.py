import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Part Number Difference Tool", layout="wide")
st.title("🔍 Exact Character-by-Character Difference Tool")

# Create tabs
tab1, tab2 = st.tabs(["🔍 Difference Tool", "📥 Download Template"])

with tab1:
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
            
            # Step 1: find first mismatch index
            def get_first_diff(part, masked):
                min_len = min(len(part), len(masked))
                for i in range(min_len):
                    if part[i] != masked[i]:
                        return i
                return min_len  # no mismatch → return end
            
            # Step 2: smarter diff_char + masked_code
            def get_diff_and_masked_code(part, masked):
                idx = get_first_diff(part, masked)
                if idx >= len(part):
                    return "no_diff", ""
                match = re.match(r"[A-Za-z]+", part[idx:])
                suffix = match.group(0) if match else part[idx:]
                masked_code = part[:idx]
                return suffix, masked_code
            
            df[['diff_char', 'masked_code']] = df.apply(
                lambda row: pd.Series(get_diff_and_masked_code(row['PartNumber'], row['MaskedText'])),
                axis=1
            )
            
            # Step 3: add length flag
            df['length'] = df.apply(
                lambda row: 'lengthIssue' if len(row['MaskedText']) > len(row['PartNumber']) else 'lengthApprove',
                axis=1
            )
            
            # Step 4: optional reconstruction - only for rows where no difference was initially found
            # Get unique suffixes from rows that had differences
            suffix_list = df.loc[df['diff_char'] != 'no_diff', 'diff_char'].dropna().unique().tolist()
            suffix_list = sorted(suffix_list, key=len, reverse=True)
            
            for suffix_item in suffix_list:
                if suffix_item:
                    # Only update rows where diff_char is 'no_diff' (no initial difference found)
                    # AND the part ends with this suffix AND masked_code is currently empty
                    mask = (df['diff_char'] == 'no_diff') & (df['PartNumber'].str.endswith(suffix_item, na=False)) & (df['masked_code'] == '')
                    if mask.any():
                        df.loc[mask, 'masked_code'] = df.loc[mask, 'PartNumber'].str[:-len(suffix_item)]
                        df.loc[mask, 'diff_char'] = suffix_item
            
            # Step 5: Add status column based on masked_code
            df['status'] = df['masked_code'].apply(lambda x: 'match' if x == '' else 'NotMatch')
            
            # Show preview
            st.subheader("📋 Differences Found")
            st.dataframe(df[['PartNumber', 'MaskedText', 'length', 'diff_char', 'masked_code', 'status']].head(20))
            
            # Download results
            to_download = io.BytesIO()
            df.to_excel(to_download, index=False)
            to_download.seek(0)
            st.download_button("📥 Download Results", to_download, file_name="part_diff_output.xlsx")
            
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
    else:
        st.info("Please upload an Excel file to begin.")

with tab2:
    st.subheader("📥 Download Excel Template")
    
    # Create template DataFrame
    template_df = pd.DataFrame(columns=["PartNumber", "CompanyName", "MaskedText"])
    template_file = io.BytesIO()
    template_df.to_excel(template_file, index=False)
    template_file.seek(0)
    
    st.download_button("📥 Download Template File", template_file, file_name="part_number_template.xlsx")
