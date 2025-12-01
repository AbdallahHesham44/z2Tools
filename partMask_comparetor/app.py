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
            
            # Don't strip trailing hyphens from MaskedText - keep original
            # df['MaskedText'] = df['MaskedText'].str.rstrip('-')  # REMOVED
            
            # Add company filter if CompanyName column exists
            if 'CompanyName' in df.columns:
                df['CompanyName'] = df['CompanyName'].fillna('').astype(str).str.strip()
                companies = ['All'] + sorted(df['CompanyName'].unique().tolist())
                selected_company = st.selectbox("🏢 Filter by Company:", companies)
                
                if selected_company != 'All':
                    df_filtered = df[df['CompanyName'] == selected_company].copy()
                    st.info(f"Showing {len(df_filtered)} rows for {selected_company}")
                else:
                    df_filtered = df.copy()
            else:
                df_filtered = df.copy()
                selected_company = 'All'
            
            # Step 1: find first mismatch index
            def get_first_diff(part, masked):
                min_len = min(len(part), len(masked))
                for i in range(min_len):
                    if part[i] != masked[i]:
                        return i
                return min_len  # no mismatch → return end
            
            # Step 2: smarter diff_char + masked_code + rule extraction
            def get_diff_and_masked_code(part, masked):
                idx = get_first_diff(part, masked)
                if idx >= len(part):
                    return "no_diff", "", "", ""
                match = re.match(r"[A-Za-z]+", part[idx:])
                suffix = match.group(0) if match else part[idx:]
                masked_code = part[:idx]
                rule = suffix  # The rule is the extracted suffix
                applied_rule = masked_code + rule if masked_code else rule  # Apply rule back
                return suffix, masked_code, rule, applied_rule
            
            # Process each company group separately
            if 'CompanyName' in df.columns and selected_company == 'All':
                # Process each company separately
                result_dfs = []
                for company in df['CompanyName'].unique():
                    company_df = df[df['CompanyName'] == company].copy()
                    
                    company_df[['diff_char', 'masked_code', 'rule', 'applied_rule']] = company_df.apply(
                        lambda row: pd.Series(get_diff_and_masked_code(row['PartNumber'], row['MaskedText'])),
                        axis=1
                    )
                    
                    # Step 3: add length flag
                    company_df['length'] = company_df.apply(
                        lambda row: 'lengthIssue' if len(row['MaskedText']) > len(row['PartNumber']) else 'lengthApprove',
                        axis=1
                    )
                    
                    # Step 4: optional reconstruction - only for rows where no difference was initially found
                    suffix_list = company_df.loc[company_df['diff_char'] != 'no_diff', 'diff_char'].dropna().unique().tolist()
                    suffix_list = sorted(suffix_list, key=len, reverse=True)
                    
                    for suffix_item in suffix_list:
                        if suffix_item:
                            mask = (company_df['diff_char'] == 'no_diff') & (company_df['PartNumber'].str.endswith(suffix_item, na=False)) & (company_df['masked_code'] == '')
                            if mask.any():
                                company_df.loc[mask, 'masked_code'] = company_df.loc[mask, 'PartNumber'].str[:-len(suffix_item)]
                                company_df.loc[mask, 'diff_char'] = suffix_item
                                company_df.loc[mask, 'rule'] = suffix_item
                                company_df.loc[mask, 'applied_rule'] = company_df.loc[mask, 'masked_code'] + suffix_item
                    
                    # Step 5: Add status column based on masked_code
                    company_df['status'] = company_df['masked_code'].apply(lambda x: 'match' if x == '' else 'NotMatch')
                    
                    result_dfs.append(company_df)
                
                df_result = pd.concat(result_dfs, ignore_index=True)
            else:
                # Process filtered data
                df_filtered[['diff_char', 'masked_code', 'rule', 'applied_rule']] = df_filtered.apply(
                    lambda row: pd.Series(get_diff_and_masked_code(row['PartNumber'], row['MaskedText'])),
                    axis=1
                )
                
                # Step 3: add length flag
                df_filtered['length'] = df_filtered.apply(
                    lambda row: 'lengthIssue' if len(row['MaskedText']) > len(row['PartNumber']) else 'lengthApprove',
                    axis=1
                )
                
                # Step 4: optional reconstruction
                suffix_list = df_filtered.loc[df_filtered['diff_char'] != 'no_diff', 'diff_char'].dropna().unique().tolist()
                suffix_list = sorted(suffix_list, key=len, reverse=True)
                
                for suffix_item in suffix_list:
                    if suffix_item:
                        mask = (df_filtered['diff_char'] == 'no_diff') & (df_filtered['PartNumber'].str.endswith(suffix_item, na=False)) & (df_filtered['masked_code'] == '')
                        if mask.any():
                            df_filtered.loc[mask, 'masked_code'] = df_filtered.loc[mask, 'PartNumber'].str[:-len(suffix_item)]
                            df_filtered.loc[mask, 'diff_char'] = suffix_item
                            df_filtered.loc[mask, 'rule'] = suffix_item
                            df_filtered.loc[mask, 'applied_rule'] = df_filtered.loc[mask, 'masked_code'] + suffix_item
                
                # Step 5: Add status column
                df_filtered['status'] = df_filtered['masked_code'].apply(lambda x: 'match' if x == '' else 'NotMatch')
                df_result = df_filtered
            
            # Show preview
            st.subheader("📋 Differences Found")
            display_cols = ['PartNumber', 'MaskedText', 'length', 'diff_char', 'rule', 'applied_rule', 'masked_code', 'status']
            if 'CompanyName' in df_result.columns:
                display_cols = ['CompanyName'] + display_cols
            st.dataframe(df_result[display_cols].head(20))
            
            # Show statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", len(df_result))
            with col2:
                st.metric("Match", len(df_result[df_result['status'] == 'match']))
            with col3:
                st.metric("Not Match", len(df_result[df_result['status'] == 'NotMatch']))
            
            # Download results
            to_download = io.BytesIO()
            df_result.to_excel(to_download, index=False)
            to_download.seek(0)
            st.download_button("📥 Download Results", to_download, file_name="part_diff_output.xlsx")
            
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
            import traceback
            st.error(traceback.format_exc())
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
