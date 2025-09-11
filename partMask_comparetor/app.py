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
            
            # Step 4: Add status column (match/NotMatch)
            def get_status(part, masked):
                if part == masked:
                    return "match"
                else:
                    return "NotMatch"
            
            df['status'] = df.apply(
                lambda row: get_status(row['PartNumber'], row['MaskedText']),
                axis=1
            )
            
            # Step 5: optional reconstruction for no_diff cases
            suffix_list = df.loc[df['diff_char'] != 'no_diff', 'diff_char'].dropna().unique().tolist()
            suffix_list = sorted(suffix_list, key=len, reverse=True)
            
            for suffix_item in suffix_list:
                if suffix_item:
                    mask = (df['diff_char'] == 'no_diff') & (df['PartNumber'].str.endswith(suffix_item, na=False))
                    if mask.any():
                        df.loc[mask, 'masked_code'] = df.loc[mask, 'PartNumber'].str[:-len(suffix_item)]
                        df.loc[mask, 'diff_char'] = suffix_item
            
            # Reorder columns to match desired output format
            column_order = ['PartNumber', 'CompanyName', 'MaskedText', 'length', 'diff_char', 'masked_code', 'status']
            
            # Ensure all columns exist
            for col in column_order:
                if col not in df.columns:
                    df[col] = ''
            
            # Select and reorder columns
            df_output = df[column_order]
            
            # Show preview
            st.subheader("📋 Differences Found")
            st.dataframe(df_output.head(20))
            
            # Show summary statistics
            st.subheader("📊 Summary")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                match_count = len(df_output[df_output['status'] == 'match'])
                st.metric("✅ Matches", match_count)
            
            with col2:
                notmatch_count = len(df_output[df_output['status'] == 'NotMatch'])
                st.metric("❌ Not Matches", notmatch_count)
            
            with col3:
                length_issues = len(df_output[df_output['length'] == 'lengthIssue'])
                st.metric("⚠️ Length Issues", length_issues)
            
            # Download results
            to_download = io.BytesIO()
            df_output.to_excel(to_download, index=False)
            to_download.seek(0)
            
            st.download_button(
                "📥 Download Results", 
                to_download, 
                file_name="part_diff_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
            st.error("Please ensure your file has the required columns: PartNumber, CompanyName, MaskedText")
    else:
        st.info("Please upload an Excel file to begin.")

with tab2:
    st.subheader("📥 Download Excel Template")
    st.write("Download this template to ensure your data is in the correct format.")
    
    # Create template DataFrame with sample data
    template_data = {
        "PartNumber": [
            "ACS730LLCTR-50AB-S",
            "ACS730KLCTR-50AB-S", 
            "ACS725LLCTR-50AB-S",
            "ACS724LLCTR-50AB-S"
        ],
        "CompanyName": [
            "Allegro MicroSystems, Inc.",
            "Allegro MicroSystems, Inc.",
            "Allegro MicroSystems, Inc.",
            "Allegro MicroSystems, Inc."
        ],
        "MaskedText": [
            "ACS730LLC_-50AB-S",
            "ACS730KLC_-50AB-S",
            "ACS725LLC_-50AB-S",
            "ACS724LLC_-50AB-S"
        ]
    }
    
    template_df = pd.DataFrame(template_data)
    
    # Show template preview
    st.write("**Template Preview:**")
    st.dataframe(template_df)
    
    # Create download file
    template_file = io.BytesIO()
    template_df.to_excel(template_file, index=False)
    template_file.seek(0)
    
    st.download_button(
        "📥 Download Template File", 
        template_file, 
        file_name="part_number_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info("💡 **Tip:** Make sure your Excel file contains columns named exactly: PartNumber, CompanyName, MaskedText")
