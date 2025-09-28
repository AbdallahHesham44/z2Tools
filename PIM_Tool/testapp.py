import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
import numpy as np
import requests
from io import BytesIO
import tempfile
import os

# Google Drive URLs for the reference files (replace with your actual Google Drive file IDs)
DRIVE_FILES = {
    "pattern_file": "https://docs.google.com/spreadsheets/d/1W4oA3BtsmWdNhSQOLceCFAfUiKE9BW1_/edit?usp=sharing&ouid=107529105221195873567&rtpof=true&sd=true",
    "preset_file": "https://docs.google.com/spreadsheets/d/1jXQvc7g7juMts5TNyXwGfBIrny_FPR6i/edit?usp=sharing&ouid=107529105221195873567&rtpof=true&sd=true"
}

def get_drive_download_url(file_id):
    """Convert Google Drive file ID to direct download URL"""
    return f"https://drive.google.com/uc?id={file_id}&export=download"

@st.cache_data
def download_drive_file(file_id):
    """Download file from Google Drive and return as BytesIO object"""
    try:
        url = get_drive_download_url(file_id)
        
        # First request to get the confirmation token for large files
        response = requests.get(url, stream=True)
        
        # Check if we need to handle the virus scan warning
        if 'download_warning' in response.cookies:
            # Get the confirmation token
            token = None
            for key, value in response.cookies.items():
                if key.startswith('download_warning'):
                    token = value
                    break
            
            # Make second request with confirmation token
            if token:
                params = {'id': file_id, 'export': 'download', 'confirm': token}
                response = requests.get('https://drive.google.com/uc', params=params, stream=True)
        
        response.raise_for_status()
        
        # Check if response is HTML (means sharing is restricted)
        content_type = response.headers.get('content-type', '')
        if 'text/html' in content_type:
            raise Exception("File is not publicly accessible. Please check sharing permissions.")
        
        return BytesIO(response.content)
        
    except Exception as e:
        st.error(f"Error downloading file from Google Drive: {e}")
        return None

def make_pattern(text):
    def replacer(match):
        integer_part = match.group(1)   # before dot
        decimal_part = match.group(2)   # after dot

        # replace integer part with single $
        integer_replaced = "$"
        # replace each digit in decimal part with $
        decimal_replaced = "$" * len(decimal_part) if decimal_part else ""

        return integer_replaced + ("." + decimal_replaced if decimal_part else "")

    return re.sub(r"(\d+)(?:\.(\d+))?", replacer, text)

def process_excel_data(df):
    """Process the uploaded Excel data"""
    # Convert to string type
    df = df.astype(str)
    
    # Create a key column
    df["key"] = (
        df["Category"].astype(str) + "|" +
        df["Sub-Category"].astype(str) + "|" +
        df["Attribute Name"].astype(str)
    )

    # Create 'pattern' column by replacing numbers with $
    df["Helper_pattern"] = df["Preset values"].astype(str).apply(
        lambda x: re.sub(r"\d+(\.\d+)?", "$", x)
    )

    df["pattern"] = df["Preset values"].astype(str).apply(make_pattern)

    # Count occurrences of the same pattern within each key
    df["count"] = df.groupby(["key", "pattern"])["pattern"].transform("count")

    return df

def sort_preset_values(df):
    """Sort preset values intelligently"""
    def is_alpha_only(s):
        """Check if value is alphabetic only (ignoring spaces)."""
        return bool(re.fullmatch(r"[A-Za-z\s]+", s))

    def sort_groups(val):
        if not isinstance(val, str):
            return val
        
        # Split by "-" to handle groups
        groups = [grp.strip() for grp in val.split(" - ")]

        sorted_groups = []
        for grp in groups:
            if ", " in grp:  
                parts = [p.strip() for p in grp.split(",")]

                # Only sort if all items are alphabetic
                if all(is_alpha_only(p) for p in parts):
                    parts_sorted = sorted(parts, key=lambda x: x.lower())
                    sorted_groups.append(", ".join(parts_sorted))
                else:
                    sorted_groups.append(", ".join(parts))  # keep original order
            else:
                sorted_groups.append(grp.strip())  

        return " - ".join(sorted_groups)

    # Apply sorting logic
    df["sorted Preset values"] = df["Preset values"].apply(sort_groups)

    # Add True/False column
    df["Was Sorted"] = df["Preset values"].astype(str).str.strip() != df["sorted Preset values"].astype(str).str.strip()

    return df

def normalize_pattern(s: str) -> str:
    """Normalize patterns by removing spaces and unifying special chars"""
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"\s+", "", s)   # remove all spaces
    s = s.replace("±", "+-")    # unify ± to a consistent form
    return s.lower()            # case-insensitive

def mark_patterns(df1, df2):
    """Mark patterns as numbers or not"""
    # Normalize both
    df1["pattern_norm"] = df1["Helper_pattern"].apply(normalize_pattern)
    df2["Pattern_norm"] = df2["Pattern"].apply(normalize_pattern)

    # Create a set for fast lookup
    patterns_set = set(df2["Pattern_norm"].unique())

    # Check membership on normalized values
    df1["IsNumber"] = df1["pattern_norm"].apply(lambda x: "Yes" if x in patterns_set else "No")

    # Drop helper column if not needed
    df1.drop(columns=["pattern_norm"], inplace=True)

    return df1

def normalize_case(s: str) -> str:
    """Remove nothing, just lowercase for case sensitivity check"""
    return s.lower()

def normalize_alpha(s: str) -> str:
    """Keep only alphabetic characters"""
    return re.sub(r'[^a-zA-Z]+', '', s).lower()

def normalize_alnum(s: str) -> str:
    """Keep alphanumeric, drop spaces/specials"""
    return re.sub(r'[^a-zA-Z0-9]+', '', s).lower()

def extract_numbers_by_pattern(value: str, pattern: str):
    """Extract numbers from value based on Helper_pattern"""
    if not isinstance(value, str) or not isinstance(pattern, str):
        return []

    regex_pattern = re.escape(pattern).replace("\\$", r"(\d*\.?\d+)")
    match = re.match(regex_pattern, value.strip())
    if not match:
        return []
    return [float(x) for x in match.groups() if x]

def compare_two_numbers(numA, numB):
    """Compare two lists of numbers and return division and percentage strings"""
    div_parts, per_parts = [], []
    for a, b in zip(numA, numB):
        if a == 0 and b == 0:
            div_parts.append("0/0")
            per_parts.append("100")
        elif b == 0:
            div_parts.append(f"{a}/0")
            per_parts.append("0")
        else:
            div_parts.append(f"{a}/{b}")
            per_parts.append(f"{min(a,b)/max(a,b)*100:.1f}")
    return " | ".join(div_parts), " | ".join(per_parts)

def calc_overall_percentage(perc_str):
    """Calculate overall percentage (average)"""
    if not perc_str:
        return ""
    nums = [float(x) for x in perc_str.split(" | ") if x.replace(".","",1).isdigit()]
    if not nums:
        return ""
    return round(sum(nums) / len(nums), 2)

def calc_worst_percentage(perc_str):
    """Calculate worst case percentage (minimum)"""
    if not perc_str:
        return ""
    nums = [float(x) for x in perc_str.split(" | ") if x.replace(".","",1).isdigit()]
    if not nums:
        return ""
    return min(nums)

def determine_comment(is_number, per_value1, per_value2_nospecial, perc_num1_worst, perc_num2_worst, comment_value1, is_caseSensetive='False'):
    """Determine the Comment based on business rules"""
    if comment_value1 == "no match found":
        return "Category not found"

    if is_number:
        # IsNumber = Yes rules
        if per_value1 == 100 and per_value2_nospecial == 100:
            return "Found Exact Number"
        elif (perc_num1_worst and perc_num1_worst > 98) or (perc_num2_worst and perc_num2_worst > 98):
            return "Found with minority Number"
        elif (perc_num1_worst and perc_num1_worst > 80) or (perc_num2_worst and perc_num2_worst > 80):
            return "Found with majority Number"
        else:
            return "Not Found Number"
    else:
        # IsNumber = No rules
        if per_value1 == 100 and per_value2_nospecial == 100 and is_caseSensetive:
            return "Found Exact String with caseSensitive"
        elif per_value1 == 100 and per_value2_nospecial == 100:
            return "Found Exact String"
        elif per_value1 == 0 and per_value2_nospecial == 0:
            return "Not Found String"
        elif per_value1 > 90 and per_value2_nospecial == 100:
            return "Found with minority String"
        elif per_value1 > 80 and per_value2_nospecial > 90:
            return "Found with majority String"
        elif per_value2_nospecial > 70:
            return "Similar String"
        else:
            return "Not Found String"

def compare_files_with_conditional_numbers(df1, df2, threshold=0, top_n=2):
    """Main comparison function"""
    results = []

    for idx, row in df1.iterrows():
        key = row["key"]
        val1 = str(row["sorted Preset values"])

        is_number = row.get("IsNumber", "").strip().lower() == "yes"

        # Candidates with same key
        candidates = df2[df2["key"] == key]

        if candidates.empty:
            row_out = row.to_dict()
            base_update = {
                "Value1": "",
                "per_Value1": 0,
                "Comment_Value1": "no match found",
                "Value2_noSpecial": "",
                "per_Value2_noSpecial": 0,
                "MatchRank": 0,
                "valueHelper_pattern": "",
                "value2Helper_pattern": "",
                "ContainMatch": ""
            }

            # Add number comparison columns if IsNumber is Yes
            if is_number:
                base_update.update({
                    "valueNumber_Value1": "",
                    "percentageNumber_Value1": "",
                    "percentageNumber_Value1_overall": "",
                    "percentageNumber_Value1_worst": 0,
                    "valueNumber_Value2": "",
                    "percentageNumber_Value2": "",
                    "percentageNumber_Value2_overall": "",
                    "percentageNumber_Value2_worst": 0,
                })

            # Determine comment based on business rules
            comment = determine_comment(
                is_number, 0, 0,
                0 if is_number else None,
                0 if is_number else None,
                "no match found"
            )
            base_update["Comment"] = comment

            row_out.update(base_update)
            results.append(row_out)
            continue

        # Collect all similarity scores (per candidate)
        sims = []
        for _, cand in candidates.iterrows():
            main_val2 = str(cand["Preset values"])
            val2 = str(cand["Preset values"]).lower()

            is_caseSensetive = row.get("isCaseSensetive", "").strip().lower() == "yes"
            
            if is_caseSensetive:
                sim = fuzz.ratio(val1.lower(), val2)
            else:
                sim = fuzz.ratio(val1, main_val2)
            
            sim_noSpecial = fuzz.ratio(normalize_alnum(val1), normalize_alnum(val2))

            sims.append({
                "val2": main_val2,
                "sim": sim,
                "sim_noSpecial": sim_noSpecial,
                "helper_pattern": cand.get("Helper_pattern", "")
            })

        # Sort by main similarity (sim), keep top_n
        sims_sorted = sorted(sims, key=lambda x: x["sim"], reverse=True)[:top_n]

        for rank, match in enumerate(sims_sorted, start=1):
            best_val = match["val2"]
            best_score = match["sim"]
            best_score_noSpecial = match["sim_noSpecial"]
            helper_pattern_val = match["helper_pattern"]

            # Comment Logic
            if normalize_case(val1) == normalize_case(best_val) and val1 != best_val:
                comment = "caseSensitive"
            elif normalize_alpha(val1) == normalize_alpha(best_val) and val1 != best_val:
                comment = "nonAlpha"
            elif normalize_alnum(val1) == normalize_alnum(best_val) and val1 != best_val:
                comment = "SpecialCharacter"
            elif normalize_alnum(val1) in normalize_alnum(best_val) and val1 != best_val:
                comment = "Contain"
            else:
                comment = ""

            contain_val = " | ".join([c["val2"] for c in sims if normalize_alnum(val1) in normalize_alnum(c["val2"])])

            row_out = row.to_dict()
            base_update = {
                "Value1": best_val,
                "per_Value1": best_score,
                "Comment_Value1": comment,
                "Value2_noSpecial": best_val,
                "per_Value2_noSpecial": best_score_noSpecial,
                "MatchRank": rank,
                "valueHelper_pattern": helper_pattern_val,
                "value2Helper_pattern": helper_pattern_val,
                "ContainMatch": contain_val
            }

            # Number Comparison Logic (only if IsNumber = "Yes")
            if is_number:
                pat_preset = row.get("Helper_pattern", "")

                value_num1, perc_num1 = "", ""
                value_num2, perc_num2 = "", ""

                if pat_preset == helper_pattern_val:
                    nums_base = extract_numbers_by_pattern(val1, pat_preset)
                    nums1 = extract_numbers_by_pattern(best_val, helper_pattern_val)
                    if nums_base and nums1 and len(nums_base) == len(nums1):
                        value_num1, perc_num1 = compare_two_numbers(nums_base, nums1)

                if pat_preset == helper_pattern_val:
                    nums2 = extract_numbers_by_pattern(best_val, helper_pattern_val)
                    if nums_base and nums2 and len(nums_base) == len(nums2):
                        value_num2, perc_num2 = compare_two_numbers(nums_base, nums2)

                perc1_overall = calc_overall_percentage(perc_num1)
                perc2_overall = calc_overall_percentage(perc_num2)
                perc1_worst = calc_worst_percentage(perc_num1)
                perc2_worst = calc_worst_percentage(perc_num2)

                base_update.update({
                    "valueNumber_Value1": value_num1,
                    "percentageNumber_Value1": perc_num1,
                    "percentageNumber_Value1_overall": perc1_overall,
                    "percentageNumber_Value1_worst": perc1_worst if perc1_worst != "" else 0,
                    "valueNumber_Value2": value_num2,
                    "percentageNumber_Value2": perc_num2,
                    "percentageNumber_Value2_overall": perc2_overall,
                    "percentageNumber_Value2_worst": perc2_worst if perc2_worst != "" else 0,
                })
            else:
                base_update.update({
                    "percentageNumber_Value1_worst": 0,
                    "percentageNumber_Value2_worst": 0,
                })

            # Determine comment based on business rules
            perc1_worst_val = base_update.get("percentageNumber_Value1_worst", 0) if is_number else None
            perc2_worst_val = base_update.get("percentageNumber_Value2_worst", 0) if is_number else None

            comment = determine_comment(
                is_number,
                best_score,
                best_score_noSpecial,
                perc1_worst_val,
                perc2_worst_val,
                comment,
                row.get("isCaseSensetive", "").strip().lower() == "yes"
            )
            base_update["Comment"] = comment

            row_out.update(base_update)
            results.append(row_out)

    df_out = pd.DataFrame(results)
    return df_out

# Streamlit App
def main():
    st.set_page_config(page_title="Excel Pattern Matcher", layout="wide")
    
    st.title("🔍 Excel Pattern Matcher")
    st.markdown("Upload your Excel file and compare it against reference patterns from GitHub.")
    
    # Sidebar for parameters
    st.sidebar.header("Parameters")
    threshold = st.sidebar.slider("Similarity Threshold", 0, 100, 50, help="Minimum similarity score for matches")
    top_n = st.sidebar.selectbox("Top N Matches", [1, 2, 3, 4, 5], index=0, help="Number of top matches to return")
    
    # Google Drive File IDs input
    st.sidebar.header("Google Drive File IDs")
    st.sidebar.markdown("""
    **How to get Google Drive File ID:**
    1. Open the file in Google Drive
    2. Click Share → Anyone with the link
    3. Copy the link: `https://drive.google.com/file/d/FILE_ID/view`
    4. Extract the FILE_ID part
    """)
    
    pattern_file_id = st.sidebar.text_input(
        "Pattern File ID",
        value=DRIVE_FILES["pattern_file"],
        help="Google Drive File ID for the pattern reference file"
    )
    preset_file_id = st.sidebar.text_input(
        "Preset File ID", 
        value=DRIVE_FILES["preset_file"],
        help="Google Drive File ID for the preset reference file"
    )
    
    # File upload
    st.header("📁 Upload Your Excel File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload the Excel file you want to process"
    )
    
    if uploaded_file is not None:
        try:
            # Load uploaded file
            with st.spinner("Loading uploaded file..."):
                df_input = pd.read_excel(uploaded_file, dtype=str)
            
            st.success(f"✅ Uploaded file loaded successfully! Shape: {df_input.shape}")
            
            # Show preview of uploaded file
            with st.expander("📊 Preview of Uploaded File", expanded=True):
                st.dataframe(df_input.head(10))
            
            # Download GitHub files
            st.header("📥 Loading Reference Files from GitHub")
            
            col1, col2 = st.columns(2)
            
            with col1:
                with st.spinner("Downloading pattern file..."):
                    pattern_file_data = download_github_file(pattern_url)
                if pattern_file_data:
                    df_pattern = pd.read_excel(pattern_file_data, dtype=str)
                    st.success(f"✅ Pattern file loaded! Shape: {df_pattern.shape}")
                else:
                    st.error("❌ Failed to load pattern file from GitHub")
                    return
            
            with col2:
                with st.spinner("Downloading preset file..."):
                    preset_file_data = download_github_file(preset_url)
                if preset_file_data:
                    df_preset = pd.read_excel(preset_file_data, dtype=str)
                    st.success(f"✅ Preset file loaded! Shape: {df_preset.shape}")
                else:
                    st.error("❌ Failed to load preset file from GitHub")
                    return
            
            # Process button
            if st.button("🚀 Process Files", type="primary"):
                with st.spinner("Processing files..."):
                    # Step 1: Process uploaded file
                    st.info("Step 1: Processing uploaded file...")
                    df_processed = process_excel_data(df_input)
                    
                    # Step 2: Sort preset values
                    st.info("Step 2: Sorting preset values...")
                    df_sorted = sort_preset_values(df_processed)
                    
                    # Step 3: Mark patterns
                    st.info("Step 3: Marking patterns...")
                    df_marked = mark_patterns(df_sorted, df_pattern)
                    
                    # Step 4: Compare files
                    st.info("Step 4: Comparing files...")
                    df_result = compare_files_with_conditional_numbers(
                        df_marked, df_preset, threshold=threshold, top_n=top_n
                    )
                
                st.success("✅ Processing completed!")
                
                # Display results
                st.header("📊 Results")
                
                # Summary statistics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Rows", len(df_result))
                with col2:
                    exact_matches = len(df_result[df_result['Comment'].str.contains('Found Exact', na=False)])
                    st.metric("Exact Matches", exact_matches)
                with col3:
                    no_matches = len(df_result[df_result['Comment'].str.contains('Not Found', na=False)])
                    st.metric("No Matches", no_matches)
                with col4:
                    similarity_matches = len(df_result[df_result['Comment'].str.contains('Similar', na=False)])
                    st.metric("Similar Matches", similarity_matches)
                
                # Results table
                with st.expander("📋 Full Results", expanded=True):
                    st.dataframe(df_result, use_container_width=True)
                
                # Download buttons
                st.header("💾 Download Results")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Full results download
                    full_csv = df_result.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📄 Download Full Results (CSV)",
                        data=full_csv,
                        file_name='pattern_matching_results.csv',
                        mime='text/csv'
                    )
                
                with col2:
                    # Summary results download
                    summary_cols = ["Category", "Sub-Category", "Attribute Name",
                                  "Preset values", "Max Value", "Unit", "key",
                                  "Value1", "Comment"]
                    existing_summary_cols = [c for c in summary_cols if c in df_result.columns]
                    df_summary = df_result[existing_summary_cols].copy()
                    
                    summary_csv = df_summary.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📋 Download Summary (CSV)",
                        data=summary_csv,
                        file_name='pattern_matching_summary.csv',
                        mime='text/csv'
                    )
                
                # Excel downloads
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                    df_result.to_excel(tmp.name, index=False)
                    with open(tmp.name, 'rb') as f:
                        excel_data = f.read()
                    
                    st.download_button(
                        label="📊 Download Full Results (Excel)",
                        data=excel_data,
                        file_name='pattern_matching_results.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                
        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
    
    else:
        st.info("👆 Please upload an Excel file to get started.")
        
        # Show example of expected file format
        st.header("📋 Expected File Format")
        st.markdown("""
        Your Excel file should contain the following columns:
        - **Category**: Product category
        - **Sub-Category**: Product sub-category  
        - **Attribute Name**: Name of the attribute
        - **Preset values**: Values to be matched
        - **Max Value**: Maximum value (if applicable)
        - **Unit**: Unit of measurement (if applicable)
        """)

if __name__ == "__main__":
    main()
