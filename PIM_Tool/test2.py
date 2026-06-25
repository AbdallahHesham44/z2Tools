import streamlit as st
import pandas as pd
import re
import tempfile
import os
from rapidfuzz import fuzz
import numpy as np
import requests
from io import BytesIO
from typing import List, Dict, Tuple, Optional
import matplotlib.pyplot as plt

# Constants
SIMILARITY_THRESHOLDS = {
    'EXACT': 100,
    'MINORITY_NUMBER': 97,
    'MAJORITY_NUMBER': 80,
    'MINORITY_STRING': 90,
    'MAJORITY_STRING_PRIMARY': 80,
    'MAJORITY_STRING_SECONDARY': 90,
    'SIMILAR_STRING': 70
}

REQUIRED_COLUMNS = ['Category', 'Sub-Category', 'Attribute Name', 'Preset values']

# Google Drive URLs for the reference files
DRIVE_FILES = {
    "pattern_file": "https://docs.google.com/spreadsheets/d/1W4oA3BtsmWdNhSQOLceCFAfUiKE9BW1_/export?format=xlsx",
    # "preset_file": "https://docs.google.com/spreadsheets/d/1IhOFscHAvOH8erop8xg8HVWn6MMvMXvo/export?format=xlsx",
    # "preset_file": "https://docs.google.com/spreadsheets/d/1hrTL_ciligFN38mlFxt439UB3l9m4Dhh/export?format=xlsx",
    # "preset_file": "https://docs.google.com/spreadsheets/d/1vf-ab9PTerh1D2qw8LC-g94jjqLdLhIK/export?format=xlsx",
    # "preset_file": "https://docs.google.com/spreadsheets/d/1vf-ab9PTerh1D2qw8LC-g94jjqLdLhIK/export?format=xlsx",
    # https://docs.google.com/spreadsheets/d/1FzQ8zYDQijq6XRAjg0BMtIvOXlwSHe6K/edit?usp=sharing&ouid=117756686149107163584&rtpof=true&sd=true
    "preset_file": "https://docs.google.com/spreadsheets/d/1FzQ8zYDQijq6XRAjg0BMtIvOXlwSHe6K/export?format=xlsx",
    
    "isnumber_file": "https://docs.google.com/spreadsheets/d/1miXOKaln_uj5x52-vQmvqVtvKUKJuOnU/export?format=xlsx"
}

def download_file_from_drive(url: str) -> Optional[BytesIO]:
    """Download file from Google Drive URL"""
    try:
        # Convert edit URL to export URL if needed
        if "/edite" in url:
            match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
            if match:
                doc_id = match.group(1)
                url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
        
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        return BytesIO(response.content)
    except Exception as e:
        st.error(f"Error downloading file: {str(e)}")
        return None

@st.cache_data
def load_reference_files() -> Tuple[Optional[BytesIO], Optional[BytesIO], Optional[BytesIO]]:
    """Load reference files from Google Drive (pattern, preset, isnumber)"""
    pattern_file = download_file_from_drive(DRIVE_FILES["pattern_file"])
    preset_file = download_file_from_drive(DRIVE_FILES["preset_file"])
    isnumber_file = download_file_from_drive(DRIVE_FILES["isnumber_file"])
    
    if pattern_file is None or preset_file is None or isnumber_file is None:
        return None, None, None
    
    return pattern_file, preset_file, isnumber_file

def validate_dataframe(df: pd.DataFrame, required_cols: List[str]) -> Tuple[bool, str]:
    """Validate that DataFrame has required columns"""
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        return False, f"Missing columns: {', '.join(missing)}"
    return True, ""

def make_pattern(text: str) -> str:
    """Convert numbers to $ pattern"""
    def replacer(match):
        integer_part = match.group(1)
        decimal_part = match.group(2)
        integer_replaced = "$"
        decimal_replaced = "$" * len(decimal_part) if decimal_part else ""
        return integer_replaced + ("." + decimal_replaced if decimal_part else "")
    return re.sub(r"(\d+)(?:\.(\d+))?", replacer, text)

def process_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Process the uploaded DataFrame with validation"""
    # Validate input
    is_valid, error_msg = validate_dataframe(df, REQUIRED_COLUMNS)
    if not is_valid:
        raise ValueError(error_msg)
    
    # Convert to string once
    str_cols = ["Category", "Sub-Category", "Attribute Name", "Preset values"]
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].astype(str)
    
    # Filter empty values
    df = df[~df["Preset values"].str.strip().isin(["", "-", "nan"])].copy()
    df.reset_index(drop=True, inplace=True)
    
    # Create key column
    df["key"] = (
        df["Category"] + "|" +
        df["Sub-Category"] + "|" +
        df["Attribute Name"]
    )
    
    # Create pattern columns
    df["Helper_pattern"] = df["Preset values"].str.replace(r"\d+(\.\d+)?", "$", regex=True)
    df["pattern"] = df["Preset values"].apply(make_pattern)
    df["count"] = df.groupby(["key", "pattern"])["pattern"].transform("count")
    
    return df

def sort_preset_values(df: pd.DataFrame) -> pd.DataFrame:
    """Sort preset values alphabetically within groups"""
    def is_alpha_only(s: str) -> bool:
        return bool(re.fullmatch(r"[A-Za-z\s]+", s))
    
    def sort_groups(val):
        if not isinstance(val, str):
            return val
        groups = [grp.strip() for grp in val.split(" - ")]
        sorted_groups = []
        for grp in groups:
            if ", " in grp:
                parts = [p.strip() for p in grp.split(",")]
                if all(is_alpha_only(p) for p in parts):
                    parts_sorted = sorted(parts, key=lambda x: x.lower())
                    sorted_groups.append(", ".join(parts_sorted))
                else:
                    sorted_groups.append(", ".join(parts))
            else:
                sorted_groups.append(grp.strip())
        return " - ".join(sorted_groups)

    df["sorted Preset values"] = df["Preset values"].apply(sort_groups)
    df["Was Sorted"] = df["Preset values"].astype(str).str.strip() != df["sorted Preset values"].astype(str).str.strip()
    return df

def normalize_case(s: str) -> str:
    """Normalize to lowercase"""
    return s.lower()

def normalize_alpha(s: str) -> str:
    """Remove non-alpha characters and lowercase"""
    return re.sub(r'[^a-zA-Z]+', '', s).lower()

def normalize_alnum(s: str) -> str:
    """Remove non-alphanumeric characters and lowercase"""
    return re.sub(r'[^a-zA-Z0-9]+', '', s).lower()

def normalize_pattern(s: str) -> str:
    """Normalize pattern for comparison"""
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"\s+", "", s)
    s = s.replace("±", "+-")
    return s.lower()

def mark_patterns(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    """Mark rows with IsNumber flag based on pattern matching"""
    df1["pattern_norm"] = df1["Helper_pattern"].apply(normalize_pattern)
    df2["Pattern_norm"] = df2["Pattern"].apply(normalize_pattern)

    patterns_set = set(df2["Pattern_norm"].unique())
    
    # Initialize IsNumber if not exists
    if "IsNumber" not in df1.columns:
        df1["IsNumber"] = None
    
    df1["IsNumber"] = df1.apply(
        lambda row: row["IsNumber"] if pd.notna(row["IsNumber"]) and str(row["IsNumber"]).strip() != ""
        else ("Yes" if row["pattern_norm"] in patterns_set else "No"),
        axis=1
    )
    df1.drop(columns=["pattern_norm"], inplace=True)
    return df1

def merge_isnumber_safely(df: pd.DataFrame, df_isnumber: pd.DataFrame) -> pd.DataFrame:
    """Safely merge IsNumber column from external file"""
    df_isnumber = df_isnumber.copy()
    df_isnumber["key"] = (
        df_isnumber["Category"].astype(str) + "|" +
        df_isnumber["Sub-Category"].astype(str) + "|" +
        df_isnumber["Attribute Name"].astype(str)
    )
    
    # Drop IsNumber if it exists in df to avoid conflicts
    if "IsNumber" in df.columns:
        df = df.drop(columns=["IsNumber"])
    
    # Merge only IsNumber column, removing duplicates first
    isnumber_merge = df_isnumber[["key", "IsNumber"]].drop_duplicates(subset=["key"])
    df = df.merge(isnumber_merge, on="key", how="left")
    
    return df

def extract_numbers_by_pattern(value: str, pattern: str) -> List[float]:
    """Extract numbers from value based on pattern"""
    if not isinstance(value, str) or not isinstance(pattern, str):
        return []
    regex_pattern = re.escape(pattern).replace("\\$", r"(\d*\.?\d+)")
    match = re.match(regex_pattern, value.strip())
    if not match:
        return []
    return [float(x) for x in match.groups() if x]

def compare_two_numbers(numA: List[float], numB: List[float]) -> Tuple[str, str]:
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

def calc_overall_percentage(perc_str: str) -> float:
    """Calculate average percentage from string"""
    if not perc_str:
        return 0.0
    nums = [float(x) for x in perc_str.split(" | ") if x.replace(".", "", 1).isdigit()]
    if not nums:
        return 0.0
    return round(sum(nums) / len(nums), 2)

def calc_worst_percentage(perc_str: str) -> float:
    """Calculate worst (minimum) percentage from string"""
    if not perc_str:
        return 0.0
    nums = [float(x) for x in perc_str.split(" | ") if x.replace(".", "", 1).isdigit()]
    if not nums:
        return 0.0
    return min(nums)

def determine_comment(is_number: bool, per_value1: float, per_value2_nospecial: float, 
                     perc_num1_worst: float, perc_num2_worst: float, 
                     comment_value1: str, is_caseSensitive: bool = False) -> str:
    """Determine comment based on matching results"""
    if comment_value1 == "no match found":
        return "Category not found"
    
    if is_number:
        # Number matching logic
        if per_value1 == SIMILARITY_THRESHOLDS['EXACT'] and \
           per_value2_nospecial == SIMILARITY_THRESHOLDS['EXACT']:
            return "Found Exact Number"
        
        worst_perc = max(perc_num1_worst or 0, perc_num2_worst or 0)
        if worst_perc > SIMILARITY_THRESHOLDS['MINORITY_NUMBER']:
            return "Found with minority Number"
        elif worst_perc > SIMILARITY_THRESHOLDS['MAJORITY_NUMBER']:
            return "Found with majority Number"
        else:
            return "Not Found Number"
    else:
        # String matching logic
        if per_value1 == SIMILARITY_THRESHOLDS['EXACT'] and \
           per_value2_nospecial == SIMILARITY_THRESHOLDS['EXACT']:
            if is_caseSensitive:
                return "Found Exact String with caseSensitive"
            return "Found Exact String"
        
        if per_value1 == 0 and per_value2_nospecial == 0:
            return "Not Found String"
        
        if per_value1 > SIMILARITY_THRESHOLDS['MINORITY_STRING'] and \
           per_value2_nospecial == SIMILARITY_THRESHOLDS['EXACT']:
            return "Found with minority String"
        
        if per_value1 > SIMILARITY_THRESHOLDS['MAJORITY_STRING_PRIMARY'] and \
           per_value2_nospecial > SIMILARITY_THRESHOLDS['MAJORITY_STRING_SECONDARY']:
            return "Found with majority String"
        
        if per_value2_nospecial > SIMILARITY_THRESHOLDS['SIMILAR_STRING']:
            return "Similar String"
        
        return "Not Found String"

def check_exact_match(val1: str, val2: str, case_sensitive: bool = False) -> bool:
    """Check if two values are exactly equal"""
    if case_sensitive:
        return val1.strip() == val2.strip()
    else:
        return val1.strip().lower() == val2.strip().lower()

def compare_files_with_conditional_numbers(df1: pd.DataFrame, df2: pd.DataFrame, 
                                           threshold: int = 0, top_n: int = 2, 
                                           filter_exact_matches: bool = False) -> pd.DataFrame:
    """Compare two dataframes with number and string matching logic"""
    results = []
    total_rows = len(df1)
    
    # Create progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Group df2 by key for faster lookups
    df2_grouped = df2.groupby("key")
    
    for idx, row in df1.iterrows():
        # Update progress every 10 rows
        if idx % 10 == 0 or idx == total_rows - 1:
            progress = (idx + 1) / total_rows
            progress_bar.progress(progress)
            status_text.text(f"Processing row {idx + 1} of {total_rows}...")
        
        key = row["key"]
        val1 = str(row["Preset values"])
        is_number = str(row.get("IsNumber", "")).strip().lower() == "yes"
        is_caseSensitive = str(row.get("isCaseSensitive", "")).strip().lower() == "yes"

        # Fast lookup using grouped data
        if key in df2_grouped.groups:
            candidates = df2_grouped.get_group(key)
        else:
            candidates = pd.DataFrame()
        
        if candidates.empty:
            row_out = row.to_dict()
            base_update = {
                "Value1": "",
                "per_Value1": 0,
                "Comment_Value1": "no match found",
                "Value2_noSpecial": "",
                "per_Value2_noSpecial": 0,
                "HigherPercentage": 0,
                "MatchRank": 0,
                "valueHelper_pattern": "",
                "value2Helper_pattern": "",
                "ContainMatch": "",
                "IsExactMatch": False
            }
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
            comment = determine_comment(is_number, 0, 0, 0, 0, "no match found")
            base_update["Comment"] = comment
            row_out.update(base_update)
            results.append(row_out)
            continue

        # Calculate similarities for all candidates
        sims = []
        for _, cand in candidates.iterrows():
            main_val2 = str(cand["Preset values"])
            val2 = str(cand["Preset values"]).lower()
            is_exact = check_exact_match(val1, main_val2, is_caseSensitive)
            
            if is_caseSensitive:
                sim = fuzz.ratio(val1.lower(), val2)
            else:
                sim = fuzz.ratio(val1, main_val2)
            
            sim_noSpecial = fuzz.ratio(normalize_alnum(val1), normalize_alnum(val2))
            
            sims.append({
                "val2": main_val2,
                "sim": sim,
                "sim_noSpecial": sim_noSpecial,
                "helper_pattern": cand.get("Helper_pattern", ""),
                "is_exact": is_exact
            })

        sims_sorted = sorted(sims, key=lambda x: x["sim"], reverse=True)[:top_n]
        has_exact_match = any(m["is_exact"] for m in sims_sorted)
        
        if filter_exact_matches and has_exact_match:
            continue

        for rank, match in enumerate(sims_sorted, start=1):
            best_val = match["val2"]
            best_score = match["sim"]
            best_score_noSpecial = match["sim_noSpecial"]
            helper_pattern_val = match["helper_pattern"]
            is_exact = match["is_exact"]

            # Determine comment type
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
            higher_percentage = max(best_score, best_score_noSpecial)
            base_update = {
                "Value1": best_val,
                "per_Value1": best_score,
                "Comment_Value1": comment,
                "Value2_noSpecial": best_val,
                "per_Value2_noSpecial": best_score_noSpecial,
                "HigherPercentage": higher_percentage,
                "MatchRank": rank,
                "valueHelper_pattern": helper_pattern_val,
                "value2Helper_pattern": helper_pattern_val,
                "ContainMatch": contain_val,
                "IsExactMatch": is_exact
            }

            if is_number:
                pat_preset = row.get("Helper_pattern", "")
                value_num1, perc_num1 = "", ""
                value_num2, perc_num2 = "", ""
                nums_base = []
                
                if pat_preset == helper_pattern_val:
                    nums_base = extract_numbers_by_pattern(val1, pat_preset)
                    nums1 = extract_numbers_by_pattern(best_val, helper_pattern_val)
                    if nums_base and nums1 and len(nums_base) == len(nums1):
                        value_num1, perc_num1 = compare_two_numbers(nums_base, nums1)
                
                if pat_preset == helper_pattern_val and nums_base:
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

            perc1_worst_val = base_update.get("percentageNumber_Value1_worst", 0) if is_number else None
            perc2_worst_val = base_update.get("percentageNumber_Value2_worst", 0) if is_number else None
            final_comment = determine_comment(
                is_number,
                best_score,
                best_score_noSpecial,
                perc1_worst_val,
                perc2_worst_val,
                comment,
                is_caseSensitive
            )
            base_update["Comment"] = final_comment
            row_out.update(base_update)
            results.append(row_out)

            if best_score == 100 and best_score_noSpecial == 100:
                break
    
    progress_bar.progress(1.0)
    status_text.text("Processing complete!")
    
    df_out = pd.DataFrame(results)
    
    # Sort results
    if any(str(row.get("IsNumber", "")).strip().lower() == "yes" for _, row in df1.iterrows()):
        sort_cols = [
            "HigherPercentage",
            "per_Value1",
            "per_Value2_noSpecial",
            "percentageNumber_Value1_worst",
            "percentageNumber_Value2_worst"
        ]
        existing_cols = [c for c in sort_cols if c in df_out.columns]
        if len(existing_cols) > 2:
            df_out = df_out.sort_values(by=existing_cols, ascending=[False]*len(existing_cols))
    
    return df_out

def duplicate_rows_for_case_sensitivity(df: pd.DataFrame) -> pd.DataFrame:
    """Duplicate rows to test both case-sensitive and case-insensitive matching"""
    df_yes = df.copy()
    df_yes['isCaseSensitive'] = 'Yes'
    df_no = df.copy()
    df_no['isCaseSensitive'] = 'No'
    df_combined = pd.concat([df_yes, df_no], ignore_index=True)
    return df_combined

def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Remove duplicates in summary file"""
    subset_cols = ["key", "Value1", "HigherPercentage", "Comment"]
    available_cols = [col for col in subset_cols if col in df.columns]
    
    if not available_cols:
        return df
    
    df_dedup = df.drop_duplicates(subset=available_cols, keep="first")
    return df_dedup.reset_index(drop=True)

def create_summary(df_result: pd.DataFrame) -> pd.DataFrame:
    """Create summary dataframe with optimized memory usage"""
    summary_cols = ["Category", "Sub-Category", "Attribute Name",
                               "Preset values", "isCaseSensitive","IsNumber", "Max Value", "Unit", "key",
                               "Value1", "HigherPercentage", "Comment", "IsExactMatch"]
    
    # Only select columns that exist
    existing_cols = [c for c in summary_cols if c in df_result.columns]
    df_summary = df_result[existing_cols].copy()
    
    # Remove duplicates - FIXED BUG
    df_summary = remove_duplicates(df_summary)
    
    # Convert to categorical for memory savings
    categorical_cols = ["Category", "Sub-Category", "Attribute Name", 
                        "isCaseSensitive", "Comment"]
    for col in categorical_cols:
        if col in df_summary.columns:
            df_summary[col] = df_summary[col].astype('category')
    
    return df_summary

def main():
    st.set_page_config(
        page_title="Data Processing App",
        page_icon="📊",
        layout="wide"
    )
    # if st.button("🔄 Refresh Reference Files"):
    #     st.cache_data.clear()
    #     st.rerun()
    
    st.title("📊 Data Processing and Comparison Tool")
    st.markdown("Upload your Excel file and let the system process it against reference data from Google Drive.")
    
    # Sidebar parameters
    st.sidebar.header("⚙️ Parameters")
    threshold = st.sidebar.slider("Similarity Threshold", 0, 100, 50, 
                                  help="Minimum similarity score for matches")
    top_n = st.sidebar.selectbox("Top N Matches", [1, 2, 3, 4, 5], index=0, 
                                 help="Number of top matches to consider")
    
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Filtering Options")
    filter_exact = st.sidebar.checkbox(
        "Filter Out Exact Matches", 
        value=False,
        help="Remove rows where an exact match is found in the comparison"
    )
    
    # File upload
    st.header("📁 File Upload")
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload the Excel file you want to process"
    )
    
    if uploaded_file is not None:
        try:
            with st.spinner("Loading uploaded file..."):
                df_uploaded = pd.read_excel(uploaded_file, dtype=str)
            
            st.success(f"✅ Uploaded file loaded successfully! ({len(df_uploaded)} rows)")
            
            with st.expander("📋 Preview of Uploaded Data"):
                st.dataframe(df_uploaded.head(10), use_container_width=True)
            
            # Process data
            st.header("🔄 Processing Data")
            
            with st.spinner("Loading reference files from Google Drive..."):
                pattern_file, preset_file, isnumber_file = load_reference_files()
            
            if pattern_file is None or preset_file is None or isnumber_file is None:
                st.error("❌ Failed to load reference files from Google Drive. Please check the URLs.")
                return
            
            try:
                df_pattern = pd.read_excel(pattern_file, dtype=str)
                # df_preset = pd.read_excel(preset_file, dtype=str)
                
                df_preset = pd.read_excel(preset_file, dtype=str)
                df_preset = pd.concat(
                        pd.read_excel(preset_file, sheet_name=None, dtype=str).values(),
                        ignore_index=True
                    )
                df_isnumber = pd.read_excel(isnumber_file, dtype=str)
                st.success("✅ Reference files loaded successfully!")
                
                with st.spinner("Processing your data..."):
                    # Process uploaded file
                    df_processed = process_excel(df_uploaded.copy())
                    
                    # Merge IsNumber flag from external file
                    df_processed = merge_isnumber_safely(df_processed, df_isnumber)
                    
                    # Sort preset values
                    df_sorted = sort_preset_values(df_processed)
                    
                    # Mark patterns
                    df_marked = mark_patterns(df_sorted, df_pattern)
                    
                    # Duplicate for case sensitivity
                    df_with_case_sensitivity = duplicate_rows_for_case_sensitivity(df_marked)
                    st.info(f"ℹ️ Created {len(df_with_case_sensitivity)} rows ({len(df_marked)} original × 2 for case sensitivity variations)")
                    
                    # Compare files
                    df_result = compare_files_with_conditional_numbers(
                        df_with_case_sensitivity, df_preset, threshold, top_n, 
                        filter_exact_matches=filter_exact
                    )
                
                st.success("🎉 Processing completed successfully!")
                
                # Display results
                st.header("📊 Results")
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    st.metric("Total Rows", len(df_result))
                with col2:
                    exact_matches = len(df_result[df_result.get('IsExactMatch', pd.Series([False]*len(df_result))) == True])
                    st.metric("Exact Matches", exact_matches)
                with col3:
                    case_yes = len(df_result[df_result['isCaseSensitive'] == 'Yes'])
                    st.metric("Case Sensitive: Yes", case_yes)
                with col4:
                    case_no = len(df_result[df_result['isCaseSensitive'] == 'No'])
                    st.metric("Case Sensitive: No", case_no)
                with col5:
                    no_matches = len(df_result[df_result['Comment'] == 'Category not found'])
                    st.metric("No Matches", no_matches)
                
                if filter_exact:
                    st.info(f"ℹ️ Exact matches have been filtered out. Showing only non-exact matches.")
                
                with st.expander("📋 Full Results", expanded=True):
                    st.dataframe(df_result, use_container_width=True)
                
                # Create summary - FIXED BUG HERE
                df_summary = create_summary(df_result)
                
                # Download buttons
                st.header("💾 Download Results")
                col1, col2 = st.columns(2)
                
                with col1:
                    output_buffer = BytesIO()
                    df_result.to_excel(output_buffer, index=False, engine='openpyxl')
                    output_buffer.seek(0)
                    st.download_button(
                        label="📥 Download Full Results",
                        data=output_buffer.getvalue(),
                        file_name=f"processed_results_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col2:
                    summary_buffer = BytesIO()
                    df_summary.to_excel(summary_buffer, index=False, engine='openpyxl')
                    summary_buffer.seek(0)
                    st.download_button(
                        label="📥 Download Summary",
                        data=summary_buffer.getvalue(),
                        file_name=f"summary_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Analysis section
                st.header("📈 Analysis")
                tab1, tab2, tab3 = st.tabs(["Comment Distribution", "Case Sensitivity Analysis", "Higher Percentage Distribution"])
                
                with tab1:
                    st.subheader("Comment Distribution")
                    comment_counts = df_result['Comment'].value_counts()
                    if not comment_counts.empty:
                        st.bar_chart(comment_counts)
                        with st.expander("📊 Comment Distribution Details"):
                            st.dataframe(
                                pd.DataFrame({
                                    'Comment': comment_counts.index,
                                    'Count': comment_counts.values,
                                    'Percentage': (comment_counts.values / len(df_result) * 100).round(2)
                                }),
                                use_container_width=True
                            )
                    else:
                        st.info("No comment data available")
                
                with tab2:
                    st.subheader("Case Sensitivity Analysis")
                    case_counts = df_result['isCaseSensitive'].value_counts()
                    if not case_counts.empty:
                        st.bar_chart(case_counts)
                        st.subheader("Average Higher Percentage by Case Sensitivity")
                        avg_by_case = df_result.groupby('isCaseSensitive')['HigherPercentage'].mean()
                        st.dataframe(
                            pd.DataFrame({
                                'Case Sensitivity': avg_by_case.index,
                                'Average Higher %': avg_by_case.values.round(2)
                            }),
                            use_container_width=True
                        )
                    else:
                        st.info("No case sensitivity data available")
                
                with tab3:
                    st.subheader("Higher Percentage Statistics")
                    if 'HigherPercentage' in df_result.columns:
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.metric("Mean", f"{df_result['HigherPercentage'].mean():.2f}%")
                        with col2:
                            st.metric("Median", f"{df_result['HigherPercentage'].median():.2f}%")
                        with col3:
                            st.metric("Min", f"{df_result['HigherPercentage'].min():.2f}%")
                        with col4:
                            st.metric("Max", f"{df_result['HigherPercentage'].max():.2f}%")
                        
                        # Histogram
                        fig, ax = plt.subplots(figsize=(10, 4))
                        ax.hist(df_result['HigherPercentage'], bins=20, edgecolor='black', alpha=0.7)
                        ax.set_xlabel('Higher Percentage')
                        ax.set_ylabel('Frequency')
                        ax.set_title('Distribution of Higher Percentage Values')
                        ax.grid(axis='y', alpha=0.3)
                        st.pyplot(fig)
                    else:
                        st.info("No percentage data available")
                
            except Exception as e:
                st.error(f"❌ Error processing reference files: {str(e)}")
                st.exception(e)
                
        except Exception as e:
            st.error(f"❌ Error reading uploaded file: {str(e)}")
            st.info("Please make sure your file is a valid Excel file with the required columns.")
            st.exception(e)
    
    else:
        st.info("👆 Please upload an Excel file to get started.")
        
        st.header("📋 Required File Format")
        st.markdown("""
        Your Excel file should contain the following columns:
        - **Category**: Main category of the item  
        - **Sub-Category**: Subcategory of the item  
        - **Attribute Name**: Name of the attribute  
        - **Preset values**: The values to be processed  
        - **Max Value**: Maximum value (if applicable)  
        - **Unit**: Unit of measurement (if applicable)  

        **Note**: Each row will be automatically duplicated to test both case-sensitive (Yes) and case-insensitive (No) matching.
        """)
        
        st.header("🔍 How It Works")
        with st.expander("Click to learn more about the processing pipeline"):
            st.markdown("""
            ### Processing Steps:
            1. **Upload**: Your Excel file is validated for required columns
            2. **Reference Loading**: Pattern, preset, and IsNumber files are loaded from Google Drive
            3. **Processing**: 
               - Creates unique keys from Category, Sub-Category, and Attribute Name
               - Generates patterns from numeric values
               - Merges IsNumber flags from reference file
               - Sorts preset values alphabetically
            4. **Case Sensitivity**: Duplicates data to test both case-sensitive and case-insensitive matching
            5. **Comparison**: 
               - Fuzzy matching using RapidFuzz library
               - Number extraction and comparison for numeric patterns
               - String similarity scoring
            6. **Results**: Detailed results with match percentages, comments, and exact match flags
            
            ### Key Features:
            - ✅ Automatic pattern recognition for numbers
            - ✅ Case-sensitive and case-insensitive matching
            - ✅ Fuzzy string matching with multiple algorithms
            - ✅ Number comparison with percentage calculations
            - ✅ Detailed comment classification
            - ✅ Progress tracking for large files
            - ✅ Duplicate removal in summary
            """)

if __name__ == "__main__":
    main()

