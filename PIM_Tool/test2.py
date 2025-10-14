import streamlit as st
import pandas as pd
import re
import tempfile
import os
from rapidfuzz import fuzz
import numpy as np
import requests
from io import BytesIO

# Google Drive URLs for the reference files
DRIVE_FILES = {
    "pattern_file": "https://docs.google.com/spreadsheets/d/1W4oA3BtsmWdNhSQOLceCFAfUiKE9BW1_/export?format=xlsx",
    "preset_file": "https://docs.google.com/spreadsheets/d/1IhOFscHAvOH8erop8xg8HVWn6MMvMXvo/export?format=xlsx",
    "isnumber_file": "https://docs.google.com/spreadsheets/d/1miXOKaln_uj5x52-vQmvqVtvKUKJuOnU/export?format=xlsx"
}

def download_file_from_drive(url):
    """Download file from Google Drive URL"""
    try:
        # Convert edit URL to export URL if needed
        if "/edit" in url:
            match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
            if match:
                doc_id = match.group(1)
                url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
        
        response = requests.get(url)
        response.raise_for_status()
        return BytesIO(response.content)
    except Exception as e:
        st.error(f"Error downloading file: {str(e)}")
        return None

@st.cache_data
def load_reference_files():
    """Load reference files from Google Drive (pattern, preset, isnumber)"""
    pattern_file = download_file_from_drive(DRIVE_FILES["pattern_file"])
    preset_file = download_file_from_drive(DRIVE_FILES["preset_file"])
    isnumber_file = download_file_from_drive(DRIVE_FILES["isnumber_file"])
    
    if pattern_file is None or preset_file is None or isnumber_file is None:
        return None, None, None
    
    return pattern_file, preset_file, isnumber_file

def make_pattern(text):
    def replacer(match):
        integer_part = match.group(1)
        decimal_part = match.group(2)
        integer_replaced = "$"
        decimal_replaced = "$" * len(decimal_part) if decimal_part else ""
        return integer_replaced + ("." + decimal_replaced if decimal_part else "")
    return re.sub(r"(\d+)(?:\.(\d+))?", replacer, text)

def process_excel(df):
    """Process the uploaded DataFrame"""
    df = df[~df["Preset values"].astype(str).str.strip().isin(["", "-"])]
    df.reset_index(drop=True, inplace=True)
    
    df["key"] = (
        df["Category"].astype(str) + "|" +
        df["Sub-Category"].astype(str) + "|" +
        df["Attribute Name"].astype(str)
    )
    
    df["Helper_pattern"] = df["Preset values"].astype(str).apply(
        lambda x: re.sub(r"\d+(\.\d+)?", "$", x)
    )
    df["pattern"] = df["Preset values"].astype(str).apply(make_pattern)
    df["count"] = df.groupby(["key", "pattern"])["pattern"].transform("count")
    
    return df

def sort_preset_values(df):
    def is_alpha_only(s):
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
    return s.lower()

def normalize_alpha(s: str) -> str:
    return re.sub(r'[^a-zA-Z]+', '', s).lower()

def normalize_alnum(s: str) -> str:
    return re.sub(r'[^a-zA-Z0-9]+', '', s).lower()

def normalize_pattern(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"\s+", "", s)
    s = s.replace("±", "+-")
    return s.lower()

def mark_patterns(df1, df2):
    df1["pattern_norm"] = df1["Helper_pattern"].apply(normalize_pattern)
    df2["Pattern_norm"] = df2["Pattern"].apply(normalize_pattern)

    patterns_set = set(df2["Pattern_norm"].unique())
    # If no IsNumber column, create it (will be filled partially from external file)
    if "IsNumber" not in df1.columns:
        df1["IsNumber"] = None
    df1["IsNumber"] = df1.apply(
        lambda row: row["IsNumber"] if pd.notna(row["IsNumber"]) and str(row["IsNumber"]).strip() != ""
        else ("Yes" if row["pattern_norm"] in patterns_set else "No"),
        axis=1
    )
    df1.drop(columns=["pattern_norm"], inplace=True)
    return df1

def extract_numbers_by_pattern(value: str, pattern: str):
    if not isinstance(value, str) or not isinstance(pattern, str):
        return []
    regex_pattern = re.escape(pattern).replace("\\$", r"(\d*\.?\d+)")
    match = re.match(regex_pattern, value.strip())
    if not match:
        return []
    return [float(x) for x in match.groups() if x]

def compare_two_numbers(numA, numB):
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
    if not perc_str:
        return ""
    nums = [float(x) for x in perc_str.split(" | ") if x.replace(".", "", 1).isdigit()]
    if not nums:
        return ""
    return round(sum(nums) / len(nums), 2)

def calc_worst_percentage(perc_str):
    if not perc_str:
        return ""
    nums = [float(x) for x in perc_str.split(" | ") if x.replace(".", "", 1).isdigit()]
    if not nums:
        return ""
    return min(nums)

def determine_comment(is_number, per_value1, per_value2_nospecial, perc_num1_worst, perc_num2_worst, comment_value1, is_caseSensitive=False):
    if comment_value1 == "no match found":
        return "Category not found"
    if is_number:
        if per_value1 == 100 and per_value2_nospecial == 100:
            return "Found Exact Number"
        elif (perc_num1_worst and perc_num1_worst > 97) or (perc_num2_worst and perc_num2_worst > 97):
            return "Found with minority Number"
        elif (perc_num1_worst and perc_num1_worst > 80) or (perc_num2_worst and perc_num2_worst > 80):
            return "Found with majority Number"
        else:
            return "Not Found Number"
    else:
        if per_value1 == 100 and per_value2_nospecial == 100 and is_caseSensitive:
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

def check_exact_match(val1, val2, case_sensitive=False):
    if case_sensitive:
        return val1.strip() == val2.strip()
    else:
        return val1.strip().lower() == val2.strip().lower()

def compare_files_with_conditional_numbers(df1, df2, threshold=0, top_n=2, filter_exact_matches=False):
    results = []
    for idx, row in df1.iterrows():
        key = row["key"]
        val1 = str(row["Preset values"])
        is_number = str(row.get("IsNumber", "")).strip().lower() == "yes"
        is_caseSensitive = str(row.get("isCaseSensitive", "")).strip().lower() == "yes"

        candidates = df2[df2["key"] == key]
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
            comment = determine_comment(is_number, 0, 0,
                                        0 if is_number else None,
                                        0 if is_number else None,
                                        "no match found")
            base_update["Comment"] = comment
            row_out.update(base_update)
            results.append(row_out)
            continue

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

            perc1_worst_val = base_update.get("percentageNumber_Value1_worst", 0) if is_number else None
            perc2_worst_val = base_update.get("percentageNumber_Value2_worst", 0) if is_number else None
            comment = determine_comment(
                is_number,
                best_score,
                best_score_noSpecial,
                perc1_worst_val,
                perc2_worst_val,
                comment,
                is_caseSensitive
            )
            base_update["Comment"] = comment
            row_out.update(base_update)
            results.append(row_out)

            if best_score == 100 and best_score_noSpecial == 100:
                break

    df_out = pd.DataFrame(results)
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

def duplicate_rows_for_case_sensitivity(df):
    df_yes = df.copy()
    df_yes['isCaseSensitive'] = 'Yes'
    df_no = df.copy()
    df_no['isCaseSensitive'] = 'No'
    df_combined = pd.concat([df_yes, df_no], ignore_index=True)
    return df_combined

def main():
    st.set_page_config(
        page_title="Data Processing App",
        page_icon="📊",
        layout="wide"
    )
    st.title("📊 Data Processing and Comparison Tool")
    st.markdown("Upload your Excel file and let the system process it against reference data from Google Drive.")
    
    st.sidebar.header("⚙️ Parameters")
    threshold = st.sidebar.slider("Similarity Threshold", 0, 100, 50, help="Minimum similarity score for matches")
    top_n = st.sidebar.selectbox("Top N Matches", [1, 2, 3, 4, 5], index=0, help="Number of top matches to consider")
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Filtering Options")
    filter_exact = st.sidebar.checkbox(
        "Filter Out Exact Matches", 
        value=False,
        help="Remove rows where an exact match is found in the comparison"
    )
    
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
                st.dataframe(df_uploaded.head(), use_container_width=True)
            
            st.header("🔄 Processing Data")
            with st.spinner("Loading reference files from Google Drive..."):
                pattern_file, preset_file, isnumber_file = load_reference_files()
            if pattern_file is None or preset_file is None or isnumber_file is None:
                st.error("❌ Failed to load reference files from Google Drive. Please check the URLs.")
                return
            
            try:
                df_pattern = pd.read_excel(pattern_file, dtype=str)
                df_preset = pd.read_excel(preset_file, dtype=str)
                df_isnumber = pd.read_excel(isnumber_file, dtype=str)
                st.success("✅ Reference files loaded successfully!")
                
                with st.spinner("Processing your data..."):
                    df_processed = process_excel(df_uploaded.copy())
                    
                    # ---- NEW: Merge IsNumber flag from isnumber file ----
                    df_isnumber["key"] = (
                        df_isnumber["Category"].astype(str) + "|" +
                        df_isnumber["Sub-Category"].astype(str) + "|" +
                        df_isnumber["Attribute Name"].astype(str)
                    )
                    # Only keep relevant columns from df_isnumber
                    df_processed = df_processed.merge(
                        df_isnumber[["key", "IsNumber"]],
                        on="key",
                        how="left"
                    )
                    # -------------------------------------------------------
                    
                    df_sorted = sort_preset_values(df_processed)
                    df_marked = mark_patterns(df_sorted, df_pattern)
                    df_with_case_sensitivity = duplicate_rows_for_case_sensitivity(df_marked)
                    st.info(f"ℹ️ Created {len(df_with_case_sensitivity)} rows ({len(df_marked)} original × 2 for case sensitivity variations)")
                    
                    df_result = compare_files_with_conditional_numbers(
                        df_with_case_sensitivity, df_preset, threshold, top_n, filter_exact_matches=filter_exact
                    )
                
                st.success("🎉 Processing completed successfully!")
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
                
                summary_cols = ["Category", "Sub-Category", "Attribute Name",
                                "Preset values", "isCaseSensitive", "Max Value", "Unit", "key",
                                "Value1", "HigherPercentage", "Comment", "IsExactMatch"]
                existing_summary_cols = [c for c in summary_cols if c in df_result.columns]
                df_summary = df_result[existing_summary_cols].copy()
                
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
                
                st.header("📈 Analysis")
                tab1, tab2, tab3 = st.tabs(["Comment Distribution", "Case Sensitivity Analysis", "Higher Percentage Distribution"])
                with tab1:
                    comment_counts = df_result['Comment'].value_counts()
                    if not comment_counts.empty:
                        st.bar_chart(comment_counts)
                        with st.expander("📊 Comment Distribution Details"):
                            st.dataframe(
                                pd.DataFrame({
                                    'Comment': comment_counts.index,
                                    'Count': comment_counts.values,
                                    'Percentage': (comment_counts.values / len(df_result) * 100).round(2)
                                })
                            )
                with tab2:
                    case_counts = df_result['isCaseSensitive'].value_counts()
                    if not case_counts.empty:
                        st.bar_chart(case_counts)
                        st.subheader("Average Higher Percentage by Case Sensitivity")
                        avg_by_case = df_result.groupby('isCaseSensitive')['HigherPercentage'].mean()
                        st.dataframe(
                            pd.DataFrame({
                                'Case Sensitivity': avg_by_case.index,
                                'Average Higher %': avg_by_case.values.round(2)
                            })
                        )
                with tab3:
                    st.subheader("Higher Percentage Statistics")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Mean", f"{df_result['HigherPercentage'].mean():.2f}%")
                    with col2:
                        st.metric("Median", f"{df_result['HigherPercentage'].median():.2f}%")
                    with col3:
                        st.metric("Min", f"{df_result['HigherPercentage'].min():.2f}%")
                    with col4:
                        st.metric("Max", f"{df_result['HigherPercentage'].max():.2f}%")
                    import matplotlib.pyplot as plt
                    fig, ax = plt.subplots(figsize=(10, 4))
                    ax.hist(df_result['HigherPercentage'], bins=20, edgecolor='black')
                    ax.set_xlabel('Higher Percentage')
                    ax.set_ylabel('Frequency')
                    ax.set_title('Distribution of Higher Percentage Values')
                    st.pyplot(fig)
                
            except Exception as e:
                st.error(f"❌ Error processing reference files: {str(e)}")
        except Exception as e:
            st.error(f"❌ Error reading uploaded file: {str(e)}")
            st.info("Please make sure your file is a valid Excel file with the required columns.")
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

if __name__ == "__main__":
    main()
