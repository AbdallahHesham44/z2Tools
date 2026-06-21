# --- START OF FILE result_combiner.py ---
#FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
#############################################
# MODULE: RESULT COMBINER
# Purpose: Merges the output of the fixed processing pipeline
#          with the output of the detailed analysis pipeline.
#          Adds results from the regex-based classifier.
#############################################
import pandas as pd
import os
import traceback
import streamlit as st
import re
from analysis_helpers import update_number_with_identifier      #  <-- add this line

from regex_classifier import add_mixed_types_detail_dvt
from prefix_utils import STRING_KEYWORDS
# Import from other project modules
from mapping_utils import read_mapping_file
from config import MAPPING_FILE_LOCAL # Use config for mapping file path
from regex_classifier import build_mixed_types_detail 
# Import the new regex classifier function
try:
    from regex_classifier import run_regex_classification
except ImportError:
    st.error("Combine Results: Failed to import run_regex_classification.")
    # Define a dummy function to prevent crashing if import fails
    def run_regex_classification(df, mapping_file_path):
        st.warning("Regex Classifier module not found. Skipping regex classification.")
        df['Classification_New'] = "Not Run"
        df['Reason_New'] = "Not Run"
        df['DetailedValueType_New'] = "Not Run"
        return df

from value_table import (
    compute_value__, compute_multiplication, compute_space_separator,
    compute_value_pattern, compute_normalized_value__,
    compute_integar_value, compute_fraction_value, compute_fraction_digits
)

KEYWORDS = {kw.lower() for kw in STRING_KEYWORDS}   # ensure lower-case set
     # put at top if you prefer
# ----------------------------------------------------------------------
# 3-NEW)  Inject “/String 1,2” (or similar) exactly where it belongs
# ----------------------------------------------------------------------
def _inject_string_suffix(label: str, col_name: str, chunk_list: list[int]) -> str:
    """
    Converts:
        Multiple (Single Value) [1][0] x4
    into:
        Multiple (Single Value/String 1,2) [1][0] x4
    using the same placement rules you had for prefix injection.
    """
    if not chunk_list or "/String" in str(label):
        return label                     # nothing to add / already done

    suffix = "/String " + ", ".join(map(str, chunk_list))
    is_dvt = "DetailedValueType" in col_name
    label   = str(label)

    if label.startswith("Multiple ("):
        close = label.find(")")
        if close != -1:
            return f"{label[:close]}{suffix}{label[close:]}"
        return label + suffix

    if is_dvt:                             # keep counters at far right
        anno = label.find(" [")
        return (f"{label[:anno]}{suffix}{label[anno:]}" if anno != -1
                else label + suffix)
    return label + suffix



# ----------------------------------------------------------------------

def _mixed_row_detail(row):
    if not row.get("MixedTypesFlag"):                      # fast-path for non-mixed rows
        return row.get("DetailedValueType", "")
    # guard against missing/nulls
    if pd.isna(row.get("MixedTypesDetail")):
        return ""
    return build_mixed_types_detail(row["Value_Normalized"], row["MixedTypesDetail"])


def _make_string_code(code: str) -> str:
    """
    Correctly turns a "Unit" code into a "String" code by only modifying
    the attribute part, preserving the original type.
    e.g., M2-RV-Un0  →  M2-RV-Sn0
          M0-SV-U0   →  M0-SV-S0
    """
    m_chunk, block_type, attribute_suffix = code.split('-', 2)
    
    # Only change the 'U' to an 'S' in the attribute suffix.
    # This correctly preserves 'n0', 'x0', '0', '1', etc.
    new_attribute_suffix = attribute_suffix.replace('U', 'S', 1)
    
    return f"{m_chunk}-{block_type}-{new_attribute_suffix}"

def _post_process_string_keywords(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean up disguised keyword rows. The new code will now be like M0-SV-S0.
    """
    st.write("DEBUG (Post-Processing): Starting universal clean-up (corrected code gen)...")

    value_col_name = 'Value_processed' if 'Value_processed' in df.columns else 'Value'
    required = {'Attribute', value_col_name, 'Code', 'Main Key'}
    if not required.issubset(df.columns):
        st.warning(f"Post-Processing: required cols missing, skipping. Missing: {required - set(df.columns)}")
        return df

    val_lower = df[value_col_name].astype(str).str.lower()
    kw_mask = (df['Attribute'] == 'Unit') & val_lower.isin(KEYWORDS)

    if not kw_mask.any():
        st.write("DEBUG: no disguised keywords found; nothing to do.")
        return df

    idx_kw_rows   = df[kw_mask].index
    idx_to_drop   = []
    modifications = []

    for idx in idx_kw_rows:
        row       = df.loc[idx]
        code      = row['Code']
        main_key  = row['Main Key']

        twin_code = re.sub(r'-U', '-V', code, count=1) # Correctly finds twin value code
        twin_mask = (
            (df['Main Key'] == main_key) &
            (df['Code'] == twin_code) &
            (df[value_col_name].astype(str) == '5')
        )
        twin_rows = df[twin_mask]

        if not twin_rows.empty:
            idx_to_drop.append(twin_rows.index[0])
            
            # <<< CHANGE: No more 'hybrid' logic needed here >>>
            new_code = _make_string_code(code) # Just pass the original code
            # <<< END OF CHANGE >>>

            modifications += [
                (idx, 'Attribute', 'String'),
                (idx, 'Category',  'String'), # Category becomes 'String' for clarity
                (idx, 'Code',      new_code)
            ]
            if 'New_FeatureCode' in df.columns:
                modifications.append((idx, 'New_FeatureCode', code_to_new_feature(new_code)))
        else:
            st.warning(f"Could not find twin row for keyword '{row[value_col_name]}' ({code}).")
            continue

    for row_idx, col, new_val in modifications:
        df.at[row_idx, col] = new_val
        
    # Rename 'Value_processed' back to 'Value' *before* dropping.
    if 'Value_processed' in df.columns and value_col_name == 'Value_processed':
        if 'Value' in df.columns and 'Value' != 'Value_processed': # Avoid error if 'Value' already gone
             df = df.drop(columns=['Value'])
        df = df.rename(columns={'Value_processed': 'Value'})

    df_clean = df.drop(index=idx_to_drop).reset_index(drop=True)

    st.write(f"DEBUG: rewrote {len(idx_kw_rows)} keyword rows; removed {len(idx_to_drop)} fake '5' rows.")
    return df_clean



def code_to_new_feature(code: str) -> str:
    """
    Turn a FeatureCode into its human-readable form.
    Now understands codes where the attribute is 'S' (e.g., M0-SV-S0).
    """
    if not isinstance(code, str) or '-' not in code:
        return ""
    
    parts = code.split('-')
    if len(parts) != 3:
        return ""
    first, second, third = parts
    
    p_idx = first[1:]
    m = re.search(r'(\d+)', third)
    n_idx = m.group(1) if m else ""
    suffix = f"P{p_idx}/M{n_idx}"
    
    # Root word logic is unchanged
    root = "Condition" if len(second) > 1 and second[1] == "C" else "Main"
    
    # Qualifier logic is unchanged
    if third.startswith(("Un", "Vn", "Sn", "Pn")): # Added Sn, Pn for future
        qualifier = "Min. "
    elif third.startswith(("Ux", "Vx", "Sx", "Px")): # Added Sx, Px for future
        qualifier = "Max. "
    else:
        qualifier = ""
    
    # <<< START OF MODIFICATION: Dimension logic is simplified >>>
    # The second part (SV, RV, RC) now correctly tells us the type.
    # The third part tells us the dimension (Value, Unit, Prefix, or String).
    if third.startswith("S"):  # S0, Sn0, Sx0, S1, S2...
        dimension = "string"
    elif third.startswith("P"):
        dimension = "prefix"
    elif third.startswith("U"):
        dimension = "unit"
    else: # Default to value (covers V0, Vn0, Vx0, V1, V2...)
        dimension = "value"
    # <<< END OF MODIFICATION >>>
    
    return f"{qualifier}{root} {dimension} - {suffix}"

def _regenerate_thousands_separators_rc(s_val: str) -> str:
    s_val = str(s_val)
    # Regex: optional sign, optional integer part, optional decimal part (dot included), optional non-numeric suffix
    match = re.fullmatch(r"([+-]?)(\d*)((?:\.\d+)?)([a-zA-Zµ]*)", s_val)
    if match:
        sign_part, integer_part_str, decimal_part_str, suffix_part = match.groups()
        
        sign_part = sign_part or ""
        integer_part_str = integer_part_str or ""
        decimal_part_str = decimal_part_str or "" 
        suffix_part = suffix_part or ""

        if not integer_part_str and not decimal_part_str.replace('.', ''):
            return s_val 

        try:
            if integer_part_str: 
                formatted_integer_part = f"{int(integer_part_str):,}"
            else: 
                formatted_integer_part = ""
            
            return f"{sign_part}{formatted_integer_part}{decimal_part_str}{suffix_part}"
        except ValueError:
            return s_val 
    return s_val

# MODIFIED combine_results
def combine_results(processed_df: pd.DataFrame,
                    analysis_file: str,
                    output_file="final_combined.xlsx"):
    """
    Combines processed data (from fixed pipeline DataFrame) with detailed analysis
    results (read from Excel file). Adds regex-based classification results.
    Merges based on the preprocessed value string ('Main Key').
    Error reporting is done via individual error columns instead of a single 'ErrorReason' column.

    Args:
        processed_df (pd.DataFrame): DataFrame output from the fixed pipeline.
                                     Expected to have 'Main Key' (preprocessed value)
                                     and potentially 'Value_E', 'Value_Normalized', etc.
        analysis_file (str): Path to the Excel file generated by the detailed analysis pipeline.
                             Expected to have a 'Value' column matching 'Main Key'.
        output_file (str): Path to save the final combined Excel file.

    Returns:
        str or None: The path to the output file on success, None on failure.
    """
    st.write("DEBUG: Starting combine_results function...")
    try:
        # --- Input Validation ---
        if not isinstance(processed_df, pd.DataFrame):
            st.error("Combine Error: Processed data input must be a pandas DataFrame.")
            return None
        if processed_df.empty:
            st.warning("Combine Warning: Processed DataFrame (from fixed pipeline) is empty.")
            # Decide if returning an empty structure or None is better.
            # Let's try to proceed, merge might result in empty anyway.
        if 'Main Key' not in processed_df.columns:
            st.error("Combine Error: 'Main Key' column missing in processed data DataFrame. Cannot merge.")
            return None
        # Ensure Value_Normalized exists for the regex classifier
        if 'Value_Normalized' not in processed_df.columns:
            st.error("Combine Error: 'Value_Normalized' column missing in processed data DataFrame. Cannot run regex classification.")
            # Attempt to use 'Main Key' as a fallback? Or fail? Let's fail for now.
            return None

        try:
            if not os.path.exists(analysis_file):
                st.error(f"Combine Error: Analysis file not found at '{analysis_file}'.")
                return None
            df_analysis = pd.read_excel(analysis_file)
            st.write(f"DEBUG: Read analysis file '{analysis_file}'. Shape: {df_analysis.shape}")
        except Exception as e:
            st.error(f"Combine Error: Failed to read analysis file '{analysis_file}': {e}")
            return None

        if df_analysis.empty:
            st.warning("Combine Warning: Analysis DataFrame read from file is empty.")
        if 'Value' not in df_analysis.columns and not df_analysis.empty:
            st.error("Combine Error: 'Value' column missing in non-empty analysis data. Cannot use it as merge key.")
            return None

        st.write("DEBUG: Columns in processed_df (fixed pipeline output):", processed_df.columns.tolist())
        st.write("DEBUG: Columns in df_analysis (detailed analysis output):", df_analysis.columns.tolist())

        # --- Define Columns to Merge from Analysis ---
        analysis_cols_to_merge = [
            "Value", "Classification", "DetailedValueType", "ExtractedPrefix", "Identifiers",
            "SubValueCount", "ConditionCount", "MainItemCount", "HasRangeInMain",
            "HasMultiValueInMain", "HasRangeInCondition", "HasMultipleConditions",
            "Normalized Unit", "Absolute Unit", "MainUnits", "MainDistinctUnitCount",
            "MainUnitsConsistent", "ConditionUnits", "ConditionDistinctUnitCount",
            "ConditionUnitsConsistent", "OverallUnitConsistency", "ParsingErrorFlag",
            "SubValueUnitVariationSummary", "MainNumericValues", "ConditionNumericValues",
            "MainMultipliers", "ConditionMultipliers", "MainBaseUnits", "ConditionBaseUnits",
            "NormalizedMainValues", "NormalizedConditionValues", "MinNormalizedValue",
            "MaxNormalizedValue", "SingleUnitForAllSubs", "AllDistinctUnitsUsed"
        ]
        # V START: Fix from user suggestion to ensure all analysis columns are present even if empty
        # Ensure all expected analysis columns are present in df_analysis, adding them with None if missing
        for col in analysis_cols_to_merge:
            if col not in df_analysis.columns:
                df_analysis[col] = None # Or pd.NA, or appropriate empty value
        # V END
        existing_analysis_cols = [col for col in analysis_cols_to_merge if col in df_analysis.columns]


        if not df_analysis.empty and ("Value" not in existing_analysis_cols):
            st.error("Combine Error: Analysis data is missing the merge key 'Value'.") #Should not happen with fix above
            return None
        st.write(f"DEBUG: Merging analysis columns: {existing_analysis_cols}")

        df_analysis_subset = pd.DataFrame()
        if not df_analysis.empty and existing_analysis_cols:
            df_analysis_subset = df_analysis[existing_analysis_cols].copy()
            # Ensure 'Value' column is suitable for merge key (string type, no NaNs if possible for key)
            # df_analysis_subset['Value'] = df_analysis_subset['Value'].astype(str).fillna('') # Example if issues
            df_analysis_subset.drop_duplicates(subset=['Value'], keep='first', inplace=True)
        else:
            st.warning("Analysis data is empty or missing key columns, merge will be one-sided.")
            # If df_analysis is empty, create an empty df_analysis_subset with expected columns to prevent merge errors
            df_analysis_subset = pd.DataFrame(columns=existing_analysis_cols)


        # --- Perform the Merge ---
        # Ensure 'Main Key' in processed_df is suitable for merge key
        # processed_df['Main Key'] = processed_df['Main Key'].astype(str).fillna('') # Example
        df_merged = processed_df.merge(
            df_analysis_subset,
            how="left",
            left_on="Main Key",
            right_on="Value",
            suffixes=("_processed", "_analysis")
        )
        st.write(f"DEBUG: Merged DataFrame shape: {df_merged.shape}")

        # --- Clean up Merged Columns ---
        col_to_drop_from_analysis = None
        if 'Value_analysis' in df_merged.columns:
            col_to_drop_from_analysis = 'Value_analysis'
        elif 'Value' in df_merged.columns and 'Value' in existing_analysis_cols and 'Value_processed' not in df_merged.columns: 
             if 'Value' not in processed_df.columns: 
                 st.warning("Combine Warning: Ambiguous 'Value' column after merge. Assuming it's from analysis and dropping.")
                 col_to_drop_from_analysis = 'Value'
             elif processed_df['Value'].name == 'Value': 
                  pass 

        if col_to_drop_from_analysis and col_to_drop_from_analysis in df_merged.columns:
            if col_to_drop_from_analysis == 'Value' and 'Value_processed' in df_merged.columns:
                pass
            elif col_to_drop_from_analysis == 'Value' and 'Value' not in processed_df.columns: 
                st.warning(f"Combine Warning: Attempting to drop merged 'Value' column from analysis, and no original 'Value' found from processed data.")
                df_merged.drop(columns=[col_to_drop_from_analysis], inplace=True) 
                st.write(f"DEBUG: Dropped redundant merge key column '{col_to_drop_from_analysis}' from analysis side.")
            elif col_to_drop_from_analysis:
                df_merged.drop(columns=[col_to_drop_from_analysis], inplace=True)
                st.write(f"DEBUG: Dropped redundant merge key column '{col_to_drop_from_analysis}' from analysis side.")


        cols_to_drop = []
        cols_to_rename = {}
        for col in analysis_cols_to_merge: 
            if col == 'Value': continue
            processed_col_suffix = f"{col}_processed"
            analysis_col_suffix = f"{col}_analysis"
            
            # If both suffixed versions exist, drop _processed, rename _analysis to col
            if processed_col_suffix in df_merged.columns and analysis_col_suffix in df_merged.columns:
                cols_to_drop.append(processed_col_suffix)
                cols_to_rename[analysis_col_suffix] = col
            # If only _analysis version exists, rename it to col
            elif analysis_col_suffix in df_merged.columns and col not in df_merged.columns:
                cols_to_rename[analysis_col_suffix] = col
            # If only _processed version exists, rename it to col
            elif processed_col_suffix in df_merged.columns and col not in df_merged.columns:
                cols_to_rename[processed_col_suffix] = col
            # If 'col' (unsuffixed) already exists and it came from analysis_df (no conflict), do nothing.
            # If 'col' (unsuffixed) already exists and it came from processed_df (no conflict), do nothing.

        if cols_to_drop:
            df_merged.drop(columns=cols_to_drop, inplace=True, errors='ignore')
            st.write(f"DEBUG: Dropped processed-side duplicate columns: {cols_to_drop}")
        if cols_to_rename:
            df_merged.rename(columns=cols_to_rename, inplace=True)
            st.write(f"DEBUG: Renamed analysis-side columns: {cols_to_rename}")


        st.write(f"DEBUG: Columns after merge & cleanup: {df_merged.columns.tolist()}")

        # ***** Run Regex Classification *****
        try:
            if 'Value_Normalized' not in df_merged.columns:
                st.error("Combine Error: 'Value_Normalized' column lost before regex classification step.")
                if 'Main Key' in df_merged.columns:
                    st.warning("Using 'Main Key' as input for regex classification as 'Value_Normalized' is missing.")
                    # Ensure no collision if Value_Normalized_original existed
                    if 'Value_Normalized' in df_merged.columns: # Should not happen if outer if is true
                        df_merged.rename(columns={'Value_Normalized': 'Value_Normalized_original_temp'}, inplace=True, errors='ignore')

                    df_merged['Value_Normalized'] = df_merged['Main Key'] # Use Main Key for regex
                    df_merged = run_regex_classification(df_merged, MAPPING_FILE_LOCAL)
                    # Clean up: remove the temp 'Value_Normalized' and restore original if it existed
                    df_merged.drop(columns=['Value_Normalized'], inplace=True, errors='ignore')
                    if 'Value_Normalized_original_temp' in df_merged.columns:
                         df_merged.rename(columns={'Value_Normalized_original_temp':'Value_Normalized'}, inplace=True)
                else:
                    raise ValueError("'Value_Normalized' and 'Main Key' are missing for regex input.")
            else:
                 # Ensure 'Value_Normalized' is used correctly by regex function
                 df_merged = run_regex_classification(df_merged, MAPPING_FILE_LOCAL)

            st.write("DEBUG: Regex classification step completed.")
            st.write(f"DEBUG: Columns after regex classification: {df_merged.columns.tolist()}")
        except Exception as regex_err:
            st.error(f"Combine Error: Failed during regex classification step: {regex_err}")
            st.error(traceback.format_exc())
            # Ensure these columns exist if regex fails
            for r_col in ["Classification_New", "Reason_New", "DetailedValueType_New"]:
                if r_col not in df_merged.columns: df_merged[r_col] = ""
            df_merged['Classification_New'] = "Error: Regex Step Failed"
            df_merged['Reason_New'] = str(regex_err)
            # DetailedValueType_New can be left empty or marked
        # ***** END: Run Regex Classification *****

        # --- Define and Reorder Final Columns (Incorporating New Error Columns) ---
        core_fixed_pipeline_cols = {"Main Key", "Category", "Attribute", "Code", "Value", "Sheet", "OverallChunkPrefix"} # Added OverallChunkPrefix
        preprocessing_cols = {"Value_E", "Value_Normalized", "ExceptionFlag", "ExceptionTypes", "ExceptionCount", "ExceptionTypeCount", "ConflictFlag", "ConflictCount"}
        
        # Use the original analysis_cols_to_merge as the canonical list of what was intended from analysis
        analysis_cols_added_after_merge = set(analysis_cols_to_merge) - {'Value'} # 'Value' was merge key
        
        regex_cols_added = {"Classification_New", "Reason_New", "DetailedValueType_New", "Distinct_Abs_Pattern_Count"} # Added Distinct_Abs_Pattern_Count

        new_error_columns = [
            "Error_Is_AbsoluteUnit",
            "Error_Is_BlankUnit_NonSNRN",
            "Error_Is_UnmappedUnit_NonSNRN",
            "Error_OriginalClassification_Details",
            "Error_RegexClassification_Details",
            "Error_Mismatch_Classification_Details",
            "Error_Mismatch_DVT_Details"
        ]

        known_added_cols = core_fixed_pipeline_cols.union(preprocessing_cols).union(analysis_cols_added_after_merge).union(regex_cols_added).union(set(new_error_columns))
        
        # Identify original input columns from processed_df that aren't part of the standard pipeline output
        # This requires knowing columns of processed_df before merge
        original_input_cols_from_processed_df = [
            col for col in processed_df.columns 
            if col not in core_fixed_pipeline_cols and col not in preprocessing_cols
        ]
        # Filter this list to ensure they are still in df_merged and not part of other known sets
        original_input_cols = [
            col for col in original_input_cols_from_processed_df 
            if col in df_merged.columns and col not in known_added_cols
        ]


        final_columns_order = [
            "Value_E", "Value_Normalized", "ExceptionFlag", "ExceptionTypes", "ExceptionCount", "ExceptionTypeCount",
            "ConflictFlag", "ConflictCount", "Main Key", "Category", "Attribute", "Code", "Value", "OverallChunkPrefix", # Added OverallChunkPrefix
            "Classification", "DetailedValueType", "Normalized Unit", "Absolute Unit", "Identifiers", "ExtractedPrefix", #Added ExtractedPrefix
            "Classification_New","MixedTypesFlag", "MixedTypesDetail", "MixedTypesDetail_DVT", "Reason_New", "DetailedValueType_New", "Distinct_Abs_Pattern_Count", # Added Distinct_Abs_Pattern_Count
            # New error columns will be inserted here by logic below
            "SubValueCount", "ConditionCount", "MainItemCount", "HasRangeInMain", "HasMultiValueInMain",
            "HasRangeInCondition", "HasMultipleConditions", "MainUnits", "MainDistinctUnitCount",
            "MainUnitsConsistent", "ConditionUnits", "ConditionDistinctUnitCount",
            "ConditionUnitsConsistent", "OverallUnitConsistency", "SingleUnitForAllSubs", "AllDistinctUnitsUsed", "SubValueUnitVariationSummary",
            "MainNumericValues", "ConditionNumericValues", "MainMultipliers", "ConditionMultipliers",
            "MainBaseUnits", "ConditionBaseUnits","MainBaseUnits_New", "ConditionBaseUnits_New", "NormalizedMainValues", "NormalizedConditionValues",
            "MinNormalizedValue", "MaxNormalizedValue", "ParsingErrorFlag", "Sheet",
            *sorted(original_input_cols) # Add original unprocessed columns from input file
        ]
        
        current_cols_in_df = df_merged.columns.tolist()
        final_ordered_cols = [col for col in final_columns_order if col in current_cols_in_df]
        
        missed_cols = [col for col in current_cols_in_df if col not in final_ordered_cols and col not in new_error_columns]
        final_ordered_cols.extend(sorted(missed_cols))
        final_ordered_cols = list(dict.fromkeys(final_ordered_cols)) 

        insertion_point_col = "Distinct_Abs_Pattern_Count" if "Distinct_Abs_Pattern_Count" in final_ordered_cols else "DetailedValueType_New"
        if insertion_point_col in final_ordered_cols:
            idx = final_ordered_cols.index(insertion_point_col)
            final_ordered_cols = final_ordered_cols[:idx+1] + new_error_columns + final_ordered_cols[idx+1:]
        elif "Sheet" in final_ordered_cols: # Fallback insertion point
            idx = final_ordered_cols.index("Sheet")
            final_ordered_cols = final_ordered_cols[:idx+1] + new_error_columns + final_ordered_cols[idx+1:]
        else:
            final_ordered_cols.extend(new_error_columns)
        final_ordered_cols = list(dict.fromkeys(final_ordered_cols)) 

        for err_col in new_error_columns:
            if err_col not in df_merged.columns: 
                 if err_col.startswith("Error_Is_"):
                     df_merged[err_col] = False
                 else:
                     df_merged[err_col] = ""
            else: 
                 if err_col.startswith("Error_Is_"):
                     df_merged[err_col] = df_merged[err_col].fillna(False) 
                 else:
                     df_merged[err_col] = df_merged[err_col].fillna("")   

        if 'New_FeatureCode' not in final_ordered_cols:
            if 'Code' in final_ordered_cols:
                pos = final_ordered_cols.index('Code')
                final_ordered_cols.insert(pos + 1, 'New_FeatureCode')
            else: 
                final_ordered_cols.append('New_FeatureCode')
        
        # Ensure 'Code' column exists before creating 'New_FeatureCode'
        if 'Code' not in df_merged.columns:
            df_merged['Code'] = "" # Add empty 'Code' column if missing
            st.warning("Combine Warning: 'Code' column was missing. Added as empty for New_FeatureCode generation.")
            
        df_merged['New_FeatureCode'] = df_merged['Code'].astype(str).apply(code_to_new_feature)
        # ────────────────────────────────────────────────────────────────
        #  Build MainBaseUnits_New  &  ConditionBaseUnits_New
        # ────────────────────────────────────────────────────────────────
        MAIN_PREFIX_ORDER = [
            "Main unit",
            "Min. Main unit",
            "Max. Main unit",
        ]
        
        COND_PREFIX_ORDER = [
            "Condition unit",
            "Min. Condition unit",
            "Max. Condition unit",
        ]
        
        def _collect_units_ordered(
                group: pd.DataFrame,
                prefixes: list[str]
            ) -> str | None:
            """
            Return a comma-joined string of ALL unit tokens that belong to *prefixes*
            in the order they appear in the DataFrame (which is already sorted by Code).
        
            • Row must have Attribute == "Unit"
            • New_FeatureCode must start with one of *prefixes*
            • Skip rows whose raw value contains '@'  (they’re part of another block)
            • Clean raw text so only the trailing unit token (letters/µ/Ω/°/slash) remains
            • Keeps duplicates; if nothing found → return None
            """
            import re
        
            # regex for trailing unit token
            unit_re = re.compile(r'([A-Za-zµΩ°/]+)$')
        
            # decide which column holds the raw unit text:
            col_priority = ["Value_processed", "Value"]
            col_priority += [c for c in group.columns
                             if "value" in c.lower() and c not in col_priority]
            raw_col = next((c for c in col_priority if c in group.columns), None)
            if raw_col is None:
                return None
        
            units = []
            for _, row in group.iterrows():          # keep DataFrame order
                if row["Attribute"] != "Unit":
                    continue
                if not any(str(row["New_FeatureCode"]).startswith(p) for p in prefixes):
                    continue
        
                raw = str(row[raw_col]).strip()
                if not raw or "@" in raw:
                    continue
        
                m = unit_re.search(raw)
                if m:
                    tok = m.group(1)
                    if tok.lower() != "false":       # weed out placeholders
                        units.append(tok)
        
            return ", ".join(units) if units else None




                
        MAIN_PREFIXES = ["Main unit", "Min. Main unit", "Max. Main unit"]
        COND_PREFIXES = ["Condition unit",
                         "Min. Condition unit", "Max. Condition unit"]
        
        main_unit_map  = {}
        cond_unit_map  = {}
        
        for mk, grp in df_merged.groupby("Main Key", sort=False):
            main_unit_map[mk] = _collect_units_ordered(grp, MAIN_PREFIXES)
            cond_unit_map[mk] = _collect_units_ordered(grp, COND_PREFIXES)
        
        df_merged["MainBaseUnits_New"]       = df_merged["Main Key"].map(main_unit_map)
        df_merged["ConditionBaseUnits_New"]  = df_merged["Main Key"].map(cond_unit_map)
        df_merged["MainBaseUnits_New"].replace("", None, inplace=True)
        df_merged["ConditionBaseUnits_New"].replace("", None, inplace=True)

        # ───────────────────────────────────────────────────────────────────
        
        # make sure the columns survive the final re-index
        for _c in ("MainBaseUnits_New", "ConditionBaseUnits_New"):
            if _c not in final_ordered_cols:
                final_ordered_cols.append(_c)
        final_ordered_cols = list(dict.fromkeys(final_ordered_cols)) 

        st.write(f"DEBUG: Final column order being applied: {final_ordered_cols}")
        
        # Ensure all columns in final_ordered_cols exist in df_merged before reindexing
        for col_final_order in final_ordered_cols:
            if col_final_order not in df_merged.columns:
                df_merged[col_final_order] = pd.NA # Or appropriate fill value like "" or 0
                st.warning(f"Combine Warning: Column '{col_final_order}' from final_columns_order was not in df_merged. Added with NA.")
        
        df_final = df_merged.reindex(columns=final_ordered_cols, fill_value='')
        df_final = _post_process_string_keywords(df_final)
        chunk_re = re.compile(r"^M(\d+)-")         # grabs the leading M<idx>-

        df_final['ChunkIdx'] = (                      # NaN for ERR/UNK rows
            df_final['Code'].str.extract(chunk_re)[0].astype('Int64', errors='ignore')
        )
        
        string_chunks_per_key = (                     #  e.g.  {'1A,2B': [1, 2], …}
            df_final[df_final['Category'].eq('String')]
                .dropna(subset=['ChunkIdx'])
                .groupby('Main Key')['ChunkIdx']
                .apply(lambda s: sorted(set(int(x) for x in s)))
        )
        df_final = add_mixed_types_detail_dvt(df_final)
       # df_final = update_number_with_identifier(df_final)

        # ─── NEW: keep the column order list in sync with any renames ───
        clean_cols = df_final.columns.tolist()
        
        # 1) replace the old name with the new one
        final_ordered_cols = [
            'Value' if c == 'Value_processed' else c
            for c in final_ordered_cols
        ]
        
        # 2) drop names that vanished and append any fresh ones,
        #    preserving the original relative order as far as possible
        final_ordered_cols = (
            [c for c in final_ordered_cols if c in clean_cols] +
            [c for c in clean_cols         if c not in final_ordered_cols]
        )
        # ────────────────────────────────────────────────────────────────


        # +++ START: MODIFICATION FOR CLASSIFICATION UPDATES (REVISED) +++
# In result_combiner.py -> inside the combine_results function
        # This block comes after _post_process_string_keywords has been called and df_final is cleaned.

        # +++ START: UNIFIED MODIFICATION FOR PREFIXES AND STRINGS (SCOPE FIX) +++
        # ──────────────────────────────────────────────────────────────────────────────
        #  Unified classification re-writer
        #      – adds “ with prefix” if OverallChunkPrefix is present
        #      – adds “ with String” if the Main-Key contains a String-keyword row
        # ──────────────────────────────────────────────────────────────────────────────
        st.write("DEBUG: Updating classification columns for prefixes and String keywords...")
        
        classification_cols_to_update = [
            "Classification",
            "DetailedValueType",
            "Classification_New",
            "DetailedValueType_New",
        ]
        
        # ----------------------------------------------------------------------
        # 1) Generalised helper – identical to your old prefix routine
        # ----------------------------------------------------------------------
        def _add_suffix(val_str: str, col_name: str, suffix: str) -> str:
            """
            Insert <suffix> into *val_str* using the same rules you used for prefixes.
            Works for both plain Classification and DetailedValueType strings.
            """
            val_str = str(val_str)
        
            if not val_str or suffix in val_str:
                return val_str  # skip NaN/empty/already processed
        
            is_detailed = "DetailedValueType" in col_name
        
            # --- 'Multiple ( … )' pattern ------------------------------------
            if val_str.startswith("Multiple ("):
                close_idx = val_str.find(")")
                if close_idx != -1:
                    return f"{val_str[:close_idx]}{suffix}{val_str[close_idx:]}"
                return val_str
        
            # --- DetailedValueType with trailing annotation -------------------
            if is_detailed:
                anno_idx = val_str.find(" [")        # beginning of "[1][0] x1"
                if anno_idx != -1:
                    return f"{val_str[:anno_idx]}{suffix}{val_str[anno_idx:]}"
                return val_str + suffix              # no annotation – just append
        
            # --- Plain Classification ----------------------------------------
            return val_str + suffix
        
        
        # ----------------------------------------------------------------------
        # 2)  Apply “with prefix” (unchanged logic)
        # ----------------------------------------------------------------------
        if "OverallChunkPrefix" not in df_final.columns:
            st.warning("Combine Warning: 'OverallChunkPrefix' column not found. Skipping prefix updates.")
        else:
            prefix_mask = (
                df_final["OverallChunkPrefix"].astype(str).str.strip().ne("")
                & df_final["OverallChunkPrefix"].notna()
            )
        
            if prefix_mask.any():
                for col in classification_cols_to_update:
                    if col in df_final.columns:
                        df_final.loc[prefix_mask, col] = (
                            df_final.loc[prefix_mask, col]
                            .apply(lambda x, c=col: _add_suffix(x, c, " with prefix"))
                        )
                st.write(f"DEBUG: Added 'with prefix' to {prefix_mask.sum()} rows.")
        
        
        # ----------------------------------------------------------------------
        # 3)  Apply “with String” (new, but same insertion rules)
        # ----------------------------------------------------------------------
        # 3-a)  Find every Main-Key that contains at least one row classified as Category == 'String'
        string_row_mask = df_final["Category"].eq("String")
# 3) Inject precise “/String 1,2” suffixes
        def _inject_string_suffix(label: str, col_name: str, chunk_list: list[int]) -> str:
            if not chunk_list or "/String" in str(label):
                return label
        
            suffix = "/String " + ", ".join(map(str, chunk_list))
            is_dvt = "DetailedValueType" in col_name
            label = str(label)
        
            if label.startswith("Multiple ("):
                close = label.find(")")
                if close != -1:
                    return f"{label[:close]}{suffix}{label[close:]}"
                return label + suffix
        
            if is_dvt:
                anno = label.find(" [")
                return f"{label[:anno]}{suffix}{label[anno:]}" if anno != -1 else label + suffix
        
            return label + suffix
        
        for mk, chunk_nums in string_chunks_per_key.items():
            row_mask = df_final['Main Key'].eq(mk)
            for col in classification_cols_to_update:
                if col in df_final.columns:
                    df_final.loc[row_mask, col] = (
                        df_final.loc[row_mask, col]
                            .apply(_inject_string_suffix,
                                   args=(col, chunk_nums))
                    )
        st.write(f"DEBUG: Injected '/String …' into {len(string_chunks_per_key)} Main Keys.")

        # +++ END: UNIFIED MODIFICATION +++
        # +++ END: UNIFIED MODIFICATION +++

        # --- Regenerate thousands separators when the ONLY exception is ThousandsSeparator ---
        value_col_for_regen = "Value_processed" if "Value_processed" in df_final.columns else "Value"
        
        required_cols_for_regen = ["Attribute", "ExceptionFlag", "ExceptionTypes", value_col_for_regen]
        
        if all(col in df_final.columns for col in required_cols_for_regen):
        
            attr_mask = df_final["Attribute"].astype(str) == "Value"
            flag_mask = df_final["ExceptionFlag"].astype(str).str.upper() == "YES"
            # types_mask = df_final["ExceptionTypes"].astype(str).apply( #Make change hereeeeeee
            #     lambda x: [et.strip() for et in x.split(",") if et.strip()] == ["ThousandsSeparator"]
            # )
            def _is_thousands_separator_only(value) -> bool:
                # Be defensive with mixed/dirty types in ExceptionTypes (e.g. float/NaN).
                if pd.isna(value):
                    return False
                parsed = [et.strip() for et in str(value).split(",") if et.strip()]
                return parsed == ["ThousandsSeparator"]
 
            types_mask = df_final["ExceptionTypes"].apply(_is_thousands_separator_only)
        
            target_rows = df_final[attr_mask & flag_mask & types_mask].index
        
            if not target_rows.empty:
                st.write(
                    f"DEBUG: Regenerating thousands separators in "
                    f"'{value_col_for_regen}' for {len(target_rows)} rows."
                )
                df_final.loc[target_rows, value_col_for_regen] = (
                    df_final.loc[target_rows, value_col_for_regen]
                    .apply(_regenerate_thousands_separators_rc)
                )
            else:
                st.write("DEBUG: No rows matched criteria for thousands-separator regeneration.")
        else:
            missing = [c for c in required_cols_for_regen if c not in df_final.columns]
            st.warning(
                "Combine Warning: Required columns "
                f"{missing} not found in df_final; skipping thousands-separator regeneration."
            )
        
        # Error checker section starts here
        value_col_for_error_check = 'Value'
        if 'Value_processed' in df_final.columns: value_col_for_error_check = 'Value_processed'
        elif 'Value' not in df_final.columns:
            st.warning("Combine Warning: No parsed-value column ('Value' or 'Value_processed') found; using 'Main Key' for some error checks.")
            value_col_for_error_check = 'Main Key'
        if value_col_for_error_check not in df_final.columns:
            st.error(f"Critical error: Column '{value_col_for_error_check}' for error checking not found. Adding dummy.")
            df_final[value_col_for_error_check] = ""

        base_units_for_check, _ = read_mapping_file(MAPPING_FILE_LOCAL)
        _SKIP_UNIT_ERROR = r"(SN-U|RN-U)"

        # --- Define ORIGINAL boolean error conditions (used for error_mask AND to populate new columns) ---
        abs_err = pd.Series(False, index=df_final.index)
        if "Absolute Unit" in df_final.columns:
            abs_err = df_final["Absolute Unit"].astype(str).str.contains("error", case=False, na=False)
            df_final.loc[abs_err, 'Error_Is_AbsoluteUnit'] = True
        else:
            st.warning("Combine Warning: 'Absolute Unit' column not found for 'abs_err' check.")

        blank_unit_condition = (
            (df_final["Attribute"].astype(str) == "Unit") &
            df_final[value_col_for_error_check].astype(str).str.strip().eq("") &
            ~df_final["Code"].astype(str).str.contains(_SKIP_UNIT_ERROR, na=False)
        )
        df_final.loc[blank_unit_condition, "Error_Is_BlankUnit_NonSNRN"] = True
        unit_error_keys_blank = df_final.loc[blank_unit_condition, "Main Key"].unique()
        main_keys_with_blank_unit_error = df_final["Main Key"].isin(unit_error_keys_blank)

        unmapped_unit_condition = (
            (df_final["Attribute"].astype(str) == "Unit") &
            ~df_final[value_col_for_error_check].astype(str).isin(base_units_for_check) & 
            ~df_final["Code"].astype(str).str.contains(_SKIP_UNIT_ERROR, na=False) &
            df_final[value_col_for_error_check].astype(str).str.strip().ne("")
        )
        df_final.loc[unmapped_unit_condition, "Error_Is_UnmappedUnit_NonSNRN"] = True
        unit_error_keys_unmapped = df_final.loc[unmapped_unit_condition, "Main Key"].unique()
        main_keys_with_unmapped_unit_error = df_final["Main Key"].isin(unit_error_keys_unmapped)

        classification_error = pd.Series(False, index=df_final.index)
        # Ensure 'Classification' column exists before checks
        if "Classification" in df_final.columns:
            classification_error_condition = df_final["Classification"].astype(str).str.contains(r"Mixed|Pairs|Error|Unknown", na=False, regex=True) # Added Unknown
            classification_error = classification_error_condition
            df_final.loc[classification_error_condition, 'Error_OriginalClassification_Details'] = df_final.loc[classification_error_condition, 'Classification']
        else:
            st.warning("Combine Warning: 'Classification' column not found for 'classification_error' check.")
            df_final['Error_OriginalClassification_Details'] = "Column Missing" # Initialize if not present

        regex_classification_error = pd.Series(False, index=df_final.index)
        # Ensure 'Classification_New' and 'Reason_New' exist
        if "Classification_New" in df_final.columns:
            regex_classification_error_condition = df_final["Classification_New"].astype(str).str.contains(r"Unknown|Mixed|Error", na=False, regex=True)
            regex_classification_error = regex_classification_error_condition
            if "Reason_New" in df_final.columns:
                reason_new_values = df_final.loc[regex_classification_error_condition, 'Reason_New'].fillna('')
                df_final.loc[regex_classification_error_condition, 'Error_RegexClassification_Details'] = reason_new_values
            else: # Fallback if Reason_New is missing
                 df_final['Error_RegexClassification_Details'] = "" # Initialize
                 df_final.loc[regex_classification_error_condition, 'Error_RegexClassification_Details'] = df_final.loc[regex_classification_error_condition, 'Classification_New'] 
        else:
            st.warning("Combine Warning: 'Classification_New' column not found for 'regex_classification_error' check.")
            df_final['Error_RegexClassification_Details'] = "Column Missing" # Initialize

        mismatch_error = pd.Series(False, index=df_final.index)
        # Ensure all columns for mismatch check exist
        mismatch_check_cols = ['Classification_New', 'Classification', 'DetailedValueType_New', 'DetailedValueType']
        if all(m_col in df_final.columns for m_col in mismatch_check_cols):
            class_new_str = df_final['Classification_New'].astype(str).fillna('')
            class_orig_str = df_final['Classification'].astype(str).fillna('')
            dvt_new_str = df_final['DetailedValueType_New'].astype(str).fillna('')
            dvt_orig_str = df_final['DetailedValueType'].astype(str).fillna('')

            # Identify mismatches, excluding cases where original is problematic and new might be an improvement or different error
            # This logic can be refined: a mismatch isn't an error if original was "Error" and new is something valid.
            # For now, any difference is flagged.
            class_mismatch_rows = (class_new_str != class_orig_str)
            dvt_mismatch_rows = (dvt_new_str != dvt_orig_str)
            mismatch_condition = class_mismatch_rows | dvt_mismatch_rows
            mismatch_error = mismatch_condition

            details_class_mismatch = "Orig: " + class_orig_str + ", New: " + class_new_str
            df_final.loc[class_mismatch_rows, 'Error_Mismatch_Classification_Details'] = details_class_mismatch[class_mismatch_rows]
            
            details_dvt_mismatch = "Orig: " + dvt_orig_str + ", New: " + dvt_new_str
            df_final.loc[dvt_mismatch_rows, 'Error_Mismatch_DVT_Details'] = details_dvt_mismatch[dvt_mismatch_rows]
        else:
            st.warning(f"Combine Warning: One or more columns for mismatch check missing: {mismatch_check_cols}")
            # Initialize if not present
            if 'Error_Mismatch_Classification_Details' not in df_final.columns: df_final['Error_Mismatch_Classification_Details'] = ""
            if 'Error_Mismatch_DVT_Details' not in df_final.columns: df_final['Error_Mismatch_DVT_Details'] = ""
        
        # --- Final error mask (Uses ORIGINAL aggregated conditions to maintain functionality) ---
        error_mask = (
            abs_err |
            main_keys_with_blank_unit_error | 
            main_keys_with_unmapped_unit_error | 
            classification_error |
            regex_classification_error |
            mismatch_error
        )
        
        error_df = pd.DataFrame(columns=df_final.columns) 
        if error_mask.any():
            error_df = df_final.loc[error_mask].drop_duplicates().copy() 
            df_final = df_final.loc[~error_mask].reset_index(drop=True)
            st.write(f"DEBUG: Identified {len(error_df)} error rows. Moved to 'error' sheet.")
        else:
            st.write("DEBUG: No specific error rows identified for the 'error' sheet.")
        # ── FINAL: update numbers with identifier (run after combine is fully done)
        try:
            from analysis_helpers import update_number_with_identifier
            # Ensure the helper's expected working column exists without changing semantics
            if "Value_processed" not in df_final.columns and "Value" in df_final.columns:
                df_final["Value_processed"] = df_final["Value"]
        
            df_final = update_number_with_identifier(df_final)
        
            # If your downstream expects 'Value' to reflect the updated content, un-comment:
            # if "Value" in df_final.columns:
            #     df_final["Value"] = df_final["Value_processed"]
        
            st.write("DEBUG: Post-combine update_number_with_identifier completed.")
        except Exception as e:
            st.error(f"Post-combine update_number_with_identifier failed: {e}")
            st.error(traceback.format_exc())

        # --- Prepare DataFrames for Output Sheets ---
        value_processed_col_name = 'Value'
        if 'Value_processed' in df_final.columns: value_processed_col_name = 'Value_processed'
        elif 'Value' not in df_final.columns:
            st.error("Combine Error: Cannot find suitable column for 'Value_processed' for Value Table if 'Value' also missing.")
            df_final['Value_processed_dummy'] = "" 
            value_processed_col_name = 'Value_processed_dummy'
        
        if value_processed_col_name not in df_final.columns:
            st.error(f"Critical error preparing for Value Table: Column '{value_processed_col_name}' not found. Adding dummy.")
            df_final[value_processed_col_name] = ""


        value_details_cols = [
            "Value_E", "DetailedValueType", "Normalized Unit",
            "Attribute", "Code", value_processed_col_name
        ]
        
        if "ValueID" in df_final.columns: # ValueID might come from original input
            value_details_cols.append("ValueID")
        if "New_FeatureCode" in df_final.columns:
            value_details_cols.append("New_FeatureCode")

        missing_vd_cols = [c for c in value_details_cols if c not in df_final.columns]
        if missing_vd_cols:
            st.warning(f"Combine Warning: Missing columns for 'value_details' sheet: {missing_vd_cols}. Sheet might be incomplete.")
            for mc in missing_vd_cols:
                df_final[mc] = "" # Add missing columns with empty string

        value_details_df = df_final[value_details_cols].copy()

        
        new_vd_colnames = [
            "AcceptedValue", "DetailedValueType", "Pattern",
            "AttributeName", "FeatureCode", "AttributeValue"
        ]
        if "ValueID" in df_final.columns: # if ValueID was included above
            new_vd_colnames.append("ValueID")
        if "New_FeatureCode" in df_final.columns: # if New_FeatureCode was included
            new_vd_colnames.append("New_FeatureCode")

        value_details_df.columns = new_vd_colnames

        mask_value_table = (
            (df_final['Attribute'].astype(str) == 'Value') & 
            df_final[value_processed_col_name].notna() &
            df_final[value_processed_col_name].astype(str).str.strip().ne('')
        )
        value_df_source = df_final.loc[mask_value_table, [value_processed_col_name]] 
        
        value_df = pd.DataFrame() 
        if not value_df_source.empty:
            value_df = value_df_source.drop_duplicates(subset=[value_processed_col_name], keep='first').reset_index(drop=True).copy()
            value_df['Value__'] = value_df[value_processed_col_name].apply(compute_value__)
            value_df['Multiplication'] = value_df[value_processed_col_name].apply(compute_multiplication)
            value_df['Space separator'] = value_df[value_processed_col_name].apply(compute_space_separator)
            value_df['Value pattern'] = value_df[value_processed_col_name].apply(compute_value_pattern)
            value_df['Normalized value__'] = value_df[value_processed_col_name].apply(compute_normalized_value__)
            value_df['Integar value'] = value_df[value_processed_col_name].apply(compute_integar_value)
            value_df['Fraction value'] = value_df[value_processed_col_name].apply(compute_fraction_value)
            value_df['Fraction digits'] = value_df[value_processed_col_name].apply(compute_fraction_digits)
            value_df.rename(columns={value_processed_col_name: 'Value_processed'}, inplace=True)
        else: 
            expected_value_table_cols = [
                'Value_processed', 'Value__', 'Multiplication', 'Space separator', 
                'Value pattern', 'Normalized value__', 'Integar value', 
                'Fraction value', 'Fraction digits'
            ]
            value_df = pd.DataFrame(columns=expected_value_table_cols)

        pattern_cols = ["Normalized Unit", "Absolute Unit", "Classification", "DetailedValueType","MainBaseUnits", "ConditionBaseUnits"]
        missing_pt_cols = [c for c in pattern_cols if c not in df_final.columns]
        if missing_pt_cols:
            st.warning(f"Combine Warning: Missing columns for 'PatternTable' sheet: {missing_pt_cols}. Sheet might be incomplete.")
            for mc in missing_pt_cols: df_final[mc] = "" 
        pattern_df = df_final[pattern_cols].copy()
        pattern_df.columns = [
            "Pattern", "AbsolutePattern", "ValueType", 
            "DetailedValueType", "MainBaseUnits", "ConditionBaseUnits"
        ]
        pattern_df = pattern_df.drop_duplicates(
            subset=["Pattern","AbsolutePattern","ValueType","DetailedValueType","MainBaseUnits", "ConditionBaseUnits"], keep="first"
        ).reset_index(drop=True)

        feature_cols = ["DetailedValueType", "Attribute", "Code"]
        if "New_FeatureCode" in df_final.columns:
            feature_cols.append("New_FeatureCode")

        missing_ft_cols = [c for c in feature_cols if c not in df_final.columns]
        if missing_ft_cols:
            st.warning(f"Combine Warning: Missing columns for 'FeatureTable' sheet: {missing_ft_cols}. Sheet might be incomplete.")
            for mc in missing_ft_cols:
                df_final[mc] = ""

        feature_df = df_final[feature_cols].copy()

        new_ft_colnames = ["DetailedValueType", "AttributeName", "FeatureCode"]
        if "New_FeatureCode" in df_final.columns:
            new_ft_colnames.append("New_FeatureCode")

        feature_df.columns = new_ft_colnames

        feature_df = feature_df.drop_duplicates(
            subset=new_ft_colnames, keep="first"
        ).reset_index(drop=True)


        # --- Write all sheets to Excel ---
        st.write(f"DEBUG: Writing final output to {output_file}...")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Combined')
            if not error_df.empty:
                error_df_ordered = error_df.reindex(columns=final_ordered_cols, fill_value='') 
                error_df_ordered.to_excel(writer, index=False, sheet_name='error')
            value_df.to_excel(writer, index=False, sheet_name='Value_table')
            value_details_df.to_excel(writer, index=False, sheet_name='value_details')
            pattern_df.to_excel(writer, index=False, sheet_name='PatternTable')
            feature_df.to_excel(writer, index=False, sheet_name='FeatureTable')

        st.write(f"DEBUG: Sheets written — Combined ({len(df_final)} rows), error ({len(error_df)} rows), "
                 f"Value_table ({len(value_df)} rows), value_details ({len(value_details_df)} rows), "
                 f"PatternTable ({len(pattern_df)} rows), FeatureTable ({len(feature_df)} rows)")
        return output_file

    except Exception as e:
        st.error(f"An unexpected error occurred during combine_results: {e}")
        st.error(traceback.format_exc())
        return None

# --- END OF FILE result_combiner.py ---
