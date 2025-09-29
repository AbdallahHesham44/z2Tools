import streamlit as st
import pandas as pd
import numpy as np
import re
from typing import Dict, List, Tuple, Optional
import tempfile
import os

# ============================================================================
# (KEEPING YOUR FUNCTIONS EXACTLY AS IS)
# ============================================================================
# ============================================================================
# STEP 1: Process and validate cross-reference files
# ============================================================================

def process_files(crossfile, parametricData, pakageAndPinout, recipe, output_path=None):
    """Process cross-reference files with validation logic"""
    # Load files efficiently with proper dtypes
    cross_df = pd.read_excel(crossfile, dtype=str)
    recipe_df = pd.read_excel(recipe, dtype=str)
    pakage_df = pd.read_excel(pakageAndPinout, dtype=str)
    
    # Read and concatenate parametric sheets
    parametric_sheets = pd.read_excel(parametricData, sheet_name=None, dtype=str)
    parametric_df = pd.concat(parametric_sheets.values(), ignore_index=True)
    
    # Strip whitespace from key columns once
    for col in ["PLc", "PLx"]:
        if col in cross_df.columns:
            cross_df[col] = cross_df[col].astype(str).str.strip()
    
    recipe_df["ZProductValue"] = recipe_df["ZProductValue"].astype(str).str.strip()
    
    # Step 1: Compare PLc and PLx (vectorized)
    cross_df["SamePLs"] = np.where(
        cross_df["PLc"] == cross_df["PLx"], "same", "not"
    )
    
    # Step 2: Add similarity columns
    cross_df["IsPackaeSimilar"] = "NA"
    cross_df["IsPinoutSimilar"] = "NA"
    cross_df["PackaePinout"] = "NA"
    
    # Filter rows where SamePLs == "same" for efficiency
    same_pl_mask = cross_df["SamePLs"] == "same"
    same_pl_indices = cross_df[same_pl_mask].index
    
    # Group recipe by ZProductValue for faster lookups
    recipe_grouped = recipe_df.groupby("ZProductValue")
    
    for idx in same_pl_indices:
        plc = cross_df.at[idx, "PLc"]
        
        if plc in recipe_grouped.groups:
            subset = recipe_grouped.get_group(plc)
            
            # Check similarity flags
            is_pack = "FALSE" if (subset["IsPackaeSimilar"].str.upper() == "FALSE").all() else "TRUE"
            is_pin = "FALSE" if (subset["IsPinoutSimilar"].str.upper() == "FALSE").all() else "TRUE"
            
            cross_df.at[idx, "IsPackaeSimilar"] = is_pack
            cross_df.at[idx, "IsPinoutSimilar"] = is_pin
            cross_df.at[idx, "PackaePinout"] = "TRUE" if (is_pack == "TRUE" or is_pin == "TRUE") else "FALSE"
    
    # Step 3: Check FoundData (optimized with sets)
    parametric_parts = set(parametric_df["PartNumber"].values)
    pakage_parts = set(pakage_df["PartNumber"].values)
    
    cross_df["FoundData"] = "FALSE"
    
    for idx in same_pl_indices:
        part_c = str(cross_df.at[idx, "PartNumberC"]).strip()
        part_x = str(cross_df.at[idx, "PartNumberX"]).strip()
        pack_pinout = cross_df.at[idx, "PackaePinout"]
        
        if pack_pinout == "TRUE":
            exists_c = part_c in parametric_parts or part_c in pakage_parts
            exists_x = part_x in parametric_parts or part_x in pakage_parts
        elif pack_pinout == "FALSE":
            exists_c = part_c in parametric_parts
            exists_x = part_x in parametric_parts
        else:
            continue
        
        if exists_c and exists_x:
            cross_df.at[idx, "FoundData"] = "TRUE"
    
    if output_path:
        cross_df.to_excel(output_path, index=False)
    
    return cross_df


# ============================================================================
# STEP 2: Merge data from multiple files
# ============================================================================

def deduplicate_columns(columns):
    """Make duplicate column names unique by adding suffixes"""
    seen = {}
    new_cols = []
    for col in columns:
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
    return new_cols


def load_all_excel_sheets(file_list):
    """Load all sheets from all Excel files once - returns cache"""
    cache = {}
    for file in file_list:
        try:
            xl = pd.ExcelFile(file)
            cache[file] = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}
        except Exception as e:
            print(f"❌ Error reading {file}: {e}")
            cache[file] = None
    return cache


def find_matching_row(df, part_value, company_value):
    """Search for matching row in dataframe"""
    df.columns = df.columns.str.strip()
    cols_lower = df.columns.str.lower()
    
    # Find company and part columns
    company_col = None
    part_col = None
    
    for idx, col in enumerate(cols_lower):
        if col in ["company", "companyname", "company name", "companynamec"]:
            company_col = df.columns[idx]
        if col in ["partnumber", "part number", "part_number", "partnumberc"]:
            part_col = df.columns[idx]
    
    if not company_col or not part_col:
        return None
    
    # Filter with vectorized operations
    mask = (df[part_col].astype(str).str.strip() == part_value) & \
           (df[company_col].astype(str).str.strip() == company_value)
    
    filtered = df[mask]
    
    if not filtered.empty:
        row = filtered.iloc[0].drop([part_col, company_col], errors='ignore')
        return row.to_frame().T.reset_index(drop=True)
    
    return None


def get_single_row_all_files(file_cache, part_value, company_value, PackaePinout_val):
    """Merge one row from multiple Excel files horizontally"""
    merged_parts = [pd.DataFrame({
        "PartNumber": [part_value],
        "Company": [company_value],
        "Comments": [""]
    })]
    
    PackaePinout_bool = str(PackaePinout_val).strip().lower() in ["true", "1", "yes"]
    
    file_list = list(file_cache.keys())
    if not PackaePinout_bool:
        file_list = file_list[:-1]
    
    for file in file_list:
        file_sheets = file_cache.get(file)
        
        if file_sheets is None:
            if PackaePinout_bool:
                merged_parts[0].loc[0, "Comments"] += f"Error reading {file}; "
            continue
        
        found = False
        for sheet_name, df in file_sheets.items():
            row = find_matching_row(df, part_value, company_value)
            if row is not None:
                row.columns = deduplicate_columns(row.columns)
                merged_parts.append(row)
                found = True
                break
        
        if not found and PackaePinout_bool:
            merged_parts[0].loc[0, "Comments"] += f"Not found in {file}; "
    
    final_row = pd.concat(merged_parts, axis=1)
    final_row.columns = deduplicate_columns(final_row.columns)
    
    return final_row


def merge_from_crossesparts(crosses_file, target_files, output_path=None):
    """Reads cross-reference file and merges corresponding rows"""
    df_crosses = pd.read_excel(crosses_file)
    
    required_cols = ["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex", "PackaePinout"]
    for col in required_cols:
        if col not in df_crosses.columns:
            raise ValueError(f"❌ Missing required column: {col}")
    
    # Load all files once
    file_cache = load_all_excel_sheets(target_files)
    all_merged_rows = []
    
    for idx, row in df_crosses.iterrows():
        part_c = str(row["PartNumberC"]).strip()
        company_c = str(row["CompanyNameC"]).strip()
        part_x = str(row["PartNumberX"]).strip()
        company_x = str(row["CompanyNamex"]).strip()
        pack_pinout = str(row["PackaePinout"]).strip()
        
        print(f"➡️ Merging {part_c} from {company_c}")
        merged_c = get_single_row_all_files(file_cache, part_c, company_c, pack_pinout)
        
        print(f"➡️ Merging {part_x} from {company_x}")
        merged_x = get_single_row_all_files(file_cache, part_x, company_x, pack_pinout)
        
        all_merged_rows.extend([merged_c, merged_x])
    
    final_df = pd.concat(all_merged_rows, ignore_index=True)
    
    if output_path:
        final_df.to_excel(output_path, index=False)
    
    return final_df


# ============================================================================
# STEP 3: Validation and comparison
# ============================================================================

def percent_diff_base_a(a, b):
    """Calculate percentage difference based on value a"""
    if a == 0:
        return float('inf') if b != 0 else 0.0
    return round(abs(a - b) / abs(a) * 100, 2)


def fc_sort_key(fc):
    """Sort key for FeatureCodes (V before U, numeric order)"""
    match = re.search(r"([VU])(\d+)$", fc)
    if match:
        t, num = match.groups()
        return int(num) * 2 + (0 if t == "V" else 1)
    return float("inf")


def compare_parts(df, value, value2):
    """Compare two parts by AcceptedValue with sorted FeatureCodes"""
    # Filter rows
    part1 = df[df["AcceptedValue"] == value].reset_index(drop=True)
    part2 = df[df["AcceptedValue"] == value2].reset_index(drop=True)
    
    if part1.empty or part2.empty:
        return {
            "value": value, "nValue": [], "nUnit": [],
            "FeatureCode": "NotFound",
            "state": "NotFound In File",
            "value2": value2, "nValue2": [], "nUnit2": []
        }
    
    # Filter by first ValueID
    chosen_value_id = part1.loc[0, "ValueID"]
    chosen_value_id2 = part2.loc[0, "ValueID"]
    
    part1 = part1[part1["ValueID"] == chosen_value_id].reset_index(drop=True)
    part2 = part2[part2["ValueID"] == chosen_value_id2].reset_index(drop=True)
    
    # Compare FeatureCode sets
    fc1 = set(part1["FeatureCode"].dropna().unique())
    fc2 = set(part2["FeatureCode"].dropna().unique())
    
    if fc1 != fc2:
        return {
            "value": value, "nValue": [], "nUnit": [],
            "FeatureCode": "Different",
            "state": "Different FeatureCode",
            "value2": value2, "nValue2": [], "nUnit2": []
        }
    
    # Sort by FeatureCode
    part1_sorted = part1.sort_values(by="FeatureCode", key=lambda x: x.map(fc_sort_key))
    part2_sorted = part2.sort_values(by="FeatureCode", key=lambda x: x.map(fc_sort_key))
    
    return {
        "value": value,
        "nValue": part1_sorted["NormalizedValue"].dropna().tolist(),
        "nUnit": part1_sorted["Unit"].dropna().tolist(),
        "FeatureCode": "Match",
        "state": "Match",
        "value2": value2,
        "nValue2": part2_sorted["NormalizedValue"].dropna().tolist(),
        "nUnit2": part2_sorted["Unit"].dropna().tolist()
    }


def validate_core_upgrade_equal(newcross_file, finalmerged_file, lookup_file1, 
                                lookup_file2, loadLookUp, output_path=None):
    """Validate core features, upgrades, and equal features"""
    # Load files
    rules_df = pd.read_excel(newcross_file, dtype=str)
    data = pd.read_excel(finalmerged_file, dtype=str)
    
    # Load lookup values if needed
    global values
    if loadLookUp == "YES":
        print("Loading lookup tables...")
        values1 = pd.read_excel(lookup_file1, dtype=str)
        values2 = pd.read_excel(lookup_file2, dtype=str)
        values = pd.concat([values1, values2], ignore_index=True)
    
    if "ProductName" not in data.columns:
        raise ValueError("❌ 'ProductName' column not found")
    
    # Filter rules by PLs present in data
    PLs = data["ProductName"].astype(str).unique().tolist()
    rules_df = rules_df[rules_df["ZProductValue"].astype(str).isin(PLs)]
    
    # Keep only relevant rules
    rules_df = rules_df[
        (rules_df["FeaturesType"].str.lower() == "core") |
        (rules_df["UpgradeFeature"].str.upper() == "TRUE") |
        (rules_df["IsEqualFeature"].str.upper() == "TRUE") |
        (rules_df["FeaturesType"].str.lower() == "tolerance")
    ]
    
    # Convert tolerance columns to numeric
    for col in ["G1", "G2", "G3"]:
        if col in rules_df.columns:
            rules_df[col] = pd.to_numeric(rules_df[col], errors="coerce")
    
    results = []
    grouped_data = data.groupby("ProductName")
    
    for PL, data_subset in grouped_data:
        print(f"Processing PL: {PL}")
        data_subset = data_subset.reset_index(drop=True)
        rules_subset = rules_df[rules_df["ZProductValue"].astype(str).str.strip() == PL]
        
        # Check pinout flag
        PinOut_flag = "TRUE" if rules_df["IsPinoutSimilar"].iloc[1] == "True" else "FALSE"
        
        # Process pairs
        for i in range(0, len(data_subset), 2):
            if i + 1 >= len(data_subset):
                break
            
            row1, row2 = data_subset.iloc[i], data_subset.iloc[i+1]
            feature_statuses, feature_grades, flags = [], [], []
            count = 1
            
            for _, rule in rules_subset.iterrows():
                feature = rule["Features"]
                ftype = str(rule["FeaturesType"]).lower()
                upgrade_flag = str(rule.get("UpgradeFeature", "")).upper()
                equal_flag = str(rule.get("IsEqualFeature", "")).upper()
                
                # Pinout check
                if PinOut_flag == "TRUE" and count:
                    val1 = str(row1['NormalizedPinName']).strip()
                    val2 = str(row2['NormalizedPinName']).strip()
                    count = 0
                    
                    if val1 == val2:
                        status, grade = f"✅ PinOut Match", "A"
                    else:
                        status, grade = f"❌ Different PinOut", "DiffPinOut"
                    
                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)
                    status, grade = "", ""
                
                # Core/Equal check
                if ftype == "core" or equal_flag == "TRUE":
                    if feature not in data_subset.columns:
                        status, grade = f"⚠ Missing column {feature}", "Fail"
                    else:
                        val1 = str(row1[feature]).strip()
                        val2 = str(row2[feature]).strip()
                        
                        if val1 == val2:
                            status, grade = f"✅ Match", "A"
                        else:
                            status, grade = f"❌ Mismatch ({val1} vs {val2})", "Fail"
                            if ftype == "core":
                                flags.append(f"missMatchAtCore {feature}")
                                grade = "FailInCore"
                    
                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)
                
                elif upgrade_flag == "TRUE":
                    continue
                
                # Tolerance check
                elif ftype == "tolerance":
                    if feature not in data_subset.columns:
                        status, grade = f"⚠ Missing column {feature}", "Fail"
                    else:
                        val1 = str(row1[feature]).strip()
                        val2 = str(row2[feature]).strip()
                        
                        if val1 == val2:
                            status, grade = f"✅ Match", "A"
                        else:
                            result = compare_parts(values, val1, val2)
                            
                            if result['nUnit'] != result['nUnit2']:
                                status, grade = f"Different Unit", "unitFail"
                                flags.append(f"Different Unit {feature}")
                            
                            elif result['state'] == "Match":
                                for j in range(min(len(result['nValue']), len(result['nValue2']))):
                                    v1 = float(result['nValue'][j])
                                    v2 = float(result['nValue2'][j])
                                    diff = percent_diff_base_a(v1, v2)
                                    
                                    g1, g2, g3 = rule.get("G1"), rule.get("G2"), rule.get("G3")
                                    
                                    if diff < g1:
                                        status, grade = f"✅ Within G1 ({diff}%)", "A"
                                    elif diff < g2:
                                        status, grade = f"✅ Within G2 ({diff}%)", "B"
                                    elif diff < g3 + 10:
                                        status, grade = f"✅ Within G3 ({diff}%)", "C"
                                    else:
                                        status, grade = f"❌ Outside tolerance ({diff}%)", "Fail"
                                        flags.append(f"Outside tolerance {feature}")
                            
                            elif result['state'] == "Different DetailedValueType":
                                status, grade = "Different DetailedValueType", "FailInDetailedValueType"
                                flags.append(f"Different DetailedValueType {feature}")
                            
                            elif result['state'] == "Different FeatureCode":
                                status, grade = "Different FeatureCode", "FailInFeatureCode"
                                flags.append(f"Different FeatureCode {feature}")
                            
                            else:
                                status, grade = f"Missing Values", "Fail"
                    
                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)
            
            # Determine overall grade
            if "DiffPinOut" in feature_grades:
                overall_grade = "DiffPinOut"
            elif "FailInCore" in feature_grades:
                overall_grade = "Not Drop-in"
            elif "FailInDetailedValueType" in feature_grades:
                overall_grade = "Detailed Value Type Fail Not Drop-in"
            elif "FailInFeatureCode" in feature_grades:
                overall_grade = "FeatureCode FAIL Not Drop-in"
            elif "unitFail" in feature_grades:
                overall_grade = "Unit FAIL Not Drop-in"
            elif "Fail" in feature_grades:
                overall_grade = "Drop-in D"
            elif "C" in feature_grades:
                overall_grade = "Drop-in C"
            elif "B" in feature_grades:
                overall_grade = "Drop-in B"
            else:
                overall_grade = "Drop-in A"
            
            results.append({
                "Part Number C": row2.get("PartNumber", ""),
                "Company Name C": row2.get("Company", ""),
                "Part Number X": row1.get("PartNumber", ""),
                "Company Name X": row1.get("Company", ""),
                "PL Name": row1.get("PL Name", ""),
                "Feature": "OVERALL",
                "Status": " | ".join(feature_statuses),
                "flags": " | ".join(flags) if flags else "",
                "Grade": overall_grade
            })
    
    df_report = pd.DataFrame(results)
    
    if output_path:
        df_report.to_excel(output_path, index=False)
    
    return df_report


# ============================================================================
# STEP 4: Final merge and organization
# ============================================================================

def merge_files(file1, file2, file3, output_path=None):
    """Merge three files into final output"""
    f1 = pd.read_excel(file1, dtype=str)
    f2 = pd.read_excel(file2, dtype=str)
    f3 = pd.read_excel(file3, dtype=str)
    
    # Strip column names
    for df in [f1, f2, f3]:
        df.columns = df.columns.str.strip()
    
    # Rename for consistency
    f2 = f2.rename(columns={
        "Part Number C": "PartNumberC",
        "Company Name C": "CompanyNameC",
        "Part Number X": "PartNumberX",
        "Company Name X": "CompanyNamex"
    })
    
    f3 = f3.rename(columns=lambda x: x.replace(" ", ""))
    
    # Merge operations
    merged = pd.merge(
        f3,
        f1[["PartNumberX", "CompanyNamex", "PartNumberC", "CompanyNameC",
            "SamePLs", "IsPackaeSimilar", "IsPinoutSimilar", "PackaePinout", "FoundData"]],
        on=["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex"],
        how="left"
    )
    
    merged = pd.merge(
        merged,
        f2[["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex", 
            "Status", "flags", "Grade"]],
        on=["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex"],
        how="left"
    )
    
    # Try swapped orientation
    swapped = f2.rename(columns={
        "PartNumberC": "PartNumberX",
        "CompanyNameC": "CompanyNamex",
        "PartNumberX": "PartNumberC",
        "CompanyNamex": "CompanyNameC"
    })
    
    merged = pd.merge(
        merged,
        swapped[["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex",
                "Status", "flags", "Grade"]],
        on=["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex"],
        how="left",
        suffixes=("", "_swapped")
    )
    
    # Fill from swapped
    for col in ["Status", "flags", "Grade"]:
        merged[col] = merged[col].fillna(merged[f"{col}_swapped"])
        merged.drop(columns=[f"{col}_swapped"], inplace=True, errors='ignore')
    
    if output_path:
        merged.to_excel(output_path, index=False)
    
    return merged


def organize_file(input_path, output_path=None):
    """Post-process and organize final output"""
    df = pd.read_excel(input_path, dtype=str)
    df.columns = df.columns.str.strip()
    
    # Rename columns
    df = df.rename(columns={
        "Status": "Match Feature",
        "flags": "Different Features"
    })
    
    # Add Status column
    def compute_status(row):
        grade = str(row.get("Grade", "")).strip()
        same_pls = str(row.get("SamePLs", "")).strip().lower()
        found_data = str(row.get("FoundData", "")).strip().upper()
        pack_pinout = str(row.get("PackaePinout", "")).strip()
        
        if grade in ["Drop-in D", "Drop-in A", "Drop-in B", "Drop-in C"]:
            return "Cross"
        elif grade in ["Not Drop-in", "Detailed Value Type Fail Not Drop-in",
                      "Unit FAIL Not Drop-in", "FeatureCode FAIL Not Drop-in"] and found_data == "TRUE":
            return "Not Cross"
        elif same_pls == "not":
            return "Different PLs"
        elif found_data == "FALSE" and pack_pinout == "":
            return "Not Found Data"
        else:
            return ""
    
    df["Status"] = df.apply(compute_status, axis=1)
    
    if output_path:
        df.to_excel(output_path, index=False)
    
    return df

# All your functions (process_files, deduplicate_columns, load_all_excel_sheets, 
# find_matching_row, get_single_row_all_files, merge_from_crossesparts, 
# percent_diff_base_a, fc_sort_key, compare_parts, validate_core_upgrade_equal, 
# merge_files, organize_file) remain UNCHANGED.

# Copy-paste your full function code here without modification
# ---------------------------------------------------------------------------
# >>> PASTE ALL FUNCTIONS FROM YOUR CODE ABOVE HERE <<<
# ---------------------------------------------------------------------------


# ============================================================================
# STREAMLIT APP
# ============================================================================

st.set_page_config(page_title="Cross-Reference Validation System", page_icon="🔄", layout="wide")

st.title("🔄 Optimized Cross-Reference Validation System")

st.markdown("Upload required files to run the validation pipeline.")

# File inputs
crosses_file = st.file_uploader("Upload Crosses File", type=["xlsx"])
parametric_file = st.file_uploader("Upload Parametric File", type=["xlsx"])
pakageAndPinout_file = st.file_uploader("Upload Package & Pinout File", type=["xlsx"])
recipe_file = st.file_uploader("Upload Recipe File", type=["xlsx"])
lookUpFile1 = st.file_uploader("Upload Lookup File 1", type=["xlsx"])
lookUpFile2 = st.file_uploader("Upload Lookup File 2", type=["xlsx"])

if st.button("🚀 Run Validation Pipeline"):
    if not all([crosses_file, parametric_file, pakageAndPinout_file, recipe_file, lookUpFile1, lookUpFile2]):
        st.error("⚠️ Please upload all required files.")
    else:
        with st.spinner("Processing... Please wait."):
            # Save uploaded files to temp dir for compatibility
            with tempfile.TemporaryDirectory() as tmpdir:
                crosses_path = os.path.join(tmpdir, "crosses.xlsx")
                parametric_path = os.path.join(tmpdir, "parametric.xlsx")
                pakage_path = os.path.join(tmpdir, "pakage.xlsx")
                recipe_path = os.path.join(tmpdir, "recipe.xlsx")
                lookup1_path = os.path.join(tmpdir, "lookup1.xlsx")
                lookup2_path = os.path.join(tmpdir, "lookup2.xlsx")

                for uploaded, path in [
                    (crosses_file, crosses_path),
                    (parametric_file, parametric_path),
                    (pakageAndPinout_file, pakage_path),
                    (recipe_file, recipe_path),
                    (lookUpFile1, lookup1_path),
                    (lookUpFile2, lookup2_path),
                ]:
                    with open(path, "wb") as f:
                        f.write(uploaded.read())

                # Output paths
                output_file_1 = os.path.join(tmpdir, "step1.xlsx")
                output_file_merged = os.path.join(tmpdir, "merged.xlsx")
                output_file_2 = os.path.join(tmpdir, "validation.xlsx")
                final_path = os.path.join(tmpdir, "FINAL.xlsx")
                final_org_path = os.path.join(tmpdir, "FINAL_ORGANIZED.xlsx")

                # Step 1
                st.write("### Step 1: Processing files...")
                process_files(crosses_path, parametric_path, pakage_path, recipe_path, output_path=output_file_1)

                # Step 2
                st.write("### Step 2: Merging data...")
                target_files = [parametric_path, pakage_path]
                merge_from_crossesparts(output_file_1, target_files, output_path=output_file_merged)

                # Step 3
                st.write("### Step 3: Validating...")
                validate_core_upgrade_equal(recipe_path, output_file_merged, lookup1_path, lookup2_path,
                                            loadLookUp="YES", output_path=output_file_2)

                # Step 4
                st.write("### Step 4: Final merge...")
                merge_files(output_file_1, output_file_2, crosses_path, output_path=final_path)

                # Step 5
                st.write("### Step 5: Organizing output...")
                final_df = organize_file(final_path, final_org_path)

                st.success("✅ Process completed successfully!")

                # Show sample of results
                st.subheader("📊 Final Results Preview")
                st.dataframe(final_df.head(50))

                # Download link
                with open(final_org_path, "rb") as f:
                    st.download_button("⬇️ Download Final Organized File", f, file_name="FINAL_ORGANIZED.xlsx")
