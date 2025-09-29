# app.py
import streamlit as st
import pandas as pd
import re
import os
import tempfile
from io import BytesIO

# ---------------------------
# Page configuration (always before other Streamlit calls)
# ---------------------------
st.set_page_config(
    page_title="Cross-Reference Validation System",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------
# CUSTOM CSS
# Place CSS here (top of file) so it applies globally.
# ---------------------------
st.markdown(
    """
    <style>
        .main-header {
            font-size: 2.5rem;
            font-weight: bold;
            color: #1f2937;
            margin-bottom: 0.5rem;
        }
        .sub-header {
            font-size: 1rem;
            color: #6b7280;
            margin-bottom: 2rem;
        }
        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 1.5rem;
            border-radius: 0.5rem;
            color: white;
            text-align: center;
        }
        .success-box {
            background-color: #d1fae5;
            border-left: 4px solid #10b981;
            padding: 1rem;
            border-radius: 0.25rem;
            margin: 1rem 0;
        }
        .error-box {
            background-color: #fee2e2;
            border-left: 4px solid #ef4444;
            padding: 1rem;
            border-radius: 0.25rem;
            margin: 1rem 0;
        }
        /* Progressbar color tweak */
        .stProgress > div > div > div > div {
            background-color: #667eea;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# ====================================================================
# Processing helper functions
# ====================================================================

def clean_to_floats(lst):
    cleaned = []
    for x in lst:
        try:
            if pd.notna(x) and str(x).strip().lower() != "nan" and str(x).strip() != "":
                cleaned.append(float(x))
        except ValueError:
            pass
    return cleaned

def percent_diff_base_a(a, b):
    if a == 0:
        return float('inf') if b != 0 else 0.0
    return round(abs(a - b) / abs(a) * 100, 2)

def deduplicate_columns(columns):
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

def process_files(crossfile, parametricData, pakageAndPinout, recipe):
    cross_df = pd.read_excel(crossfile, dtype=str)
    recipe_df = pd.read_excel(recipe, dtype=str)
    pakage_df = pd.read_excel(pakageAndPinout, dtype=str)

    parametric_sheets = pd.read_excel(parametricData, sheet_name=None, dtype=str)
    parametric_df = pd.concat(parametric_sheets.values(), ignore_index=True)

    # Compare PLc and PLx
    cross_df["SamePLs"] = cross_df.apply(
        lambda r: "same" if str(r.get("PLc","")).strip() == str(r.get("PLx","")).strip() else "not",
        axis=1
    )

    # Prepare similarity columns
    cross_df["IsPackaeSimilar"] = "NA"
    cross_df["IsPinoutSimilar"] = "NA"
    cross_df["PackaePinout"] = "NA"

    # Fill similarity values based on recipe
    for idx, row in cross_df.iterrows():
        if row["SamePLs"] == "same":
            plc = str(row["PLc"]).strip()
            subset = recipe_df[recipe_df["ZProductValue"].astype(str).str.strip() == plc]

            if not subset.empty:
                is_pack = "FALSE" if (subset["IsPackaeSimilar"].astype(str).str.upper() == "FALSE").all() else "TRUE"
                is_pin = "FALSE" if (subset["IsPinoutSimilar"].astype(str).str.upper() == "FALSE").all() else "TRUE"

                cross_df.at[idx, "IsPackaeSimilar"] = is_pack
                cross_df.at[idx, "IsPinoutSimilar"] = is_pin
                cross_df.at[idx, "PackaePinout"] = "TRUE" if (is_pack == "TRUE" or is_pin == "TRUE") else "FALSE"

    # Check FoundData
    cross_df["FoundData"] = "FALSE"
    for idx, row in cross_df.iterrows():
        part_c = str(row.get("PartNumberC","")).strip()
        part_x = str(row.get("PartNumberX","")).strip()

        if row["SamePLs"] == "same":
            if row["PackaePinout"] == "TRUE":
                exists_c = ((parametric_df["PartNumber"] == part_c).any() or (pakage_df["PartNumber"] == part_c).any())
                exists_x = ((parametric_df["PartNumber"] == part_x).any() or (pakage_df["PartNumber"] == part_x).any())
            elif row["PackaePinout"] == "FALSE":
                exists_c = (parametric_df["PartNumber"] == part_c).any()
                exists_x = (parametric_df["PartNumber"] == part_x).any()
            else:
                exists_c = exists_x = False

            if exists_c and exists_x:
                cross_df.at[idx, "FoundData"] = "TRUE"

    return cross_df

def load_all_excel_sheets(file_list):
    cache = {}
    for file in file_list:
        try:
            xl = pd.ExcelFile(file)
            cache[file] = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}
        except Exception as e:
            st.warning(f"Error reading {file}: {e}")
            cache[file] = None
    return cache

def find_matching_row(df, part_value, company_value):
    df = df.copy()
    df.columns = df.columns.str.strip()
    cols = [c.lower() for c in df.columns]

    company_col = next((df.columns[i] for i, c in enumerate(cols) if c in ["company", "companyname", "company name", "companynamec"]), None)
    part_col = next((df.columns[i] for i, c in enumerate(cols) if c in ["partnumber", "part number", "part_number", "partnumberc"]), None)

    if not company_col or not part_col:
        return None

    filtered = df[
        (df[part_col].astype(str).str.strip() == part_value) &
        (df[company_col].astype(str).str.strip() == company_value)
    ]

    if not filtered.empty:
        row = filtered.iloc[0].copy()
        row = row.drop([part_col, company_col], errors='ignore')
        row = row.to_frame().T.reset_index(drop=True)
        return row

    return None

def get_single_row_all_files(file_cache, part_value, company_value, PackaePinout_val):
    merged_parts = [pd.DataFrame({"PartNumber": [part_value], "Company": [company_value], "Comments": [""]})]
    PackaePinout_bool = str(PackaePinout_val).strip().lower() in ["true", "1", "yes"]

    file_list = list(file_cache.keys())
    # If PackaePinout == False we skip the last file (as in your original logic)
    if not PackaePinout_bool and len(file_list) > 0:
        file_list = file_list[:-1]

    for file in file_list:
        file_sheets = file_cache.get(file)
        if file_sheets is None:
            if PackaePinout_bool:
                merged_parts[0].loc[0, "Comments"] += f"Error reading {os.path.basename(file)}; "
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
            merged_parts[0].loc[0, "Comments"] += f"Not found in {os.path.basename(file)}; "

    final_row = pd.concat(merged_parts, axis=1)
    final_row.columns = deduplicate_columns(final_row.columns)

    return final_row

def merge_from_crossesparts(crosses_file, target_files):
    df_crosses = pd.read_excel(crosses_file)

    required_cols = ["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex", "PackaePinout"]
    for col in required_cols:
        if col not in df_crosses.columns:
            raise ValueError(f"Missing required column: {col}")

    file_cache = load_all_excel_sheets(target_files)
    all_merged_rows = []

    progress_bar = st.progress(0)
    status_text = st.empty()

    total_rows = len(df_crosses)

    for idx, row in df_crosses.iterrows():
        progress = (idx + 1) / total_rows if total_rows > 0 else 1.0
        progress_bar.progress(progress)
        status_text.text(f"Merging data: {idx + 1}/{total_rows} records processed")

        part_val_c = str(row["PartNumberC"]).strip()
        company_val_c = str(row["CompanyNameC"]).strip()
        part_val_x = str(row["PartNumberX"]).strip()
        company_val_x = str(row["CompanyNamex"]).strip()
        PackaePinout_val = str(row["PackaePinout"]).strip()

        merged_c = get_single_row_all_files(file_cache, part_val_c, company_val_c, PackaePinout_val)
        merged_x = get_single_row_all_files(file_cache, part_val_x, company_val_x, PackaePinout_val)

        all_merged_rows.extend([merged_c, merged_x])

    progress_bar.empty()
    status_text.empty()

    if len(all_merged_rows) == 0:
        return pd.DataFrame()
    final_df = pd.concat(all_merged_rows, ignore_index=True)
    return final_df

def compare_parts(df, value, value2):
    part1 = df[df["AcceptedValue"] == value].reset_index(drop=True)
    part2 = df[df["AcceptedValue"] == value2].reset_index(drop=True)

    if part1.empty or part2.empty:
        return {
            "value": value, "nValue": [], "nUnit": [],
            "FeatureCode": "NotFound",
            "state": "NotFound In File",
            "value2": value2, "nValue2": [], "nUnit2": []
        }

    chosen_value_id = part1.loc[0, "ValueID"]
    chosen_value_id2 = part2.loc[0, "ValueID"]

    part1 = part1[part1["ValueID"] == chosen_value_id].reset_index(drop=True)
    part2 = part2[part2["ValueID"] == chosen_value_id2].reset_index(drop=True)

    fc1 = set(part1["FeatureCode"].dropna().unique())
    fc2 = set(part2["FeatureCode"].dropna().unique())
    if fc1 != fc2:
        return {
            "value": value, "nValue": [], "nUnit": [],
            "FeatureCode": "Different",
            "state": "Different FeatureCode",
            "value2": value2, "nValue2": [], "nUnit2": []
        }

    def fc_sort_key(fc):
        match = re.search(r"([VU])(\d+)$", fc)
        if match:
            t, num = match.groups()
            return int(num) * 2 + (0 if t == "V" else 1)
        return float("inf")

    part1_sorted = part1.sort_values(by="FeatureCode", key=lambda x: x.map(fc_sort_key))
    part2_sorted = part2.sort_values(by="FeatureCode", key=lambda x: x.map(fc_sort_key))

    nValue = part1_sorted["NormalizedValue"].dropna().tolist()
    nUnit = part1_sorted["Unit"].dropna().tolist()
    nValue2 = part2_sorted["NormalizedValue"].dropna().tolist()
    nUnit2 = part2_sorted["Unit"].dropna().tolist()

    return {
        "value": value, "nValue": nValue, "nUnit": nUnit,
        "FeatureCode": "Match", "state": "Match",
        "value2": value2, "nValue2": nValue2, "nUnit2": nUnit2
    }

def validate_core_upgrade_equal(newcross_file, finalmerged_file, lookup_file1, lookup_file2, loadLookUp):
    rules_df = pd.read_excel(newcross_file, dtype=str)
    data = pd.read_excel(finalmerged_file, dtype=str)

    values = None
    if loadLookUp == "YES" and lookup_file1 and lookup_file2:
        values1 = pd.read_excel(lookup_file1, dtype=str)
        values2 = pd.read_excel(lookup_file2, dtype=str)
        values = pd.concat([values1, values2], ignore_index=True)

    if "ProductName" not in data.columns:
        raise ValueError("'ProductName' column not found in merged file")

    PLs = data["ProductName"].astype(str).unique().tolist()
    rules_df = rules_df[rules_df["ZProductValue"].astype(str).isin(PLs)]

    rules_df = rules_df[
        (rules_df["FeaturesType"].str.lower() == "core") |
        (rules_df["UpgradeFeature"].str.upper() == "TRUE") |
        (rules_df["IsEqualFeature"].str.upper() == "TRUE") |
        (rules_df["FeaturesType"].str.lower() == "tolerance")
    ]

    for col in ["G1", "G2", "G3"]:
        if col in rules_df.columns:
            rules_df[col] = pd.to_numeric(rules_df[col], errors="coerce")

    results = []
    grouped_data = data.groupby("ProductName")

    progress_bar = st.progress(0)
    status_text = st.empty()
    total_pls = len(grouped_data)
    current_pl = 0

    for PL, data_subset in grouped_data:
        current_pl += 1
        progress = current_pl / total_pls if total_pls > 0 else 1.0
        progress_bar.progress(progress)
        status_text.text(f"Validating: Processing PL {current_pl}/{total_pls}")

        data_subset = data_subset.reset_index(drop=True)
        rules_subset = rules_df[rules_df["ZProductValue"].astype(str).str.strip() == PL].reset_index(drop=True)

        PinOut_flag = "FALSE"
        if not rules_df.empty and "IsPinoutSimilar" in rules_df.columns:
            if rules_df["IsPinoutSimilar"].iloc[0] == "True":
                PinOut_flag = "TRUE"

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

                # Pinout check (only do it once per pair)
                if PinOut_flag == "TRUE" and count:
                    val1, val2 = str(row1.get('NormalizedPinName', '')).strip(), str(row2.get('NormalizedPinName', '')).strip()
                    count = 0
                    if val1 == val2:
                        status, grade = f"✅ PinOut Match", "A"
                    else:
                        status, grade = f"❌ Different PinOut", "DiffPinOut"
                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)

                # Core or equal features
                if ftype == "core" or equal_flag == "TRUE":
                    if feature not in data_subset.columns:
                        status, grade = f"⚠ Missing column {feature}", "Fail"
                    else:
                        val1, val2 = str(row1[feature]).strip(), str(row2[feature]).strip()
                        if val1 == val2:
                            status, grade = f"✅ Match", "A"
                        else:
                            status, grade = f"❌ Mismatch ({val1} vs {val2})", "Fail"
                            if ftype == "core":
                                flags.append(f"Core mismatch in {feature}")
                                grade = "FailInCore"
                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)

                elif upgrade_flag == "TRUE":
                    # skip upgrade-only features in current logic
                    continue

                elif ftype == "tolerance":
                    if feature not in data_subset.columns:
                        status, grade = f"⚠ Missing column {feature}", "Fail"
                    else:
                        val1, val2 = str(row1[feature]).strip(), str(row2[feature]).strip()
                        if val1 == val2:
                            status, grade = f"✅ Match", "A"
                        else:
                            if loadLookUp == "YES" and values is not None:
                                result = compare_parts(values, val1, val2)
                                if result['nUnit'] != result['nUnit2']:
                                    status, grade = f"❌ Different units", "unitFail"
                                    flags.append(f"Unit mismatch in {feature}")
                                elif result['state'] == "Match":
                                    # apply G1/G2/G3 thresholds if present
                                    for j in range(min(len(result['nValue']), len(result['nValue2']))):
                                        try:
                                            v1 = float(result['nValue'][j])
                                            v2 = float(result['nValue2'][j])
                                        except Exception:
                                            continue
                                        diff = percent_diff_base_a(v1, v2)
                                        g1, g2, g3 = rule.get("G1"), rule.get("G2"), rule.get("G3")
                                        # if g1/g2/g3 are NaN, treat as very large
                                        try:
                                            g1 = float(g1) if pd.notna(g1) else float("inf")
                                            g2 = float(g2) if pd.notna(g2) else float("inf")
                                            g3 = float(g3) if pd.notna(g3) else float("inf")
                                        except Exception:
                                            g1 = g2 = g3 = float("inf")

                                        if diff < g1:
                                            status, grade = f"✅ Within G1 tolerance", "A"
                                        elif diff < g2:
                                            status, grade = f"✅ Within G2 tolerance", "B"
                                        elif diff < (g3 + 10):
                                            status, grade = f"✅ Within G3 tolerance", "C"
                                        else:
                                            status, grade = f"❌ Outside tolerance (diff={diff}%)", "Fail"
                                            flags.append(f"Tolerance exceeded in {feature}")
                                else:
                                    status, grade = f"❌ {result['state']}", "Fail"
                            else:
                                status, grade = f"⚠ Lookup not loaded", "Fail"
                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)

            # Determine overall grade
            if "DiffPinOut" in feature_grades:
                overall_grade = "DiffPinOut"
            elif "FailInCore" in feature_grades:
                overall_grade = "Not Drop-in"
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

    progress_bar.empty()
    status_text.empty()

    df_report = pd.DataFrame(results)
    return df_report

def merge_files(file1, file2, file3):
    f1 = pd.read_excel(file1, dtype=str)
    f2 = pd.read_excel(file2, dtype=str)
    f3 = pd.read_excel(file3, dtype=str)

    f1.columns = f1.columns.str.strip()
    f2.columns = f2.columns.str.strip()
    f3.columns = f3.columns.str.strip()

    f2 = f2.rename(columns={
        "Part Number C": "PartNumberC",
        "Company Name C": "CompanyNameC",
        "Part Number X": "PartNumberX",
        "Company Name X": "CompanyNamex"
    })

    f3 = f3.rename(columns=lambda x: x.replace(" ", ""))

    merged = pd.merge(
        f3,
        f1[["PartNumberX", "CompanyNamex", "PartNumberC", "CompanyNameC",
            "SamePLs", "IsPackaeSimilar", "IsPinoutSimilar", "PackaePinout", "FoundData"]],
        on=["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex"],
        how="left"
    )

    merged = pd.merge(
        merged,
        f2[["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex", "Status", "flags", "Grade"]],
        on=["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex"],
        how="left"
    )

    return merged

def organize_file(df):
    df = df.copy()
    df.columns = df.columns.str.strip()

    rename_map = {
        "Status": "Match Feature",
        "flags": "Different Features"
    }
    df = df.rename(columns=rename_map)

    def compute_status(row):
        grade = str(row.get("Grade", "")).strip()
        same_pls = str(row.get("SamePLs", "")).strip().lower()
        found_data = str(row.get("FoundData", "")).strip().upper()

        if grade in ["Drop-in D", "Drop-in A", "Drop-in B", "Drop-in C"]:
            return "Cross"
        elif grade in ["Not Drop-in", "Detailed Value Type Fail Not Drop-in",
                       "Unit FAIL Not Drop-in", "FeatureCode FAIL Not Drop-in"] and found_data == "TRUE":
            return "Not Cross"
        elif same_pls == "not":
            return "Different PLs"
        elif found_data == "FALSE":
            return "Not Found Data"
        else:
            return ""

    df["Status"] = df.apply(compute_status, axis=1)
    return df

def save_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ====================================================================
# STREAMLIT UI
# ====================================================================

def main():
    # Header
    st.markdown('<p class="main-header">🔄 Cross-Reference Validation System</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Upload your data files to validate part cross-references and generate comprehensive reports</p>', unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.header("📋 Instructions")
        st.markdown("""
        1. Upload all required files (marked with *)
        2. Optionally upload lookup tables
        3. Click 'Process Files' to start
        4. Download the generated reports

        **File Requirements:**
        - Excel format (.xlsx or .xls)
        - Proper column structure
        - Valid data types
        """)
        st.divider()
        st.header("📊 About")
        st.markdown("""
        This system validates cross-references between parts by:
        - Comparing product lines
        - Checking package similarity
        - Validating core features
        - Calculating tolerance grades
        """)

    # Initialize session state
    if 'processed' not in st.session_state:
        st.session_state.processed = False
    if 'results' not in st.session_state:
        st.session_state.results = {}

    # File upload section
    st.header("📁 Upload Files")
    col1, col2 = st.columns(2)
    with col1:
        crosses_file = st.file_uploader("Cross Reference File *", type=['xlsx', 'xls'], key="crosses")
        parametric_file = st.file_uploader("Parametric Data File *", type=['xlsx', 'xls'], key="parametric")
        package_file = st.file_uploader("Package & Pinout File *", type=['xlsx', 'xls'], key="package")
    with col2:
        recipe_file = st.file_uploader("Recipe File *", type=['xlsx', 'xls'], key="recipe")
        lookup1_file = st.file_uploader("Lookup Table 1 (Optional)", type=['xlsx', 'xls'], key="lookup1")
        lookup2_file = st.file_uploader("Lookup Table 2 (Optional)", type=['xlsx', 'xls'], key="lookup2")

    # check required
    all_required = all([crosses_file, parametric_file, package_file, recipe_file])
    st.divider()

    # Process button
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        process_btn = st.button("🚀 Process Files", type="primary", disabled=not all_required, use_container_width=True)

    if not all_required:
        st.warning("⚠️ Please upload all required files (marked with *) to continue.")

    # Processing flow
    if process_btn:
        try:
            with st.spinner("Processing files..."):
                with tempfile.TemporaryDirectory() as tmpdir:
                    # Save uploaded files temporarily
                    crosses_path = os.path.join(tmpdir, "crosses.xlsx")
                    parametric_path = os.path.join(tmpdir, "parametric.xlsx")
                    package_path = os.path.join(tmpdir, "package.xlsx")
                    recipe_path = os.path.join(tmpdir, "recipe.xlsx")

                    with open(crosses_path, 'wb') as f:
                        f.write(crosses_file.getvalue())
                    with open(parametric_path, 'wb') as f:
                        f.write(parametric_file.getvalue())
                    with open(package_path, 'wb') as f:
                        f.write(package_file.getvalue())
                    with open(recipe_path, 'wb') as f:
                        f.write(recipe_file.getvalue())

                    # lookup files (optional)
                    use_lookup = "NO"
                    lookup1_path = None
                    lookup2_path = None
                    if lookup1_file and lookup2_file:
                        use_lookup = "YES"
                        lookup1_path = os.path.join(tmpdir, "lookup1.xlsx")
                        lookup2_path = os.path.join(tmpdir, "lookup2.xlsx")
                        with open(lookup1_path, 'wb') as f:
                            f.write(lookup1_file.getvalue())
                        with open(lookup2_path, 'wb') as f:
                            f.write(lookup2_file.getvalue())

                    # Step 1
                    st.info("Step 1/5: Validating input and comparing PLs...")
                    cross_df = process_files(crosses_path, parametric_path, package_path, recipe_path)
                    step1_path = os.path.join(tmpdir, "step1_output.xlsx")
                    cross_df.to_excel(step1_path, index=False)

                    # Step 2
                    st.info("Step 2/5: Merging data from multiple sources...")
                    target_files = [parametric_path, package_path]
                    merged_df = merge_from_crossesparts(step1_path, target_files)
                    merged_path = os.path.join(tmpdir, "merged_output.xlsx")
                    merged_df.to_excel(merged_path, index=False)

                    # Step 3
                    st.info("Step 3/5: Validating core features and tolerances...")
                    report_df = validate_core_upgrade_equal(
                        recipe_path, merged_path, lookup1_path, lookup2_path, use_lookup
                    )
                    report_path = os.path.join(tmpdir, "validation_report.xlsx")
                    report_df.to_excel(report_path, index=False)

                    # Step 4
                    st.info("Step 4/5: Merging all results...")
                    final_df = merge_files(step1_path, report_path, crosses_path)
                    final_path = os.path.join(tmpdir, "final_merged.xlsx")
                    final_df.to_excel(final_path, index=False)

                    # Step 5
                    st.info("Step 5/5: Organizing final output...")
                    organized_df = organize_file(final_df)

                    # Store results
                    st.session_state.results = {
                        'step1': cross_df,
                        'merged': merged_df,
                        'validation': report_df,
                        'final': final_df,
                        'organized': organized_df
                    }
                    st.session_state.processed = True

            st.markdown('<div class="success-box">✅ Processing completed successfully!</div>', unsafe_allow_html=True)

        except Exception as e:
            st.markdown(f'<div class="error-box">❌ Error during processing: {str(e)}</div>', unsafe_allow_html=True)
            st.error(f"Full error details: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

    # Display results
    if st.session_state.processed:
        st.divider()
        st.header("📊 Results Summary")

        organized_df = st.session_state.results['organized']

        # stats (safe checks if dataframe empty)
        total_records = len(organized_df) if organized_df is not None else 0
        cross_matches = len(organized_df[organized_df['Status'] == 'Cross']) if total_records > 0 and 'Status' in organized_df.columns else 0
        not_cross = len(organized_df[organized_df['Status'] == 'Not Cross']) if total_records > 0 and 'Status' in organized_df.columns else 0
        different_pls = len(organized_df[organized_df['Status'] == 'Different PLs']) if total_records > 0 and 'Status' in organized_df.columns else 0
        not_found = len(organized_df[organized_df['Status'] == 'Not Found Data']) if total_records > 0 and 'Status' in organized_df.columns else 0

        grade_counts = organized_df['Grade'].value_counts() if total_records > 0 and 'Grade' in organized_df.columns else pd.Series(dtype=int)
        drop_in_a = int(grade_counts.get('Drop-in A', 0))
        drop_in_b = int(grade_counts.get('Drop-in B', 0))
        drop_in_c = int(grade_counts.get('Drop-in C', 0))
        drop_in_d = int(grade_counts.get('Drop-in D', 0))
        not_drop_in = int(sum([grade_counts[g] for g in grade_counts.index if 'Not Drop-in' in str(g)])) if not grade_counts.empty else 0

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Records", total_records, help="Total number of cross-reference records processed")
        with col2:
            st.metric("Cross Matches", cross_matches, help="Number of valid cross-references found")
        with col3:
            st.metric("Drop-in A", drop_in_a, help="Highest quality matches")
        with col4:
            st.metric("Not Found Data", not_found, help="Records with missing data")

        # Detailed breakdown
        st.subheader("📈 Detailed Breakdown")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Status Distribution**")
            status_data = {
                "Cross": cross_matches,
                "Not Cross": not_cross,
                "Different PLs": different_pls,
                "Not Found Data": not_found
            }
            for status, count in status_data.items():
                percentage = (count / total_records * 100) if total_records > 0 else 0
                st.metric(status, f"{count} ({percentage:.1f}%)")
        with col2:
            st.markdown("**Grade Distribution**")
            grade_data = {
                "Drop-in A": drop_in_a,
                "Drop-in B": drop_in_b,
                "Drop-in C": drop_in_c,
                "Drop-in D": drop_in_d,
                "Not Drop-in": not_drop_in
            }
            for grade, count in grade_data.items():
                percentage = (count / total_records * 100) if total_records > 0 else 0
                st.metric(grade, f"{count} ({percentage:.1f}%)")

        # Data preview
        st.divider()
        st.subheader("🔍 Data Preview")
        tab1, tab2, tab3, tab4 = st.tabs(["Final Organized", "Validation Report", "Merged Data", "Step 1 Output"])
        with tab1:
            st.dataframe(organized_df.head(100) if total_records > 0 else pd.DataFrame(), use_container_width=True)
        with tab2:
            st.dataframe(st.session_state.results['validation'].head(100) if 'validation' in st.session_state.results else pd.DataFrame(), use_container_width=True)
        with tab3:
            st.dataframe(st.session_state.results['merged'].head(100) if 'merged' in st.session_state.results else pd.DataFrame(), use_container_width=True)
        with tab4:
            st.dataframe(st.session_state.results['step1'].head(100) if 'step1' in st.session_state.results else pd.DataFrame(), use_container_width=True)

        # Download section
        st.divider()
        st.header("💾 Download Reports")
        col1, col2, col3 = st.columns(3)
        with col1:
            organized_excel = save_to_excel(organized_df if organized_df is not None else pd.DataFrame())
            st.download_button(
                label="📥 Download Final Organized Report",
                data=organized_excel,
                file_name="final_organized_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col2:
            validation_excel = save_to_excel(st.session_state.results['validation'] if 'validation' in st.session_state.results else pd.DataFrame())
            st.download_button(
                label="📥 Download Validation Report",
                data=validation_excel,
                file_name="validation_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col3:
            merged_excel = save_to_excel(st.session_state.results['merged'] if 'merged' in st.session_state.results else pd.DataFrame())
            st.download_button(
                label="📥 Download Merged Data",
                data=merged_excel,
                file_name="merged_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # Reset button
        st.divider()
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("🔄 Process New Files", use_container_width=True):
                st.session_state.processed = False
                st.session_state.results = {}
                st.experimental_rerun()

if __name__ == "__main__":
    main()
