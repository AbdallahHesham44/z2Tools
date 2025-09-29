import streamlit as st
import pandas as pd
import os
from io import BytesIO
import time

# Set page config
st.set_page_config(
    page_title="Component Cross-Reference Validator",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main > div {
        padding-top: 2rem;
    }
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        color: #856404;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
    }
</style>
""", unsafe_allow_html=True)

# ==================== PROCESSING FUNCTIONS ====================

def deduplicate_columns(columns):
    """Make duplicate column names unique by adding suffixes _1, _2, etc."""
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
    """Load all sheets from all Excel files once and return a cache."""
    cache = {}
    for file in file_list:
        try:
            xl = pd.ExcelFile(file)
            cache[file] = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}
        except Exception as e:
            st.error(f"❌ Error reading {file}: {e}")
            cache[file] = None
    return cache


def find_matching_row(df, part_value, company_value):
    """Search for the row in df that matches part and company."""
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
    """Merge one row from multiple Excel files into a single row (horizontal merge)."""
    merged_parts = [pd.DataFrame({"PartNumber": [part_value], "Company": [company_value], "Comments": [""]})]
    PackaePinout_bool = str(PackaePinout_val).strip().lower() in ["true", "1", "yes"]

    file_list = list(file_cache.keys())
    if not PackaePinout_bool:
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
            merged_parts[0].loc[0, "Comments"] += f"Not found in {os.path.basename(file)} for {part_value},{company_value}; "

    final_row = pd.concat(merged_parts, axis=1)
    final_row.columns = deduplicate_columns(final_row.columns)

    return final_row


def process_files(crossfile, parametricData, pakageAndPinout, recipe):
    """Step 1: Validate input files and add comparison columns."""
    cross_df = pd.read_excel(crossfile, dtype=str)
    recipe_df = pd.read_excel(recipe, dtype=str)
    pakage_df = pd.read_excel(pakageAndPinout, dtype=str)

    parametric_sheets = pd.read_excel(parametricData, sheet_name=None, dtype=str)
    parametric_df = pd.concat(parametric_sheets.values(), ignore_index=True)

    cross_df["SamePLs"] = cross_df.apply(
        lambda r: "same" if str(r.get("PLc","")).strip() == str(r.get("PLx","")).strip() else "not",
        axis=1
    )

    cross_df["IsPackaeSimilar"] = "NA"
    cross_df["IsPinoutSimilar"] = "NA"
    cross_df["PackaePinout"] = "NA"

    for idx, row in cross_df.iterrows():
        if row["SamePLs"] == "same":
            plc = str(row["PLc"]).strip()
            subset = recipe_df[recipe_df["ZProductValue"].astype(str).str.strip() == plc]

            if not subset.empty:
                is_pack = "FALSE" if (subset["IsPackaeSimilar"].str.upper() == "FALSE").all() else "TRUE"
                is_pin = "FALSE" if (subset["IsPinoutSimilar"].str.upper() == "FALSE").all() else "TRUE"

                cross_df.at[idx, "IsPackaeSimilar"] = is_pack
                cross_df.at[idx, "IsPinoutSimilar"] = is_pin
                cross_df.at[idx, "PackaePinout"] = "TRUE" if (is_pack == "TRUE" or is_pin == "TRUE") else "FALSE"

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


def merge_from_crossesparts(crosses_file, target_files):
    """Step 2: Merge data from all files for each part."""
    df_crosses = pd.read_excel(crosses_file)

    required_cols = ["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex", "PackaePinout"]
    for col in required_cols:
        if col not in df_crosses.columns:
            raise ValueError(f"❌ Missing required column: {col}")

    file_cache = load_all_excel_sheets(target_files)
    all_merged_rows = []

    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total_rows = len(df_crosses)
    
    for idx, row in df_crosses.iterrows():
        part_val_c = str(row["PartNumberC"]).strip()
        company_val_c = str(row["CompanyNameC"]).strip()
        part_val_x = str(row["PartNumberX"]).strip()
        company_val_x = str(row["CompanyNamex"]).strip()
        PackaePinout_val = str(row["PackaePinout"]).strip()

        status_text.text(f"Processing {idx+1}/{total_rows}: {part_val_c} & {part_val_x}")
        
        merged_c = get_single_row_all_files(file_cache, part_val_c, company_val_c, PackaePinout_val)
        merged_x = get_single_row_all_files(file_cache, part_val_x, company_val_x, PackaePinout_val)

        all_merged_rows.extend([merged_c, merged_x])
        
        progress_bar.progress((idx + 1) / total_rows)

    status_text.text("Merging complete!")
    final_df = pd.concat(all_merged_rows, ignore_index=True)

    return final_df


def percent_diff_base_a(a, b):
    """Calculate percentage difference based on first value."""
    if a == 0:
        return float('inf') if b != 0 else 0.0
    return round(abs(a - b) / abs(a) * 100, 2)


def compare_parts(df, value, value2):
    """Compare two parts in dataframe by AcceptedValue."""
    part1_unique = df[df["AcceptedValue"] == value]
    part2_unique = df[df["AcceptedValue"] == value2]
    
    unique_cols = [
        "AcceptedValue", "AttributeName", "DetailedValueType", "Pattern",
        "FeatureCode", "AttributeValue", "NormalizedValue", "Unit", "Identifier"
    ]

    part1 = part1_unique.drop_duplicates(subset=unique_cols).reset_index(drop=True)
    part2 = part2_unique.drop_duplicates(subset=unique_cols).reset_index(drop=True)

    if part1.empty or part2.empty:
        return {
            "value": value, "nValue": [], "nUnit": [],
            "state": "NotFound In File ",
            "value2": value2, "nValue2": [], "nUnit2": []
        }

    dvt1 = set(part1["DetailedValueType"].dropna().unique())
    dvt2 = set(part2["DetailedValueType"].dropna().unique())
    if dvt1 != dvt2:
        return {
            "value": value, "nValue": [], "nUnit": [], "DetailedValueType": dvt1,
            "state": "Different DetailedValueType",
            "value2": value2, "nValue2": [], "nUnit2": [], "DetailedValueType2": dvt2
        }

    fc1 = set(part1["FeatureCode"].dropna().unique())
    fc2 = set(part2["FeatureCode"].dropna().unique())
    if fc1 != fc2:
        return {
            "value": value, "nValue": [], "nUnit": [],
            "state": "Different FeatureCode",
            "value2": value2, "nValue2": [], "nUnit2": []
        }

    def extract_values(part):
        nvals = part.loc[part["FeatureCode"].str.contains(r"V\d+$", na=False), "Normalized value"].tolist()
        nunits = part.loc[part["FeatureCode"].str.contains(r"U\d+$", na=False), "Unit"].tolist()
        return nvals, nunits

    nValue, nUnit = extract_values(part1)
    nValue2, nUnit2 = extract_values(part2)

    return {
        "value": value, "nValue": nValue, "nUnit": nUnit,
        "state": "Match",
        "value2": value2, "nValue2": nValue2, "nUnit2": nUnit2
    }


def validate_core_upgrade_equal(recipe_file, finalmerged_file, lookup_file1, lookup_file2, loadLookUp):
    """Step 3: Validate features and assign grades."""
    rules_df = pd.read_excel(recipe_file, dtype=str)
    data = pd.read_excel(finalmerged_file, dtype=str)

    values = None
    if loadLookUp == "YES" and lookup_file1 and lookup_file2:
        values1 = pd.read_excel(lookup_file1, dtype=str)
        values2 = pd.read_excel(lookup_file2, dtype=str)
        values = pd.concat([values1, values2], ignore_index=True)

    if "ProductName" not in data.columns:
        raise ValueError("❌ 'ProductName' column not found in merged file")

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
    
    for pl_idx, (PL, data_subset) in enumerate(grouped_data):
        status_text.text(f"Validating PL {pl_idx+1}/{total_pls}: {PL}")
        data_subset = data_subset.reset_index(drop=True)
        rules_subset = rules_df[rules_df["ZProductValue"].astype(str).str.strip() == PL].reset_index(drop=True)

        for i in range(0, len(data_subset), 2):
            if i + 1 >= len(data_subset):
                break

            row1, row2 = data_subset.iloc[i], data_subset.iloc[i+1]
            feature_statuses, feature_grades, flags = [], [], []

            for _, rule in rules_subset.iterrows():
                feature = rule["Features"]
                ftype = str(rule["FeaturesType"]).lower()
                upgrade_flag = str(rule.get("UpgradeFeature", "")).upper()
                equal_flag = str(rule.get("IsEqualFeature", "")).upper()

                if ftype == "core" or equal_flag == "TRUE":
                    if feature not in data_subset.columns:
                        status, grade = f"⚠ Missing column {feature}", "Fail"
                    else:
                        val1, val2 = str(row1[feature]).strip(), str(row2[feature]).strip()
                        if val1 == val2:
                            status, grade = f"✅ Match ({feature})", "A"
                        else:
                            status, grade = f"❌ Mismatch ({feature}: {val1} vs {val2})", "Fail"
                            if ftype == "core":
                                flags.append("missMatchAtCore")
                                grade = "FailInCore"

                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)

                elif ftype == "tolerance" and values is not None:
                    if feature not in data_subset.columns:
                        status, grade = f"⚠ Missing column {feature}", "Fail"
                    else:
                        val1, val2 = str(row1[feature]).strip(), str(row2[feature]).strip()
                        result = compare_parts(values, val1, val2)
                        
                        if result['state'] == "Match":
                            for j in range(len(result['nValue'])):
                                v1, v2 = float(result['nValue'][j]), float(result['nValue2'][j])
                                diff = percent_diff_base_a(v1, v2)
                                g1, g2, g3 = rule.get("G1"), rule.get("G2"), rule.get("G3")
                                
                                if diff < g1:
                                    status, grade = f"✅ Within {g1}% tolerance ({v1} vs {v2})", "A"
                                elif diff < g2:
                                    status, grade = f"✅ Within {g2}% tolerance ({v1} vs {v2})", "B"
                                elif diff < g3 + 10:
                                    status, grade = f"✅ Within {g3}% tolerance ({v1} vs {v2})", "C"
                                else:
                                    status, grade = f"❌ Outside tolerance ({v1} vs {v2}) diff={diff}", "Fail"
                                    flags.append(f"missMatch At Feature {feature}({v1},{v2}) diff={diff}")
                        else:
                            status, grade = f"⚠ {result['state']}", "Fail"

                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)

            if "FailInCore" in feature_grades:
                overall_grade = "Not Drop-in"
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
        
        progress_bar.progress((pl_idx + 1) / total_pls)

    status_text.text("Validation complete!")
    return pd.DataFrame(results)


def merge_files(file1, file2, file3):
    """Step 4: Merge all processed files together."""
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
    """Step 5: Final organization and status assignment."""
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
        elif grade == "Not Drop-in" and found_data == "TRUE":
            return "Not Cross"
        elif same_pls == "not":
            return "Different PLs"
        elif found_data == "FALSE":
            return "Not Found Data"
        else:
            return ""

    df["Status"] = df.apply(compute_status, axis=1)

    return df


def to_excel(df):
    """Convert dataframe to Excel bytes for download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


# ==================== STREAMLIT UI ====================

def main():
    st.title("🔬 Component Cross-Reference Validator")
    st.markdown("Upload your data files and validate component cross-references with automated grading")
    
    # Sidebar for file uploads
    st.sidebar.header("📁 Upload Files")
    
    crosses_file = st.sidebar.file_uploader("Cross Reference File *", type=['xlsx', 'xls'], key='crosses')
    parametric_file = st.sidebar.file_uploader("Parametric Data File *", type=['xlsx', 'xls'], key='parametric')
    package_file = st.sidebar.file_uploader("Package & Pinout File *", type=['xlsx', 'xls'], key='package')
    recipe_file = st.sidebar.file_uploader("Recipe File *", type=['xlsx', 'xls'], key='recipe')
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Optional Files")
    lookup1_file = st.sidebar.file_uploader("Lookup Table 1", type=['xlsx', 'xls'], key='lookup1')
    lookup2_file = st.sidebar.file_uploader("Lookup Table 2", type=['xlsx', 'xls'], key='lookup2')
    
    load_lookup = st.sidebar.checkbox("Load Lookup Tables", value=True)
    
    # Check if all required files are uploaded
    required_files_uploaded = all([crosses_file, parametric_file, package_file, recipe_file])
    
    # Main content area
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Required Files", "4/4" if required_files_uploaded else f"{sum([crosses_file is not None, parametric_file is not None, package_file is not None, recipe_file is not None])}/4")
    
    with col2:
        st.metric("Optional Files", f"{sum([lookup1_file is not None, lookup2_file is not None])}/2")
    
    with col3:
        st.metric("Ready to Process", "✅" if required_files_uploaded else "❌")
    
    st.markdown("---")
    
    # Process button
    if st.button("🚀 Start Processing", disabled=not required_files_uploaded, use_container_width=True):
        try:
            with st.spinner("Processing your files..."):
                # Initialize session state for results
                if 'results' not in st.session_state:
                    st.session_state.results = {}
                
                # Step 1: Process files
                st.info("📝 Step 1/5: Validating input files...")
                step1_result = process_files(crosses_file, parametric_file, package_file, recipe_file)
                st.session_state.results['step1'] = step1_result
                st.success("✅ Step 1 complete!")
                
                # Step 2: Merge data
                st.info("🔗 Step 2/5: Merging data from all sources...")
                target_files = [parametric_file, package_file]
                
                # Save step1 result to temp file for merging
                temp_step1 = BytesIO()
                step1_result.to_excel(temp_step1, index=False)
                temp_step1.seek(0)
                
                step2_result = merge_from_crossesparts(temp_step1, target_files)
                st.session_state.results['step2'] = step2_result
                st.success("✅ Step 2 complete!")
                
                # Step 3: Validate and grade
                st.info("✔️ Step 3/5: Validating features and assigning grades...")
                
                # Save step2 result to temp file
                temp_step2 = BytesIO()
                step2_result.to_excel(temp_step2, index=False)
                temp_step2.seek(0)
                
                loadLookUp = "YES" if load_lookup and lookup1_file and lookup2_file else "NO"
                step3_result = validate_core_upgrade_equal(
                    recipe_file, temp_step2, lookup1_file, lookup2_file, loadLookUp
                )
                st.session_state.results['step3'] = step3_result
                st.success("✅ Step 3 complete!")
                
                # Step 4: Merge all results
                st.info("📊 Step 4/5: Merging validation results...")
                
                # Save step3 result to temp file
                temp_step3 = BytesIO()
                step3_result.to_excel(temp_step3, index=False)
                temp_step3.seek(0)
                
                step4_result = merge_files(temp_step1, temp_step3, crosses_file)
                st.session_state.results['step4'] = step4_result
                st.success("✅ Step 4 complete!")
                
                # Step 5: Final organization
                st.info("🎯 Step 5/5: Organizing final report...")
                final_result = organize_file(step4_result)
                st.session_state.results['final'] = final_result
                st.success("✅ Step 5 complete!")
                
                st.balloons()
                st.success("🎉 All processing complete!")
        
        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
            st.exception(e)
    
    # Display results if available
    if 'results' in st.session_state and 'final' in st.session_state.results:
        st.markdown("---")
        st.header("📊 Results Summary")
        
        final_df = st.session_state.results['final']
        
        # Statistics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Records", len(final_df))
        
        with col2:
            cross_count = len(final_df[final_df['Status'] == 'Cross'])
            st.metric("Cross Matches", cross_count)
        
        with col3:
            if 'Grade' in final_df.columns:
                drop_in_a = len(final_df[final_df['Grade'] == 'Drop-in A'])
                st.metric("Drop-in A", drop_in_a)
        
        with col4:
            if 'Grade' in final_df.columns:
                not_drop_in = len(final_df[final_df['Grade'] == 'Not Drop-in'])
                st.metric("Not Drop-in", not_drop_in)
        
        # Grade distribution
        if 'Grade' in final_df.columns:
            st.subheader("Grade Distribution")
            grade_counts = final_df['Grade'].value_counts()
            st.bar_chart(grade_counts)
        
        # Preview results
        st.subheader("Preview Results")
        st.dataframe(final_df.head(20), use_container_width=True)
        
        # Download button
        st.subheader("📥 Download Results")
        
        col1, col2 = st.columns(2)
        
        with col1:
            final_excel = to_excel(final_df)
            st.download_button(
                label="⬇️ Download Final Report",
                data=final_excel,
                file_name="final_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col2:
            if 'step3' in st.session_state.results:
                validation_excel = to_excel(st.session_state.results['step3'])
                st.download_button(
                    label="⬇️ Download Validation Report",
                    data=validation_excel,
                    file_name="validation_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
    # Info section
    with st.expander("ℹ️ Processing Pipeline Information"):
        st.markdown("""
        ### Processing Steps:
        
        1. **Validation**: Compares PLc and PLx values, checks package/pinout similarity, and validates data availability
        2. **Data Merging**: Merges parametric data and package information for each part number from all sources
        3. **Feature Validation**: Validates core features, checks tolerances, and compares upgrade features
        4. **Results Merging**: Combines validation results with original cross-reference data
        5. **Final Organization**: Organizes the report and assigns final status to each cross-reference
        
        ### Grade Definitions:
        
        - **Drop-in A**: Perfect match within G1 tolerance
        - **Drop-in B**: Good match within G2 tolerance
        - **Drop-in C**: Acceptable match within G3 tolerance
        - **Drop-in D**: Match with some features outside tolerance
        - **Not Drop-in**: Core features don't match or critical differences exist
        
        ### Status Definitions:
        
        - **Cross**: Valid cross-reference with acceptable grade
        - **Not Cross**: Invalid cross-reference (core mismatch or tolerance failure)
        - **Different PLs**: Product lines don't match
        - **Not Found Data**: Required data not found in source files
        """)
if __name__ == "__main__":
    main()
