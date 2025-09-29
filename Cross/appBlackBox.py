import streamlit as st
import pandas as pd
import os
import tempfile
import io
from pathlib import Path

# Set page config
st.set_page_config(page_title="Parts Validation Pipeline", layout="wide")

# Title
st.title("Parts Cross-Validation and Upgrade Pipeline")
st.markdown("Upload input files in the sidebar and run the pipeline to generate validation reports.")

# Sidebar for file uploads
st.sidebar.header("Upload Input Files")
crosses_file = st.sidebar.file_uploader("Crosses File (e.g., crossTestWithoutALU.xlsx)", type=["xlsx"])
parametric_file = st.sidebar.file_uploader("Parametric Data File (multi-sheet, e.g., plmultitabs_input_fullWithoutALU.xlsx)", type=["xlsx"])
pakageAndPinout_file = st.sidebar.file_uploader("Package and Pinout File (e.g., qapackagepinoutexporter_input1_withoutALU.xlsx)", type=["xlsx"])
recipe_file = st.sidebar.file_uploader("Recipe File (e.g., newcrossexproductnewrecipeinputtemplate4-5f6cafa5-cab9-416a-95bb-50f012ab4b31.xlsx)", type=["xlsx"])
lookUpFile1 = st.sidebar.file_uploader("Lookup File 1 (e.g., lookupTableForSample.xlsx)", type=["xlsx"])
lookUpFile2 = st.sidebar.file_uploader("Lookup File 2 (e.g., lookupTableForSample2.xlsx)", type=["xlsx"])

# Load lookup option
loadLookUp = st.sidebar.selectbox("Load Lookup Files?", ["YES", "NO"])

# Temporary directory for file handling
@st.cache_data
def get_temp_dir():
    return tempfile.mkdtemp()

temp_dir = get_temp_dir()

# Function to save uploaded file to temp path
def save_uploaded_file(uploaded_file, temp_path):
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return temp_path

# Your original code starts here (unchanged, except for global handling)
# =============================================================================

# step 1
# validate input (LOG FILE )
def process_files(crossfile, parametricData, pakageAndPinout, recipe, output_path=None):
    # --- Load files ---
    cross_df = pd.read_excel(crossfile, dtype=str)
    recipe_df = pd.read_excel(recipe, dtype=str)
    pakage_df = pd.read_excel(pakageAndPinout, dtype=str)

    # parametricData has multiple sheets → read all
    parametric_sheets = pd.read_excel(parametricData, sheet_name=None, dtype=str)
    parametric_df = pd.concat(parametric_sheets.values(), ignore_index=True)

    # --- Step 1: Compare PLc and PLx ---
    cross_df["SamePLs"] = cross_df.apply(
        lambda r: "same" if str(r.get("PLc","")).strip() == str(r.get("PLx","")).strip() else "not",
        axis=1
    )

    # --- Step 2: Add IsPackaeSimilar, IsPinoutSimilar, PackaePinout ---
    cross_df["IsPackaeSimilar"] = "NA"
    cross_df["IsPinoutSimilar"] = "NA"
    cross_df["PackaePinout"] = "NA"

    for idx, row in cross_df.iterrows():
        if row["SamePLs"] == "same":
            plc = str(row["PLc"]).strip()
            # filter recipe by PLc
            subset = recipe_df[recipe_df["ZProductValue"].astype(str).str.strip() == plc]

            if not subset.empty:
                # Check IsPackaeSimilar
                is_pack = "FALSE" if (subset["IsPackaeSimilar"].str.upper() == "FALSE").all() else "TRUE"
                is_pin = "FALSE" if (subset["IsPinoutSimilar"].str.upper() == "FALSE").all() else "TRUE"

                cross_df.at[idx, "IsPackaeSimilar"] = is_pack
                cross_df.at[idx, "IsPinoutSimilar"] = is_pin
                cross_df.at[idx, "PackaePinout"] = "TRUE" if (is_pack == "TRUE" or is_pin == "TRUE") else "FALSE"
            else:
                cross_df.at[idx, "IsPackaeSimilar"] = "NA"
                cross_df.at[idx, "IsPinoutSimilar"] = "NA"
                cross_df.at[idx, "PackaePinout"] = "NA"

    # --- Step 3: Check FoundData ---
    cross_df["FoundData"] = "FALSE"

    for idx, row in cross_df.iterrows():
        part_c = str(row.get("PartNumberC","")).strip()
        part_x = str(row.get("PartNumberX","")).strip()

        if row["SamePLs"] == "same":
            if row["PackaePinout"] == "TRUE":
                # Check in parametricData OR pakageAndPinout
                exists_c = ((parametric_df["PartNumber"] == part_c).any() or (pakage_df["PartNumber"] == part_c).any())
                exists_x = ((parametric_df["PartNumber"] == part_x).any() or (pakage_df["PartNumber"] == part_x).any())
            elif row["PackaePinout"] == "FALSE":
                # Check only in parametricData
                exists_c = (parametric_df["PartNumber"] == part_c).any()
                exists_x = (parametric_df["PartNumber"] == part_x).any()
            else:
                exists_c = exists_x = False

            if exists_c and exists_x:
                cross_df.at[idx, "FoundData"] = "TRUE"

    # --- Save full cross_df with new columns ---
    if output_path:
        cross_df.to_excel(output_path, index=False)

    return cross_df

# step 2
# get file cotaion all data each part and his cross under it
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
    """
    Load all sheets from all Excel files once and return a cache.
    Returns: {filename: {sheetname: dataframe}}
    """
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
    """
    Search for the row in df that matches part and company.
    """
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


def get_single_row_all_files(file_cache, part_value, company_value, PackaePinout_val, output_path=None):
    """
    Merge one row from multiple Excel files into a single row (horizontal merge).
    """
    merged_parts = [pd.DataFrame({"PartNumber": [part_value], "Company": [company_value], "Comments": [""]})]
    PackaePinout_bool = str(PackaePinout_val).strip().lower() in ["true", "1", "yes"]

    file_list = list(file_cache.keys())
    if not PackaePinout_bool:
        file_list = file_list[:-1]  # remove last file if needed

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

    if output_path:
        final_row.to_excel(output_path, index=False)

    return final_row


def merge_from_crossesparts(crosses_file, target_files, output_path=None):
    """
    Reads part/company values from crosses_file and merges corresponding rows from target files.
    """
    df_crosses = pd.read_excel(crosses_file)

    # Check required columns
    required_cols = ["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex", "PackaePinout"]
    for col in required_cols:
        if col not in df_crosses.columns:
            raise ValueError(f"❌ Missing required column: {col}")

    file_cache = load_all_excel_sheets(target_files)
    all_merged_rows = []

    for _, row in df_crosses.iterrows():
        part_val_c = str(row["PartNumberC"]).strip()
        company_val_c = str(row["CompanyNameC"]).strip()
        part_val_x = str(row["PartNumberX"]).strip()
        company_val_x = str(row["CompanyNamex"]).strip()
        PackaePinout_val = str(row["PackaePinout"]).strip()

        st.write(f"➡️ Merging {part_val_c} from {company_val_c}")
        merged_c = get_single_row_all_files(file_cache, part_val_c, company_val_c, PackaePinout_val)

        st.write(f"➡️ Merging {part_val_x} from {company_val_x}")
        merged_x = get_single_row_all_files(file_cache, part_val_x, company_val_x, PackaePinout_val)

        # Append both parts (C and X)
        all_merged_rows.extend([merged_c, merged_x])

    final_df = pd.concat(all_merged_rows, ignore_index=True)

    if output_path:
        final_df.to_excel(output_path, index=False)

    return final_df

#test
#new
def clean_to_floats(lst):
    cleaned = []
    for x in lst:
        try:
            if pd.notna(x) and str(x).strip().lower() != "nan" and str(x).strip() != "":
                cleaned.append(float(x))
        except ValueError:
            pass
    return cleaned

def percent_diff(a, b):
    if a == 0 or b == 0:  # avoid division by zero
        return float('inf') if a != b else 0.0
    return round(abs(a - b) / min(abs(a), abs(b)) * 100, 2)

def percent_diff_base_a(a, b):
    if a == 0:
        return float('inf') if b != 0 else 0.0
    return round(abs(a - b) / abs(a) * 100, 2)

def compare_parts(df, value, value2):
    """
    Compare two parts in dataframe by AcceptedValue.
    Returns values/units sorted by FeatureCode order.
    """

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
    #need to filter by one value at column ValueID

    # Step 2: Take one ValueID from part1 (first row)
    chosen_value_id = part1.loc[0, "ValueID"]
    chosen_value_id2 = part2.loc[0, "ValueID"]
    # print(f"chosen_value_id {chosen_value_id}")
    # print(f"chosen_value_id2 {chosen_value_id2}")
    # print(f"part1 {part1}")
    # print(f"part2 {part2}")
    # Step 3: Filter both parts with the same ValueID
    part1 = part1[part1["ValueID"] == chosen_value_id].reset_index(drop=True)
    part2 = part2[part2["ValueID"] == chosen_value_id2].reset_index(drop=True)

    # print("*"*50)
    # print(f"part1 {part1}")
    # print(f"part2 {part2}")
    # print("*"*50)
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

    # Sort FeatureCodes (V before U, numeric order)
    def fc_sort_key(fc):
        match = re.search(r"([VU])(\d+)$", fc)
        if match:
            t, num = match.groups()
            return int(num) * 2 + (0 if t == "V" else 1)
        return float("inf")  # push unexpected codes last

    part1_sorted = part1.sort_values(by="FeatureCode", key=lambda x: x.map(fc_sort_key))
    part2_sorted = part2.sort_values(by="FeatureCode", key=lambda x: x.map(fc_sort_key))

    nValue  = part1_sorted["NormalizedValue"].dropna().tolist()
    nUnit   = part1_sorted["Unit"].dropna().tolist()
    nValue2 = part2_sorted["NormalizedValue"].dropna().tolist()
    nUnit2  = part2_sorted["Unit"].dropna().tolist()

    return {
        "value": value,
        "nValue": nValue,
        "nUnit": nUnit,
        "FeatureCode": "Match",
        "state": "Match",
        "value2": value2,
        "nValue2": nValue2,
        "nUnit2": nUnit2
    }

# Global for values (handled in session state if needed)
values = None

def validate_core_upgrade_equal(newcross_file, finalmerged_file, lookup_file1, lookup_file2, loadLookUp, output_path=None):
    global values
    # --- Load rules (newCross) ---
    rules_df = pd.read_excel(newcross_file, dtype=str)
    # --- Load data (final_merged) ---
    data = pd.read_excel(finalmerged_file, dtype=str)
    # # use lookup normalization
    if loadLookUp=="YES":
        st.info("Reading lookup files...")
        values1 = pd.read_excel(lookup_file1, dtype=str)
        values2 = pd.read_excel(lookup_file2, dtype=str)
        values = pd.concat([values1, values2], ignore_index=True)
        st.success("Lookup files loaded.")

    if "ProductName" not in data.columns:
        raise ValueError("❌ 'ProductName' column not found in final_merged.xlsx")

    # ✅ take PL list from final_merged
    PLs = data["ProductName"].astype(str).unique().tolist() # Get unique PLs
    st.write(f"Rules DF: {rules_df.shape}")
    rules_df = rules_df[rules_df["ZProductValue"].astype(str).isin(PLs)]
    st.write(f"Rules after filtering: {rules_df.shape}")
    st.write(f"PLs: {PLs}")

    # # ✅ Keep only rules of interest
    # rules_df = rules_df[
    #     (rules_df["FeaturesType"].str.lower() == "core") |
    #     (rules_df["UpgradeFeature"].str.upper() == "TRUE") |
    #     (rules_df["IsEqualFeature"].str.upper() == "TRUE") |
    #     (rules_df["FeaturesType

        # ✅ Keep only rules of interest
    rules_df = rules_df[
        (rules_df["FeaturesType"].str.lower() == "core") |
        (rules_df["UpgradeFeature"].str.upper() == "TRUE") |
        (rules_df["IsEqualFeature"].str.upper() == "TRUE") |
        (rules_df["FeaturesType"].str.lower() == "tolerance")
    ]

    # --- Ensure numeric tolerance values ---
    for col in ["G1", "G2", "G3"]:
        if col in rules_df.columns:
            rules_df[col] = pd.to_numeric(rules_df[col], errors="coerce")

    results = []

    # Group data by ProductName to process each PL efficiently
    grouped_data = data.groupby("ProductName")
    PinOut_flag="FALSE"
    for PL, data_subset in grouped_data:
        st.write(f"Processing PL: {PL}")
        data_subset = data_subset.reset_index(drop=True)
        # Filter rules for the current PL
        rules_subset = rules_df[rules_df["ZProductValue"].astype(str).str.strip() == PL].reset_index(drop=True)
        st.write(f"Rules for PL {PL}: {len(rules_subset)} rules")
        # print(f"rules_df[\"IsPinoutSimilar\"].iloc[0]{rules_df["IsPinoutSimilar"].iloc[0]}")
        # print(f"PinOut_flag {PinOut_flag}******************************************")
        if len(rules_df) > 0 and rules_df["IsPinoutSimilar"].iloc[0] == "TRUE":  # Adjusted for iloc[0] safety
            # print(f"v1 {row1['NormalizedPinName']}, v2 {row2['NormalizedPinName']}")
            PinOut_flag="TRUE"

        for i in range(0, len(data_subset), 2):
            if i + 1 >= len(data_subset):
                break

            row1, row2 = data_subset.iloc[i], data_subset.iloc[i+1]
            feature_statuses, feature_grades, flags = [], [], []
            count=1
            for _, rule in rules_subset.iterrows(): # Iterate over the filtered rules
                feature = rule["Features"]
                ftype = str(rule["FeaturesType"]).lower()
                upgrade_flag = str(rule.get("UpgradeFeature", "")).upper()
                equal_flag = str(rule.get("IsEqualFeature", "")).upper()
                # PinOut_flag = str(rule.get("IsPinoutSimilar", "")).upper()
                # print(f"feature {feature}")
                # breakpoint()
                # -------------------------
                # ✅ Core / Upgrade / Equal check
                # -------------------------
                # print(f"PL is {PL}\n")
                # print(f"PinOut_flag  {PinOut_flag} count {count} ")
                if PinOut_flag == "TRUE" and count:
                    # print(f"v1 {row1['NormalizedPinName']}, v2 {row2['NormalizedPinName']}")
                    val1, val2 = str(row1['NormalizedPinName']).strip(), str(row2['NormalizedPinName']).strip()
                    # print(f"value1 is {val1} ,value2 is {val2}")
                    count=0
                    if val1 == val2:
                        status, grade = f"✅ PinOut Match ({feature} == ==)", "A"
                    else:
                        status, grade = f"❌ Diffrent pinOuT ({feature}: ", "DiffPinOut"
                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)
                    status, grade="",""
                if ftype == "core"  or equal_flag == "TRUE":
                    if feature not in data_subset.columns: # Check in the subset
                        status, grade = f"⚠ Missing column {feature}", "Fail"
                    else:
                        val1, val2 = str(row1[feature]).strip(), str(row2[feature]).strip()
                        if val1 == val2:
                            status, grade = f"✅ Match ({feature})  value1 {val1}  value2 {val2}", "A"
                        else:
                            status, grade = f"❌ Mismatch ({feature}: {val1} vs {val2})", "Fail"
                            if ftype == "core":
                                flags.append("missMatchAtCore ({feature}: {val1} vs {val2})")
                                grade="FailInCore"
                                # Removed redundant flag assignment

                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)

                elif upgrade_flag == "TRUE":
                    continue
                elif ftype == "tolerance":
                    if feature not in data_subset.columns: # Check in the subset
                        status, grade = f"⚠ Missing column {feature}", "Fail"
                    else:
                        val1, val2 = str(row1[feature]).strip(), str(row2[feature]).strip()
                        if val1 == val2:
                            status, grade = f"✅ Match ({feature})*", "A"
                        else:
                          result = compare_parts(values, val1, val2)
                          if result['nUnit'] != result['nUnit2']:
                            st.warning(f"Difference in Units for {feature}: {result['state']}, values ({val1},{val2}), Units ({result['nUnit']},{result['nUnit2']})")
                            status, grade = f"Difference Unit {result['state']} , value({val1},{val2}) Units({result['nUnit']},{result['nUnit2']})", "unitFail"
                            flags.append(f"Difference Unit  {feature} ({val1},{val2}) Units({result['nUnit']},{result['nUnit2']})")

                          elif result['state']=="Match":
                            # st.write(f"result['nValue']  {result['nValue']}")
                            # st.write(f"len(result['nValue'])  {len(result['nValue'])}")

                            # for j in range(len(result['nValue'])): # Changed variable name to j
                            for j in range(min(len(result['nValue']), len(result['nValue2']))):
                              # print(f"result['nValue'][{i}]  {result['nValue'][i]}, result['nValue2'][{i}]  , {result['nValue2'][i]}")
                              v1, v2 = float(result['nValue'][j]), float(result['nValue2'][j]) # Changed variable name to j
                              diff = percent_diff_base_a(v1,v2)
                              g1, g2, g3 = rule.get("G1"), rule.get("G2"), rule.get("G3")
                              if diff < g1:
                                  status, grade = f"✅ Within {g1}% G1  diff is {diff} tolerance ({v1} vs {v2})", "A"
                              elif diff < g2:
                                  status, grade = f"✅ Within {g2}% G2 diff is {diff} tolerance ({v1} vs {v2})", "B"
                              elif diff < g3 + 10:
                                  status, grade = f"✅ Within {g3}% G3 diff is {diff} tolerance ({v1} vs {v2})", "C"
                              else:
                                  status, grade = f"❌ Outside tolerance ({v1} vs {v2}) diff={diff}", "Fail"
                                  flags.append(f"miss Match value5 At Feature  {feature}({v1},{v2}) diff={diff}")
                          elif result['state']=="Different DetailedValueType":
                            status, grade = f" Different DetailedValueType column {result['state']} , value({val1},{val2})", "FailInDetailedValueType"
                            flags.append(f"Different DetailedValueType At Feature  {feature} ({val1},{val2})")
                          elif result['state']=="Different FeatureCode":
                            status, grade = f"Different FeatureCode column {result['state']} , value({val1},{val2})", "FailInFeatureCode"
                            flags.append(f"Different FeatureCode At Feature  {feature} ({val1},{val2})")
                          else:

                            status, grade = f" Missing Values column {result['state']} , value({val1},{val2})", "Fail"


                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)

            # -------------------------
            # ✅ Overall grade
            # -------------------------

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
                "Part Number C": row2.get("PartNumber",""),
                "Company Name C": row2.get("Company",""),
                "Part Number X": row1.get("PartNumber",""),
                "Company Name X": row1.get("Company",""),
                "PL Name": row1.get("PL Name",""),
                "Feature": "OVERALL",
                "Status": " | ".join(feature_statuses),
                "flags": " | ".join(flags) if flags else "",
                "Grade": overall_grade
            })

    # Build final report
    df_report = pd.DataFrame(results)
    for res in results:
        st.write(res["Status"])
        st.write(res["flags"])
        st.write(res["Grade"])
    if output_path:
        df_report.to_excel(output_path, index=False)

    return df_report

def merge_files(file1, file2, file3, output_path=None):
    """
    Merge File1 + File2 into File3.

    File1 columns: [PartNumberC, CompanyNameC, PLc, PartNumberX, CompanyNamex, PLx,
                    CrossGrade, SamePLs, IsPackaeSimilar, IsPinoutSimilar, PackaePinout, FoundData]
    File2 columns: [Part Number C, Company Name C, Part Number X, Company Name X,
                    PL Name, Feature, Status, flags, Grade]
    File3 is the base input.

    File4 (output) should look like File3 +
        - From File1: [SamePLs, IsPackaeSimilar, IsPinoutSimilar, PackaePinout, FoundData]
        - From File2: [Status, flags, Grade]
    """

    # Load
    f1 = pd.read_excel(file1, dtype=str)
    f2 = pd.read_excel(file2, dtype=str)
    f3 = pd.read_excel(file3, dtype=str)

    # Strip spaces in column names
    f1.columns = f1.columns.str.strip()
    f2.columns = f2.columns.str.strip()
    f3.columns = f3.columns.str.strip()

    # Align file2 keys with file1/file3
    f2 = f2.rename(columns={
        "Part Number C": "PartNumberC",
        "Company Name C": "CompanyNameC",
        "Part Number X": "PartNumberX",
        "Company Name X": "CompanyNamex"
    })

    # Align file3 keys (if needed)
    f3 = f3.rename(columns=lambda x: x.replace(" ", ""))

    # --- Merge File3 with File1 (add 5 cols) ---
    merged = pd.merge(
        f3,
        f1[[
            "PartNumberX","CompanyNamex","PartNumberC","CompanyNameC",
            "SamePLs","IsPackaeSimilar","IsPinoutSimilar","PackaePinout","FoundData"
        ]],
        on=["PartNumberC","CompanyNameC","PartNumberX","CompanyNamex"],
        how="left"
    )

    # --- Merge result with File2 (add 3 cols) ---
    merged = pd.merge(
        merged,
        f2[["PartNumberC","CompanyNameC","PartNumberX","CompanyNamex","Status","flags","Grade"]],
        on=["PartNumberC","CompanyNameC","PartNumberX","CompanyNamex"],
        how="left"
    )

    # Also try swapped orientation for File2
    swapped = f2.rename(columns={
        "PartNumberC":"PartNumberX",
        "CompanyNameC":"CompanyNamex",
        "PartNumberX":"PartNumberC",
        "CompanyNamex":"CompanyNameC"
    })

    merged = pd.merge(
        merged,
        swapped[["PartNumberC","CompanyNameC","PartNumberX","CompanyNamex","Status","flags","Grade"]],
        on=["PartNumberC","CompanyNameC","PartNumberX","CompanyNamex"],
        how="left",
        suffixes=("","_swapped")
    )

    # Fill blanks from swapped
    for col in ["Status","flags","Grade"]:
        merged[col] = merged[col].fillna(merged[f"{col}_swapped"])
        merged.drop(columns=[f"{col}_swapped"], inplace=True)

    # Save if requested
    if output_path:
        merged.to_excel(output_path, index=False)

    return merged

def organize_file(input_path, output_path=None):
    """
    Post-process merged file:
    - Rename columns
    - Add derived Status column
    """
    df = pd.read_excel(input_path, dtype=str)

    # --- Clean column names ---
    df.columns = df.columns.str.strip()

    # --- Rename columns ---
    rename_map = {
        "Status": "Match Feature",
        "flags": "Different Features"
    }
    df = df.rename(columns=rename_map)

    # --- Add new Status column ---
    def compute_status(row):
        grade = str(row.get("Grade", "")).strip()
        same_pls = str(row.get("SamePLs", "")).strip().lower()
        found_data = str(row.get("FoundData", "")).strip().upper()
        pack_pinout = str(row.get("PackaePinout", "")).strip()

        if grade in ["Drop-in D", "Drop-in A", "Drop-in B", "Drop-in C"]:
            return "Cross"
        elif grade in ["Not Drop-in","Detailed Value Type Fail Not Drop-in","Unit FAIL Not Drop-in","FeatureCode FAIL Not Drop-in"] and found_data == "TRUE":
            return "Not Cross"
        elif grade == "Detailed Value Type Fail Not Drop-in" and found_data == "TRUE":
            return "Not Cross"
        elif same_pls == "not":
            return "Different PLs"
        elif found_data == "FALSE" and pack_pinout == "":
            return "Not Found Data"
        else:
            return ""

    df["Status"] = df.apply(compute_status, axis=1)

    # Save if requested
    if output_path:
        df.to_excel(output_path, index=False)

    return df

# =============================================================================
# Streamlit Main Logic (Execution Pipeline)
# =============================================================================

# Initialize session state for storing generated files
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = {}

# Check if all required files are uploaded
required_files = [crosses_file, parametric_file, pakageAndPinout_file, recipe_file]
if lookUp == "YES":
    required_files.extend([lookUpFile1, lookUpFile2])

if all(f is not None for f in required_files):
    st.sidebar.success("All files uploaded! Ready to run.")

    # Button to run full pipeline
    if st.sidebar.button("Run Full Pipeline", type="primary"):
        with st.spinner("Running full pipeline..."):
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                # Step 1: Save uploaded files to temp paths
                temp_crosses = save_uploaded_file(crosses_file, os.path.join(temp_dir, "crosses.xlsx"))
                temp_parametric = save_uploaded_file(parametric_file, os.path.join(temp_dir, "parametric.xlsx"))
                temp_pakage = save_uploaded_file(pakageAndPinout_file, os.path.join(temp_dir, "pakage.xlsx"))
                temp_recipe = save_uploaded_file(recipe_file, os.path.join(temp_dir, "recipe.xlsx"))

                temp_lookup1 = None
                temp_lookup2 = None
                if loadLookUp == "YES":
                    temp_lookup1 = save_uploaded_file(lookUpFile1, os.path.join(temp_dir, "lookup1.xlsx"))
                    temp_lookup2 = save_uploaded_file(lookUpFile2, os.path.join(temp_dir, "lookup2.xlsx"))

                # Output paths (temp)
                output_file_1 = os.path.join(temp_dir, "crossTestWithoutALU_1_with_results.xlsx")
                output_file_merged = os.path.join(temp_dir, "final_merged_fullWithoutALU.xlsx")
                output_file_2 = os.path.join(temp_dir, "validation_reportWithoutALU.xlsx")
                final_output = os.path.join(temp_dir, "FINAL.xlsx")
                organized_output = os.path.join(temp_dir, "FINAL_ORGANIZED_fullWithoutALU.xlsx")

                # Step 1: process_files
                status_text.text("Step 1: Processing files and validating inputs...")
                result1 = process_files(temp_crosses, temp_parametric, temp_pakage, temp_recipe, output_path=output_file_1)
                st.session_state.generated_files['step1'] = output_file_1
                progress_bar.progress(20)
                st.success("Step 1 completed!")

                # # Step 2: merge_from_crossesparts
                # status_text.text("Step 2: Merging data from crosses...")
                # target_files = [temp_parametric, temp_pakage]
                # result2 = merge_from_crosses
                              # Step 2: merge_from_crossesparts
                status_text.text("Step 2: Merging data from crosses...")
                target_files = [temp_parametric, temp_pakage]
                result2 = merge_from_crossesparts(output_file_1, target_files, output_path=output_file_merged)
                st.session_state.generated_files['step2'] = output_file_merged
                progress_bar.progress(40)
                st.success("Step 2 completed!")

                # Step 3: validate_core_upgrade_equal
                status_text.text("Step 3: Validating core upgrades and tolerances...")
                if loadLookUp == "YES":
                    result3 = validate_core_upgrade_equal(temp_recipe, output_file_merged, temp_lookup1, temp_lookup2, loadLookUp, output_path=output_file_2)
                else:
                    # If no lookup, pass None (adjust function if needed; assuming it handles gracefully)
                    result3 = validate_core_upgrade_equal(temp_recipe, output_file_merged, None, None, loadLookUp, output_path=output_file_2)
                st.session_state.generated_files['step3'] = output_file_2
                progress_bar.progress(60)
                st.success("Step 3 completed!")

                # Step 4: merge_files
                status_text.text("Step 4: Merging all results...")
                result4 = merge_files(output_file_1, output_file_2, temp_crosses, output_path=final_output)
                st.session_state.generated_files['step4'] = final_output
                progress_bar.progress(80)
                st.success("Step 4 completed!")

                # Step 5: organize_file
                status_text.text("Step 5: Organizing final output...")
                result5 = organize_file(final_output, organized_output)
                st.session_state.generated_files['step5'] = organized_output
                progress_bar.progress(100)
                st.success("Full pipeline completed! Check downloads below.")

                # Display sample results
                st.subheader("Sample Results")
                if os.path.exists(output_file_merged):
                    st.write("**Merged Data (Step 2) Head:**")
                    df_merged = pd.read_excel(output_file_merged)
                    st.dataframe(df_merged.head())

                if os.path.exists(output_file_2):
                    st.write("**Validation Report (Step 3) Head:**")
                    df_report = pd.read_excel(output_file_2)
                    st.dataframe(df_report.head())

                if os.path.exists(organized_output):
                    st.write("**Final Organized Output (Step 5) Head:**")
                    df_final = pd.read_excel(organized_output)
                    st.dataframe(df_final.head())

            except Exception as e:
                st.error(f"❌ Pipeline failed: {str(e)}")
                st.exception(e)
                progress_bar.empty()
                status_text.empty()

        progress_bar.empty()
        status_text.empty()

    # Download Section
    st.sidebar.header("Download Generated Files")
    if st.session_state.generated_files:
        for step, file_path in st.session_state.generated_files.items():
            if os.path.exists(file_path):
                with open(file_path, "rb") as f:
                    st.sidebar.download_button(
                        label=f"Download {step.replace('step', 'Output ')} ({os.path.basename(file_path)})",
                        data=f.read(),
                        file_name=os.path.basename(file_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.sidebar.warning(f"File for {step} not generated yet.")

else:
    st.warning("⚠️ Please upload all required files in the sidebar to run the pipeline.")
    if lookUp == "YES" and (lookUpFile1 is None or lookUpFile2 is None):
        st.info("Note: Lookup files are required if 'Load Lookup Files?' is set to 'YES'.")

# Optional: Individual Step Buttons (for debugging)
st.sidebar.header("Run Individual Steps (Advanced)")
if all(f is not None for f in [crosses_file, parametric_file, pakageAndPinout_file, recipe_file]):
    if st.sidebar.button("Run Step 1 Only (Process Files)"):
        with st.spinner("Running Step 1..."):
            temp_crosses = save_uploaded_file(crosses_file, os.path.join(temp_dir, "crosses.xlsx"))
            temp_parametric = save_uploaded_file(parametric_file, os.path.join(temp_dir, "parametric.xlsx"))
            temp_pakage = save_uploaded_file(pakageAndPinout_file, os.path.join(temp_dir, "pakage.xlsx"))
            temp_recipe = save_uploaded_file(recipe_file, os.path.join(temp_dir, "recipe.xlsx"))
            output_file_1 = os.path.join(temp_dir, "crossTestWithoutALU_1_with_results.xlsx")
            result1 = process_files(temp_crosses, temp_parametric, temp_pakage, temp_recipe, output_path=output_file_1)
            st.session_state.generated_files['step1'] = output_file_1
            st.success("Step 1 completed! Download available in sidebar.")
            st.dataframe(result1.head())

    # Add similar buttons for other steps if needed (omitted for brevity; follow the same pattern)

# Footer
st.markdown("---")
st.markdown("**Notes:**")
st.markdown("- This app processes parts validation data using your provided pipeline.")
st.markdown("- Outputs are generated as Excel files for download.")
st.markdown("- If errors occur (e.g., missing columns), check the console or file formats.")
st.markdown("- For large files, processing may take time—progress is shown above.")
