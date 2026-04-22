import io
import re
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# ============================================================================
# TIMING UTILITIES
# ============================================================================

LOOKUP_CACHE: Dict[Tuple[str, str, Tuple[str, str]], str] = {}


def _get_timing_log() -> List[Dict[str, str]]:
    if "timing_log" not in st.session_state:
        st.session_state["timing_log"] = []
    return st.session_state["timing_log"]


def reset_runtime_state() -> None:
    LOOKUP_CACHE.clear()
    st.session_state["timing_log"] = []


def log_timing(func_name: str, elapsed: float) -> Dict[str, str]:
    """Store timing info for display in Streamlit."""
    record = {
        "function": func_name,
        "elapsed_sec": round(float(elapsed), 4),
        "timestamp": str(datetime.now()),
    }
    _get_timing_log().append(record)
    return record


def get_timing_df() -> pd.DataFrame:
    """Return timing data as DataFrame for Streamlit display."""
    timing_log = _get_timing_log()
    if not timing_log:
        return pd.DataFrame()
    return pd.DataFrame(timing_log)


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================


def safe_int(value):
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return None



def clean_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()



def truncate_sheet_name(pl_name):
    pl_name = str(pl_name)
    return pl_name[:31] if len(pl_name) > 31 else pl_name



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



def percent_diff_base_a(a, b):
    if a == 0:
        return float("inf") if b != 0 else 0.0
    return round(abs(a - b) / abs(a) * 100, 2)



def fc_sort_key(fc):
    match = re.search(r"([VU])(\d+)\$", str(fc))
    if match:
        t, num = match.groups()
        return int(num) * 2 + (0 if t == "V" else 1)
    return float("inf")



def ensure_columns(df: pd.DataFrame, required_cols: List[str], df_name: str) -> None:
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in {df_name}: {missing}")



def normalize_boolish(value: str) -> str:
    value = clean_str(value).lower()
    if value in {"true", "1", "yes", "y"}:
        return "TRUE"
    if value in {"false", "0", "no", "n"}:
        return "FALSE"
    return value.upper()



def build_lookup_index(df: pd.DataFrame) -> Dict[str, List[Tuple[str, str]]]:
    """
    Build a feature -> [(real_group_string_lower, parent_group_id_str)] index.
    This preserves the original 'contains' behavior more accurately than exact-value indexing.
    """
    ensure_columns(
        df,
        ["FeatureName", "RealGroupString", "ParentGroupID"],
        "lookUp Core file",
    )

    index: Dict[str, List[Tuple[str, str]]] = {}
    for _, row in df.iterrows():
        feature = clean_str(row.get("FeatureName", "")).lower()
        real_group = clean_str(row.get("RealGroupString", "")).lower()
        parent = clean_str(row.get("ParentGroupID", ""))
        index.setdefault(feature, []).append((real_group, parent))
    return index



def find_feature_CoreLookUp(PL_name, df, feature1, value1, value2):
    t0 = time.perf_counter()
    key = (
        clean_str(PL_name),
        clean_str(feature1).lower(),
        tuple(sorted([clean_str(value1).lower(), clean_str(value2).lower()])),
    )

    if key in LOOKUP_CACHE:
        return LOOKUP_CACHE[key]

    required_cols = [
        "Product",
        "GroupID",
        "ParentGroupID",
        "RealGroupString",
        "GroupType",
        "ModifiedDate",
        "FeatureName",
    ]
    ensure_columns(df, required_cols, "lookUp Core file")

    feature1 = clean_str(feature1)
    value1 = clean_str(value1)
    value2 = clean_str(value2)

    if not feature1 or not value1 or not value2:
        LOOKUP_CACHE[key] = "Not Found"
        return "Not Found"

    core_df = df[df["FeatureName"].astype(str).str.contains(feature1, case=False, na=False)]
    has_value1 = core_df[core_df["RealGroupString"].astype(str).str.contains(value1, case=False, na=False)]
    has_value2 = core_df[core_df["RealGroupString"].astype(str).str.contains(value2, case=False, na=False)]

    parent_ids_with_both = set(has_value1["ParentGroupID"].astype(str)).intersection(
        set(has_value2["ParentGroupID"].astype(str))
    )
    result = "Found" if parent_ids_with_both else "Not Found"

    LOOKUP_CACHE[key] = result
    log_timing("find_feature_CoreLookUp", time.perf_counter() - t0)
    return result



def find_feature_CoreLookUp_fast(index, PL_name, feature1, value1, value2):
    key = (
        clean_str(PL_name),
        clean_str(feature1).lower(),
        tuple(sorted([clean_str(value1).lower(), clean_str(value2).lower()])),
    )
    if key in LOOKUP_CACHE:
        return LOOKUP_CACHE[key]

    f = clean_str(feature1).lower()
    v1 = clean_str(value1).lower()
    v2 = clean_str(value2).lower()

    if not f or not v1 or not v2 or v1 == "nan" or v2 == "nan":
        LOOKUP_CACHE[key] = "Not Found"
        return "Not Found"

    rows = index.get(f, [])
    parents1 = {parent for real_group, parent in rows if v1 in real_group}
    parents2 = {parent for real_group, parent in rows if v2 in real_group}

    result = "Found" if parents1.intersection(parents2) else "Not Found"
    LOOKUP_CACHE[key] = result
    return result


# ============================================================================
# FILE LOADING
# ============================================================================


@st.cache_data(show_spinner=False)
def read_excel_bytes(file_bytes: bytes, sheet_name=0, keep_default_na: bool = False) -> pd.DataFrame:
    return pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        dtype=str,
        keep_default_na=keep_default_na,
        na_values=[],
    )


@st.cache_data(show_spinner=False)
def load_files_from_excel(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    """Load required sheets from the uploaded combined Excel file."""
    t0 = time.perf_counter()

    xl_file = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xl_file.sheet_names

    files_data: Dict[str, pd.DataFrame] = {}

    if "lookUp" in sheet_names:
        lookup_df = pd.read_excel(xl_file, sheet_name="lookUp", dtype=str)
        files_data["lookUpFile1"] = lookup_df.copy()
        files_data["lookUpFile2"] = lookup_df.copy()
    else:
        raise ValueError("'lookUp' sheet not found in Combined Data file. LookUp data is required.")

    data_sheet_name = "Data"
    if data_sheet_name in sheet_names:
        data_df = pd.read_excel(
            xl_file,
            sheet_name=data_sheet_name,
            dtype=str,
            keep_default_na=False,
            na_values=[],
        )
        files_data["parametric_file"] = data_df.copy()
        files_data["pakageAndPinout_file"] = data_df.copy()
        files_data["qualification_file"] = data_df.copy()
    else:
        raise ValueError(f"'{data_sheet_name}' sheet not found in Combined Data file.")

    log_timing("load_files_from_excel", time.perf_counter() - t0)
    return files_data


# ============================================================================
# CORE PROCESSING FUNCTIONS
# ============================================================================



def process_single_pl(pl_name, cross_df_pl, parametric_df, pakage_df, recipe_df):
    t0 = time.perf_counter()

    print(f"\n{'=' * 60}")
    print(f"Processing PL: {pl_name}")
    print(f"{'=' * 60}")

    _ = truncate_sheet_name(pl_name)

    for col in ["PLc", "PLx"]:
        if col in cross_df_pl.columns:
            cross_df_pl[col] = cross_df_pl[col].astype(str).str.strip()

    cross_df_pl["SamePLs"] = np.where(cross_df_pl["PLc"] == cross_df_pl["PLx"], "same", "not")

    cross_df_pl["IsPackaeSimilar"] = "NA"
    cross_df_pl["IsPinoutSimilar"] = "NA"
    cross_df_pl["PackaePinout"] = "NA"

    same_pl_mask = cross_df_pl["SamePLs"] == "same"
    same_pl_indices = cross_df_pl[same_pl_mask].index

    ensure_columns(recipe_df, ["ZProductValue"], "recipe file")
    recipe_pl = recipe_df[recipe_df["ZProductValue"].astype(str).str.strip() == pl_name]

    for idx in same_pl_indices:
        plc = cross_df_pl.at[idx, "PLc"]
        subset = recipe_pl[recipe_pl["ZProductValue"].astype(str).str.strip() == clean_str(plc)]

        if not subset.empty:
            is_pack = "FALSE" if (subset.get("IsPackaeSimilar", pd.Series(dtype=str)).astype(str).str.upper() == "FALSE").all() else "TRUE"
            is_pin = "FALSE" if (subset.get("IsPinoutSimilar", pd.Series(dtype=str)).astype(str).str.upper() == "FALSE").all() else "TRUE"

            cross_df_pl.at[idx, "IsPackaeSimilar"] = is_pack
            cross_df_pl.at[idx, "IsPinoutSimilar"] = is_pin
            cross_df_pl.at[idx, "PackaePinout"] = "TRUE" if (is_pack == "TRUE" or is_pin == "TRUE") else "FALSE"

    parametric_parts = (
        set(parametric_df["PartNumber"].astype(str).str.strip().values)
        if not parametric_df.empty and "PartNumber" in parametric_df.columns
        else set()
    )
    pakage_parts = (
        set(pakage_df["PartNumber"].astype(str).str.strip().values)
        if not pakage_df.empty and "PartNumber" in pakage_df.columns
        else set()
    )

    cross_df_pl["FoundData"] = "FALSE"

    for idx in same_pl_indices:
        part_c = clean_str(cross_df_pl.at[idx, "PartNumberC"])
        part_x = clean_str(cross_df_pl.at[idx, "PartNumberX"])
        pack_pinout = clean_str(cross_df_pl.at[idx, "PackaePinout"])

        if pack_pinout == "TRUE":
            exists_c = part_c in parametric_parts or part_c in pakage_parts
            exists_x = part_x in parametric_parts or part_x in pakage_parts
        elif pack_pinout == "FALSE":
            exists_c = part_c in parametric_parts
            exists_x = part_x in parametric_parts
        else:
            continue

        if exists_c and exists_x:
            cross_df_pl.at[idx, "FoundData"] = "TRUE"

    log_timing(f"process_single_pl [{pl_name}]", time.perf_counter() - t0)
    return cross_df_pl



def merge_single_pl(pl_name, cross_df_pl, parametric_df, pakage_df, qualification_df):
    t0 = time.perf_counter()

    all_merged_rows = []

    for _, row in cross_df_pl.iterrows():
        part_c = clean_str(row.get("PartNumberC", ""))
        company_c = clean_str(row.get("CompanyNameC", ""))
        part_x = clean_str(row.get("PartNumberX", ""))
        company_x = clean_str(row.get("CompanyNamex", ""))
        pack_pinout = clean_str(row.get("PackaePinout", ""))

        packae_pinout_bool = pack_pinout.lower() in ["true", "1", "yes"]

        merged_c = pd.DataFrame({"PartNumber": [part_c], "Company": [company_c], "Comments": [""]})

        if not parametric_df.empty:
            row_data = find_matching_row(parametric_df, part_c, company_c)
            if row_data is not None:
                row_data.columns = deduplicate_columns(row_data.columns)
                merged_c = pd.concat([merged_c, row_data], axis=1)
            elif packae_pinout_bool:
                merged_c.loc[0, "Comments"] += "Not found in parametric; "

        if packae_pinout_bool:
            row_data = find_matching_row(pakage_df, part_c, company_c)
            if row_data is not None:
                row_data.columns = deduplicate_columns(row_data.columns)
                merged_c = pd.concat([merged_c, row_data], axis=1)
            else:
                merged_c.loc[0, "Comments"] += "Not found in package/pinout; "

        if not qualification_df.empty:
            row_data = find_matching_row(qualification_df, part_c, company_c)
            if row_data is not None:
                col_match = [c for c in row_data.columns if c.strip().lower() == "automotive"]
                merged_c["Automotive Qualified"] = row_data[col_match[0]].iloc[0] if col_match else ""
            else:
                merged_c["Automotive Qualified"] = ""
        else:
            merged_c["Automotive Qualified"] = ""

        merged_c.columns = deduplicate_columns(merged_c.columns)

        merged_x = pd.DataFrame({"PartNumber": [part_x], "Company": [company_x], "Comments": [""]})

        if not parametric_df.empty:
            row_data = find_matching_row(parametric_df, part_x, company_x)
            if row_data is not None:
                row_data.columns = deduplicate_columns(row_data.columns)
                merged_x = pd.concat([merged_x, row_data], axis=1)
            elif packae_pinout_bool:
                merged_x.loc[0, "Comments"] += "Not found in parametric; "

        if packae_pinout_bool:
            row_data = find_matching_row(pakage_df, part_x, company_x)
            if row_data is not None:
                row_data.columns = deduplicate_columns(row_data.columns)
                merged_x = pd.concat([merged_x, row_data], axis=1)
            else:
                merged_x.loc[0, "Comments"] += "Not found in package/pinout; "

        if not qualification_df.empty:
            row_data = find_matching_row(qualification_df, part_x, company_x)
            if row_data is not None:
                col_match = [c for c in row_data.columns if c.strip().lower() == "automotive"]
                merged_x["Automotive Qualified"] = row_data[col_match[0]].iloc[0] if col_match else ""
                temp_col = [c for c in row_data.columns if c.strip().lower() == "ztemperaturegrade"]
                merged_x["ZTemperatureGrade"] = row_data[temp_col[0]].iloc[0] if temp_col else ""
            else:
                merged_x["Automotive Qualified"] = ""
                merged_x["ZTemperatureGrade"] = ""
        else:
            merged_x["Automotive Qualified"] = ""
            merged_x["ZTemperatureGrade"] = ""

        merged_x.columns = deduplicate_columns(merged_x.columns)

        all_merged_rows.extend([merged_c, merged_x])

    final_df = pd.concat(all_merged_rows, ignore_index=True) if all_merged_rows else pd.DataFrame()

    log_timing(f"merge_single_pl [{pl_name}]", time.perf_counter() - t0)
    return final_df



def validate_single_pl(pl_name, merged_df, recipe_df, lookup_values, lookup_index):
    t0 = time.perf_counter()

    if merged_df.empty:
        log_timing(f"validate_single_pl [{pl_name}]", time.perf_counter() - t0)
        return pd.DataFrame(
            columns=[
                "Part Number C",
                "Company Name C",
                "Part Number X",
                "Company Name X",
                "PL Name",
                "Feature",
                "Status",
                "flags",
                "Grade",
                "UpDown",
            ]
        )

    rules_subset = recipe_df[recipe_df["ZProductValue"].astype(str).str.strip() == pl_name].copy()

    rules_subset = rules_subset[
        (rules_subset["FeaturesType"].astype(str).str.lower() == "core")
        | (rules_subset["UpgradeFeature"].astype(str).str.upper() == "TRUE")
        | (rules_subset["IsEqualFeature"].astype(str).str.upper() == "TRUE")
        | (rules_subset["FeaturesType"].astype(str).str.lower() == "tolerance")
    ].copy()

    for col in ["G1", "G2", "G3"]:
        if col in rules_subset.columns:
            rules_subset[col] = pd.to_numeric(rules_subset[col], errors="coerce")

    results = []
    pinout_flag = (
        "TRUE"
        if not rules_subset.empty
        and clean_str(rules_subset.get("IsPinoutSimilar", pd.Series(dtype=str)).iloc[0]).lower() == "true"
        else "FALSE"
    )

    for i in range(0, len(merged_df), 2):
        if i + 1 >= len(merged_df):
            break

        row1, row2 = merged_df.iloc[i], merged_df.iloc[i + 1]
        feature_statuses, feature_grades, flags, updown_notes = [], [], [], []
        max_flag, min_flag, auto_flag, temp_grade = "", "", "", ""
        pinout_check_pending = 1
        rohs_check_pending = True

        for _, rule in rules_subset.iterrows():
            status = ""
            grade = ""
            feature = clean_str(rule.get("Features", ""))
            ftype = clean_str(rule.get("FeaturesType", "")).lower()
            upgrade_flag = clean_str(rule.get("UpgradeFeature", "")).upper()
            equal_flag = clean_str(rule.get("IsEqualFeature", "")).upper()

            val1 = val2 = nunit1 = nunit2 = None

            if pinout_flag == "TRUE" and pinout_check_pending:
                val1 = clean_str(row1.get("NormalizedPinOutName", ""))
                val2 = clean_str(row2.get("NormalizedPinOutName", ""))
                pinout_check_pending = 0
                if val1 == val2 and val1:
                    status, grade = "✅ PinOut Match", "A"
                else:
                    status, grade = "❌ Different PinOut", "DiffPinOut"
                feature_statuses.append(f"{feature}: {status}" if feature else status)
                feature_grades.append(grade)
                status, grade = "", ""

            if rohs_check_pending:
                rohs_check_pending = False
                val1 = clean_str(row1.get("RoHSComplianceStatus", ""))
                val2 = clean_str(row2.get("RoHSComplianceStatus", ""))
                if val1 != val2 and val1:
                    status, grade = "❌ Different Rohs", "Diff_ROHs"
                    flags.append(f"Diff In ROHs {val1} , {val2}")
                else:
                    status, grade = f"RoHSComplianceStatus: ✅ Same Rohs {val1} ,{val2}", "Diff_ROHs"
                feature_statuses.append(f"{feature}: {status}" if feature else status)
                feature_grades.append(grade)
                status, grade = "", ""

            if ftype == "core":
                if feature not in merged_df.columns:
                    status, grade = f"⚠ Missing column {feature}", "Fail"
                else:
                    val1 = clean_str(row1.get(feature, ""))
                    val2 = clean_str(row2.get(feature, ""))
                    if val1 == "nan" and val2 == "nan":
                        status, grade = "⚠ Missing Value", "nan"
                    elif ((val2 == "N/A") and (val1 != "N/A")) or ((val1 == "N/A") and (val2 != "N/A")):
                        status, grade = "⚠ Value is N/A", "B"
                    elif val1 == val2:
                        status, grade = f"✅ C_Match {val1}", "A"
                    else:
                        found = find_feature_CoreLookUp_fast(lookup_index, pl_name, feature, val1, val2)
                        if found == "Found":
                            status, grade = "✅ Match By LookUp", "A"
                        else:
                            status, grade = f"❌ Mismatch and not in LookUp ({val1} vs {val2})", "Fail"
                            flags.append(f"missMatchAtCore {feature}")
                            grade = "DiffPKG" if feature == "Normalized Package Name" else "FailInCore"
                feature_statuses.append(f"{feature}: {status}")
                feature_grades.append(grade)

            elif equal_flag == "TRUE":
                if feature not in merged_df.columns:
                    status, grade = f"⚠ Missing column {feature}", "MissingColumn"
                else:
                    val1 = clean_str(row1.get(feature, ""))
                    val2 = clean_str(row2.get(feature, ""))
                    if val1 == "nan" and val2 == "nan":
                        status, grade = "⚠ Missing Value", "B"
                    elif val1 == "N/A" and val2 == "N/A":
                        status, grade = "⚠ Value is N/A", "N/A"
                    elif ((val2 == "N/A") and (val1 != "N/A")) or ((val1 == "N/A") and (val2 != "N/A")):
                        status, grade = "⚠ Value is N/A", "B"
                    elif val1 == val2:
                        status, grade = "✅ E_Match", "A"
                    else:
                        found = find_feature_CoreLookUp_fast(lookup_index, pl_name, feature, val1, val2)
                        if found == "Found":
                            status, grade = "✅ Match By LookUp", "A"
                        else:
                            status, grade = f"❌ E_Mismatch ({val1} vs {val2})", "B"
                feature_statuses.append(f"{feature}: {status}")
                feature_grades.append(grade)

            elif upgrade_flag == "TRUE":
                if feature not in merged_df.columns:
                    status, grade = f"⚠ Missing column {feature}", "Fail"
                    feature_statuses.append(f"{feature}: {status}")
                    feature_grades.append(grade)
                else:
                    val1 = clean_str(row1.get(feature, ""))
                    val2 = clean_str(row2.get(feature, ""))

                    if feature == "Automotive Qualified":
                        a1 = val1.lower()
                        a2 = val2.lower()
                        if a1 == a2:
                            auto_flag = "Same"
                            updown_notes.append(f"Same At {feature},value1 {val1},value2 {val2}")
                        elif a1 == "yes" and a2 == "no":
                            auto_flag = "DownGrade"
                            updown_notes.append(f"DownGrade At {feature},value1 {val1},value2 {val2}")
                        elif a1 == "no" and a2 == "yes":
                            auto_flag = "UpGrade"
                            updown_notes.append(f"UpGrade At {feature},value1 {val1},value2 {val2}")
                        else:
                            updown_notes.append(f"Missing/Unknown Value At {feature},value1 {val1},value2 {val2}")

                    elif feature == "ZTemperatureGrade":
                        if val2 == val1:
                            temp_grade = "Same"
                            updown_notes.append(f"Same At {feature},value1 {val1},value2 {val2}")
                        else:
                            grade_hierarchy = {
                                "nan": 0,
                                None: 0,
                                "": 0,
                                "N/R": 0,
                                "Commercial": 1,
                                "Industrial": 1,
                                "Ex. Industrial": 1,
                                "Ex. Commercial": 1,
                                "High Rel": 2,
                                "Hi-REL": 2,
                                "Ex. High Rel": 2,
                                "Automotive (Grade 4)": 3,
                                "Automotive (Grade 3)": 3,
                                "Automotive (Grade 2)": 3,
                                "Automotive (Grade 1)": 3,
                                "Automotive (Grade 0)": 3,
                                "Military": 5,
                            }
                            val1_level = grade_hierarchy.get(val1, -1)
                            val2_level = grade_hierarchy.get(val2, -1)

                            if val1_level == -1 and val2_level == -1:
                                updown_notes.append(f"Unknown Grade At {feature},value1 {val1},value2 {val2}")
                                temp_grade = "Unknown"
                            elif val2_level > val1_level:
                                updown_notes.append(f"UpGrade At {feature},value1 {val1},value2 {val2}")
                                temp_grade = "UpGrade"
                            elif val2_level < val1_level:
                                updown_notes.append(f"DownGrade At {feature},value1 {val1},value2 {val2}")
                                temp_grade = "DownGrade"
                            else:
                                temp_grade = "Same"
                                updown_notes.append(f"Same At {feature},value1 {val1},value2 {val2}")

                    elif "Temperature" in feature:
                        if val1 == "nan" or val2 == "nan":
                            updown_notes.append(f"Missing Value At {feature},value1 {val1},value2 {val2}")
                        elif val1 == val2:
                            updown_notes.append(f"Same At {feature},value1 {val1},value2 {val2}")
                            flag = "Same"
                            if "Maximum" in feature:
                                max_flag = flag
                            elif "Minimum" in feature:
                                min_flag = flag
                        elif val1 == "N/A" and val2 != "N/A":
                            updown_notes.append(f"UpGrade At {feature},value1 {val1},value2 {val2}")
                            flag = "UpGrade"
                            if "Maximum" in feature:
                                max_flag = flag
                            elif "Minimum" in feature:
                                min_flag = flag
                        elif val1 != "N/A" and val2 == "N/A":
                            updown_notes.append(f"DownGrade At {feature},value1 {val1},value2 {val2}")
                            flag = "DownGrade"
                            if "Maximum" in feature:
                                max_flag = flag
                            elif "Minimum" in feature:
                                min_flag = flag
                        else:
                            result = compare_parts(lookup_values, val1, val2)
                            if result["state"] == "Match" and result["nValue"] and result["nValue2"]:
                                try:
                                    n_val1 = safe_int(result["nValue"][0])
                                    n_val2 = safe_int(result["nValue2"][0])
                                except Exception:
                                    n_val1 = n_val2 = None
                                if n_val1 is None or n_val2 is None:
                                    updown_notes.append(f"Missing Value At {feature},value1 {val1},value2 {val2}")
                                else:
                                    flag = None
                                    if "Maximum" in feature:
                                        if n_val2 > n_val1:
                                            flag = "UpGrade"
                                        elif n_val2 < n_val1:
                                            flag = "DownGrade"
                                        else:
                                            flag = "Same"
                                        max_flag = flag
                                    elif "Minimum" in feature:
                                        if n_val2 < n_val1:
                                            flag = "UpGrade"
                                        elif n_val2 > n_val1:
                                            flag = "DownGrade"
                                        else:
                                            flag = "Same"
                                        min_flag = flag
                                    if flag:
                                        updown_notes.append(f"{flag} At {feature},value1 {val1},value2 {val2}")
                            elif result["state"] == "Different DetailedValueType":
                                updown_notes.append(f"Different DetailedValueType {feature},value1 {val1},value2 {val2}")
                            elif result["state"] == "Different FeatureCode":
                                updown_notes.append(f"Different FeatureCode {feature},value1 {val1},value2 {val2}")
                            else:
                                updown_notes.append(f"Not Found {feature},value1 {val1},value2 {val2}")

            elif ftype == "tolerance":
                if feature not in merged_df.columns:
                    status, grade = f"⚠ Missing column {feature}", "Fail"
                elif feature == "Pin Pitch_mm" or "Calculated_" in feature:
                    val1 = clean_str(row1.get(feature, ""))
                    val2 = clean_str(row2.get(feature, ""))
                    if val1 == "nan" and val2 == "nan":
                        status, grade = "⚠ Missing Value", "nan"
                    elif val1 == "N/A" or val2 == "N/A":
                        status, grade = "⚠ Value is N/A", "B"
                    elif val1 == val2:
                        status, grade = "✅T_Match", "A"
                    else:
                        if val1 in {"nan", ""} or val2 in {"nan", ""}:
                            status, grade = "⚠ Missing Value", "nan"
                        else:
                            try:
                                vv1 = float(val1)
                                vv2 = float(val2)
                                g1, g2, g3 = rule.get("G1"), rule.get("G2"), rule.get("G3")
                                g1 = 0 if pd.isna(g1) else float(g1)
                                g2 = 0 if pd.isna(g2) else float(g2)
                                g3 = 0 if pd.isna(g3) else float(g3)
                                diff = percent_diff_base_a(vv1, vv2)
                                if diff <= (g1 + 1) or diff == 0.0:
                                    status, grade = f"❓ Within G1 {g1} ({diff}%) value1: {vv1} value2: {vv2}", "A"
                                elif diff <= g3:
                                    status, grade = f"❓ Within G2 (less) {g2} | {g3} ({diff}%) value1: {vv1} value2: {vv2}", "B"
                                elif diff <= g3 + 10:
                                    status, grade = f"❓ Within G3 {g3} ({diff}%) value1: {vv1} value2: {vv2}", "C"
                                else:
                                    status, grade = f"❌ Outside tolerance ({diff}%) value1: {vv1} value2: {vv2}", "Fail"
                                    flags.append(f"Outside tolerance {feature} value1: {vv1} value2: {vv2}")
                            except Exception as exc:
                                status, grade = f"⚠ Error processing tolerance: {exc}", "Fail"
                else:
                    val1 = clean_str(row1.get(feature, ""))
                    val2 = clean_str(row2.get(feature, ""))
                    if val1 == "nan" and val2 == "nan":
                        status, grade = "⚠ Missing Value", "nan"
                    elif ((val1 == "N/A") and (val2 != "N/A")) or ((val2 == "N/A") and (val1 != "N/A")):
                        status, grade = "⚠ Value is N/A", "B"
                    elif val1 == val2:
                        status, grade = "✅T_Match", "A"
                    else:
                        result = compare_parts(lookup_values, val1, val2)
                        if result["state"] == "Match":
                            if result["nValue"] and result["nValue2"]:
                                try:
                                    nunit1 = result["nUnit"][0] if result["nUnit"] else None
                                    nunit2 = result["nUnit2"][0] if result["nUnit2"] else None
                                except Exception:
                                    nunit1 = nunit2 = None

                            if nunit1 and nunit2 and nunit1 != nunit2:
                                status, grade = "Different Unit", "unitFail"
                                flags.append(f"Different Unit {feature}")
                            else:
                                g1, g2, g3 = rule.get("G1"), rule.get("G2"), rule.get("G3")
                                g1 = 0 if pd.isna(g1) else float(g1)
                                g2 = 0 if pd.isna(g2) else float(g2)
                                g3 = 0 if pd.isna(g3) else float(g3)
                                status = ""
                                grade = ""
                                for j in range(min(len(result["nValue"]), len(result["nValue2"]))):
                                    try:
                                        vv1 = float(result["nValue"][j])
                                        vv2 = float(result["nValue2"][j])
                                    except Exception:
                                        continue
                                    diff = percent_diff_base_a(vv1, vv2)
                                    if diff <= (g1 + 1) or diff == 0.0:
                                        status, grade = f"❓ Within G1 {g1} ({diff}%) value1: {vv1} value2: {vv2}", "A"
                                    elif diff <= g3:
                                        status, grade = f"❓ Within G2 (less) {g2} | {g3} ({diff}%) value1: {vv1} value2: {vv2}", "B"
                                    elif diff <= g3 + 10:
                                        status, grade = f"❓ Within G3 {g3} ({diff}%) value1: {vv1} value2: {vv2}", "C"
                                    else:
                                        status, grade = f"❌ Outside tolerance ({diff}%) value1: {vv1} value2: {vv2}", "Fail"
                                        flags.append(f"Outside tolerance {feature} value1: {vv1} value2: {vv2}")

                                if not status:
                                    try:
                                        vv1 = float(val1)
                                        vv2 = float(val2)
                                        diff = percent_diff_base_a(vv1, vv2)
                                        if diff <= (g1 + 1) or diff == 0.0:
                                            status, grade = f"❓ Within G1 {g1} ({diff}%) value1: {vv1} value2: {vv2}", "A"
                                        elif diff <= g3:
                                            status, grade = f"❓ Within G2 (less) {g2} | {g3} ({diff}%) value1: {vv1} value2: {vv2}", "B"
                                        elif diff <= g3 + 10:
                                            status, grade = f"❓ Within G3 {g3} ({diff}%) value1: {vv1} value2: {vv2}", "C"
                                        else:
                                            status, grade = f"❌ Outside tolerance ({diff}%) value1: {vv1} value2: {vv2}", "Fail"
                                            flags.append(f"Outside tolerance {feature} value1: {vv1} value2: {vv2}")
                                    except Exception:
                                        status, grade = f"⚠ Missing Values at LookUp Table {val1} vs {val2}", "B"

                        elif result["state"] == "Different DetailedValueType":
                            status, grade = "Different DetailedValueType", "FailInDetailedValueType"
                            flags.append(f"Different DetailedValueType {feature}")
                        elif result["state"] == "Different FeatureCode":
                            status, grade = "Different FeatureCode", "FailInFeatureCode"
                            flags.append(f"Different FeatureCode {feature}")
                        else:
                            status, grade = f"⚠ Missing Values at LookUp Table {val1} vs {val2}", "B"

                feature_statuses.append(f"{feature}: {status}")
                feature_grades.append(grade)

        overall_grade = determine_overall_grade(feature_grades, auto_flag, max_flag, min_flag)

        results.append(
            {
                "Part Number C": row2.get("PartNumber", ""),
                "Company Name C": row2.get("Company", ""),
                "Part Number X": row1.get("PartNumber", ""),
                "Company Name X": row1.get("Company", ""),
                "PL Name": pl_name,
                "Feature": "OVERALL",
                "Status": " | ".join(feature_statuses),
                "flags": " | ".join(flags) if flags else "",
                "Grade": overall_grade,
                "UpDown": " | ".join(updown_notes) if updown_notes else temp_grade,
            }
        )

    log_timing(f"validate_single_pl [{pl_name}]", time.perf_counter() - t0)
    return pd.DataFrame(results)



def compare_parts(df, value, value2):
    required_cols = ["AcceptedValue", "ValueID", "FeatureCode", "NormalizedValue", "Unit"]
    if any(col not in df.columns for col in required_cols):
        return {
            "value": value,
            "nValue": [],
            "nUnit": [],
            "FeatureCode": "NotFound",
            "state": "NotFound In File",
            "value2": value2,
            "nValue2": [],
            "nUnit2": [],
        }

    accepted_clean = df["AcceptedValue"].astype(str).str.strip().str.lower()
    value_clean = clean_str(value).lower()
    value2_clean = clean_str(value2).lower()

    part1 = df[accepted_clean == value_clean].reset_index(drop=True)
    part2 = df[accepted_clean == value2_clean].reset_index(drop=True)

    if part1.empty or part2.empty:
        return {
            "value": value,
            "nValue": [],
            "nUnit": [],
            "FeatureCode": "NotFound",
            "state": "NotFound In File",
            "value2": value2,
            "nValue2": [],
            "nUnit2": [],
        }

    chosen_value_id = part1.loc[0, "ValueID"]
    chosen_value_id2 = part2.loc[0, "ValueID"]

    part1 = part1[part1["ValueID"] == chosen_value_id].reset_index(drop=True)
    part2 = part2[part2["ValueID"] == chosen_value_id2].reset_index(drop=True)

    fc1 = set(part1["FeatureCode"].dropna().unique())
    fc2 = set(part2["FeatureCode"].dropna().unique())

    if fc1 != fc2:
        len_fc1 = sorted([len(str(x)) for x in fc1])
        len_fc2 = sorted([len(str(x)) for x in fc2])
        if len_fc1 == len_fc2:
            feature_state = "MatchByLength"
        else:
            return {
                "value": value,
                "nValue": [],
                "nUnit": [],
                "FeatureCode": "Different",
                "state": "Different FeatureCode",
                "value2": value2,
                "nValue2": [],
                "nUnit2": [],
            }
    else:
        feature_state = "Match"

    sorted_fc = sorted(fc1, key=fc_sort_key)

    def prepare_sorted_part(part, sorted_feature_codes):
        try:
            part = part.set_index("FeatureCode").loc[sorted_feature_codes].reset_index()
        except Exception:
            part = part.copy()
        part["NumericValue"] = pd.to_numeric(part["NormalizedValue"], errors="coerce")
        part = part.sort_values(by=["NumericValue"], ascending=True, na_position="last")
        return part

    part1_sorted = prepare_sorted_part(part1, sorted_fc)
    part2_sorted = prepare_sorted_part(part2, sorted_fc)

    return {
        "value": value,
        "nValue": part1_sorted["NormalizedValue"].dropna().tolist(),
        "nUnit": part1_sorted["Unit"].dropna().tolist(),
        "FeatureCode": feature_state,
        "state": "Match",
        "value2": value2,
        "nValue2": part2_sorted["NormalizedValue"].dropna().tolist(),
        "nUnit2": part2_sorted["Unit"].dropna().tolist(),
    }



def determine_overall_grade(feature_grades, Auto, maxFlag, minFlag):
    critical_failures = {
        "DiffPinOut": "DiffPinOut",
        "DiffPKG": "Diff. Package",
        "FailInCore": "Not Drop-in",
        "FailInDetailedValueType": "Detailed Value Type Fail Not Drop-in",
        "FailInFeatureCode": "FeatureCode FAIL Not Drop-in",
        "unitFail": "Unit FAIL Not Drop-in",
        "nan": "In-Complete Data",
        "MissingColumn": "Missing Column",
    }
    for failure_type, grade in critical_failures.items():
        if failure_type in feature_grades:
            return grade

    grade_levels = ["Fail", "C", "B", "A"]
    for grade_level in grade_levels:
        if grade_level in feature_grades:
            base_grade = "Drop-in H" if grade_level == "Fail" else f"Drop-in {grade_level}"
            modifier = determine_grade_modifier(Auto, maxFlag, minFlag)
            if modifier:
                return f"{base_grade} / {modifier}"
            if base_grade == "Drop-in A" and "Diff_ROHs" in feature_grades:
                return "Drop-in B"
            return base_grade

    return "Unknown Grade"



def determine_grade_modifier(Auto, maxFlag, minFlag):
    return determine_modifier_from_flags(Auto, maxFlag, minFlag)



def determine_modifier_from_flags(Auto, maxFlag, minFlag):
    """Determine modifier based on min/max flags."""
    upgrade_conditions = [
        (maxFlag == "UpGrade" and minFlag in ["Same", "UpGrade"] and Auto in ["Same", "UpGrade"]),
        (minFlag == "UpGrade" and maxFlag in ["Same", "UpGrade"] and Auto in ["Same", "UpGrade"]),
    ]
    if any(upgrade_conditions):
        return "Upgrade"

    downgrade_conditions = [
        (maxFlag == "DownGrade" or minFlag == "DownGrade" or Auto == "DownGrade"),
        (minFlag == "DownGrade" or maxFlag == "DownGrade" or Auto == "DownGrade"),
    ]
    if any(downgrade_conditions):
        return "Downgrade"

    mixed_conditions = [
        (maxFlag == "DownGrade" and minFlag in ["Same", "UpGrade"]),
        (minFlag == "DownGrade" and maxFlag in ["Same", "UpGrade"]),
    ]
    if any(mixed_conditions):
        return "Downgrade"

    return ""



def organize_single_pl(cross_df_pl, validation_df_pl):
    t0 = time.perf_counter()

    if validation_df_pl.empty:
        validation_df_pl = pd.DataFrame(
            columns=[
                "Part Number C",
                "Company Name C",
                "Part Number X",
                "Company Name X",
                "Status",
                "flags",
                "Grade",
                "UpDown",
            ]
        )

    validation_df_pl = validation_df_pl.rename(
        columns={
            "Part Number C": "PartNumberC",
            "Company Name C": "CompanyNameC",
            "Part Number X": "PartNumberX",
            "Company Name X": "CompanyNamex",
            "Status": "Match Feature",
            "flags": "Different Features",
        }
    )

    val_cols = [
        "PartNumberC",
        "CompanyNameC",
        "PartNumberX",
        "CompanyNamex",
        "Match Feature",
        "Different Features",
        "Grade",
        "UpDown",
    ]
    for col in val_cols:
        if col not in validation_df_pl.columns:
            validation_df_pl[col] = ""

    merge_keys = ["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex"]
    merged = pd.merge(cross_df_pl, validation_df_pl[val_cols], on=merge_keys, how="left")

    unmatched_mask = merged["Grade"].isna() | (merged["Grade"].astype(str).str.strip() == "")

    if unmatched_mask.any():
        swapped = validation_df_pl[val_cols].rename(
            columns={
                "PartNumberC": "PartNumberX",
                "CompanyNameC": "CompanyNamex",
                "PartNumberX": "PartNumberC",
                "CompanyNamex": "CompanyNameC",
            }
        )

        matched_rows = merged[~unmatched_mask].copy()
        unmatched_rows = merged[unmatched_mask].drop(
            columns=["Match Feature", "Different Features", "Grade", "UpDown"], errors="ignore"
        )

        filled_rows = pd.merge(unmatched_rows, swapped[val_cols], on=merge_keys, how="left")
        merged = pd.concat([matched_rows, filled_rows], ignore_index=True)

    for col in ["Match Feature", "Different Features", "Grade", "UpDown"]:
        swapped_col = f"{col}_swapped"
        if swapped_col in merged.columns:
            merged[col] = merged[col].fillna(merged[swapped_col])
            merged.drop(columns=[swapped_col], inplace=True, errors="ignore")

    def compute_status(row):
        grade = clean_str(row.get("Grade", ""))
        same_pls = clean_str(row.get("SamePLs", "")).lower()
        found_data = clean_str(row.get("FoundData", "")).upper()
        pack_pinout = clean_str(row.get("PackaePinout", ""))
        if grade in ["Drop-in H", "Drop-in A", "Drop-in B", "Drop-in C"] or grade.startswith("Drop-in "):
            return "Cross"
        if grade in [
            "Not Drop-in",
            "Detailed Value Type Fail Not Drop-in",
            "Unit FAIL Not Drop-in",
            "FeatureCode FAIL Not Drop-in",
            "DiffPinOut",
            "Diff. Package",
        ] and found_data == "TRUE":
            return "Not Cross"
        if same_pls == "not":
            return "Different PLs"
        if found_data == "FALSE" and pack_pinout == "":
            return "Not Found Data"
        return ""

    merged["Status"] = merged.apply(compute_status, axis=1)

    def compare_grades(row):
        grade = clean_str(row.get("Grade", "")).lower()
        cross_grade = clean_str(row.get("CrossGrade", "")).lower()
        if not grade or not cross_grade:
            return "Not"

        grade_match = re.search(r"drop[-\s]?in\s*([a-z])", grade)
        cross_match = re.search(r"drop[-\s]?in\s*([a-z])", cross_grade)

        grade_letter = grade_match.group(1) if grade_match else None
        cross_letter = cross_match.group(1) if cross_match else None

        grade_q = re.search(r"/\s*(upgrade|downgrade)", grade)
        cross_q = re.search(r"/\s*(upgrade|downgrade)", cross_grade)

        grade_result_q = grade_q.group(1) if grade_q else None
        cross_result_q = cross_q.group(1) if cross_q else None

        if grade == cross_grade:
            return "Same"
        if grade_letter == "h" and cross_letter == "c" and grade_result_q == cross_result_q:
            return "Pass"
        if grade_letter == "h" and cross_letter == "c" and grade_result_q != cross_result_q:
            return "Qualification Issue"

        dropin_levels = ["a", "b", "c", "h"]
        if grade_letter in dropin_levels and cross_letter in dropin_levels:
            if grade_letter != cross_letter:
                return "Grading Issue"
            if grade_letter == cross_letter and grade != cross_grade:
                return "Qualification Issue"

        if (
            "upgrade" in grade
            and "downgrade" in cross_grade
            and (grade_letter == cross_letter or (grade_letter == "h" and cross_letter == "c"))
        ) or (
            "downgrade" in grade and "upgrade" in cross_grade and grade_letter == cross_letter
        ):
            return "Qualification Issue"
        return "Not"

    merged["Compare"] = merged.apply(compare_grades, axis=1)

    log_timing("organize_single_pl", time.perf_counter() - t0)
    return merged



def find_matching_row(df, part_value, company_value):
    if df.empty:
        return None

    cols_lower = df.columns.str.lower()
    company_col = None
    part_col = None

    for idx, col in enumerate(cols_lower):
        if col in ["company", "companyname", "company name", "companynamec"]:
            company_col = df.columns[idx]
        if col in ["partnumber", "part number", "part_number", "partnumberc"]:
            part_col = df.columns[idx]

    if not company_col or not part_col:
        return None

    mask = (
        df[part_col].astype(str).str.strip() == clean_str(part_value)
    ) & (
        df[company_col].astype(str).str.strip() == clean_str(company_value)
    )

    filtered = df[mask]

    if not filtered.empty:
        row = filtered.iloc[0].drop([part_col, company_col], errors="ignore")
        return row.to_frame().T.reset_index(drop=True)

    return None



def process_all_pls_from_combined(
    crosses_df: pd.DataFrame,
    files_data: Dict[str, pd.DataFrame],
    recipe_df: pd.DataFrame,
    lookUpCore_df: pd.DataFrame,
    selected_pls: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Main processing function adapted for Streamlit."""
    t0_total = time.perf_counter()

    progress_bar = st.progress(0.0)
    status_text = st.empty()

    parametric_df = files_data["parametric_file"]
    pakage_df = files_data["pakageAndPinout_file"]
    qualification_df = files_data["qualification_file"]
    lookup_values = pd.concat([files_data["lookUpFile1"], files_data["lookUpFile2"]], ignore_index=True)
    lookup_index = build_lookup_index(lookUpCore_df)

    ensure_columns(crosses_df, ["PLc", "PLx", "PartNumberC", "PartNumberX", "CompanyNameC", "CompanyNamex"], "crosses file")
    ensure_columns(recipe_df, ["ZProductValue", "Features", "FeaturesType", "UpgradeFeature", "IsEqualFeature"], "recipe file")

    available_pls = crosses_df["PLc"].dropna().astype(str).str.strip().unique().tolist()
    if selected_pls:
        selected_set = {clean_str(x) for x in selected_pls}
        pl_list = [pl for pl in available_pls if clean_str(pl) in selected_set]
    else:
        pl_list = available_pls

    st.info(f"Found {len(pl_list)} unique PLs to process")

    all_results = []

    for i, pl_name in enumerate(pl_list):
        try:
            status_text.text(f"Processing PL {i + 1}/{len(pl_list)}: {pl_name}")
            progress_bar.progress((i + 1) / max(len(pl_list), 1))

            cross_df_pl = crosses_df[crosses_df["PLc"].astype(str).str.strip() == pl_name].copy()
            if cross_df_pl.empty:
                continue

            cross_df_pl = process_single_pl(pl_name, cross_df_pl, parametric_df, pakage_df, recipe_df)
            merged_df_pl = merge_single_pl(pl_name, cross_df_pl, parametric_df, pakage_df, qualification_df)
            validation_df_pl = validate_single_pl(pl_name, merged_df_pl, recipe_df, lookup_values, lookup_index)
            final_df_pl = organize_single_pl(cross_df_pl, validation_df_pl)

            if not final_df_pl.empty:
                all_results.append(final_df_pl)

        except Exception as exc:
            st.error(f"Error processing PL {pl_name}: {exc}")
            continue

    status_text.empty()
    progress_bar.empty()

    if all_results:
        combined_df = pd.concat(all_results, ignore_index=True).drop_duplicates()
        log_timing("process_all_pls_from_combined [TOTAL]", time.perf_counter() - t0_total)
        return combined_df
    return pd.DataFrame()


# ============================================================================
# STREAMLIT UI
# ============================================================================

st.set_page_config(
    page_title="PL Cross-Reference Validation System",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("PL-by-PL Cross-Reference Validation System")
st.markdown("### Enhanced Cross-Validation with Upgrade/Downgrade and Qualification Analysis")

st.sidebar.header("Upload Files")
crosses_file = st.sidebar.file_uploader("Crosses File (.xlsx)", type="xlsx")
combined_file = st.sidebar.file_uploader("Combined Data File (.xlsx)", type="xlsx")
recipe_file = st.sidebar.file_uploader("Recipe File (.xlsx)", type="xlsx")
lookup_core_file = st.sidebar.file_uploader("LookUp Core File (.xlsx)", type="xlsx")

st.sidebar.header("Options")
process_all = st.sidebar.checkbox("Process all PLs", value=True)

available_pls: List[str] = []
if crosses_file is not None:
    try:
        preview_crosses = read_excel_bytes(crosses_file.getvalue())
        if "PLc" in preview_crosses.columns:
            available_pls = preview_crosses["PLc"].dropna().astype(str).str.strip().unique().tolist()
    except Exception:
        available_pls = []

selected_pls: List[str] = []
if not process_all and available_pls:
    selected_pls = st.sidebar.multiselect(
        "Select specific PLs",
        options=available_pls,
        default=available_pls[: min(5, len(available_pls))],
    )
elif not process_all:
    st.sidebar.info("Upload the Crosses file first to load PL options.")

if st.sidebar.button("Run Analysis", type="primary"):
    if not all([crosses_file, combined_file, recipe_file, lookup_core_file]):
        st.error("Please upload all 4 required files.")
    else:
        with st.spinner("Processing files. This may take several minutes."):
            try:
                reset_runtime_state()

                st.info("Loading files...")
                crosses_df = read_excel_bytes(crosses_file.getvalue())
                recipe_df = read_excel_bytes(recipe_file.getvalue())
                lookup_core_df = read_excel_bytes(lookup_core_file.getvalue())
                files_data = load_files_from_excel(combined_file.getvalue())

                if not process_all and selected_pls:
                    st.info(f"Running analysis for {len(selected_pls)} selected PLs...")
                else:
                    st.info("Running full analysis...")

                result_df = process_all_pls_from_combined(
                    crosses_df=crosses_df,
                    files_data=files_data,
                    recipe_df=recipe_df,
                    lookUpCore_df=lookup_core_df,
                    selected_pls=None if process_all else selected_pls,
                )

                if not result_df.empty:
                    st.success(f"Analysis complete. Processed {len(result_df)} rows.")

                    st.subheader("Results Summary")
                    col1, col2, col3, col4 = st.columns(4)

                    with col1:
                        st.metric("Total Rows", len(result_df))
                    with col2:
                        cross_count = len(result_df[result_df.get("Status", pd.Series(dtype=str)) == "Cross"])
                        st.metric("Cross Matches", cross_count)
                    with col3:
                        not_cross_count = len(result_df[result_df.get("Status", pd.Series(dtype=str)) == "Not Cross"])
                        st.metric("Not Cross", not_cross_count)
                    with col4:
                        grade_series = result_df.get("Grade", pd.Series(dtype=str)).astype(str)
                        drop_in_a_count = len(result_df[grade_series.str.contains("Drop-in A", na=False)])
                        st.metric("Drop-in A", drop_in_a_count)

                    st.subheader("Sample Results")
                    st.dataframe(result_df.head(10), use_container_width=True)

                    st.subheader("Full Results")
                    st.dataframe(result_df, use_container_width=True, height=500)

                    csv_buffer = io.StringIO()
                    result_df.to_csv(csv_buffer, index=False)
                    st.download_button(
                        "Download Results (CSV)",
                        csv_buffer.getvalue(),
                        f"pl_cross_validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        "text/csv",
                    )

                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                        result_df.to_excel(writer, sheet_name="Results", index=False)
                    st.download_button(
                        "Download Results (Excel)",
                        excel_buffer.getvalue(),
                        f"pl_cross_validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                    timing_df = get_timing_df()
                    if not timing_df.empty:
                        st.subheader("Processing Timing")
                        st.dataframe(timing_df, use_container_width=True)
                else:
                    st.warning("No results generated.")

            except Exception as exc:
                st.error(f"Processing failed: {str(exc)}")
                st.exception(exc)

with st.expander("Instructions"):
    st.markdown(
        """
### Required Files
1. **Crosses File** - Contains `PLc`, `PLx`, `PartNumberC`, `PartNumberX`, and related columns.
2. **Combined Data File** - Excel workbook with `lookUp` and `Data` sheets.
3. **Recipe File** - Contains validation rules per PL.
4. **LookUp Core File** - Core lookup reference data.

### Features
- Full upgrade and downgrade analysis
- Automotive qualification checks
- Tolerance validation with G1, G2, and G3
- Package and pinout similarity checks
- Real-time progress tracking
- Detailed timing analysis
        """
    )

st.markdown("---")
st.markdown("PL-by-PL Cross-Reference Validation System v2.2")
