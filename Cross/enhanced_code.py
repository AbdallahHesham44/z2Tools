
from __future__ import annotations

import os
import re
import shutil
import time
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd


# ============================================================================
# CONSTANTS
# ============================================================================

TIMING_LOG: List[Dict[str, float]] = []
LOOKUP_CACHE: Dict[Tuple[str, str, Tuple[str, str]], str] = {}

CORE_LOOKUP_REQUIRED_COLS = [
    "Product",
    "GroupID",
    "ParentGroupID",
    "RealGroupString",
    "GroupType",
    "ModifiedDate",
    "FeatureName",
]

LOOKUP_INDEX_REQUIRED_COLS = ["FeatureName", "RealGroupString", "ParentGroupID"]

COMPANY_COL_CANDIDATES = {"company", "companyname", "company name", "companynamec"}
PART_COL_CANDIDATES = {"partnumber", "part number", "part_number", "partnumberc"}

CRITICAL_FAILURE_GRADES = {
    "DiffPinOut": "DiffPinOut",
    "DiffPKG": "Diff. Package",
    "FailInCore": "Not Drop-in",
    "FailInDetailedValueType": "Detailed Value Type Fail Not Drop-in",
    "FailInFeatureCode": "FeatureCode FAIL Not Drop-in",
    "unitFail": "Unit FAIL Not Drop-in",
    "nan": "In-Complete Data",
    "MissingColumn": "Missing Column",
}

GRADE_HIERARCHY = {
    "nan": 0,
    None: 0,
    " ": 0,
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

TRUTHY_VALUES = {"true", "1", "yes"}


# ============================================================================
# TIMING UTILITIES
# ============================================================================

def log_timing(func_name: str, elapsed: float) -> None:
    """Store and print timing info for a function call."""
    record = {"function": func_name, "elapsed_sec": round(elapsed, 4)}
    TIMING_LOG.append(record)
    print(f"⏱  [{func_name}] took {elapsed:.4f}s")


def print_timing_summary() -> None:
    """Print a formatted summary table of all recorded timings."""
    if not TIMING_LOG:
        print("No timing data recorded.")
        return

    print("\n" + "=" * 60)
    print("TIMING SUMMARY")
    print("=" * 60)
    print(f"{'Function':<45} {'Time (s)':>10}")
    print("-" * 60)

    total = 0.0
    for record in TIMING_LOG:
        print(f"{record['function']:<45} {record['elapsed_sec']:>10.4f}")
        total += record["elapsed_sec"]

    print("-" * 60)
    print(f"{'TOTAL':<45} {total:>10.4f}")
    print("=" * 60)


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def safe_int(value: object) -> Optional[int]:
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return None


def as_text(value: object) -> str:
    """Preserve original string conversion semantics used throughout the script."""
    return str(value).strip()


def strip_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of a dataframe with stripped column names."""
    out = df.copy()
    out.columns = out.columns.map(lambda col: str(col).strip())
    return out


def truncate_sheet_name(pl_name: object) -> str:
    """Truncate sheet name to 31 characters (Excel limit)."""
    pl_text = str(pl_name)
    return pl_text[:31] if len(pl_text) > 31 else pl_text


def deduplicate_columns(columns: Sequence[object]) -> List[str]:
    """Make duplicate column names unique by adding suffixes."""
    seen: Dict[str, int] = {}
    new_cols: List[str] = []

    for col in columns:
        col_text = str(col)
        if col_text not in seen:
            seen[col_text] = 0
            new_cols.append(col_text)
        else:
            seen[col_text] += 1
            new_cols.append(f"{col_text}_{seen[col_text]}")
    return new_cols


def percent_diff_base_a(a: float, b: float) -> float:
    """Calculate percentage difference based on value a."""
    if a == 0:
        return float("inf") if b != 0 else 0.0
    return round(abs(a - b) / abs(a) * 100, 2)


def fc_sort_key(fc: object) -> float:
    """Sort key for FeatureCodes (V before U, numeric order)."""
    match = re.search(r"([VU])(\d+)$", as_text(fc))
    if match:
        code_type, number = match.groups()
        return int(number) * 2 + (0 if code_type == "V" else 1)
    return float("inf")


# ============================================================================
# NORMALIZATION HELPERS
# ============================================================================

def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    if text in {"nan", "none", "nat"}:
        return ""
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text)
    return text


def compact_text(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", normalize_text(value))


def is_truthy(value: object) -> bool:
    return normalize_text(value) in TRUTHY_VALUES


def _append_feature_result(
    feature: object,
    status: str,
    grade: str,
    feature_statuses: List[str],
    feature_grades: List[str],
) -> None:
    feature_statuses.append(f"{feature}: {status}")
    feature_grades.append(grade)


def _grade_tolerance_result(
    vv1: float,
    vv2: float,
    g1: float,
    g2: float,
    g3: float,
) -> Tuple[str, str, float]:
    diff = percent_diff_base_a(vv1, vv2)
    if diff <= (g1 + 1) or diff == 0.0:
        return f"❓  Within 𝗚1 {g1} ({diff}%) value1: {vv1} value2: {vv2}", "A", diff
    if diff <= (g3 + 1):
        return f"❓  Within 𝗚2 (less) {g2} | {g3} ({diff}%) value1: {vv1} value2: {vv2}", "B", diff
    if diff <= g3 + 10:
        return f"❓  Within 𝗚3 {g3} ({diff}%) value1: {vv1} value2: {vv2}", "C", diff
    return f"❌ Outside tolerance ({diff}%)  value1: {vv1} value2: {vv2}", "Fail", diff


# ============================================================================
# LOOKUP HELPERS
# ============================================================================

def find_feature_CoreLookUp(
    PL_name: object,
    df: pd.DataFrame,
    feature1: object,
    value1: object,
    value2: object,
) -> str:
    """
    Original lookup logic with cache.
    """
    t0 = time.perf_counter()

    key = (
        str(PL_name).lower(),
        str(feature1).lower(),
        tuple(sorted([str(value1).lower(), str(value2).lower()])),
    )

    if key in LOOKUP_CACHE:
        return LOOKUP_CACHE[key]

    missing = [column for column in CORE_LOOKUP_REQUIRED_COLS if column not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {missing}")

    core_df = df[df["FeatureName"].astype(str).str.contains(feature1, case=False, na=False)]

    has_value1 = core_df[
        core_df["RealGroupString"].astype(str).str.contains(value1, case=False, na=False)
    ]
    has_value2 = core_df[
        core_df["RealGroupString"].astype(str).str.contains(value2, case=False, na=False)
    ]

    parent_ids_with_both = set(has_value1["ParentGroupID"]).intersection(set(has_value2["ParentGroupID"]))
    result = "Found" if parent_ids_with_both else "Not Found"

    LOOKUP_CACHE[key] = result
    log_timing("find_feature_CoreLookUp", time.perf_counter() - t0)
    return result


def build_lookup_index(df: pd.DataFrame) -> Dict[str, List[Tuple[str, str, str]]]:
    missing = [column for column in LOOKUP_INDEX_REQUIRED_COLS if column not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {missing}")

    tmp = strip_dataframe_columns(df)

    if "GroupType" in tmp.columns:
        group_type = tmp["GroupType"].astype(str).str.strip().str.lower()
        core_mask = group_type.isin(["coreid", "core id"])
        if core_mask.any():
            tmp = tmp[core_mask].copy()

    tmp["__feature_norm"] = tmp["FeatureName"].map(normalize_text)
    tmp["__value_norm"] = tmp["RealGroupString"].map(normalize_text)
    tmp["__value_compact"] = tmp["RealGroupString"].map(compact_text)
    tmp["__parent_norm"] = tmp["ParentGroupID"].map(normalize_text)

    index: Dict[str, List[Tuple[str, str, str]]] = {}
    for row in tmp[["__feature_norm", "__value_norm", "__value_compact", "__parent_norm"]].itertuples(index=False):
        feature_norm, value_norm, value_compact, parent_norm = row
        if not feature_norm:
            continue
        index.setdefault(feature_norm, []).append((value_norm, value_compact, parent_norm))

    return index


def find_feature_CoreLookUp_fast(
    index: Dict[str, List[Tuple[str, str, str]]],
    PL_name: object,
    feature1: object,
    value1: object,
    value2: object,
) -> str:
    feature_norm = normalize_text(feature1)
    feature_compact = compact_text(feature1)

    value1_norm = normalize_text(value1)
    value2_norm = normalize_text(value2)
    value1_compact = compact_text(value1)
    value2_compact = compact_text(value2)

    key = (
        normalize_text(PL_name),
        feature_norm,
        tuple(sorted([value1_compact or value1_norm, value2_compact or value2_norm])),
    )

    if key in LOOKUP_CACHE:
        return LOOKUP_CACHE[key]

    if not feature_norm or not value1_norm or not value2_norm:
        LOOKUP_CACHE[key] = "Not Found"
        return "Not Found"

    candidates = list(index.get(feature_norm, []))

    if not candidates:
        for feature_key, rows in index.items():
            feature_key_compact = compact_text(feature_key)
            if (
                (feature_norm and feature_norm in feature_key)
                or (feature_compact and feature_compact == feature_key_compact)
                or (feature_compact and feature_compact in feature_key_compact)
            ):
                candidates.extend(rows)

    parents1: set[str] = set()
    parents2: set[str] = set()

    for value_norm, value_compact, parent in candidates:
        if (value1_norm and value1_norm in value_norm) or (
            value1_compact and (value1_compact == value_compact or value1_compact in value_compact)
        ):
            parents1.add(parent)

        if (value2_norm and value2_norm in value_norm) or (
            value2_compact and (value2_compact == value_compact or value2_compact in value_compact)
        ):
            parents2.add(parent)

    result = "Found" if parents1.intersection(parents2) else "Not Found"
    LOOKUP_CACHE[key] = result
    return result


# ============================================================================
# FILE LOADERS
# ============================================================================

def load_files_from_excel(combined_file_path: str | os.PathLike[str]) -> Dict[str, pd.DataFrame]:
    """
    Load all required files from a single Excel file with the required sheets.
    """
    t0 = time.perf_counter()
    combined_file_path = str(combined_file_path)

    print(f"Loading files from: {combined_file_path}")
    xl_file = pd.ExcelFile(combined_file_path)
    sheet_names = xl_file.sheet_names
    print(f"Available sheets: {sheet_names}")

    files_data: Dict[str, pd.DataFrame] = {}

    if "lookUp" in sheet_names:
        print("Loading LookUp data from lookUp...")
        lookup_df = pd.read_excel(combined_file_path, sheet_name="lookUp", dtype=str)
        lookup_df = strip_dataframe_columns(lookup_df)
        files_data["lookUpFile1"] = lookup_df.copy()
        files_data["lookUpFile2"] = lookup_df.copy()
        print(f"✓ Loaded LookUp data: {len(lookup_df)} rows")
    else:
        raise ValueError("lookUp sheet not found! LookUp data is required.")

    data_sheet_name = "Data"
    print(f"\nLoading all data from sheet: {data_sheet_name}")
    data_df = pd.read_excel(
        combined_file_path,
        sheet_name=data_sheet_name,
        dtype=str,
        keep_default_na=False,
        na_values=[],
    )
    data_df = strip_dataframe_columns(data_df)

    files_data["parametric_file"] = data_df.copy()
    files_data["pakageAndPinout_file"] = data_df.copy()
    files_data["qualification_file"] = data_df.copy()

    print(f"✓ Loaded data: {len(data_df)} rows")
    if len(data_df.columns) > 10:
        print(f"✓ Columns: {', '.join(data_df.columns.tolist()[:10])}...")
    else:
        print(f"✓ Columns: {', '.join(data_df.columns.tolist())}")

    log_timing("load_files_from_excel", time.perf_counter() - t0)
    return files_data


# ============================================================================
# STEP 1: Process files for single PL
# ============================================================================

def process_single_pl(
    pl_name: str,
    cross_df_pl: pd.DataFrame,
    parametric_df: pd.DataFrame,
    pakage_df: pd.DataFrame,
    recipe_df: pd.DataFrame,
) -> pd.DataFrame:
    """Process a single PL through the validation pipeline."""
    t0 = time.perf_counter()

    print(f"\n{'=' * 60}")
    print(f"Processing PL: {pl_name}")
    print(f"{'=' * 60}")

    cross_df_pl = cross_df_pl.copy()
    _ = truncate_sheet_name(pl_name)  # preserved behavior point, even if not used later

    for col in ["PLc", "PLx"]:
        if col in cross_df_pl.columns:
            cross_df_pl[col] = cross_df_pl[col].astype(str).str.strip()

    cross_df_pl["SamePLs"] = np.where(cross_df_pl["PLc"] == cross_df_pl["PLx"], "same", "not")
    cross_df_pl["IsPackaeSimilar"] = "NA"
    cross_df_pl["IsPinoutSimilar"] = "NA"
    cross_df_pl["PackaePinout"] = "NA"

    same_pl_mask = cross_df_pl["SamePLs"] == "same"
    same_pl_indices = cross_df_pl[same_pl_mask].index

    recipe_pl = recipe_df[recipe_df["ZProductValue"].astype(str).str.strip() == pl_name]

    for idx in same_pl_indices:
        plc = cross_df_pl.at[idx, "PLc"]
        subset = recipe_pl[recipe_pl["ZProductValue"] == plc]

        if not subset.empty:
            is_pack = "FALSE" if (subset["IsPackaeSimilar"].str.upper() == "FALSE").all() else "TRUE"
            is_pin = "FALSE" if (subset["IsPinoutSimilar"].str.upper() == "FALSE").all() else "TRUE"

            cross_df_pl.at[idx, "IsPackaeSimilar"] = is_pack
            cross_df_pl.at[idx, "IsPinoutSimilar"] = is_pin
            cross_df_pl.at[idx, "PackaePinout"] = "TRUE" if (is_pack == "TRUE" or is_pin == "TRUE") else "FALSE"

    parametric_parts = (
        set(parametric_df["PartNumber"].values)
        if not parametric_df.empty and "PartNumber" in parametric_df.columns
        else set()
    )
    pakage_parts = (
        set(pakage_df["PartNumber"].values)
        if not pakage_df.empty and "PartNumber" in pakage_df.columns
        else set()
    )

    cross_df_pl["FoundData"] = "FALSE"

    for idx in same_pl_indices:
        part_c = as_text(cross_df_pl.at[idx, "PartNumberC"])
        part_x = as_text(cross_df_pl.at[idx, "PartNumberX"])
        pack_pinout = cross_df_pl.at[idx, "PackaePinout"]

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


# ============================================================================
# STEP 2: Merge data for single PL
# ============================================================================

def find_matching_row(
    df: pd.DataFrame,
    part_value: object,
    company_value: object,
) -> Optional[pd.DataFrame]:
    """Search for matching row in dataframe."""
    if df.empty:
        return None

    stripped_columns = [str(col).strip() for col in df.columns]
    columns_lower = [col.lower() for col in stripped_columns]

    company_col = None
    part_col = None

    for idx, col_lower in enumerate(columns_lower):
        if col_lower in COMPANY_COL_CANDIDATES:
            company_col = df.columns[idx]
        if col_lower in PART_COL_CANDIDATES:
            part_col = df.columns[idx]

    if company_col is None or part_col is None:
        return None

    mask = (
        df[part_col].astype(str).str.strip() == as_text(part_value)
    ) & (
        df[company_col].astype(str).str.strip() == as_text(company_value)
    )

    filtered = df[mask]
    if filtered.empty:
        return None

    row = filtered.iloc[0].drop([part_col, company_col], errors="ignore")
    return row.to_frame().T.reset_index(drop=True)


def _merge_single_part(
    part_number: str,
    company_name: str,
    parametric_df: pd.DataFrame,
    pakage_df: pd.DataFrame,
    qualification_df: pd.DataFrame,
    include_package_pinout: bool,
    include_temp_grade: bool,
) -> pd.DataFrame:
    """
    Merge one part's data while preserving the original merge logic.
    """
    merged = pd.DataFrame({"PartNumber": [part_number], "Company": [company_name], "Comments": [""]})

    if not parametric_df.empty:
        row_data = find_matching_row(parametric_df, part_number, company_name)
        if row_data is not None:
            row_data.columns = deduplicate_columns(row_data.columns)
            merged = pd.concat([merged, row_data], axis=1)
        elif include_package_pinout:
            merged.loc[0, "Comments"] += "Not found in parametric; "

    if include_package_pinout:
        row_data = find_matching_row(pakage_df, part_number, company_name)
        if row_data is not None:
            row_data.columns = deduplicate_columns(row_data.columns)
            merged = pd.concat([merged, row_data], axis=1)
        else:
            merged.loc[0, "Comments"] += "Not found in package/pinout; "

    if not qualification_df.empty:
        row_data = find_matching_row(qualification_df, part_number, company_name)
        if row_data is not None:
            automotive_col = [c for c in row_data.columns if c.strip().lower() == "automotive"]
            merged["Automotive Qualified"] = row_data[automotive_col[0]].iloc[0] if automotive_col else ""

            if include_temp_grade:
                temp_col = [c for c in row_data.columns if c.strip().lower() == "ztemperaturegrade"]
                merged["ZTemperatureGrade"] = row_data[temp_col[0]].iloc[0] if temp_col else ""
        else:
            merged["Automotive Qualified"] = ""
            if include_temp_grade:
                merged["ZTemperatureGrade"] = ""
    else:
        merged["Automotive Qualified"] = ""
        if include_temp_grade:
            merged["ZTemperatureGrade"] = ""

    merged.columns = deduplicate_columns(merged.columns)
    return merged


def merge_single_pl(
    pl_name: str,
    cross_df_pl: pd.DataFrame,
    parametric_df: pd.DataFrame,
    pakage_df: pd.DataFrame,
    qualification_df: pd.DataFrame,
) -> pd.DataFrame:
    """Merge data for a single PL, including Qualification info."""
    t0 = time.perf_counter()

    all_merged_rows: List[pd.DataFrame] = []

    for _, row in cross_df_pl.iterrows():
        part_c = as_text(row["PartNumberC"])
        company_c = as_text(row["CompanyNameC"])
        part_x = as_text(row["PartNumberX"])
        company_x = as_text(row["CompanyNamex"])
        pack_pinout = as_text(row["PackaePinout"])

        include_package_pinout = pack_pinout.lower() in TRUTHY_VALUES

        merged_c = _merge_single_part(
            part_number=part_c,
            company_name=company_c,
            parametric_df=parametric_df,
            pakage_df=pakage_df,
            qualification_df=qualification_df,
            include_package_pinout=include_package_pinout,
            include_temp_grade=False,  # keep original logic
        )

        merged_x = _merge_single_part(
            part_number=part_x,
            company_name=company_x,
            parametric_df=parametric_df,
            pakage_df=pakage_df,
            qualification_df=qualification_df,
            include_package_pinout=include_package_pinout,
            include_temp_grade=True,  # keep original logic
        )

        all_merged_rows.extend([merged_c, merged_x])

    final_df = pd.concat(all_merged_rows, ignore_index=True) if all_merged_rows else pd.DataFrame()

    log_timing(f"merge_single_pl [{pl_name}]", time.perf_counter() - t0)
    return final_df


# ============================================================================
# STEP 3: Validation for single PL
# ============================================================================

def compare_parts(df: pd.DataFrame, value: object, value2: object) -> Dict[str, object]:
    """Compare two parts by AcceptedValue with Excel-style equality."""
    accepted_clean = df["AcceptedValue"].astype(str).str.strip().str.lower()
    value_clean = as_text(value).lower()
    value2_clean = as_text(value2).lower()

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
        len_fc1 = sorted(len(str(x)) for x in fc1)
        len_fc2 = sorted(len(str(x)) for x in fc2)
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

    def prepare_sorted_part(part: pd.DataFrame, feature_codes: List[object]) -> pd.DataFrame:
        out = part.set_index("FeatureCode").loc[feature_codes].reset_index()
        out["NumericValue"] = pd.to_numeric(out["NormalizedValue"], errors="coerce")
        out = out.sort_values(by=["NumericValue"], ascending=True, na_position="last")
        return out

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


def validate_single_pl(
    pl_name: str,
    merged_df: pd.DataFrame,
    recipe_df: pd.DataFrame,
    lookup_values: pd.DataFrame,
    lookup_index: Dict[str, List[Tuple[str, str, str]]],
) -> pd.DataFrame:
    """Validate features for a single PL."""
    t0 = time.perf_counter()

    rules_subset = recipe_df[recipe_df["ZProductValue"].astype(str).str.strip() == pl_name].copy()
    rules_subset = rules_subset[
        (rules_subset["FeaturesType"].str.lower() == "core")
        | (rules_subset["UpgradeFeature"].str.upper() == "TRUE")
        | (rules_subset["IsEqualFeature"].str.upper() == "TRUE")
        | (rules_subset["FeaturesType"].str.lower() == "tolerance")
    ].copy()

    for col in ["G1", "G2", "G3"]:
        if col in rules_subset.columns:
            rules_subset[col] = pd.to_numeric(rules_subset[col], errors="coerce")

    results: List[Dict[str, str]] = []
    overallUpDown = ""
    PinOut_flag = (
        "TRUE"
        if not rules_subset.empty and rules_subset["IsPinoutSimilar"].iloc[0] == "True"
        else "FALSE"
    )

    for i in range(0, len(merged_df), 2):
        if i + 1 >= len(merged_df):
            break

        row1 = merged_df.iloc[i]
        row2 = merged_df.iloc[i + 1]

        feature_statuses: List[str] = []
        feature_grades: List[str] = []
        flags: List[str] = []
        UpDown: List[str] = []

        overallUpDown, maxFlag, minFlag, AutoFlag, TempGrade, ConditionGrade = "", "", "", "", "", []
        count = 1
        Rohs_flag = ""
        temp, Auto = "False", ""

        for _, rule in rules_subset.iterrows():
            result = None
            feature = rule["Features"]
            ftype = str(rule["FeaturesType"]).lower()
            upgrade_flag = str(rule.get("UpgradeFeature", "")).upper()
            equal_flag = str(rule.get("IsEqualFeature", "")).upper()

            v1 = v2 = nunit1 = nunit2 = val1 = val2 = None

            # Pinout check
            if PinOut_flag == "TRUE" and count:
                val1 = as_text(row1.get("NormalizedPinOutName", ""))
                val2 = as_text(row2.get("NormalizedPinOutName", ""))
                count = 0
                if val1 == val2 and val1 is not None:
                    status, grade = "✅ PinOut Match", "A"
                else:
                    status, grade = "❌ Different PinOut", "DiffPinOut"
                _append_feature_result(feature, status, grade, feature_statuses, feature_grades)
                status, grade, overallUpDown = "", "", ""

            if not Rohs_flag:
                Rohs_flag = "TRUE"
                val1 = as_text(row1.get("RoHSComplianceStatus", ""))
                val2 = as_text(row2.get("RoHSComplianceStatus", ""))
                if val1 != val2 and val1 is not None:
                    status, grade = f" ❌ Different Rohs", "Diff_ROHs"
                    flags.append(f"Diff In ROHs {val1} , {val2}")
                else:
                    status, grade = f"RoHSComplianceStatus: ✅ Same Rohs {val1} ,{val2} ", "Diff_ROHs"
                _append_feature_result(feature, status, grade, feature_statuses, feature_grades)
                status, grade, overallUpDown = "", "", ""

            # Core/Equal check
            if ftype == "core":
                if feature not in merged_df.columns:
                    status, grade = f"⚠ Missing column {feature}", "Fail"
                else:
                    val1 = as_text(row1.get(feature, ""))
                    val2 = as_text(row2.get(feature, ""))
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
                            if ftype == "core":
                                flags.append(f"missMatchAtCore {feature}")
                                grade = "DiffPKG" if feature == "Normalized Package Name" else "FailInCore"
                _append_feature_result(feature, status, grade, feature_statuses, feature_grades)

            elif equal_flag == "TRUE":
                if feature not in merged_df.columns:
                    status, grade = f"⚠ Missing column {feature}", "MissingColumn"
                else:
                    val1 = as_text(row1.get(feature, ""))
                    val2 = as_text(row2.get(feature, ""))
                    if val1 == "nan" and val2 == "nan":
                        status, grade = "⚠ Missing Value", "B"
                    elif val1 == "N/A" and val2 == "N/A":
                        status, grade = "⚠ Value is N/A", "N/A"
                    elif ((val2 == "N/A") and (val1 != "N/A")) or ((val1 == "N/A") and (val2 != "N/A")):
                        status, grade = "⚠ Value is N/A", "B"
                    elif val1 == val2:
                        status, grade = "✅ E_Match ", "A"
                    else:
                        found = find_feature_CoreLookUp_fast(lookup_index, pl_name, feature, val1, val2)
                        if found == "Found":
                            status, grade = "✅ Match By LookUp", "A"
                        else:
                            status, grade = f"❌ E_Mismatch ({val1} vs {val2})", "B"
                            if ftype == "core":
                                flags.append(f"missMatchAtEqual {feature}")
                                grade = "FailInCore"
                _append_feature_result(feature, status, grade, feature_statuses, feature_grades)

            elif upgrade_flag == "TRUE":
                if feature not in merged_df.columns:
                    status, grade = f"⚠ Missing column {feature}", "Fail"
                else:
                    val1 = as_text(row1.get(feature, ""))
                    val2 = as_text(row2.get(feature, ""))

                    if feature == "Automotive Qualified":
                        a1 = (val1 or "").strip().lower()
                        a2 = (val2 or "").strip().lower()
                        if a1 == a2:
                            Auto = "Same"
                            UpDown.append(f"Same At {feature},value1 {val1},value2 {val2} ")
                        elif a1 == "yes" and a2 == "no":
                            temp = "False"
                            Auto = "DownGrade"
                            UpDown.append(f"DownGrade At {feature},value1 {val1},value2 {val2} ")
                            overallUpDown = "DownGrade"
                        elif a1 == "no" and a2 == "yes":
                            UpDown.append(f"UpGrade At {feature},value1 {val1},value2 {val2} ")
                            overallUpDown = "UpGrade"
                            Auto = "UpGrade"
                        else:
                            UpDown.append(f"Missing/Unknown Value At {feature},value1 {val1},value2 {val2} ")

                    elif feature == "ZTemperatureGrade":
                        if val2 == val1:
                            TempGrade = "Same"
                            UpDown.append(f"Same At {feature},value1 {val1},value2 {val2} ")
                        else:
                            val1_level = GRADE_HIERARCHY.get(val1, -1)
                            val2_level = GRADE_HIERARCHY.get(val2, -1)

                            if val1_level == -1 and val2_level == -1:
                                UpDown.append(f"Unknown Grade At {feature},value1 {val1},value2 {val2} ")
                                TempGrade = "Unknown"
                            elif val2_level > val1_level:
                                UpDown.append(f"UpGrade At {feature},value1 {val1},value2 {val2} ")
                                TempGrade = "UpGrade"
                            elif val2_level < val1_level:
                                UpDown.append(f"DownGrade At {feature},value1 {val1},value2 {val2} ")
                                TempGrade = "DownGrade"
                            else:
                                TempGrade = "Same"
                                UpDown.append(f"Same At {feature},value1 {val1},value2 {val2} ")

                    elif "Temperature" in feature:
                        if val1 == "nan" or val2 == "nan":
                            UpDown.append(f"Missing Value At {feature},value1 {val1},value2 {val2}")
                            continue
                        elif val1 == val2:
                            UpDown.append(f"Same At {feature},value1 {val1},value2 {val2}")
                            flag = "Same"
                        elif val1 == "N/A" and val2 != "N/A":
                            UpDown.append(f"UpGrade At {feature},value1 {val1},value2 {val2}")
                            flag = "UpGrade"
                        elif val1 != "N/A" and val2 == "N/A":
                            UpDown.append(f"DownGrade At {feature},value1 {val1},value2 {val2}")
                            flag = "DownGrade"
                        else:
                            result = compare_parts(lookup_values, val1, val2)
                            if result["state"] == "Match" and result["nValue"] and result["nValue2"]:
                                try:
                                    v1 = float(result["nValue"][0])
                                    v2 = float(result["nValue2"][0])
                                except Exception:
                                    v1 = v2 = None

                                n_val1 = safe_int(result["nValue"][0])
                                n_val2 = safe_int(result["nValue2"][0])

                                if n_val1 is None or n_val2 is None:
                                    UpDown.append(f"Missing Value At {feature},value1 {val1},value2 {val2}")
                                else:
                                    if "Maximum" in feature:
                                        if n_val2 > n_val1:
                                            flag = "UpGrade"
                                            maxFlag = "UpGrade"
                                        elif n_val2 < n_val1:
                                            flag = "DownGrade"
                                            maxFlag = "DownGrade"
                                        else:
                                            flag = "Same"
                                            maxFlag = "Same"
                                    elif "Minimum" in feature:
                                        if n_val2 < n_val1:
                                            flag = "UpGrade"
                                            minFlag = "UpGrade"
                                        elif n_val2 > n_val1:
                                            flag = "DownGrade"
                                            minFlag = "DownGrade"
                                        else:
                                            flag = "Same"
                                            minFlag = "Same"
                                    else:
                                        UpDown.append(f"Not Found temp {feature},value1 {val1},value2 {val2}")
                                        flag = None

                                    if flag:
                                        UpDown.append(f"{flag} At {feature},value1 {val1},value2 {val2}")
                            elif result["state"] == "Different DetailedValueType":
                                UpDown.append(f"Different DetailedValueType {feature},value1 {val1},value2 {val2}")
                            elif result["state"] == "Different FeatureCode":
                                UpDown.append(f"Different FeatureCode {feature},value1 {val1},value2 {val2}")
                            else:
                                UpDown.append(f"Not Found {feature},value1 {val1},value2 {val2}")

            elif ftype == "tolerance":
                if feature not in merged_df.columns:
                    status, grade = f"⚠ Missing column {feature}", "Fail"
                elif feature == "Pin Pitch_mm" or "Calculated_" in feature:
                    val1 = as_text(row1.get(feature, ""))
                    val2 = as_text(row2.get(feature, ""))
                    if val1 == "nan" and val2 == "nan":
                        status, grade = "⚠ Missing Value", "nan"
                    elif val1 == "N/A" or val2 == "N/A":
                        status, grade = "⚠ Value is N/A", "B"
                    elif val1 == val2:
                        status, grade = "✅T_Match", "A"
                    else:
                        if val1 == "nan" or val2 == "nan" or val1 == "" or val2 == "":
                            status, grade = "⚠ Missing Value", "nan"
                            continue
                        try:
                            vv1 = float(val1)
                            vv2 = float(val2)
                            g1, g2, g3 = rule.get("G1"), rule.get("G2"), rule.get("G3")
                            status, grade, diff = _grade_tolerance_result(vv1, vv2, g1, g2, g3)
                            if grade == "Fail":
                                _append_feature_result(feature, status, grade, feature_statuses, feature_grades)
                                flags.append(f"Outside tolerance {feature}  value1: {vv1} value2: {vv2} ")
                        except Exception as exc:
                            print(f"✗ Error processing PL {pl_name}: {exc}")
                else:
                    val1 = as_text(row1.get(feature, ""))
                    val2 = as_text(row2.get(feature, ""))
                    if val1 == "nan" and val2 == "nan":
                        status, grade = "⚠ Missing Value", "nan"
                    elif ((val1 == "N/A") and (val2 != "N/A")) or ((val1 == "N/A") and (val2 != "N/A")):
                        status, grade = "⚠ Value is N/A", "B"
                    elif val1 == val2:
                        status, grade = "✅T_Match", "A"
                    else:
                        result = compare_parts(lookup_values, val1, val2)

                        if result["state"] == "Match":
                            if result["nValue"] and result["nValue2"]:
                                try:
                                    v1 = float(result["nValue"][0])
                                    v2 = float(result["nValue2"][0])
                                except Exception:
                                    v1 = v2 = None
                                nunit1 = result["nUnit"][0] if result["nUnit"] else None
                                nunit2 = result["nUnit2"][0] if result["nUnit2"] else None

                            if nunit1 and nunit2 and nunit1 != nunit2:
                                status, grade = "Different Unit", "unitFail"
                                flags.append(f"Different Unit {feature}")
                            else:
                                g1, g2, g3 = rule.get("G1"), rule.get("G2"), rule.get("G3")
                                status = ""
                                grade = ""
                                for j in range(min(len(result["nValue"]), len(result["nValue2"]))):
                                    try:
                                        vv1 = float(result["nValue"][j])
                                        vv2 = float(result["nValue2"][j])
                                    except Exception:
                                        continue

                                    status, grade, diff = _grade_tolerance_result(vv1, vv2, g1, g2, g3)
                                    if grade == "Fail":
                                        _append_feature_result(feature, status, grade, feature_statuses, feature_grades)
                                        flags.append(f"Outside tolerance {feature}  value1: {vv1} value2: {vv2} ")

                                if not status:
                                    try:
                                        vv1 = float(val1)
                                        vv2 = float(val2)
                                        diff_status, diff_grade, diff = _grade_tolerance_result(vv1, vv2, g1, g2, g3)
                                        status, grade = diff_status, diff_grade
                                        if grade == "Fail":
                                            _append_feature_result(feature, status, grade, feature_statuses, feature_grades)
                                            flags.append(f"Outside tolerance {feature}  value1: {vv1} value2: {vv2} ")
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

                _append_feature_result(feature, status, grade, feature_statuses, feature_grades)

        overall_grade = determine_overall_grade(feature_grades, Auto, maxFlag, minFlag)
        Auto, maxFlag, minFlag = "", "", ""

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
                "UpDown": " | ".join(UpDown) if UpDown else "",
            }
        )

    log_timing(f"validate_single_pl [{pl_name}]", time.perf_counter() - t0)
    return pd.DataFrame(results)


# ============================================================================
# GRADE DETERMINATION
# ============================================================================

def determine_overall_grade(
    feature_grades: Sequence[str],
    Auto: str,
    maxFlag: str,
    minFlag: str,
) -> str:
    """Determine overall grade based on feature grades and upgrade/downgrade flags."""
    for failure_type, grade in CRITICAL_FAILURE_GRADES.items():
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


def determine_grade_modifier(Auto: str, maxFlag: str, minFlag: str) -> str:
    """Determine the upgrade/downgrade modifier for the grade."""
    return determine_modifier_from_flags(Auto, maxFlag, minFlag)


def determine_modifier_from_flags(Auto: str, maxFlag: str, minFlag: str) -> str:
    """Determine modifier based on min/max flags."""
    upgrade_conditions = [
        maxFlag == "UpGrade" and minFlag in ["Same", "UpGrade"] and Auto in ["Same", "UpGrade"],
        minFlag == "UpGrade" and maxFlag in ["Same", "UpGrade"] and Auto in ["Same", "UpGrade"],
    ]
    if any(upgrade_conditions):
        return "Upgrade"

    downgrade_conditions = [
        maxFlag == "DownGrade" or minFlag == "DownGrade" or Auto == "DownGrade",
        minFlag == "DownGrade" or maxFlag == "DownGrade" or Auto == "DownGrade",
    ]
    if any(downgrade_conditions):
        return "Downgrade"

    mixed_conditions = [
        maxFlag == "DownGrade" and minFlag in ["Same", "UpGrade"],
        minFlag == "DownGrade" and maxFlag in ["Same", "UpGrade"],
    ]
    if any(mixed_conditions):
        return "Downgrade"

    return ""


# ============================================================================
# STEP 4: Organize final output for single PL
# ============================================================================

def organize_single_pl(cross_df_pl: pd.DataFrame, validation_df_pl: pd.DataFrame) -> pd.DataFrame:
    """Organize final output for a single PL."""
    t0 = time.perf_counter()

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
    merge_keys = ["PartNumberC", "CompanyNameC", "PartNumberX", "CompanyNamex"]

    merged = pd.merge(cross_df_pl, validation_df_pl[val_cols], on=merge_keys, how="left")

    unmatched_mask = merged["Grade"].isna()
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
            columns=["Match Feature", "Different Features", "Grade", "UpDown"],
            errors="ignore",
        )

        filled_rows = pd.merge(unmatched_rows, swapped[val_cols], on=merge_keys, how="left")
        merged = pd.concat([matched_rows, filled_rows], ignore_index=True)

    for col in ["Match Feature", "Different Features", "Grade"]:
        swapped_col = f"{col}_swapped"
        if swapped_col in merged.columns:
            merged[col] = merged[col].fillna(merged[swapped_col])
            merged.drop(columns=[swapped_col], inplace=True, errors="ignore")

    def compute_status(row: pd.Series) -> str:
        grade = as_text(row.get("Grade", ""))
        same_pls = as_text(row.get("SamePLs", "")).lower()
        found_data = as_text(row.get("FoundData", "")).upper()
        pack_pinout = as_text(row.get("PackaePinout", ""))

        if grade in ["Drop-in H", "Drop-in A", "Drop-in B", "Drop-in C"]:
            return "Cross"
        if grade in [
            "Not Drop-in",
            "Detailed Value Type Fail Not Drop-in",
            "Unit FAIL Not Drop-in",
            "FeatureCode FAIL Not Drop-in",
        ] and found_data == "TRUE":
            return "Not Cross"
        if same_pls == "not":
            return "Different PLs"
        if found_data == "FALSE" and pack_pinout == "":
            return "Not Found Data"
        return ""

    merged["Status"] = merged.apply(compute_status, axis=1)

    def compare_grades(row: pd.Series) -> str:
        grade = as_text(row.get("Grade", "")).lower()
        cross_grade = as_text(row.get("CrossGrade", "")).lower()

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


# ============================================================================
# MAIN EXECUTION: Process all PLs
# ============================================================================

def process_all_pls_from_combined(
    crosses_file: str | os.PathLike[str],
    combined_data_file: str | os.PathLike[str],
    recipe_file: str | os.PathLike[str],
    lookUpCore: str | os.PathLike[str],
    output_dir: str | os.PathLike[str] = "output",
) -> Optional[pd.DataFrame]:
    """Process all PLs from start to finish using the combined Excel file."""
    t0_total = time.perf_counter()

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    print("\n" + "=" * 60)
    print("LOADING DATA FROM COMBINED FILE")
    print("=" * 60)

    files_data = load_files_from_excel(combined_data_file)

    parametric_df = strip_dataframe_columns(files_data["parametric_file"])
    pakage_df = strip_dataframe_columns(files_data["pakageAndPinout_file"])
    qualification_df = strip_dataframe_columns(files_data["qualification_file"])

    lookup_values = pd.concat(
        [files_data["lookUpFile1"], files_data["lookUpFile2"]],
        ignore_index=True,
    )
    lookup_values = strip_dataframe_columns(lookup_values)

    print("\nLoading main files...")
    t_load = time.perf_counter()

    cross_df = pd.read_excel(crosses_file, dtype=str, keep_default_na=False, na_values=[])
    recipe_df = pd.read_excel(recipe_file, dtype=str, keep_default_na=False, na_values=[])

    cross_df = strip_dataframe_columns(cross_df)
    recipe_df = strip_dataframe_columns(recipe_df)
    recipe_df["ZProductValue"] = recipe_df["ZProductValue"].astype(str).str.strip()

    log_timing("load_crosses_and_recipe", time.perf_counter() - t_load)

    df_lookUpCore = pd.read_excel(lookUpCore, dtype=str, keep_default_na=False, na_values=[])
    df_lookUpCore = strip_dataframe_columns(df_lookUpCore)

    pl_list = cross_df["PLc"].dropna().astype(str).str.strip().unique().tolist()
    print(f"\nFound {len(pl_list)} unique PLs to process")

    all_results: List[pd.DataFrame] = []

    LOOKUP_CACHE.clear()
    lookup_index = build_lookup_index(df_lookUpCore)

    for pl_name in pl_list:
        try:
            t0_pl = time.perf_counter()

            cross_df_pl = cross_df[cross_df["PLc"].astype(str).str.strip() == pl_name].copy()
            if cross_df_pl.empty:
                print(f"⚠ No data for PL: {pl_name}")
                continue

            # Step 1
            cross_df_pl = process_single_pl(pl_name, cross_df_pl, parametric_df, pakage_df, recipe_df)
            cross_df_pl.columns = deduplicate_columns(cross_df_pl.columns)

            # Step 2
            merged_df_pl = merge_single_pl(pl_name, cross_df_pl, parametric_df, pakage_df, qualification_df)

            # Step 3
            validation_df_pl = validate_single_pl(
                pl_name,
                merged_df_pl,
                recipe_df,
                lookup_values,
                lookup_index,
            )

            # Step 4
            final_df_pl = organize_single_pl(cross_df_pl, validation_df_pl)

            pl_output_file = output_dir / f"PL_{pl_name}_result.xlsx"
            print(f"Saving results for PL: {pl_name} to {pl_output_file}")

            t_save = time.perf_counter()
            final_df_pl.to_excel(pl_output_file, index=False)
            log_timing(f"save_excel [{pl_name}]", time.perf_counter() - t_save)

            print(f"✓ Saved results for PL: {pl_name}")
            log_timing(f"TOTAL for PL [{pl_name}]", time.perf_counter() - t0_pl)

            all_results.append(final_df_pl)

        except Exception as exc:
            print(f"✗ Error processing PL {pl_name}: {exc}")
            import traceback

            traceback.print_exc()

    if all_results:
        print("\n✅ Combining all results...")
        t_combine = time.perf_counter()

        combined_df = pd.concat(all_results, ignore_index=True).drop_duplicates()

        combined_output = output_dir / "ALL_PLs_COMBINED.xlsx"
        combined_df.to_excel(combined_output, index=False)

        create_compare_links(
            combined_output,
            output_dir / "ALL_PLs_COMBINED_new.xlsx",
        )
        log_timing("combine_and_save_all", time.perf_counter() - t_combine)

        print(f"\n✅ All PLs processed! Combined results saved to: {combined_output}")
        log_timing("process_all_pls_from_combined [TOTAL]", time.perf_counter() - t0_total)
        print_timing_summary()
        return combined_df

    print("\n⚠ No results to combine")
    log_timing("process_all_pls_from_combined [TOTAL]", time.perf_counter() - t0_total)
    print_timing_summary()
    return None


# ============================================================================
# RECIPE / OUTPUT HELPERS
# ============================================================================

def add_and_sort_features(
    input_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str],
) -> None:
    """Add ZTemperatureGrade rows and sort by ZProductValue + Features."""
    t0 = time.perf_counter()

    df = pd.read_excel(input_path, dtype=str)
    df = strip_dataframe_columns(df)

    required_cols = ["ZProductValue", "Features", "FeaturesType", "Abbreviations"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"❌ Missing required column(s): {missing}")

    new_rows: List[Dict[str, Optional[str]]] = []
    for zvalue in df["ZProductValue"].unique():
        subset = df[df["ZProductValue"] == zvalue]
        has_true = (subset["UpgradeFeature"] == "True").any()
        zUpgradeFeature = "True" if has_true else ""

        new_row = {col: None for col in df.columns}
        new_row["ZProductValue"] = zvalue
        new_row["Features"] = "ZTemperatureGrade"
        new_row["FeaturesType"] = "Tolerance"
        new_row["Abbreviations"] = "Qualified"
        new_row["Source"] = "Qualifications"
        new_row["UpgradeFeature"] = zUpgradeFeature
        new_row["FeatureCompareType"] = "AH"
        new_rows.append(new_row)

    df_extended = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
    df_sorted = df_extended.sort_values(by=["ZProductValue", "Features"], ascending=[True, True])
    df_sorted.to_excel(output_path, index=False)

    log_timing("add_and_sort_features", time.perf_counter() - t0)


def create_compare_links(
    input_file: str | os.PathLike[str],
    output_file: str | os.PathLike[str],
) -> None:
    """Add Compare_link column based on ParIDC and ParIDX."""
    t0 = time.perf_counter()

    df = pd.read_excel(input_file)
    df = strip_dataframe_columns(df)

    required_cols = ["ParIDC", "ParIDX"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s) for compare link: {missing}")

    base_url = "https://app.z2data.com/compare/Parts/alldata?PartIds="
    df["Compare_link"] = base_url + df["ParIDC"].astype(str) + "," + df["ParIDX"].astype(str)
    df.to_excel(output_file, index=False)

    print("✅ Compare links created successfully!")
    log_timing("create_compare_links", time.perf_counter() - t0)


def zip_output_folder(output_dir: str | os.PathLike[str], archive_base: Optional[str | os.PathLike[str]] = None) -> str:
    """Zip an output folder using the standard library."""
    output_dir = Path(output_dir)
    archive_base_path = Path(archive_base) if archive_base else output_dir
    archive_path = shutil.make_archive(
        str(archive_base_path),
        "zip",
        root_dir=str(output_dir.parent),
        base_dir=output_dir.name,
    )
    print(f"✅ Output archived successfully: {archive_path}")
    return archive_path


def try_colab_download(file_path: str | os.PathLike[str]) -> None:
    """Download file automatically when running inside Google Colab."""
    try:
        from google.colab import files  # type: ignore

        files.download(str(file_path))
        print(f"✅ Download started: {file_path}")
    except Exception:
        print(f"ℹ Archive ready at: {file_path}")


# ============================================================================
# EXECUTION
# ============================================================================

def main() -> None:
    crosses_file = "/content/Crosses_not.xlsx"
    combined_data_file = "/content/qacrossp2panalysisinput-2026-04-15t155859075-03566359-1401-4d59-a5e7-e4aae1a87e1e.xlsx"
    recipe_file = "/content/newcrossexproductnewrecipeinputtemplate52-4c496b88-9708-441a-a7bd-077953039f5c.xlsx"
    lookUpCore = "/content/rdcrossstringsexinputtemplate70-0ef6730c-92c8-40c1-8aac-e01b7d0724b8_Removed.xlsx"
    qa_recipe = "QA_recipe.xlsx"
    output_dir = "/content/New_output_by_pl_CROSS_2_2_Missing"

    add_and_sort_features(input_path=recipe_file, output_path=qa_recipe)

    process_all_pls_from_combined(
        crosses_file=crosses_file,
        combined_data_file=combined_data_file,
        recipe_file=qa_recipe,
        lookUpCore=lookUpCore,
        output_dir=output_dir,
    )

    print("\n✅ Process completed successfully!")

    archive_path = zip_output_folder(output_dir, f"{output_dir}")
    try_colab_download(archive_path)

    print("\n✅ Process saved successfully!")


if __name__ == "__main__":
    main()