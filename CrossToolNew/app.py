import streamlit as st
import pandas as pd
import numpy as np
import re
import time
import os
from typing import Dict, List, Tuple, Optional
import io
from datetime import datetime

# ============================================================================
# TIMING UTILITIES
# ============================================================================
_timing_log = []
LOOKUP_CACHE = {}

@st.cache_data
def log_timing(func_name: str, elapsed: float):
    record = {"function": func_name, "elapsed_sec": round(elapsed, 4), "timestamp": str(datetime.now())}
    _timing_log.append(record)
    return record

def get_timing_df():
    if not _timing_log:
        return pd.DataFrame()
    return pd.DataFrame(_timing_log)

def print_timing_summary():
    timing_df = get_timing_df()
    if not timing_df.empty:
        st.subheader("⏱️ Processing Timing Summary")
        st.dataframe(timing_df, use_container_width=True)

# ============================================================================
# ALL HELPER FUNCTIONS (COMPLETE)
# ============================================================================

def safe_int(value):
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return None

def truncate_sheet_name(pl_name):
    return str(pl_name)[:31] if len(str(pl_name)) > 31 else str(pl_name)

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
        return float('inf') if b != 0 else 0.0
    return round(abs(a - b) / abs(a) * 100, 2)

def fc_sort_key(fc):
    match = re.search(r"([VU])(\d+)$", fc)
    if match:
        t, num = match.groups()
        return int(num) * 2 + (0 if t == "V" else 1)
    return float("inf")

def find_feature_CoreLookUp(PL_name, df, feature1, value1, value2):
    t0 = time.perf_counter()
    key = (PL_name, feature1.lower(), tuple(sorted([str(value1).lower(), str(value2).lower()])))
    
    if key in LOOKUP_CACHE:
        elapsed = time.perf_counter() - t0
        log_timing("find_feature_CoreLookUp (cache)", elapsed)
        return LOOKUP_CACHE[key]

    required_cols = ['Product', 'GroupID', 'ParentGroupID', 'RealGroupString', 'GroupType', 'ModifiedDate', 'FeatureName']
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing columns in lookup: {missing}")
        return "Error"

    core_df = df[df['FeatureName'].astype(str).str.contains(feature1, case=False, na=False)]
    has_value1 = core_df[core_df['RealGroupString'].astype(str).str.contains(value1, case=False, na=False)]
    has_value2 = core_df[core_df['RealGroupString'].astype(str).str.contains(value2, case=False, na=False)]
    
    parent_ids_with_both = set(has_value1['ParentGroupID']).intersection(set(has_value2['ParentGroupID']))
    result = "Found" if parent_ids_with_both else "Not Found"
    
    LOOKUP_CACHE[key] = result
    elapsed = time.perf_counter() - t0
    log_timing("find_feature_CoreLookUp", elapsed)
    return result

def build_lookup_index(df):
    index = {}
    for _, row in df.iterrows():
        feature = str(row['FeatureName']).lower()
        value = str(row['RealGroupString']).lower()
        parent = row['ParentGroupID']
        if feature not in index:
            index[feature] = {}
        if value not in index[feature]:
            index[feature][value] = set()
        index[feature][value].add(parent)
    return index

def find_feature_CoreLookUp_fast(index, PL_name, feature1, value1, value2):
    t0 = time.perf_counter()
    key = (PL_name, feature1.lower(), tuple(sorted([str(value1).lower(), str(value2).lower()])))
    if key in LOOKUP_CACHE:
        elapsed = time.perf_counter() - t0
        log_timing("find_feature_CoreLookUp_fast (cache)", elapsed)
        return LOOKUP_CACHE[key]
    
    f = feature1.lower()
    v1 = str(value1).lower()
    v2 = str(value2).lower()
    
    parents1 = index.get(f, {}).get(v1, set())
    parents2 = index.get(f, {}).get(v2, set())
    
    result = "Found" if parents1.intersection(parents2) else "Not Found"
    LOOKUP_CACHE[key] = result
    elapsed = time.perf_counter() - t0
    log_timing("find_feature_CoreLookUp_fast", elapsed)
    return result

def find_matching_row(df, part_value, company_value):
    """Search for matching row in dataframe"""
    if df.empty:
        return None
        
    df.columns = df.columns.str.strip()
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

    mask = (df[part_col].astype(str).str.strip() == part_value) & \
           (df[company_col].astype(str).str.strip() == company_value)

    filtered = df[mask]

    if not filtered.empty:
        row = filtered.iloc[0].drop([part_col, company_col], errors='ignore')
        return row.to_frame().T.reset_index(drop=True)

    return None

# ============================================================================
# CORE PROCESSING FUNCTIONS (COMPLETE & FIXED)
# ============================================================================

@st.cache_data
def load_files_from_excel(uploaded_file):
    """Load files from uploaded Excel file."""
    t0 = time.perf_counter()
    
    xl_file = pd.ExcelFile(io.BytesIO(uploaded_file.read()))
    sheet_names = xl_file.sheet_names
    
    files_data = {}
    
    if "lookUp" in sheet_names:
        lookup_df = pd.read_excel(xl_file, sheet_name="lookUp", dtype=str)
        files_data['lookUpFile1'] = lookup_df.copy()
        files_data['lookUpFile2'] = lookup_df.copy()
        st.success(f"✓ Loaded LookUp data: {len(lookup_df)} rows")
    else:
        st.error("❌ 'lookUp' sheet not found!")
        st.stop()
    
    data_sheet_name = "Data"
    if data_sheet_name in sheet_names:
        data_df = pd.read_excel(xl_file, sheet_name=data_sheet_name, dtype=str,
                               keep_default_na=False, na_values=[])
        files_data['parametric_file'] = data_df.copy()
        files_data['pakageAndPinout_file'] = data_df.copy()
        files_data['qualification_file'] = data_df.copy()
        st.success(f"✓ Loaded data: {len(data_df)} rows")
    else:
        st.error(f"❌ '{data_sheet_name}' sheet not found!")
        st.stop()
    
    log_timing("load_files_from_excel", time.perf_counter() - t0)
    return files_data

def process_single_pl(pl_name, cross_df_pl, parametric_df, pakage_df, recipe_df):
    t0 = time.perf_counter()
    
    for col in ["PLc", "PLx"]:
        if col in cross_df_pl.columns:
            cross_df_pl[col] = cross_df_pl[col].astype(str).str.strip()

    cross_df_pl["SamePLs"] = np.where(
        cross_df_pl["PLc"] == cross_df_pl["PLx"], "same", "not"
    )

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

    parametric_parts = (set(parametric_df["PartNumber"].values)
                       if not parametric_df.empty and "PartNumber" in parametric_df.columns else set())
    pakage_parts = (set(pakage_df["PartNumber"].values)
                   if not pakage_df.empty and "PartNumber" in pakage_df.columns else set())

    cross_df_pl["FoundData"] = "FALSE"

    for idx in same_pl_indices:
        part_c = str(cross_df_pl.at[idx, "PartNumberC"]).strip()
        part_x = str(cross_df_pl.at[idx, "PartNumberX"]).strip()
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

def merge_single_pl(pl_name, cross_df_pl, parametric_df, pakage_df, qualification_df):
    t0 = time.perf_counter()

    all_merged_rows = []

    for idx, row in cross_df_pl.iterrows():
        part_c = str(row["PartNumberC"]).strip()
        company_c = str(row["CompanyNameC"]).strip()
        part_x = str(row["PartNumberX"]).strip()
        company_x = str(row["CompanyNamex"]).strip()
        pack_pinout = str(row["PackaePinout"]).strip()

        PackaePinout_bool = pack_pinout.lower() in ["true", "1", "yes"]

        # --- Part C ---
        merged_c = pd.DataFrame({"PartNumber": [part_c], "Company": [company_c], "Comments": [""]})

        if not parametric_df.empty:
            row_data = find_matching_row(parametric_df, part_c, company_c)
            if row_data is not None:
                row_data.columns = deduplicate_columns(row_data.columns)
                merged_c = pd.concat([merged_c, row_data], axis=1)
            elif PackaePinout_bool:
                merged_c.loc[0, "Comments"] += "Not found in parametric; "

        if PackaePinout_bool:
            row_data = find_matching_row(pakage_df, part_c, company_c
