#############################################
# MODULE: MAPPING UTILITIES
# Purpose: Defines mapping constants and functions
#          for reading/writing the local mapping file.
#############################################

import streamlit as st
import pandas as pd
import os

#############################################
# 2) GLOBAL MULTIPLIER MAPPING + READ MAPPING
#############################################

MULTIPLIER_MAPPING = {
    'y': 1e-24, 'z': 1e-21, 'a': 1e-18, 'f': 1e-15,
    'p': 1e-12, 'n': 1e-9,  'µ': 1e-6,  'm': 1e-3,
    'c': 1e-2,    'h': 1e2,
    'k': 1e3,   'M': 1e6,   'G': 1e9,   'T': 1e12,
    'P': 1e15,  'E': 1e18,  'Z': 1e21,  'Y': 1e24
}

LOCAL_BASE_UNITS = set()  # If you want any built-ins, put them here.

def save_mapping_to_disk(df: pd.DataFrame, filename="mapping.xlsx"):
    """
    Helper: saves the current mapping DataFrame to a local 'mapping.xlsx',
    so the pipeline can read it from disk.
    """
    try:
        df.to_excel(filename, index=False, engine='openpyxl')
        st.write(f"DEBUG: Saved mapping to {filename}")
    except Exception as e:
        st.error(f"Error saving mapping to disk ({filename}): {e}")
        st.stop() # Stop if crucial mapping cannot be saved


def read_mapping_file(mapping_file_path: str):
    """
    Reads the local 'mapping.xlsx', checks for required columns and unknown multipliers.
    Returns tuple: (base_units_set, multiplier_mapping_dict)
    """
    if not os.path.exists(mapping_file_path):
        raise FileNotFoundError(f"Error: Mapping file '{mapping_file_path}' not found. Ensure it was downloaded or saved.")
    try:
        df = pd.read_excel(mapping_file_path)
    except Exception as e:
        raise ValueError(f"Error reading mapping file '{mapping_file_path}': {e}")

    required_cols = {'Base Unit Symbol', 'Multiplier Symbol'}
    if not required_cols.issubset(df.columns):
        raise ValueError(f"Mapping file '{mapping_file_path}' must have columns: {required_cols}")

    # Extract base units, ensuring they are strings and handling potential NaN/None
    base_units = set(str(x).strip() for x in df['Base Unit Symbol'].dropna().unique() if str(x).strip())

    # Extract multipliers mentioned in the file (for checking, not for value mapping)
    file_multipliers = set()
    if 'Multiplier Symbol' in df.columns: # Check if column exists
        file_multipliers = set(str(x).strip() for x in df['Multiplier Symbol'].dropna() if str(x).strip() and pd.notna(x))

    # Get the keys from the hardcoded multiplier mapping
    defined_multiplier_keys = set(MULTIPLIER_MAPPING.keys())

    # Check if multipliers mentioned in the file are defined in our hardcoded map
    undefined_in_file = file_multipliers - defined_multiplier_keys
    if undefined_in_file:
        # Allow multipliers to exist in the sheet even if not in MULTIPLIER_MAPPING for now
        # raise ValueError(f"Undefined multipliers in '{mapping_file_path}': {undefined}")
        st.warning(f"Note: Multipliers found in mapping file but not in hardcoded MULTIPLIER_MAPPING: {undefined_in_file}. They won't be used for normalization factors.")

    st.write(f"DEBUG: Read mapping '{mapping_file_path}'. Base units found: {len(base_units)}. Known normalization multipliers: {len(defined_multiplier_keys)}")

    # Return the set of base units and the hardcoded multiplier mapping dictionary
    return base_units, MULTIPLIER_MAPPING
