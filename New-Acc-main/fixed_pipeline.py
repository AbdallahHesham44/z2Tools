# --- START OF FILE fixed_pipeline.py ---

#############################################
# MODULE: FIXED PROCESSING PIPELINE
# Purpose: Implements the first pipeline logic:
#          - Classifies input strings based on structure.
#          - Extracts value/unit parts according to classification.
#          - Generates structured output rows with codes (e.g., SV-V, RV-Un).
#############################################

import pandas as pd
import re
import gc
import io # Make sure io is imported
import streamlit as st # For error/warning/debug messages
import traceback
# ─── NEW IMPORTS ──────────────────────────────────────────────
from regex_classifier import detect_value_type, build_detailed
from analysis_helpers import extract_identifiers_detailed   # keep this
import re
from regex_classifier import (
    detect_value_type,
    build_detailed,
    initialize_regex_classifier,
    classification_patterns,
    build_mixed_types_detail,          #  ← ADD THIS
    _collect_base_types_in_value_order #  ← and this
)
from config import MAPPING_FILE_LOCAL        # ← path to your mapping file
# ─── helper: break “detail DVT” into (base_class, count) tuples ──────────

_DVT_PARSER = re.compile(
    r"^\s*(?P<tag>.+?)\s*\[\d+][^\]]*]\s*x(?P<cnt>\d+)\s*$"
)


def _parse_dvt_list(detail_dvt: str) -> list[tuple[str, int]]:
    """
    "Single Value with Single Condition [1][1] x1,
     Multiple (Range Value) [1][0] x2"
        →  [("Single Value with Single Condition", 1),
            ("Range Value",                         2)]
    """
    items = []
    for frag in str(detail_dvt).split(","):
        frag = frag.strip()
        if not frag:
            continue
        m = _DVT_PARSER.match(frag)
        if not m:
            continue                          # skip malformed piece
        tag  = m.group("tag").strip()
        cnt  = int(m.group("cnt"))
        # unwrap “Multiple (…)”
        if tag.lower().startswith("multiple (") and tag.endswith(")"):
            tag = tag[len("Multiple ("):-1].strip()
        items.append((tag, cnt))
    return items

# ─── NEW HELPER ───────────────────────────────────────────────
_dvt_re = re.compile(r"\[(?P<m>[^\]]+)]\[(?P<c>[^\]]+)]\s*x(?P<s>\d+)")

def classify_with_regex(raw: str):
    """
    Return the 9-tuple that the fixed pipeline expects, using the regex
    classifier for the first element (category) and counters extracted
    from DetailedValueType.
"""
    if not classification_patterns:                 # first call only
        initialize_regex_classifier(MAPPING_FILE_LOCAL)
    raw = str(raw).strip()
    cls = detect_value_type(raw)             # e.g. "Range Value with Single Condition"
    dvt = build_detailed(raw, cls)           # e.g. "... [1][2] x3"

    m = _dvt_re.search(dvt)
    main_n = m.group("m") if m else "Mixed"
    cond_n = m.group("c") if m else "Mixed"
    sub_n  = int(m.group("s")) if m else 1

    has_range_main          = "Range Value"           in cls
    has_multi_main          = "Multi Value"           in cls
    has_range_condition     = "Range Condition"       in cls
    has_multiple_conditions = "Multiple Conditions"   in cls

    ids = ", ".join(s.strip("()") for s in extract_identifiers_detailed(raw))

    return (
        cls, ids, sub_n, cond_n,
        has_range_main, has_multi_main,
        has_range_condition, has_multiple_conditions,
        main_n
    )

# Import necessary functions and constants from other modules
# NOTE: Assuming mapping_utils.py and analysis_helpers.py exist and contain needed functions
#       EXCEPT for get_desired_order and display_mapping which are defined LOCALLY below.
try:
    from mapping_utils import read_mapping_file, LOCAL_BASE_UNITS
except ImportError as e:
    st.error(f"Failed to import from mapping_utils.py: {e}. Ensure the file exists and is accessible.")
    st.stop()

try:
    # analysis_helpers needed for classification and core parsing logic
    from analysis_helpers import (
        classify_value_type_detailed, split_outside_parens, fix_exceptions,
        MULTIPLIER_MAPPING, extract_numeric_and_unit_analysis,
        remove_parentheses_detailed, extract_identifiers_detailed
    )
except ImportError as e:
     st.error(f"Failed to import from analysis_helpers.py: {e}. Ensure the file exists and is accessible.")
     st.stop()


# --- Helper functions specific to Fixed Pipeline output generation ---
# NOTE: get_desired_order and display_mapping are defined HERE based on user clarification

def get_desired_order():
    """Returns a list defining the preferred order of codes in the output."""
    # --- UPDATED CODES with M0- prefix and 0 suffix/qualifier ---
    return [
        # Number
        "M0-SN-V0", "M0-SN-U0",
        # Single Value
        "M0-SV-V0", "M0-SV-U0",
        
        # Multi Value (within a single entry, dynamic) - Now prefixed with M0-
        "M0-MV-V1", "M0-MV-U1",
        "M0-MV-V2", "M0-MV-U2",
        "M0-MV-V3", "M0-MV-U3",
        "M0-MV-V4", "M0-MV-U4", # Add more if needed
        # Complex Single
        "M0-CX-V0", "M0-CX-U0",
        # Range Value (main) - Min then Max
        "M0-RV-Vn0", "M0-RV-Un0", "M0-RV-Vx0", "M0-RV-Ux0",
        "M0-RN-Vn0", "M0-RN-Un0", # For Range Numbers (Value Min, Unit/ID Min)
        "M0-RN-Vx0", "M0-RN-Ux0", # For Range Numbers (Value Max, Unit/ID Max)
        # Single Condition (Now consistently M0-SC-V0/U0)
        "M0-SC-V0", "M0-SC-U0",
        # Range Condition - Min then Max
        "M0-RC-Vn0", "M0-RC-Un0", "M0-RC-Vx0", "M0-RC-Ux0",
        # Multi-Condition (dynamic) - Now prefixed with M0-
        "M0-MC-V1", "M0-MC-U1",
        "M0-MC-V2", "M0-MC-U2",
        "M0-MC-V3", "M0-MC-U3",
        "M0-MC-V4", "M0-MC-U4", # Add more if needed
        # Unknown/Error Codes (Remain unchanged)
        "UNK-V", "UNK-U", # Maybe these should become M0-UNK-V0? Let's keep original for now.
        "ERR-VAL",
        "ERR-CL",
        "ERR-CHUNK",
        "ERR-PROC",
    ]



def display_mapping(mapping_dict, desired_order, category, main_key):
    """
    Formats the generated code-to-value mapping dictionary into a list of
    dictionary rows suitable for creating the output DataFrame.

    Args:
        mapping_dict (dict): The code map generated by generate_mapping.
        desired_order (list): Preferred order of codes.
        category (str): The classification category for this main_key.
        main_key (str): The original input value string (used for the 'Main Key' column).

    Returns:
        list[dict]: A list of rows, each row being a dictionary with keys
                    "Main Key", "Category", "Attribute", "Code", "Value".
    """
    def get_attribute_from_code(code):
        if "-U" in code:
            return "Unit"
        else:
            return "Value"


    output_rows = []
    processed_codes = set()

    # Add rows for codes present in the mapping, following the desired order
    for code in desired_order:
        if code in mapping_dict:
            attr = get_attribute_from_code(code)
            # Ensure value exists and handle potential None/NaN from parsing steps
            value = mapping_dict[code].get("value", "") # Default to empty string
            if pd.isna(value): value = "" # Convert NaN to empty string

            row = {
                "Main Key": main_key, # Original input string
                "Category": category, # Classification name
                "Attribute": attr,    # Value or Unit (or Info)
                "Code": code,         # The specific code (e.g., SV-V, MV-V1, SC-U0)
                "Value": str(value)   # The parsed value/unit string, ensure string type
            }
            output_rows.append(row)
            processed_codes.add(code)

    # Add any codes found in the mapping but not in the desired order (e.g., MV-V5 if desired_order stops at MV-V4)
    extra_codes = sorted([c for c in mapping_dict if c not in processed_codes])
    for code in extra_codes:
        attr = get_attribute_from_code(code)
        value = mapping_dict[code].get("value", "")
        if pd.isna(value): value = ""

        row = {
            "Main Key": main_key,
            "Category": category,
            "Attribute": attr,
            "Code": code,
            "Value": str(value)
        }
        output_rows.append(row)

    return output_rows


def extract_block_texts(main_key, category_name):
    """
    Extracts logical text blocks from a value string based on its classification category.
    These blocks correspond to the parts that will receive codes (e.g., min value, max value, condition).
    MODIFIED to correctly split Multi Value parts.

    Args:
        main_key (str): The (potentially cleaned) value string.
        category_name (str): The classification string (e.g., "Range Value with Single Condition", "Multi Value with Single Condition").

    Returns:
        list[str]: A list of extracted text blocks in logical order (main parts first, then condition parts).
                   Example for "10A, 20mV @ 5V": ['10A', '20mV', '5V']
    """
    main_key = str(main_key).strip()
    parts = []
    main_part = main_key # Default if no condition
    cond_part = ""     # Default if no condition

    # Handle structures with conditions first
    if " with " in category_name:
        # Split input string by '@' (outside parentheses)
        # Use the imported split_outside_parens from analysis_helpers
        at_split = split_outside_parens(main_key, ['@'])
        main_part = at_split[0].strip() if at_split else ""
        cond_part = "@".join(at_split[1:]).strip() if len(at_split) >= 2 else ""

        # --- Extract blocks from the main part ---
        if category_name.startswith("Range Value"):
            range_split = split_outside_parens(main_part, [' to '])
            parts.extend(p.strip() for p in range_split if p.strip())
        elif category_name.startswith("Multi Value"): # **** FIXED HERE ****
            # SPLIT the main part by comma
            multi_value_blocks = split_outside_parens(main_part, [','])
            parts.extend(p.strip() for p in multi_value_blocks if p.strip()) # Add each value part
        else: # Single Value, Number, or Complex Single
            if main_part: # Add only if non-empty
                 parts.append(main_part)

        # --- Extract blocks from the condition part ---
        if cond_part: # Process only if a condition part exists
            if "Range Condition" in category_name:
                range_split = split_outside_parens(cond_part, [' to '])
                parts.extend(p.strip() for p in range_split if p.strip())
            elif "Multiple Conditions" in category_name:
                cond_blocks = split_outside_parens(cond_part, [','])
                parts.extend(p.strip() for p in cond_blocks if p.strip())
            elif "Single Condition" in category_name:
                parts.append(cond_part) # Add the single condition block

    # Handle structures without conditions (only main_part matters)
    elif category_name == "Range Numbers":
        range_split = split_outside_parens(main_part, [' to '])
        parts.extend(p.strip() for p in range_split if p.strip())
    elif category_name.startswith("Range Value"):
        range_split = split_outside_parens(main_part, [' to '])
        parts.extend(p.strip() for p in range_split if p.strip())
    elif category_name.startswith("Multi Value"): # **** FIXED HERE ****
        # SPLIT the main part by comma
        multi_value_blocks = split_outside_parens(main_part, [','])
        parts.extend(p.strip() for p in multi_value_blocks if p.strip()) # Add each value part
    elif (category_name.startswith("Single Value") or
          category_name.startswith("Number") or
          category_name.startswith("Complex Single")):
        if main_part: # Add only if non-empty
            parts.append(main_part)
    elif category_name.startswith("Multiple"): # Top-level multiple like "1A, 2A" - should be handled by process_single_key splitting
        st.warning(f"DEBUG: extract_block_texts received top-level Multiple category '{category_name}'. This might indicate an issue in process_single_key if called directly.")
        # As a fallback, try splitting by comma anyway
        multi_value_blocks = split_outside_parens(main_part, [','])
        parts.extend(p.strip() for p in multi_value_blocks if p.strip())
    elif category_name == "Empty" or category_name.startswith("Invalid"):
        pass  # Return empty list
    else:
       # st.warning(f"DEBUG: Unknown category '{category_name}' in extract_block_texts. Treating '{main_key}' as single block (if non-empty).")
        if main_part:
             parts.append(main_part)

    # Return only non-empty, stripped parts
    return [p for p in parts if p]


# --- parse_value_unit_identifier (ensure it uses imported helpers correctly) ---
# In fixed_pipeline.py

# ... (imports and other functions) ...

def parse_value_unit_identifier(raw_chunk, base_units, multipliers_dict):
    """
    Parses a text chunk into value and unit parts:
    - Value_processed is the numeric part plus its multiplier, preserving original spacing.
    - Unit is the remainder of the original string, stripping only excess spaces before the unit,
      and ensuring exactly one space before any parenthetical identifier.
    """
    combined_token = raw_chunk  # preserve all original spacing

    # Extract numeric, multiplier, base_unit, and the exact "rest" after the number
    numeric_val, multiplier_symbol, base_unit, normalized_val, error, original_rest = \
        extract_numeric_and_unit_analysis(combined_token, base_units, multipliers_dict)

    # Fallback for unparsable tokens
    if error or numeric_val is None:
        m = re.match(r'([+\-±]?\d*(?:\.\d+)?(?:[eE][+\-]?\d+)?)(.*)', combined_token)
        if m:
            value_str = m.group(1).strip()
            unit_str = m.group(2).strip()
            # Ensure single space before '(' if present
            unit_str = re.sub(r'\s*\(', ' (', unit_str)
            return value_str, unit_str
        return combined_token, ""

    # Build the numeric string, preserving integer vs float
    try:
        f = float(numeric_val)
        numeric_str = str(int(f)) if f.is_integer() else str(f)
    except (ValueError, TypeError):
        numeric_str = str(numeric_val)

    # Append multiplier to numeric_str, preserving original spacing
    if multiplier_symbol and multiplier_symbol != "1":
        if original_rest and original_rest.startswith(' '):
            numeric_str = f"{numeric_str} {multiplier_symbol}"
        else:
            numeric_str = f"{numeric_str}{multiplier_symbol}"

    # Build unit: strip only leading/trailing whitespace from rest
    rest = original_rest or ""
    core = rest.strip()
    # Remove the multiplier prefix if it appears at the start
    if multiplier_symbol and multiplier_symbol != "1" and core.startswith(multiplier_symbol):
        core = core[len(multiplier_symbol):]
    # Ensure exactly one space before any '('
    unit_str = re.sub(r'\s*\(', ' (', core).strip()

    return numeric_str, unit_str




# ... (rest of the file, including process_single_key which calls this function) ...


def get_code_prefixes_for_category(category_name):
    """
    Maps a classification category string to a list of dictionaries,
    where each dictionary defines the *base* codes and attributes for *statically defined*
    logical blocks. Codes now follow the M0-...-0 format.
    Dynamic codes (MV-, MC-) are handled in generate_mapping.

    Args:
        category_name (str): The classification string.

    Returns:
        list[dict]: List of static block definitions. Each dict has 'prefix', 'codes', 'attributes'.
                    'codes' now start with M0- and end with 0 or n0/x0.
    """
    # --- Structures without Conditions ---
 #   is_top_level_multiple = bool(re.match(r"^Multiple\s*\(", category))
 #   normalized_category = category_name
    if category_name.lower().startswith("multiple(") or category_name.lower().startswith("multiple ("):
        category_name = category_name.split("(", 1)[1].rstrip(")").strip()

 #   cat = normalized_category      # local alias to keep code tidy
    if category_name == "Number":
        # Old: ["SN-V", "SN-U"] -> New: ["M0-SN-V0", "M0-SN-U0"]
        return [{"prefix": "SN-", "codes": ["M0-SN-V0", "M0-SN-U0"], "attributes": ["Value", "Unit"]}]
    elif category_name == "Number with Identifier":
        # Old: ["SN-V", "SN-U"] -> New: ["M0-SN-V0", "M0-SN-U0"]
        return [{"prefix": "SN-", "codes": ["M0-SN-V0", "M0-SN-U0"], "attributes": ["Value", "Unit"]}]
    elif category_name == "Single Value":
        # Old: ["SV-V", "SV-U"] -> New: ["M0-SV-V0", "M0-SV-U0"]
        return [{"prefix": "SV-", "codes": ["M0-SV-V0", "M0-SV-U0"], "attributes": ["Value", "Unit"]}]
    elif category_name == "Complex Single":
         # Old: ["CX-V", "CX-U"] -> New: ["M0-CX-V0", "M0-CX-U0"]
         return [{"prefix": "CX-", "codes": ["M0-CX-V0", "M0-CX-U0"], "attributes": ["Value", "Unit"]}]
    elif category_name == "Range Value":
        # Old: ["RV-Vn", "RV-Un"], ["RV-Vx", "RV-Ux"] -> New: ["M0-RV-Vn0", "M0-RV-Un0"], ["M0-RV-Vx0", "M0-RV-Ux0"]
        return [
            {"prefix": "RV-", "codes": ["M0-RV-Vn0", "M0-RV-Un0"], "attributes": ["Value", "Unit"]}, # Min/Start
            {"prefix": "RV-", "codes": ["M0-RV-Vx0", "M0-RV-Ux0"], "attributes": ["Value", "Unit"]}  # Max/End
        ]
    elif category_name == "Range Numbers": # Now means "Num (opt Id) to Num (opt Id)"
    # Vn0 for the first number, Un0 for its identifier (if any)
    # Vx0 for the second number, Ux0 for its identifier (if any)
        return [
            {"prefix": "RN-", "codes": ["M0-RN-Vn0", "M0-RN-Un0"], "attributes": ["Value", "Unit"]}, # Min/Start Num & its ID
            {"prefix": "RN-", "codes": ["M0-RN-Vx0", "M0-RN-Ux0"], "attributes": ["Value", "Unit"]}  # Max/End Num & its ID
    ]
    # --- Structures with Single Condition ---
    elif category_name == "Number with Single Condition":
        # Old: ["SN-V", "SN-U"], ["SC-V", "SC-U"] -> New: ["M0-SN-V0", "M0-SN-U0"], ["M0-SC-V0", "M0-SC-U0"]
        return [
            {"prefix": "SN-", "codes": ["M0-SN-V0", "M0-SN-U0"], "attributes": ["Value", "Unit"]},
            {"prefix": "SC-", "codes": ["M0-SC-V0", "M0-SC-U0"], "attributes": ["Value", "Unit"]}
        ]
    elif category_name == "Single Value with Single Condition":
        # Old: ["SV-V", "SV-U"], ["SC-V", "SC-U"] -> New: ["M0-SV-V0", "M0-SV-U0"], ["M0-SC-V0", "M0-SC-U0"]
        return [
            {"prefix": "SV-", "codes": ["M0-SV-V0", "M0-SV-U0"], "attributes": ["Value", "Unit"]},
            {"prefix": "SC-", "codes": ["M0-SC-V0", "M0-SC-U0"], "attributes": ["Value", "Unit"]}
        ]
    elif category_name == "Multi Value with Single Condition":
        # Static part is just the condition. Old: ["SC-V0", "SC-U0"] -> New: ["M0-SC-V0", "M0-SC-U0"]
        # MV parts handled dynamically by generate_mapping.
        return [
            {"prefix": "SC-", "codes": ["M0-SC-V0", "M0-SC-U0"], "attributes": ["Value", "Unit"]}
        ]
    elif category_name == "Complex Single with Single Condition":
         # Old: ["CX-V", "CX-U"], ["SC-V", "SC-U"] -> New: ["M0-CX-V0", "M0-CX-U0"], ["M0-SC-V0", "M0-SC-U0"]
         return [
             {"prefix": "CX-", "codes": ["M0-CX-V0", "M0-CX-U0"], "attributes": ["Value", "Unit"]},
             {"prefix": "SC-", "codes": ["M0-SC-V0", "M0-SC-U0"], "attributes": ["Value", "Unit"]}
         ]
    elif category_name == "Range Value with Single Condition":
        # Old: ["RV-Vn", "RV-Un"], ["RV-Vx", "RV-Ux"], ["SC-V", "SC-U"] -> New: ["M0-RV-Vn0", ...], ["M0-SC-V0", ...]
        return [
            {"prefix": "RV-", "codes": ["M0-RV-Vn0", "M0-RV-Un0"], "attributes": ["Value", "Unit"]},
            {"prefix": "RV-", "codes": ["M0-RV-Vx0", "M0-RV-Ux0"], "attributes": ["Value", "Unit"]},
            {"prefix": "SC-", "codes": ["M0-SC-V0", "M0-SC-U0"], "attributes": ["Value", "Unit"]}
        ]
    # --- Structures with Range Condition ---
    elif category_name == "Number with Range Condition":
         # Old: ["SN-V", "SN-U"], ["RC-Vn", "RC-Un"], ["RC-Vx", "RC-Ux"] -> New: ["M0-SN-V0", ...], ["M0-RC-Vn0", ...], ["M0-RC-Vx0", ...]
         return [
             {"prefix": "SN-", "codes": ["M0-SN-V0", "M0-SN-U0"], "attributes": ["Value", "Unit"]},
             {"prefix": "RC-", "codes": ["M0-RC-Vn0", "M0-RC-Un0"], "attributes": ["Value", "Unit"]},
             {"prefix": "RC-", "codes": ["M0-RC-Vx0", "M0-RC-Ux0"], "attributes": ["Value", "Unit"]}
         ]
    elif category_name == "Single Value with Range Condition":
         # Old: ["SV-V", "SV-U"], ["RC-Vn", "RC-Un"], ["RC-Vx", "RC-Ux"] -> New: ["M0-SV-V0", ...], ["M0-RC-Vn0", ...], ["M0-RC-Vx0", ...]
         return [
             {"prefix": "SV-", "codes": ["M0-SV-V0", "M0-SV-U0"], "attributes": ["Value", "Unit"]},
             {"prefix": "RC-", "codes": ["M0-RC-Vn0", "M0-RC-Un0"], "attributes": ["Value", "Unit"]},
             {"prefix": "RC-", "codes": ["M0-RC-Vx0", "M0-RC-Ux0"], "attributes": ["Value", "Unit"]}
         ]
    elif category_name == "Complex Single with Range Condition":
          # Old: ["CX-V", "CX-U"], ["RC-Vn", "RC-Un"], ["RC-Vx", "RC-Ux"] -> New: ["M0-CX-V0", ...], ["M0-RC-Vn0", ...], ["M0-RC-Vx0", ...]
          return [
              {"prefix": "CX-", "codes": ["M0-CX-V0", "M0-CX-U0"], "attributes": ["Value", "Unit"]},
              {"prefix": "RC-", "codes": ["M0-RC-Vn0", "M0-RC-Un0"], "attributes": ["Value", "Unit"]},
              {"prefix": "RC-", "codes": ["M0-RC-Vx0", "M0-RC-Ux0"], "attributes": ["Value", "Unit"]}
          ]
    elif category_name == "Range Value with Range Condition":
         # Old: ["RV-Vn", ...], ["RV-Vx", ...], ["RC-Vn", ...], ["RC-Vx", ...] -> New: ["M0-RV-Vn0", ...], etc.
         return [
             {"prefix": "RV-", "codes": ["M0-RV-Vn0", "M0-RV-Un0"], "attributes": ["Value", "Unit"]},
             {"prefix": "RV-", "codes": ["M0-RV-Vx0", "M0-RV-Ux0"], "attributes": ["Value", "Unit"]},
             {"prefix": "RC-", "codes": ["M0-RC-Vn0", "M0-RC-Un0"], "attributes": ["Value", "Unit"]},
             {"prefix": "RC-", "codes": ["M0-RC-Vx0", "M0-RC-Ux0"], "attributes": ["Value", "Unit"]}
         ]
    # --- Structures indicating Multiple Conditions (dynamic MC- needed) ---
    # Only define the static base parts. MC codes are generated dynamically.
    elif category_name == "Number with Multiple Conditions":
         # Old: ["SN-V", "SN-U"] -> New: ["M0-SN-V0", "M0-SN-U0"]
         return [{"prefix": "SN-", "codes": ["M0-SN-V0", "M0-SN-U0"], "attributes": ["Value", "Unit"]}]
    elif category_name == "Single Value with Multiple Conditions":
        # Old: ["SV-V", "SV-U"] -> New: ["M0-SV-V0", "M0-SV-U0"]
        # Note: The old definition here included MC codes statically, which was incorrect.
        # The dynamic generation in `generate_mapping` handles MC codes.
        return [{"prefix":"SV-","codes":["M0-SV-V0","M0-SV-U0"],"attributes":["Value","Unit"]}]
    elif category_name == "Complex Single with Multiple Conditions":
          # Old: ["CX-V", "CX-U"] -> New: ["M0-CX-V0", "M0-CX-U0"]
          return [{"prefix": "CX-", "codes": ["M0-CX-V0", "M0-CX-U0"], "attributes": ["Value", "Unit"]}]
    elif category_name == "Range Value with Multiple Conditions":
         # Old: ["RV-Vn", "RV-Un"], ["RV-Vx", "RV-Ux"] -> New: ["M0-RV-Vn0", ...], ["M0-RV-Vx0", ...]
         return [
             {"prefix": "RV-", "codes": ["M0-RV-Vn0", "M0-RV-Un0"], "attributes": ["Value", "Unit"]},
             {"prefix": "RV-", "codes": ["M0-RV-Vx0", "M0-RV-Ux0"], "attributes": ["Value", "Unit"]}
         ]
    elif category_name == "Multi Value with Multiple Conditions":
        return [
            {"prefix": "MV-",
             "codes": ["M0-MV-V1", "M0-MV-U1"],
             "attributes": ["Value", "Unit"]},
            {"prefix": "MV-",
             "codes": ["M0-MV-V2", "M0-MV-U2"],
             "attributes": ["Value", "Unit"]}
        ]


    # --- Multi Value (Top Level) - Fallback ---
    # This category definition is less critical as process_single_key handles M<n> prefixing.
    elif category_name.startswith("Multiple"):
         st.warning(f"DEBUG: get_code_prefixes_for_category received top-level Multiple category '{category_name}'. Static definitions here are less relevant; M<n> prefixes applied by caller. Using fallback UNK codes.")
         return [{"prefix": "UNK-", "codes": ["UNK-V", "UNK-U"], "attributes": ["Value", "Unit"]}]

    # --- Empty / Invalid / Unknown / Error --- (Keep error codes as is)
    elif category_name in ["Empty", "Invalid/Empty Structure", "Classification Error", "Unknown Chunk", "Unknown", "Processing Error", "Unknown with Condition", "Unknown Multi Value", "Unknown Range", "Unknown Single"]:
         return [{"prefix": "ERR-", "codes": ["ERR-VAL"], "attributes": ["Value"]}]

    # --- Fallback for completely unknown categories ---
    else:
        st.warning(f"Unknown category '{category_name}' in get_code_prefixes_for_category. Using default UNK- codes.")
        # Maybe return M0-UNK-V0? Let's keep original UNK for now.
        return [{"prefix": "UNK-", "codes": ["UNK-V", "UNK-U"], "attributes": ["Value", "Unit"]}]


# --- fill_mapping_for_part (minor robustness checks added, core logic unchanged) ---
def fill_mapping_for_part(part_tuple, block_info):
    """
    Fills the mapping dictionary for a single parsed part (value, unit tuple)
    using the code/attribute definitions from block_info.
    """
    (val_str, base_unit_str) = part_tuple
    result = {}
    codes = block_info.get("codes", [])
    attributes = block_info.get("attributes", [])
    prefix = block_info.get("prefix", "ERR") # Default prefix if missing

    try:
        # Ensure codes and attributes are lists even if None/missing
        if codes is None: codes = []
        if attributes is None: attributes = []

        # Find the indices for 'Value' and 'Unit' attributes
        value_idx = attributes.index("Value") if "Value" in attributes else -1
        unit_idx = attributes.index("Unit") if "Unit" in attributes else -1

        # Assign value if attribute and code exist
        if value_idx != -1 and value_idx < len(codes):
            value_code = codes[value_idx]
            result[value_code] = {"value": val_str if val_str is not None else ""} # Ensure value isn't None
        # Handle cases like ERR-VAL where only one code/attribute might exist
        elif len(codes) == 1 and len(attributes) == 1 and attributes[0] == "Value":
             result[codes[0]] = {"value": val_str if val_str is not None else ""}
        # Don't assign if Value attribute exists but code doesn't correspond
        elif value_idx != -1:
              # Commenting out warning for cleaner execution unless debugging:
              # st.warning(f"Block info for prefix '{prefix}' has 'Value' attribute at index {value_idx} but mismatched codes list length {len(codes)}.")
              pass


        # Assign unit if attribute and code exist
        if unit_idx != -1 and unit_idx < len(codes):
            unit_code = codes[unit_idx]
            # Store empty string if base_unit_str is empty or None
            result[unit_code] = {"value": base_unit_str if base_unit_str else ""}
        # Don't assign if Unit attribute exists but code doesn't correspond
        elif unit_idx != -1:
             # Commenting out warning for cleaner execution unless debugging:
             # st.warning(f"Block info for prefix '{prefix}' has 'Unit' attribute at index {unit_idx} but mismatched codes list length {len(codes)}.")
             pass


    except Exception as e:
         st.error(f"Error filling mapping for block {block_info} with part {part_tuple}: {e}")
         # Create generic error codes as fallback for this specific part
         value_code_err = f"{prefix}-V_ERR_{value_idx if 'value_idx' in locals() else 'X'}"
         unit_code_err = f"{prefix}-U_ERR_{unit_idx if 'unit_idx' in locals() else 'X'}"
         result[value_code_err] = {"value": val_str if val_str is not None else "ERROR_VAL"}
         result[unit_code_err] = {"value": base_unit_str if base_unit_str else "ERROR_UNIT"}

    return result


def generate_mapping(parsed_parts, category_name):
    """
    Generates the complete code-to-value mapping dictionary for a given value string,
    based on its parsed parts and classification category. Handles dynamic codes
    for multiple main values (MV-) and multiple conditions (MC-), applying the
    base M0- prefix.
    """
    # 1) Pull in your static block definitions (these now have M0- prefix)
    base_blocks = get_code_prefixes_for_category(category_name)
    mapping = {}

    # 2) Assign static blocks to the first N parts
    for idx, block_info in enumerate(base_blocks):
        if idx >= len(parsed_parts):
            break
        mapping.update(fill_mapping_for_part(parsed_parts[idx], block_info))

    # 3) Decide which dynamic prefix to use (MC- takes priority over MV-)
    if "Multiple Conditions" in category_name:
        dynamic_prefix_base = "MC-"
    elif category_name.startswith("Multi Value"):
        dynamic_prefix_base = "MV-"
    else:
        dynamic_prefix_base = None

    # 4) Hand out dynamic codes for the remaining parts (if any), PREPENDING M0-
    if dynamic_prefix_base:
        counter = 1
        for part in parsed_parts[len(base_blocks):]:
            value_code = f"M0-{dynamic_prefix_base}V{counter}"  # e.g. M0-MC-V1 or M0-MV-V1
            unit_code  = f"M0-{dynamic_prefix_base}U{counter}"  # e.g. M0-MC-U1 or M0-MV-U1
            dyn_block = {
                "prefix": dynamic_prefix_base,
                "codes":  [value_code, unit_code],
                "attributes": ["Value", "Unit"]
            }
            mapping.update(fill_mapping_for_part(part, dyn_block))
            counter += 1

    return mapping



'''
# --- process_single_key (Relies on updated helpers, includes fallback classification) ---
def process_single_key(main_key: str, base_units, multipliers_dict):
    """
    Processes a single 'Value' string entry: classifies it, extracts blocks,
    parses them, generates codes following the M0-/M<n>- logic,
    and returns structured output rows.
    """
    main_key_original = main_key
    main_key_clean = fix_exceptions(str(main_key).strip())

    try:
        (category, identifiers, sub_value_count, final_cond_item_count,
         has_range_main, has_multi_main, has_range_condition,
         has_multiple_conditions, final_main_item_count
        ) = classify_with_regex(main_key_clean)

        if not category or category in ["Empty", "Invalid/Empty Structure"]:
             if '@' in main_key_clean: category = "Unknown with Condition"
             elif ',' in main_key_clean: category = "Unknown Multi Value"
             elif ' to ' in main_key_clean: category = "Unknown Range"
             else: category = "Unknown Single"
             st.warning(f"Detailed classification resulted empty/invalid for '{main_key_clean}'. Using basic fallback category: '{category}'.")

    except Exception as e:
        st.error(f"Error during detailed classification for '{main_key_clean}': {e}")
        st.error(traceback.format_exc())
        category = "Classification Error"
        sub_value_count = 1
        identifiers = ""
# ─────────────────────────────────────────────────────────────────────
# SPECIAL handling for overall “Mixed Types”
# ─────────────────────────────────────────────────────────────────────
mixed_dvt = None
if category_for_struct == "Mixed Types":
    # Grab the breakdown from the regex classifier
    mixed_dvt = build_mixed_types_detail(main_key_cleaned_for_struct_class,
                                         MixedTypesDetail := ", ".join(
                                             _collect_base_types_in_value_order(
                                                 main_key_cleaned_for_struct_class
                                             )
                                         ))
    # mixed_dvt now looks like
    # "Single Value with Single Condition [1][1] x1, Range Value with Range Condition [1][1] x1"

    # Overwrite the generic counters with the real per-class pieces
    parsed = _parse_dvt_list(mixed_dvt)
    # example: [("Single Value with Single Condition", 1),
    #           ("Range Value with Range Condition", 1)]

    # we now build the list of chunks the same way the normal branch would,
    # but using the parsed class names and counts:
    chunks_to_process_individually = []
    for base_cls, how_many in parsed:
        chunks_to_process_individually.extend([base_cls] * how_many)

    # let the rest of the function know we supply the chunks ourselves
    is_top_level_multiple_type = True
else:
   

    # --- Handle Classification Error or Unknown (Keep error codes as is) ---
    if category in ["Unknown", "Classification Error", "Unknown with Condition", "Unknown Multi Value", "Unknown Range", "Unknown Single"]:
        error_value = "Could not classify structure." if category.startswith("Unknown") else f"Classification Error: Check logs."
        error_code = "ERR-CL" if category != "Classification Error" else "ERR-VAL" # Use ERR-VAL for actual errors
        error_map = {error_code: {"value": error_value}}
        desired_order_err = get_desired_order() # Gets M0- list, but error codes are unaffected
        return display_mapping(error_map, desired_order_err, category, main_key_original)

    all_output_rows = []
    is_top_level_multiple = category.startswith("Multiple ")

    if is_top_level_multiple:
        # --- Multi-Chunk Case: Apply M<n>- prefix, replacing M0- ---
        chunks = []
        # ... (chunking logic remains the same as before) ...
        is_grouped_pair_category = category in [
            "Multiple Single Value with Multiple Conditions",
            "Multiple Number with Multiple Conditions"
        ]
        raw_chunks = split_outside_parens(main_key_clean, [','])
        raw_chunks = [chk.strip() for chk in raw_chunks if chk.strip()]
        
        # ── NEW: fuse comma pieces that don’t contain an ‘@’ back to the
        #         previous chunk, so “, 1 A” stays part of the @-clause ──
        glued = []
        for part in raw_chunks:
            if '@' in part:                  # starts a new main-value group
                glued.append(part)
            else:                            # no ‘@’ → it’s an extra condition
                if glued:
                    glued[-1] = f"{glued[-1]}, {part}"
                else:                        # (shouldn’t happen) leading orphan
                    glued.append(part)
        raw_chunks = glued


        if sub_value_count is None: sub_value_count = 0
        if is_grouped_pair_category and sub_value_count > 0 and len(raw_chunks) > sub_value_count and len(raw_chunks) % sub_value_count == 0:
            items_per_chunk = len(raw_chunks) // sub_value_count
            current_grouped_chunk = []
            for i, raw_chunk in enumerate(raw_chunks):
                current_grouped_chunk.append(raw_chunk)
                if (i + 1) % items_per_chunk == 0:
                    chunks.append(", ".join(current_grouped_chunk))
                    current_grouped_chunk = []
            if current_grouped_chunk:
                 chunks.append(", ".join(current_grouped_chunk))
        else:
            chunks = raw_chunks

        if not chunks:
             st.error(f"Failed to find processable chunks for multi-value key: '{main_key_original}' (Category: {category})")
             err_map = {"ERR-CHUNK": {"value": "No chunks found for multi-value key."}}
             return display_mapping(err_map, get_desired_order(), "Processing Error", main_key_original)

        for idx, chunk in enumerate(chunks):
             chunk_prefix = f"M{idx+1}-" # Prefix like M1-, M2-

             try:
                 (chunk_cat, chunk_ids, _, _, _, _, _, _, _) = classify_with_regex(chunk)
                 if not chunk_cat or chunk_cat in ["Empty", "Invalid/Empty Structure"]:
                     chunk_cat = "Unknown Chunk"

                 if chunk_cat != "Unknown Chunk":
                     block_texts = extract_block_texts(chunk, chunk_cat)
                     parsed_parts = []
                     for bt in block_texts:
                         part_val, part_unit = parse_value_unit_identifier(bt, base_units, multipliers_dict)
                         parsed_parts.append((part_val, part_unit))

                     # Generate mapping for the chunk (yields M0- prefixed codes)
                     code_map_chunk_base = generate_mapping(parsed_parts, chunk_cat)

                     # --- Apply the M<n>- prefix, REPLACING M0- ---
                     prefixed_code_map_final = {}
                     for code, val in code_map_chunk_base.items():
                         if code.startswith("M0-"):
                             new_code = chunk_prefix + code[3:] # Replace "M0-" with "M<n>-"
                         elif code.startswith("ERR-") or code.startswith("UNK-"):
                             new_code = code # Keep error/unknown codes as is
                         else:
                             # Fallback: Prefix if it doesn't start M0- (shouldn't happen often)
                             st.warning(f"Unexpected code '{code}' found in chunk mapping for '{chunk}'. Prefixing with {chunk_prefix}.")
                             new_code = chunk_prefix + code
                         prefixed_code_map_final[new_code] = val
                     # --- End M<n>- Prefix Application ---

                     # --- Generate the desired order list with M<n>- replacing M0- ---
                     desired_order_standard = get_desired_order() # Gets the M0- list
                     desired_order_final = []
                     for code in desired_order_standard:
                         if code.startswith("M0-"):
                             desired_order_final.append(chunk_prefix + code[3:])
                         elif code.startswith("ERR-") or code.startswith("UNK-"):
                             desired_order_final.append(code) # Keep error/unknown codes
                         else:
                             desired_order_final.append(chunk_prefix + code) # Fallback prefixing
                     # --- End Desired Order Generation ---

                     chunk_rows = display_mapping(prefixed_code_map_final, desired_order_final, chunk_cat, main_key_original)
                     all_output_rows.extend(chunk_rows)
                 else:
                      # Handle error for this specific chunk (Keep error codes as is)
                      err_map = {f"ERR-CHK": {"value": f"Could not classify chunk {idx+1}: '{chunk}'"}} # No M<n> prefix on error code itself
                      all_output_rows.extend(display_mapping(err_map, get_desired_order(), f"Error in Chunk {idx+1}", main_key_original))

             except Exception as e:
                 st.error(f"Error processing chunk '{chunk}' from '{main_key_original}': {e}")
                 st.error(traceback.format_exc())
                 err_map = {"ERR-PROC": {"value": f"Processing error in chunk {idx+1}: Check logs."}} # No M<n> prefix
                 all_output_rows.extend(display_mapping(err_map, get_desired_order(), f"Error in Chunk {idx+1}", main_key_original))
        return all_output_rows

    else:
        # --- Single Chunk Case: Use M0- prefixed codes directly ---
        try:
            block_texts = extract_block_texts(main_key_clean, category)
            parsed_parts = []
            for bt in block_texts:
                part_val, part_unit = parse_value_unit_identifier(bt, base_units, multipliers_dict)
                parsed_parts.append((part_val, part_unit))

            # Generate mapping (yields M0- prefixed codes)
            code_map = generate_mapping(parsed_parts, category)
            # Get desired order (yields M0- prefixed list)
            desired_order_standard = get_desired_order()
            # Display uses the M0- codes directly
            single_rows = display_mapping(code_map, desired_order_standard, category, main_key_original)
            return single_rows
        except Exception as e:
             st.error(f"Error processing single key '{main_key_original}' (Category: {category}): {e}")
             st.error(traceback.format_exc())
             err_map = {"ERR-PROC": {"value": f"Processing error: Check logs."}} # Keep error code
             return display_mapping(err_map, get_desired_order(), "Processing Error", main_key_original)


# ...(rest of fixed_pipeline.py - process_fixed_pipeline_bytes, display_mapping etc)...

# --- END OF FILE fixed_pipeline.py ---

'''
def process_single_key(main_key: str, base_units, multipliers_dict):
    """
    Processes a single 'Value' string entry:
      1. Classifies it (detailed + fallback heuristics)
      2. Extracts block texts and parses value/unit pairs
      3. Generates M-codes (M0- for single, M<n>- for multi-chunks)
      4. Returns the structured rows expected by the downstream display layer
    """
#    import traceback
 #   import streamlit as st

    main_key_original = main_key
    main_key_clean = fix_exceptions(str(main_key).strip())

    # ───────────────────────────────────────────────────────────────
    # 1) Primary detailed classification (regex-based)
    # ───────────────────────────────────────────────────────────────
    try:
        (category, identifiers, sub_value_count, final_cond_item_count,
         has_range_main, has_multi_main, has_range_condition,
         has_multiple_conditions, final_main_item_count,
        ) = classify_with_regex(main_key_clean)
        norm_cat  = re.sub(r"\s+", " ", category).strip()   # collapse stray whitespace
        norm_cat  = re.sub(r"(?i)^Multiple\(", "Multiple (", norm_cat)  # ensure one space after “Multiple”
        category  = norm_cat
        is_top_level_multiple = bool(re.match(r"^Multiple\s*\(", category, re.I))



        
        #is_top_level_multiple = norm_cat.startswith("Multiple (")     # or keep the regex test on norm_cat
        # Fallback if the detailed classifier returns nothing usable
        if not category or category in ["Empty", "Invalid/Empty Structure"]:
            if "@" in main_key_clean:
                category = "Unknown with Condition"
            elif "," in main_key_clean:
                category = "Unknown Multi Value"
            elif " to " in main_key_clean:
                category = "Unknown Range"
            else:
                category = "Unknown Single"
            st.warning(
                f"Detailed classification was empty/invalid for "
                f"'{main_key_clean}'. Using fallback category: '{category}'."
            )

    except Exception as e:
        st.error(f"Error during detailed classification for '{main_key_clean}': {e}")
        st.error(traceback.format_exc())
        category = "Classification Error"
        sub_value_count = 1
        identifiers = ""

    # ───────────────────────────────────────────────────────────────
    # 2) OPTIONAL secondary structure classification
    #    (needed for “Mixed Types”; falls back gracefully if the
    #     supporting helper isn’t present in this module)
    # ───────────────────────────────────────────────────────────────
    try:
        # Expected helper signature:
        #   classify_structure(text) ➜ (structure_category, cleaned_text)
        category_for_struct, main_key_cleaned_for_struct_class = classify_structure(
            main_key_clean
        )
    except Exception:
        # Helper missing or failed – keep using the primary category
        category_for_struct = category
        main_key_cleaned_for_struct_class = main_key_clean

    # ───────────────────────────────────────────────────────────────
    # 3) Mixed-Types branch (delegates to per-type processing)
    # ───────────────────────────────────────────────────────────────
    mixed_dvt = None
    if category_for_struct == "Mixed Types":
        # Build a breakdown such as:
        # "Single Value with Single Condition [1][1] x1,
        #  Range Value with Range Condition [1][1] x1"
        mixed_dvt = build_mixed_types_detail(
            main_key_cleaned_for_struct_class,
            MixedTypesDetail := ", ".join(
                _collect_base_types_in_value_order(
                    main_key_cleaned_for_struct_class
                )
            ),
        )

        # Parse that breakdown into a list of (base_class, count) tuples
        parsed = _parse_dvt_list(mixed_dvt)

        # Flatten into a list where each item is an individual chunk to run
        chunks_to_process_individually = []
        for base_cls, how_many in parsed:
            chunks_to_process_individually.extend([base_cls] * how_many)

        # Down-stream logic treats this flag exactly like the
        # ordinary multi-value flag
        is_top_level_multiple = True
 

    # ───────────────────────────────────────────────────────────────
    # 4) Handle “Unknown / Error” categories early-exit
    # ───────────────────────────────────────────────────────────────
    if category in [
        "Unknown",
        "Classification Error",
        "Unknown with Condition",
        "Unknown Multi Value",
        "Unknown Range",
        "Unknown Single",
    ]:
        error_value = (
            "Could not classify structure."
            if category.startswith("Unknown")
            else "Classification Error: Check logs."
        )
        error_code = "ERR-CL" if category != "Classification Error" else "ERR-VAL"
        error_map = {error_code: {"value": error_value}}
        desired_order_err = get_desired_order()  # leaves error codes alone
        return display_mapping(error_map, desired_order_err, category, main_key_original)

    # ───────────────────────────────────────────────────────────────
    # 5) MULTI-CHUNK path  (category starts with “Multiple …” OR
    #                       we came through the Mixed-Types branch)
    # ───────────────────────────────────────────────────────────────
    if is_top_level_multiple:
        all_output_rows = []

        # a) Split into raw chunks on top-level commas
        raw_chunks = split_outside_parens(main_key_clean, [","])
        raw_chunks = [c.strip() for c in raw_chunks if c.strip()]

        # b) Glue trailing “, 1 A”-style pieces (which belong to the
        #    previous ‘@…’ chunk) back onto that chunk
        glued_chunks = []
        for part in raw_chunks:
             if not glued_chunks:
                 glued_chunks.append(part)           # always keep the first piece
                 continue
             if "@" in part:                         # starts a new main-value group
                 glued_chunks.append(part)
             else:                                   # extra text ­→ glue **only** if
                 if "@" in glued_chunks[-1]:         # the previous chunk already has @
                     glued_chunks[-1] = f"{glued_chunks[-1]}, {part}"
                 else:                               # otherwise it’s another value
                     glued_chunks.append(part)
        raw_chunks = glued_chunks

        # c) If the detailed classifier told us there are paired
        #    sub-values, fuse the correct number together
        is_grouped_pair_category = category in [
            "Multiple Single Value with Multiple Conditions",
            "Multiple Number with Multiple Conditions",
        ]
        chunks = []
        if (
            is_grouped_pair_category
            and sub_value_count
            and len(raw_chunks) > sub_value_count
            and len(raw_chunks) % sub_value_count == 0
        ):
            items_per_chunk = len(raw_chunks) // sub_value_count
            grp = []
            for i, piece in enumerate(raw_chunks, 1):
                grp.append(piece)
                if i % items_per_chunk == 0:
                    chunks.append(", ".join(grp))
                    grp = []
            if grp:
                chunks.append(", ".join(grp))
        else:
            chunks = raw_chunks

        if not chunks:
            st.error(
                f"Failed to extract processable chunks for "
                f"'{main_key_original}' (Category: {category})"
            )
            err_map = {
                "ERR-CHUNK": {"value": "No chunks found for multi-value key."}
            }
            return display_mapping(
                err_map, get_desired_order(), "Processing Error", main_key_original
            )

        # d) Process each chunk individually
        for idx, chunk in enumerate(chunks, 1):
            chunk_prefix = f"M{idx}-"  # M1-, M2-, …

            try:
                (chunk_cat, chunk_ids, *_dummy) = classify_with_regex(chunk)
                if chunk_cat in ["", "Empty", "Invalid/Empty Structure", None]:
                    chunk_cat = "Unknown Chunk"

                if chunk_cat != "Unknown Chunk":
                    block_texts = extract_block_texts(chunk, chunk_cat)
                    parsed_parts = [
                        parse_value_unit_identifier(bt, base_units, multipliers_dict)
                        for bt in block_texts
                    ]

                    # Map from (M0- codes) ➜ value dict
                    base_code_map = generate_mapping(parsed_parts, chunk_cat)

                    # ---- Rewrite codes so they start with M<n>-
                    prefixed_map = {}
                    for code, val in base_code_map.items():
                        if code.startswith("M0-"):
                            new_code = chunk_prefix + code[3:]
                        elif code.startswith(("ERR-", "UNK-")):
                            new_code = code
                        else:
                            st.warning(
                                f"Unexpected code '{code}' while processing "
                                f"chunk '{chunk}'.  Prefixing with {chunk_prefix}"
                            )
                            new_code = chunk_prefix + code
                        prefixed_map[new_code] = val
                    # ---- end prefix rewrite

                    # Desired order list with the same transformation
                    desired_order_prefixed = []
                    for code in get_desired_order():
                        if code.startswith("M0-"):
                            desired_order_prefixed.append(chunk_prefix + code[3:])
                        elif code.startswith(("ERR-", "UNK-")):
                            desired_order_prefixed.append(code)
                        else:
                            desired_order_prefixed.append(chunk_prefix + code)

                    all_output_rows.extend(
                        display_mapping(
                            prefixed_map,
                            desired_order_prefixed,
                            chunk_cat,
                            main_key_original,
                        )
                    )

                else:
                    err_map = {
                        "ERR-CHK": {
                            "value": f"Could not classify chunk {idx}: '{chunk}'"
                        }
                    }
                    all_output_rows.extend(
                        display_mapping(
                            err_map,
                            get_desired_order(),
                            f"Error in Chunk {idx}",
                            main_key_original,
                        )
                    )

            except Exception as e:
                st.error(
                    f"Exception while processing chunk '{chunk}' "
                    f"from '{main_key_original}': {e}"
                )
                st.error(traceback.format_exc())
                err_map = {
                    "ERR-PROC": {
                        "value": f"Processing error in chunk {idx}: Check logs."
                    }
                }
                all_output_rows.extend(
                    display_mapping(
                        err_map,
                        get_desired_order(),
                        f"Error in Chunk {idx}",
                        main_key_original,
                    )
                )

        return all_output_rows

    # ───────────────────────────────────────────────────────────────
    # 6) SINGLE-CHUNK path  (ordinary “Single …” categories)
    # ───────────────────────────────────────────────────────────────
    try:
        block_texts = extract_block_texts(main_key_clean, category)
        parsed_parts = [
            parse_value_unit_identifier(bt, base_units, multipliers_dict)
            for bt in block_texts
        ]

        code_map = generate_mapping(parsed_parts, category)
        desired_order = get_desired_order()

        return display_mapping(
            code_map, desired_order, category, main_key_original
        )

    except Exception as e:
        st.error(
            f"Exception while processing single key '{main_key_original}' "
            f"(Category: {category}): {e}"
        )
        st.error(traceback.format_exc())
        err_map = {"ERR-PROC": {"value": "Processing error: Check logs."}}
        return display_mapping(
            err_map, get_desired_order(), "Processing Error", main_key_original
        )

# --- process_fixed_pipeline_bytes (Add fillna before astype) ---
def process_fixed_pipeline_bytes(file_bytes: bytes, mapping_file_path: str):
    """
    Runs the first pipeline (fixed processing) on the input Excel file bytes.
    Reads mapping configuration from the specified local path.
    Produces processed data as structured rows with codes.

    Args:
        file_bytes (bytes): The content of the uploaded Excel file.
        mapping_file_path (str): Path to the local 'mapping.xlsx' file.

    Returns:
        pd.DataFrame or None: DataFrame containing the processed rows, or None on failure.
    """
    print("Running Fixed Processing Pipeline...")
    st.write("DEBUG: Starting Fixed Processing Pipeline...")
    all_processed_rows = []

    try:
        # Read mapping file from the provided disk path
        base_units_from_file, multipliers_dict = read_mapping_file(mapping_file_path)
        combined_base_units = LOCAL_BASE_UNITS.union(base_units_from_file)
        st.write(f"DEBUG: Using {len(combined_base_units)} base units for fixed pipeline.")

        try:
             xls = pd.ExcelFile(io.BytesIO(file_bytes))
             all_sheets = xls.sheet_names
             print(f"Found sheets: {all_sheets}")
             st.write(f"DEBUG: Input sheets: {all_sheets}")
        except Exception as e:
             st.error(f"Error reading input Excel file: {e}. Is the file format correct?")
             return None

        total_rows_processed = 0
        chunk_size = 500 # Process in chunks for memory efficiency

        for sheet_name in all_sheets:
            print(f"Processing sheet: '{sheet_name}'")
            st.write(f"DEBUG: Processing sheet: '{sheet_name}'")
            row_index = -1 # Initialize row index for error reporting
            try:
                sheet_df = pd.read_excel(xls, sheet_name=sheet_name)

                value_col_name = 'Value'
                if value_col_name not in sheet_df.columns:
                    st.warning(f"Sheet '{sheet_name}' skipped: Missing required column '{value_col_name}'.")
                    continue

                # *** ADDED .fillna('') before astype ***
                sheet_df[value_col_name] = sheet_df[value_col_name].fillna('').astype(str)

                # Keep track of original columns present in the sheet
                original_cols_in_sheet = sheet_df.columns.tolist()

                sheet_rows_processed = 0
                for i in range(0, len(sheet_df), chunk_size):
                    chunk_df = sheet_df.iloc[i:i + chunk_size]

                    for row_index, row_series in chunk_df.iterrows():
                        main_key = row_series.get(value_col_name, '').strip()

                        if not main_key:
                            continue

                        # Use the locally defined process_single_key
                        result_rows_for_key = process_single_key(main_key, combined_base_units, multipliers_dict)

                        # Get original data for this row
                        original_row_data = row_series.to_dict()
                        # Add processed results to the original data for each output row
                        for r_dict in result_rows_for_key:
                            final_row = original_row_data.copy() # Start with original data
                            final_row.update(r_dict) # Add/overwrite with processed info ('Main Key', 'Category', 'Code', parsed 'Value', etc.)
                            final_row["Sheet"] = sheet_name # Add sheet name identifier
                            all_processed_rows.append(final_row)

                        sheet_rows_processed += 1

                    # Optional garbage collection
                    gc.collect()

                total_rows_processed += sheet_rows_processed
                st.write(f"DEBUG: Processed {sheet_rows_processed} non-empty rows from sheet '{sheet_name}'.")

            except Exception as e:
                st.error(f"Error processing sheet '{sheet_name}' at row index approx {row_index if row_index != -1 else 'N/A'}: {e}")
                st.error(traceback.format_exc())
                # Continue processing other sheets if possible

        st.write(f"DEBUG: Fixed Processing Pipeline finished. Total input rows processed: {total_rows_processed}. Output rows generated: {len(all_processed_rows)}")

        if not all_processed_rows:
             st.warning("Fixed processing generated no output rows. Check input data and mapping.")
             return pd.DataFrame() # Return an empty DataFrame

        # Create DataFrame from the collected rows
        processed_df = pd.DataFrame(all_processed_rows)
        return processed_df

    # Catch specific errors during setup
    except FileNotFoundError as e:
         st.error(f"Pipeline Error: Mapping file not found at '{mapping_file_path}'. {e}. Cannot start fixed processing.")
         return None
    except ValueError as e: # Catch errors from read_mapping_file
         st.error(f"Pipeline Error reading mapping file: {e}. Cannot start fixed processing.")
         return None
    # Catch any other unexpected errors
    except Exception as e:
        st.error(f"An unexpected error occurred in the fixed pipeline: {e}")
        st.error(traceback.format_exc())
        return None

# --- END OF FILE fixed_pipeline.py ---
