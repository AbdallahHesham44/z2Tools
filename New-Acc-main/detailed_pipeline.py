#**`detailed_pipeline.py`**

#```python
#############################################
# MODULE: DETAILED ANALYSIS PIPELINE
# Purpose: Implements the second pipeline logic:
#          - Performs detailed classification and analysis.
#          - Extracts units, numeric values, normalized values.
#          - Generates summary columns (consistency, min/max, etc.).
#############################################

import pandas as pd
import re
import traceback
import streamlit as st # For error/warning/debug messages
from normalization_utils import SLASH_RANGE_UNITS
# Import necessary functions and constants from other modules
from mapping_utils import read_mapping_file
from analysis_helpers import (
    classify_value_type_detailed,
    transform_multiple_label,
    replace_numbers_keep_sign_all,          # keep if something else still uses it
    replace_numbers_keep_sign_outside_parens,   # ← ADD THIS
    resolve_compound_unit,
    analyze_value_units,
    extract_numeric_info_for_value,
    safe_str,
    process_unit_token
)
from normalization_utils import SLASH_RANGE_UNITS
from analysis_helpers import process_unit_token
from prefix_utils import strip_prefix_once   # add at top of file if not there
# ────────────────────────────────────────────────────────────────
#  replace_numbers_keep_sign_except_unit_digits  (token-level swap)
# ────────────────────────────────────────────────────────────────
def replace_numbers_keep_sign_except_unit_digits(
    s: str,
    digit_units: list[str],
) -> str:
    """
    Convert every *numeric token* outside parentheses to '$', UNLESS that
    digit sequence belongs to one of the supplied digit-bearing units.

    A numeric token is:
        [optional sign]  digits [ '.' digits ]   (no exponent part)

    Parameters
    ----------
    s : str
        Original text (e.g. '1.122V, 3.086V, Adj')
    digit_units : list[str]
        Every unit string from mapping.xlsx that contains at least one
        decimal digit (e.g. ['K 7-Step MacAdam Ellipse'])

    Returns
    -------
    str
        Same string with the allowed numbers mapped to '$'.
    """
    txt = str(s)
    if not txt:
        return txt

    # ---- 1) mark indices of digits that belong to a unit ------------
    keep_digit_idx: set[int] = set()
    for u in digit_units:                            # plain substring search
        pos = txt.find(u)
        while pos != -1:
            for offs, ch in enumerate(u):
                if ch.isdigit():
                    keep_digit_idx.add(pos + offs)
            pos = txt.find(u, pos + 1)

    # ---- 2) scan once and build the output --------------------------
    out = []
    i = 0
    paren_depth = 0
    n = len(txt)

    while i < n:
        ch = txt[i]

        # track parentheses nesting
        if ch == "(":
            paren_depth += 1
            out.append(ch)
            i += 1
            continue
        if ch == ")":
            paren_depth = max(paren_depth - 1, 0)
            out.append(ch)
            i += 1
            continue

        # --- candidate start of a numeric token ---------------------
        if paren_depth == 0:
            # signed number?
            if ch in "+-" and i + 1 < n and txt[i + 1].isdigit():
                sign = ch
                j = i + 1
            # unsigned number?
            elif ch.isdigit():
                sign = ""
                j = i
            else:
                out.append(ch)
                i += 1
                continue

            # j now at first digit
            if j in keep_digit_idx:
                # this digit belongs to a protected unit → copy verbatim
                out.append(ch)
                i += 1
                continue

            # walk through digits (and at most one dot + following digits)
            seen_dot = False
            while j < n and (txt[j].isdigit() or (txt[j] == "." and not seen_dot)):
                seen_dot |= txt[j] == "."
                j += 1

            # replace whole token with '$', keeping the sign if any
            out.append(f"{sign}$")
            i = j
            continue

        # ---- inside parentheses OR not a number start ---------------
        out.append(ch)
        i += 1

    return "".join(out)

# … existing code that builds df …
_NONCRITICAL_EXC = {"ThousandsSeparator"}
def _get_source_val(row):
    """
    Decide whether to analyse the cleaned Value or fall back to Value_E.
    """
    if str(row.get("ExceptionFlag")).upper() != "YES":
        return row["Value"]

    exc_types = {
        e.strip() for e in str(row.get("ExceptionTypes", "")).split(",") if e.strip()
    }
    return row["Value"] if exc_types.issubset(_NONCRITICAL_EXC) else row["Value_E"]
def _extract_special_unit(raw: str,
                          base_units: set[str],
                          multipliers: dict[str,float]
                         ) -> str | None:
    """
    If resolve_compound_unit failed (returned an Error),
    try to pull out one of your slash-range or dash-range units
    from the raw Value_E string and re-run it through your normal token logic.
    """
    # build prefix & unit alternations
    prefixes    = sorted(multipliers.keys(), key=len, reverse=True)
    prefixes_pat= "|".join(map(re.escape, prefixes))

    slash_pat   = "|".join(map(re.escape, SLASH_RANGE_UNITS))
    unit_pat    = "|".join(map(re.escape,
                         sorted(base_units, key=len, reverse=True)))

    # 1) slash-range (e.g. “-1300/ +300ppm/°C”)
    slash_re = re.compile(
        rf"/\s*[+\-]?\d+\s*"
        rf"(?P<mult>{prefixes_pat})?"
        rf"(?P<unit>{slash_pat})\b"
    )
    m = slash_re.search(raw)
    if m:
        tok = f"{m.group('mult') or ''}{m.group('unit')}"
        return process_unit_token(tok, base_units, multipliers)

    # 2) dash-range (e.g. “10-11 AWG”)
    dash_re = re.compile(
        rf"\b\d+\s*-\s*\d+\s*"
        rf"(?P<mult>{prefixes_pat})?"
        rf"(?P<unit>{unit_pat})\b"
    )
    m = dash_re.search(raw)
    if m:
        tok = f"{m.group('mult') or ''}{m.group('unit')}"
        return process_unit_token(tok, base_units, multipliers)

    return Non
def _extract_special_unit(raw: str,
                          base_units: set[str],
                          multipliers: dict[str,float]
                         ) -> str | None:
    """
    If an 'Absolute Unit' came back as an Error, look for either:
      - slash-range units in SLASH_RANGE_UNITS (e.g. 'ppm/°C','ppm/K',…)
      - dash-range units among your base_units (e.g. 'AWG','Ohm','V',…)
    and return the matched unit (with any multiplier prefix) or None.
    """
    # build an alternation of your multiplier symbols, longest first
    prefixes = sorted(multipliers.keys(), key=len, reverse=True)
    prefixes_pat = "|".join(map(re.escape, prefixes))

    # build alternations for slash-units and for all base_units
    slash_pat = "|".join(map(re.escape, SLASH_RANGE_UNITS))
    unit_pat  = "|".join(map(re.escape,
                         sorted(base_units, key=len, reverse=True)))

    # 1) slash-range, e.g.  -1300/ +300ppm/°C  or  -1300/ +300kppm/°C
    slash_re = re.compile(
        r"/\s*[+\-]?\d+\s*"
        rf"(?P<mult>(?:{prefixes_pat}))?"
        rf"(?P<unit>(?:{slash_pat}))\b"
    )
    m = slash_re.search(raw)
    if m:
        return f"{m.group('mult') or ''}{m.group('unit')}"

    # 2) dash-range, e.g.  10-11 AWG  or  10-11kAWG
    dash_re = re.compile(
        r"\b\d+\s*-\s*\d+\s*"
        rf"(?P<mult>(?:{prefixes_pat}))?"
        rf"(?P<unit>(?:{unit_pat}))\b"
    )
    m = dash_re.search(raw)
    if m:
        return f"{m.group('mult') or ''}{m.group('unit')}"

    return None
# --- detailed_analysis_pipeline function (core logic) ---
def detailed_analysis_pipeline(df, base_units, multipliers_dict):
    # ----- Add Normalized Unit column -----
    digit_units = [u for u in base_units if any(ch.isdigit() for ch in u)]

    normalized_units = []
    for _, row in df.iterrows():
        # df['Value'] is already your normalized string;
        # if the preprocessor flagged an exception, fall back to the raw input in Value_E
        source_val = _get_source_val(row)
        normalized_units.append(
            replace_numbers_keep_sign_except_unit_digits(str(source_val), digit_units)
        )
    df['Normalized Unit'] = normalized_units
    
    # ----- Compute Absolute Unit based on Normalized Unit -----
    absolute_units = []
    
    PH_RE = re.compile(r'^(\$[+\-]\$)(\s*)(.+)$')   # "$-$"  or  "$+$"
    
    for nu in df['Normalized Unit']:
    
        # 1) slash-range tokens stay untouched
        if any(slash in nu for slash in SLASH_RANGE_UNITS):
            absolute_units.append(nu)
            continue
    
        # 2) placeholder case  -------------------------------------------------
        m = PH_RE.match(nu)
        if m:
            placeholder, space, rest = m.groups()   # e.g. "$-$", " ", "mAWG"
            # strip exactly one valid multiplier prefix
            for pfx in sorted(multipliers_dict.keys(), key=len, reverse=True):
                if rest.startswith(pfx):
                    rest = rest[len(pfx):]
                    break
            absolute_units.append(f"{placeholder}{space}{rest}")
            continue
    
        # 3) everything else → normal resolver
        absolute_units.append(
            resolve_compound_unit(nu, base_units, multipliers_dict)
        )
    
    df['Absolute Unit'] = absolute_units
    classifications   = []
    identifiers_list  = []
    sub_val_counts    = []
    condition_counts  = []
    any_range_main    = []
    any_multi_main    = []
    any_range_cond    = []
    any_multi_cond    = []
    detailed_value_types = []

    for val in df['Value']:
        val_str = str(val)
        (cls, ids, sv_count, final_cond_item_count,
         rng_main, multi_main, rng_cond, multi_cond,
         final_main_item_count) = classify_value_type_detailed(val_str)
        cls = transform_multiple_label(cls)
        classifications.append(cls)
        identifiers_list.append(ids)
        sub_val_counts.append(sv_count)
        condition_counts.append(final_cond_item_count)
        any_range_main.append(rng_main)
        any_multi_main.append(multi_main)
        any_range_cond.append(rng_cond)
        any_multi_cond.append(multi_cond)
        if cls:
            dvt = f"{cls} [{final_main_item_count}][{final_cond_item_count}] x{sv_count}"
        else:
            dvt = ""
        detailed_value_types.append(dvt)

    df['Classification'] = classifications
    df['Identifiers'] = identifiers_list
    df['SubValueCount'] = sub_val_counts
    df['ConditionCount'] = condition_counts
    df['HasRangeInMain'] = any_range_main
    df['HasMultiValueInMain'] = any_multi_main
    df['HasRangeInCondition'] = any_range_cond
    df['HasMultipleConditions'] = any_multi_cond
    df['DetailedValueType'] = detailed_value_types

    # Analyze units in main vs condition
    main_units_list = []
    main_distinct_count_list = []
    main_consistent_list = []
    condition_units_list = []
    condition_distinct_count_list = []
    condition_consistent_list = []
    main_sub_analysis_list = []
    condition_sub_analysis_list = []

    for val in df['Value']:
        val_str = str(val)
        ua = analyze_value_units(val_str, base_units, multipliers_dict)
        main_units_list.append(", ".join(safe_str(x) for x in ua["main_units"]))
        main_distinct_count_list.append(len(ua["main_distinct_units"]))
        main_consistent_list.append(ua["main_units_consistent"])
        condition_units_list.append(", ".join(safe_str(x) for x in ua["condition_units"]))
        condition_distinct_count_list.append(len(ua["condition_distinct_units"]))
        condition_consistent_list.append(ua["condition_units_consistent"])
        main_sub_analysis_list.append(str(ua["main_sub_analysis"]))
        condition_sub_analysis_list.append(str(ua["condition_sub_analysis"]))

    df["MainUnits"] = main_units_list
    df["MainDistinctUnitCount"] = main_distinct_count_list
    df["MainUnitsConsistent"] = main_consistent_list
    df["ConditionUnits"] = condition_units_list
    df["ConditionDistinctUnitCount"] = condition_distinct_count_list
    df["ConditionUnitsConsistent"] = condition_consistent_list
    df["MainSubAnalysis"] = main_sub_analysis_list
    df["ConditionSubAnalysis"] = condition_sub_analysis_list

    # Numeric values, multipliers, base units, etc.
    main_numeric_values_list = []
    condition_numeric_values_list = []
    main_multiplier_list = []
    condition_multiplier_list = []
    main_base_units_list = []
    condition_base_units_list = []
    normalized_main_values_list = []
    normalized_condition_values_list = []
    overall_unit_consistency_list = []
    parsing_error_flag_list = []
    sub_value_variation_summary_list = []
    min_value_list = []
    max_value_list = []
    single_unit_list = []
    distinct_units_all_list = []

    for val in df['Value']:
        val_str = str(val)
        num_info = extract_numeric_info_for_value(val_str, base_units, multipliers_dict)
        main_numeric_values_list.append(", ".join([f"{x:.12g}" if x is not None else "" for x in num_info["main_numeric"]]))
        condition_numeric_values_list.append(", ".join([f"{x:.12g}" if x is not None else "" for x in num_info["condition_numeric"]]))
        main_multiplier_list.append(", ".join(safe_str(x) for x in num_info["main_multipliers"]))
        condition_multiplier_list.append(", ".join(safe_str(x) for x in num_info["condition_multipliers"]))
        main_base_units_list.append(", ".join(safe_str(x) for x in num_info["main_base_units"]))
        condition_base_units_list.append(", ".join(safe_str(x) for x in num_info["condition_base_units"]))
        normalized_main_values_list.append(", ".join([f"{x:.12g}" if x is not None else "" for x in num_info["normalized_main"]]))
        normalized_condition_values_list.append(", ".join([f"{x:.12g}" if x is not None else "" for x in num_info["normalized_condition"]]))

        ua = analyze_value_units(val_str, base_units, multipliers_dict)
        overall_consistency = ua["main_units_consistent"] and ua["condition_units_consistent"]
        overall_unit_consistency_list.append(overall_consistency)

        parsing_error = any(num_info["main_errors"]) or any(num_info["condition_errors"])
        parsing_error_flag_list.append(parsing_error)

        # Summarize unit variation
        main_units = sorted(list(ua["main_distinct_units"]))
        if main_units:
            if len(main_units) == 1:
                main_variation = "Uniform: " + safe_str(main_units[0])
            else:
                main_variation = "Mixed: " + ", ".join(main_units)
        else:
            main_variation = "None"

        condition_units = sorted(list(ua["condition_distinct_units"]))
        if condition_units:
            if len(condition_units) == 1:
                condition_variation = "Uniform: " + safe_str(condition_units[0])
            else:
                condition_variation = "Mixed: " + ", ".join(condition_units)
        else:
            condition_variation = "None"

        sub_value_variation_summary_list.append(f"Main: {main_variation}; Condition: {condition_variation}")

        # Min/Max normalized
        all_normalized_values = []
        all_units_used = []
        for i, val_num in enumerate(num_info["normalized_main"]):
            if not num_info["main_errors"][i] and (val_num is not None):
                all_normalized_values.append(val_num)
                all_units_used.append(num_info["main_base_units"][i])
        for i, val_num in enumerate(num_info["normalized_condition"]):
            if not num_info["condition_errors"][i] and (val_num is not None):
                all_normalized_values.append(val_num)
                all_units_used.append(num_info["condition_base_units"][i])
        if all_normalized_values:
            min_val = min(all_normalized_values)
            max_val = max(all_normalized_values)
        else:
            min_val = None
            max_val = None
        distinct_units_all = set(u for u in all_units_used if u and u.lower() != "none")
        is_single_unit = (len(distinct_units_all) <= 1)
        min_value_list.append(f"{min_val:.12g}" if min_val is not None else "")
        max_value_list.append(f"{max_val:.12g}" if max_val is not None else "")
        single_unit_list.append(is_single_unit)
        distinct_units_all_list.append(", ".join(distinct_units_all) if distinct_units_all else "")

    df["MainNumericValues"] = main_numeric_values_list
    df["ConditionNumericValues"] = condition_numeric_values_list
    df["MainMultipliers"] = main_multiplier_list
    df["ConditionMultipliers"] = condition_multiplier_list
    df["MainBaseUnits"] = main_base_units_list
    df["ConditionBaseUnits"] = condition_base_units_list
    df["NormalizedMainValues"] = normalized_main_values_list
    df["NormalizedConditionValues"] = normalized_condition_values_list
    df["OverallUnitConsistency"] = overall_unit_consistency_list
    df["ParsingErrorFlag"] = parsing_error_flag_list
    df["SubValueUnitVariationSummary"] = sub_value_variation_summary_list
    df["MinNormalizedValue"] = min_value_list
    df["MaxNormalizedValue"] = max_value_list
    df["SingleUnitForAllSubs"] = single_unit_list
    df["AllDistinctUnitsUsed"] = distinct_units_all_list

    return df



# --- Wrapper function (Entry point) ---
# MODIFIED detailed_analysis function
def detailed_analysis(input_df: pd.DataFrame, mapping_file: str, output_file: str):
    """
    Performs detailed analysis on the input DataFrame using the mapping file.
    Saves the results to the specified output Excel file.

    Args:
        input_df (pd.DataFrame): DataFrame containing the data (must have 'Value' column).
        mapping_file (str): Path to the local 'mapping.xlsx' file.
        output_file (str): Path to save the resulting Excel file.

    Returns:
        str or None: The path to the output file on success, None on failure.
    """
    st.write("DEBUG: Starting detailed_analysis function...")
    try:
        # Read mapping file - raises FileNotFoundError or ValueError on issues
        base_units, multipliers_dict = read_mapping_file(mapping_file)
        st.write(f"DEBUG: Using {len(base_units)} base units and {len(multipliers_dict)} multipliers for detailed analysis.")
    except (FileNotFoundError, ValueError) as e:
        st.error(f"Error reading mapping file '{mapping_file}': {e}")
        return None # Indicate failure: Cannot proceed without mapping
    except Exception as e:
        st.error(f"Unexpected error reading mapping file '{mapping_file}': {e}")
        return None

    # --- Input Validation ---
    if not isinstance(input_df, pd.DataFrame):
         st.error("Detailed Analysis Error: Input must be a pandas DataFrame.")
         return None
    if 'Value' not in input_df.columns:
        st.error("Detailed Analysis Error: Input DataFrame must contain a column named 'Value'.")
        return None

    # --- Run the main analysis pipeline ---
    try:
        # Pass a copy of the input df to avoid modification side effects
        analysis_df = detailed_analysis_pipeline(input_df.copy(), base_units, multipliers_dict)
    except Exception as e:
         st.error(f"Error during detailed analysis pipeline execution: {e}")
         st.error(traceback.format_exc())
         return None # Indicate failure

    # Check if the analysis produced a result
    if analysis_df is None: # Check if pipeline function itself failed critically
         st.error("Detailed analysis pipeline returned None. Aborting save.")
         return None
    if analysis_df.empty:
         # Decide if saving an empty file is desired or if it indicates an issue
         st.warning("Detailed analysis resulted in an empty DataFrame. Saving empty file.")
         # Proceed to save the empty DataFrame.

    # --- Save the results ---
    try:
        analysis_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"[✓] Detailed analysis saved to '{output_file}'.")
        st.write(f"DEBUG: Detailed analysis saved to '{output_file}'. Shape: {analysis_df.shape}")
        return output_file # Return path on success
    except Exception as e:
        st.error(f"Error writing detailed analysis to '{output_file}': {e}")
        return None # Indicate failure
