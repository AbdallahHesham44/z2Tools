#############################################
# MODULE: ANALYSIS HELPERS
# Purpose: Contains core, reusable functions for parsing,
#          classification, unit/numeric extraction.
#          Used by both fixed and detailed pipelines.
#############################################

import re
import pandas as pd
import streamlit as st # For warnings/debug, consider replacing with logging
from prefix_utils import STRING_KEYWORDS
# Import constants from mapping_utils
# Assumes mapping_utils.py is in the same directory or accessible via Python path
try:
    from mapping_utils import MULTIPLIER_MAPPING
except ImportError:
    # Fallback if run in a context where mapping_utils isn't directly importable
    # This is less ideal but provides a default.
    st.error("Could not import MULTIPLIER_MAPPING from mapping_utils. Using default empty map.")
    MULTIPLIER_MAPPING = {}
# In analysis_helpers.py
def transform_multiple_label(label: str) -> str:
    """
    If the label starts with "Multiple", wraps the descriptive text after "Multiple"
    in parentheses while keeping any trailing numeric annotations outside.
    
    For example:
       "Multiple Single Value [1][0] x5" => "Multiple (Single Value) [1][0] x5"
       "Multiple Single Value" => "Multiple (Single Value)"
    
    Args:
        label (str): The original classification string or DetailedValueType string.
    
    Returns:
        str: The transformed string if applicable, otherwise the original.
    """
    # Check for None or empty string first for robustness
    if not isinstance(label, str) or not label:
        return label

    multiple_prefix = "Multiple "
    if label.startswith(multiple_prefix):
        # Extract the part AFTER "Multiple "
        remainder = label[len(multiple_prefix):]

        # Separate descriptive part from annotations (assuming a space before '[')
        descriptive = remainder
        annotations = ""
        split_index = remainder.find(" [")  # Look for the pattern " ["
        if split_index != -1:
            descriptive = remainder[:split_index].strip()  # Text before " ["
            annotations = remainder[split_index:].strip()   # Everything from " [" onward
        else:
            descriptive = remainder.strip()

        # If the descriptive text exists, wrap it in parentheses
        if descriptive:
            new_label = f"{multiple_prefix}({descriptive})"
            if annotations:
                new_label += " " + annotations
            return new_label
        else:
            return label
    return label

def split_outside_parens_preserve(text, delimiters):
    """
    Splits text by the given delimiters, ignoring delimiters inside parentheses,
    and always preserving the delimiter tokens AND adjacent spacing.

    Args:
        text (str): Text to split.
        delimiters (list[str]): List of delimiter strings.

    Returns:
        list[str]: List of tokens including delimiters and original spacing.
    """
    text = str(text)
    tokens = []
    current = ""
    i = 0
    depth = 0
    # Sort delimiters by length descending to match longest first (e.g., ' to ' before 'to')
    # No, the current logic checks exact matches, so sorting standard delimiters is fine.
    # Let's keep the original sorting logic for now, it prioritizes longer delimiter matches if overlapping.
    sorted_delims = sorted(delimiters, key=len, reverse=True)

    while i < len(text):
        char = text[i]
        if char == '(':
            depth += 1
            current += char
            i += 1
        elif char == ')':
            depth = max(0, depth - 1)
            current += char
            i += 1
        elif depth == 0:
            matched_delim = None
            for delim in sorted_delims:
                 # Check if the delimiter exists at the current position
                if i + len(delim) <= len(text) and text[i:i+len(delim)] == delim:
                    matched_delim = delim
                    break

            if matched_delim:
                # Append the segment BEFORE the delimiter (without stripping)
                if current: # Append only if there's something
                    tokens.append(current)
                # Append the delimiter itself
                tokens.append(matched_delim)
                current = "" # Reset current segment
                i += len(matched_delim) # Move index past the delimiter
            else:
                # Not a delimiter, add char to current segment
                current += char
                i += 1
        else: # Inside parentheses
            current += char
            i += 1

    # Append any remaining part after the last delimiter (without stripping)
    if current:
        tokens.append(current)

    # Filter out potential empty strings resulting from adjacent delimiters if necessary
    # Example: "A,,B" -> ["A", ",", "", ",", "B"]. Usually desirable to keep empty.
    # Let's return all parts for now.
    return tokens

# --- Utility functions (Mostly from original Section 3 & 4 helpers) ---

# In analysis_helpers.py

# ... (imports and other functions) ...

def extract_numeric_and_unit_analysis(token, base_units, multipliers_dict):
    """
    Analyzes a token to extract numeric value, multiplier symbol, base unit,
    normalized value, error flag, and the original rest string.
    Handles numbers, units, and combinations.

    Returns:
        tuple: (numeric_val, multiplier_symbol, base_unit, normalized_value, error_flag, original_rest_string)
    """
    token = str(token).strip()  # Ensure string and strip whitespace
    if not token:
        # Return None for the new 'original_rest' field as well
        return None, None, None, None, False, None

    # Regex to capture optional sign, numeric part, and the rest (unit, etc.)
    pattern = re.compile(r'^(?P<numeric>[+\-±]?\d*(?:\.\d+)?(?:[eE][+\-]?\d+)?)(?P<rest>.*)$')
    m = pattern.match(token)

    # Handle case where token might be only a unit
    if not m or not m.group("numeric"):
        original_rest = token # The whole token is the rest
        if token in base_units:
            return None, "1", token, None, False, original_rest
        else:
            sorted_prefixes = sorted(multipliers_dict.keys(), key=len, reverse=True)
            for prefix in sorted_prefixes:
                if token.startswith(prefix):
                    possible_base = token[len(prefix):].strip()
                    if possible_base in base_units:
                        return None, prefix, possible_base, None, False, original_rest
        # Return None for the new 'original_rest' field on error
        return None, None, None, None, True, None

    numeric_str_raw = m.group("numeric")
    original_rest = m.group("rest") # <-- STORE THE ORIGINAL REST STRING
    rest = original_rest.strip() # Use stripped version for unit logic

    # NEW: Instead of immediately removing parentheses, first check if the rest starts
    # with a known multiplier prefix and that the leftover is exactly a recognized unit.
    prefix_found = False
    cleaned_rest = rest # Default to stripped rest
    for prefix in sorted(multipliers_dict.keys(), key=len, reverse=True):
        if rest.startswith(prefix):
            possible_base = rest[len(prefix):].strip()
            if possible_base in base_units:
                # Keep the original rest (with the prefix and parentheses intact)
                # No, keep using the stripped rest for logic below
                cleaned_rest = rest # Use the stripped version found
                prefix_found = True
                break
    if not prefix_found:
        if rest in base_units:
            cleaned_rest = rest
        else:
            # We still remove parentheses for the unit *identification* part if it wasn't a clear prefix+unit match initially
            cleaned_rest = remove_parentheses_detailed(rest)


    try:
        numeric_val = float(numeric_str_raw.replace('±',''))
    except ValueError:
        # Return None for the new 'original_rest' field on error
        return None, None, None, None, True, None

    # Case: Only a number was found.
    if not cleaned_rest:
        # Return the original_rest (which might be spaces)
        return numeric_val, "1", None, numeric_val, False, original_rest

    # Try to match a known prefix + base unit
    multiplier_symbol = None
    base_unit = None
    multiplier_factor = 1.0
    found_unit_structure = False

    sorted_prefixes = sorted(multipliers_dict.keys(), key=len, reverse=True)
    for prefix in sorted_prefixes:
        # Use cleaned_rest (parentheses removed if necessary) for matching
        if cleaned_rest.startswith(prefix):
            possible_base = cleaned_rest[len(prefix):].strip()
            if possible_base in base_units:
                multiplier_symbol = prefix
                base_unit = possible_base
                multiplier_factor = multipliers_dict[prefix]
                found_unit_structure = True
                break

    if not found_unit_structure:
        # If no prefix was detected, check if cleaned_rest itself is a base unit.
        if cleaned_rest in base_units:
            multiplier_symbol = "1" # Explicitly set multiplier to 1
            base_unit = cleaned_rest
            found_unit_structure = True
        else:
            # Return None for the new 'original_rest' field on error
            return numeric_val, None, None, None, True, None

    normalized_value = numeric_val * multiplier_factor
    final_multiplier_symbol = multiplier_symbol if multiplier_symbol is not None else "1"

    # Return the extracted parts AND the original_rest string
    return numeric_val, final_multiplier_symbol, base_unit, normalized_value, False, original_rest

# ... (rest of the file)


def remove_parentheses_detailed(text: str) -> str:
    """Removes content within the outermost parentheses."""
    return re.sub(r'\([^()]*\)', '', str(text)) # Ensure input is string

def extract_identifiers_detailed(text: str):
    """Finds all content within non-nested parentheses."""
    return re.findall(r'\(([^()]*)\)', str(text)) # Finds content inside (...)


def split_outside_parens(text, delimiters):
    text = str(text)  # Ensure string input
    tokens = []
    current = ""
    i = 0
    depth = 0
    # Sort delimiters by length descending to match longest first
    sorted_delims = sorted(delimiters, key=len, reverse=True)

    while i < len(text):
        char = text[i]
        if char == '(':
            depth += 1
            current += char
            i += 1
        elif char == ')':
            depth = max(0, depth - 1)
            current += char
            i += 1
        elif depth == 0:
            matched_delim = None
            for delim in sorted_delims:
                if i + len(delim) <= len(text) and text[i:i+len(delim)] == delim:
                    matched_delim = delim
                    break

            if matched_delim:
                if current.strip():
                    tokens.append(current.strip())
                # Do not append the delimiter token; use it only as a boundary.
                current = ""
                i += len(matched_delim)
            else:
                current += char
                i += 1
        else:
            current += char
            i += 1

    if current.strip():
        tokens.append(current.strip())

    return [token for token in tokens if token]



def extract_numeric_info(part_text, base_units, multipliers_dict):
    """
    Extracts numeric information (values, multipliers, units, normalized)
    from a string potentially containing single, range, or multiple values.

    Args:
        part_text (str): The text segment to analyze (e.g., "10k to 20k Ohm", "5A", "1, 2, 3").
        base_units (set): Set of known base units.
        multipliers_dict (dict): Dictionary of multipliers.

    Returns:
        dict: Dictionary containing lists of extracted info and the detected type ('single', 'range', 'multiple', 'none').
              Keys: "numeric_values", "multipliers", "base_units", "normalized_values", "error_flags", "type".
    """
    # First remove content within parentheses for analysis
    text = remove_parentheses_detailed(part_text).strip()

    if not text:
        return {
            "numeric_values": [], "multipliers": [],
            "base_units": [], "normalized_values": [],
            "error_flags": [], "type": "none"
        }

    # Determine structure: range, multiple values, or single
    # Use robust splitting for ' to ' to avoid issues with substrings
    # Regex checks for ' to ' surrounded by whitespace or start/end of string parts
    is_range = bool(re.search(r'(?:^|\s)to(?:\s|$)', text, re.IGNORECASE)) and len(split_outside_parens(text, [','])) == 1 # Avoid "A, B to C" being range

    # Split by comma first if not a simple range
    tokens = split_outside_parens(text, [','])

    if len(tokens) > 1:
        info_type = "multiple"
        # Further split parts containing ' to ' if necessary? No, treat "A, B to C" as multiple.
    elif len(tokens) == 1:
         # Now check the single token for ' to ' range
         range_parts = split_outside_parens(tokens[0], [' to ']) # Crude split for ' to '
         # Refine range check: needs values on both sides?
         if len(range_parts) > 1 and re.search(r'\s+to\s+', tokens[0], re.IGNORECASE):
              info_type = "range"
              tokens = range_parts # Use the parts split by ' to '
         else:
              info_type = "single"
              tokens = [tokens[0]] # Keep the single token
    else: # No tokens found after splitting (e.g., empty string after paren removal)
        info_type = "none"
        tokens = []


    # Initialize lists to store results for each token
    numeric_values = []
    multipliers = []
    base_units_list = []
    normalized_values = []
    error_flags = []

    # Process each token found
    for token in tokens:
        token_strip = token.strip()
        if not token_strip: continue # Skip empty tokens resulting from splits

        num_val, multiplier_symbol, base_unit, norm_val, err_flag, _ = extract_numeric_and_unit_analysis(
            token_strip, base_units, multipliers_dict
        )

        numeric_values.append(num_val)
        # Use "1" if no multiplier symbol was identified but parsing was ok
        multipliers.append(multiplier_symbol if multiplier_symbol else ("1" if not err_flag and num_val is not None else None))
         # Base unit might be None if only number, or if error
        base_units_list.append(base_unit if base_unit else None)
        normalized_values.append(norm_val)
        error_flags.append(err_flag)

    # Return dictionary of results
    return {
        "numeric_values": numeric_values,
        "multipliers": multipliers,
        "base_units": base_units_list,
        "normalized_values": normalized_values,
        "error_flags": error_flags,
        "type": info_type # Indicates if it was parsed as 'single', 'range', or 'multiple' parts
    }

def safe_str(item, placeholder="None"):
    """Converts item to string, using placeholder if None."""
    return str(item) if item is not None else placeholder

def extract_numeric_info_for_value(raw_value, base_units, multipliers_dict):
    """
    Extracts numeric information for a potentially complex value string,
    splitting it into main and condition parts based on '@'.

    Args:
        raw_value (str): The input value string (e.g., "10A @ 5V", "50 Ohm").
        base_units (set): Set of known base units.
        multipliers_dict (dict): Dictionary of multipliers.

    Returns:
        dict: Dictionary containing numeric info for main and condition parts.
              Keys like "main_numeric", "condition_base_units", "normalized_main", etc.
    """
    # This function processes ONE logical value string which might contain '@' internally.
    # It assumes the input `raw_value` represents one entry (potentially complex).
    # Splitting multiple entries (like "10A, 20A") should happen *before* calling this.

    raw_value = str(raw_value).strip() # Ensure string type
    main_part = raw_value
    cond_part = ""

    # Split only on the first '@' found outside parentheses
    at_split = split_outside_parens(raw_value, ['@'])
    if len(at_split) > 1:
        main_part = at_split[0].strip()
        cond_part = "@".join(at_split[1:]).strip() # Rejoin if multiple '@' were present (unlikely?)
    elif len(at_split) == 1:
         main_part = at_split[0].strip() # No '@' found outside parens
         # Check if '@' exists inside parentheses (should be ignored by split_outside_parens)


    # Process the main part and the condition part separately
    main_info = extract_numeric_info(main_part, base_units, multipliers_dict)
    # Process cond_part only if it's not empty
    cond_info = extract_numeric_info(cond_part, base_units, multipliers_dict) if cond_part else {
            "numeric_values": [], "multipliers": [],
            "base_units": [], "normalized_values": [],
            "error_flags": [], "type": "none"
        }

    # Combine results. Note these lists contain results for *all* tokens within main/condition parts
    # e.g., for "10 to 20A @ 5V", main_info will have two entries, cond_info one entry.
    return {
        "main_numeric": main_info["numeric_values"],
        "main_multipliers": main_info["multipliers"],
        "main_base_units": main_info["base_units"],
        "normalized_main": main_info["normalized_values"],
        "main_errors": main_info["error_flags"],
        "main_type": main_info["type"], # Add type ('single', 'range', 'multiple')

        "condition_numeric": cond_info["numeric_values"],
        "condition_multipliers": cond_info["multipliers"],
        "condition_base_units": cond_info["base_units"],
        "normalized_condition": cond_info["normalized_values"],
        "condition_errors": cond_info["error_flags"],
        "condition_type": cond_info["type"], # Add type
    }

# --- change1 seems to have been incorporated into process_unit_token_no_paren ---
# analysis_helpers.py
def process_unit_token_no_paren(core: str,
                                base_units: set[str],
                                multipliers_dict: dict[str, float]) -> str:
    """
    Fallback unit resolver used by the detailed-splitter path.
    Accepts:
      • exact match on full token (incl. (…) )
      • exact match after stripping parentheses
      • prefix + base-unit (with or without parentheses)
    Returns just the base unit (no “$ ” prefix).  On failure returns
    'Error: Undefined core unit …'.
    """
    import re

    full    = core.strip()                          # e.g. 'Nm (3.5-4.4 Lb-In)'
    noparen = re.sub(r"\([^)]*\)", "", full).strip()  # e.g. 'Nm'

    # 1. exact hit (incl. parentheses)
    if full and full in base_units:
        return full

    # 1b. exact hit after removing (…) text
    if noparen and noparen in base_units:
        return noparen

    # 2. prefix + unit  (first on full, then on noparen)
    for pfx in sorted(multipliers_dict.keys(), key=len, reverse=True):
        if full.startswith(pfx):
            tail = full[len(pfx):]
            if tail in base_units:
                return tail
            tail_np = re.sub(r"\([^)]*\)", "", tail).strip()
            if tail_np in base_units:
                return tail_np

    # 3. still nothing ⇒ error
    return f"Error: Undefined core unit '{full}'"




def analyze_unit_part(part_text, base_units, multipliers_dict):
    """
    Analyzes the unit(s) present in a text segment (which might be single, range, or multiple).
    Identifies distinct base units and checks for consistency.

    Args:
        part_text (str): The text segment (e.g., "10k to 20k Ohm", "5A", "1V, 2V").
        base_units (set): Set of known base units.
        multipliers_dict (dict): Dictionary of multipliers.

    Returns:
        dict: Contains lists/sets of units, consistency flag, count, and type.
              Keys: "units", "distinct_units", "is_consistent", "count", "type".
    """
    # Remove parentheses content first
    text = remove_parentheses_detailed(part_text).strip()
    if not text:
        return {
            "units": [], "distinct_units": set(),
            "is_consistent": True, "count": 0,
            "type": "none"
        }

    # Determine structure and split into tokens using the same logic as extract_numeric_info
    is_range = bool(re.search(r'(?:^|\s)to(?:\s|$)', text, re.IGNORECASE)) and len(split_outside_parens(text, [','])) == 1
    tokens = split_outside_parens(text, ['to',',', '@'])

    if len(tokens) > 1:
        part_type = "multiple"
    elif len(tokens) == 1:
         range_parts = split_outside_parens(tokens[0], [' to '])
         if len(range_parts) > 1 and re.search(r'\s+to\s+', tokens[0], re.IGNORECASE):
              part_type = "range"
              tokens = range_parts
         else:
              part_type = "single"
              tokens = [tokens[0]]
    else:
        part_type = "none"
        tokens = []

    units = []
    for token in tokens:
        token_strip = token.strip()
        if not token_strip: continue

        # Extract only the unit part from the token (e.g., "10kOhm" -> "Ohm")
        _, _, base_unit, _, err_flag, _ = extract_numeric_and_unit_analysis(
            token_strip, base_units, multipliers_dict
        )

        if not err_flag and base_unit:
            units.append(base_unit)
        elif token_strip in base_units: # Handle cases like just "V"
             units.append(token_strip)
        else:
             # If no unit was resolved by extraction, add None or placeholder
             units.append(None) # Represent absence of recognized unit

    # Filter out None before creating distinct set and checking consistency
    valid_units = [u for u in units if u is not None]
    distinct_units = set(valid_units)
    # Consistency means 0 or 1 distinct *valid* unit found
    is_consistent = (len(distinct_units) <= 1)
    count = len(tokens) # Count based on number of tokens processed

    return {
        "units": units, # List of resolved base units (or None) for each token
        "distinct_units": distinct_units, # Set of unique *valid* base units found
        "is_consistent": is_consistent, # True if zero or one distinct valid base unit
        "count": count, # Number of tokens processed
        "type": part_type # How the part was structured
    }


def analyze_value_units(raw_value, base_units, multipliers_dict):
    """
    Analyzes units in a potentially complex value string (main @ condition).

    Args:
        raw_value (str): Input value string.
        base_units (set): Known base units.
        multipliers_dict (dict): Multiplier map.

    Returns:
        dict: Aggregated unit analysis results for main and condition parts.
    """
    # Similar structure to extract_numeric_info_for_value
    raw_value = str(raw_value).strip()
    main_part = raw_value
    cond_part = ""

    # Split by '@' outside parentheses
    at_split = split_outside_parens(raw_value, ['@'])
    if len(at_split) > 1:
        main_part = at_split[0].strip()
        cond_part = "@".join(at_split[1:]).strip()
    elif len(at_split) == 1:
        main_part = at_split[0].strip()

    # Analyze units in main and condition parts
    main_analysis = analyze_unit_part(main_part, base_units, multipliers_dict)
    cond_analysis = analyze_unit_part(cond_part, base_units, multipliers_dict) if cond_part else {
            "units": [], "distinct_units": set(),
            "is_consistent": True, "count": 0,
            "type": "none"
        }


    # Combine unit information
    all_main_units = main_analysis["units"] # List including None
    all_condition_units = cond_analysis["units"] # List including None

    main_distinct = main_analysis["distinct_units"] # Set excluding None
    condition_distinct = cond_analysis["distinct_units"] # Set excluding None

    main_consistent = main_analysis["is_consistent"]
    condition_consistent = cond_analysis["is_consistent"]

    # Overall consistency considers *distinct valid* units across both parts
    all_distinct_units = main_distinct.union(condition_distinct)
    # Overall is consistent if 0 or 1 distinct valid unit is found across all parts
    overall_consistent = (len(all_distinct_units) <= 1)


    return {
        "main_units": all_main_units,
        "main_distinct_units": main_distinct,
        "main_units_consistent": main_consistent,
        "main_unit_count": main_analysis["count"], # Count of tokens in main part
        # Keep sub_analysis if needed for detailed debugging, maybe as string
        "main_sub_analysis": str(main_analysis),

        "condition_units": all_condition_units,
        "condition_distinct_units": condition_distinct,
        "condition_units_consistent": condition_consistent,
        "condition_unit_count": cond_analysis["count"], # Count of tokens in condition part
        "condition_sub_analysis": str(cond_analysis),

        "all_distinct_units": all_distinct_units, # Set of all unique valid units found
        "overall_consistent": overall_consistent # True if <= 1 unique valid unit across main and condition
    }


def process_unit_token(token: str,
                       base_units: set[str],
                       multipliers_dict: dict[str, float]) -> str:
    """
    Resolve one unit-token to its absolute form, preserving any leading sign.

    Works for: "$V", "-$V", "+$kΩ", "$mA", "±$GΩ", "kΩ", "MOhm (Typ)"
    """
    import re

    tok = str(token)

    # ------------------------------------------------------------------ #
    # 0)  Build a trimmed, case-insensitive lookup set *once per call*   #
    # ------------------------------------------------------------------ #
    _norm = lambda s: s.strip()              # keep only leading/trailing spaces out
    bu_trim = {_norm(u) for u in base_units} # exact spellings, case-sensitive
    
    def is_unit(u: str) -> bool:
        """
        Return True only when *u* matches a base-unit symbol **exactly**
        (same spelling, same capitalisation) after whitespace is stripped.
        """
        return _norm(u) in bu_trim
    # ------------------------------------------------------------------ #
    # 1)  Separate leading whitespace, sign and "$" prefix               #
    # ------------------------------------------------------------------ #
    lead_ws_match = re.match(r'^\s*', tok)
    lead_ws = lead_ws_match.group(0) if lead_ws_match else ""
    body = tok[len(lead_ws):]

    sign = ""
    if body and body[0] in "+-±":
        sign, body = body[0], body[1:]

    leading_dollar = ""
    if body.startswith("$"):
        leading_dollar, body = "$", body[1:]

    # ── NEW: blanks immediately after sign / "$" that must be kept ─────
    inner_ws_match = re.match(r'^\s*', body)
    inner_ws = inner_ws_match.group(0) if inner_ws_match else ""
    body = body[len(inner_ws):]

    # token was just whitespace / sign / '$'
    if not body.strip():
        return lead_ws + sign + leading_dollar

    unit_full_for_kw_check = body.strip()
    if unit_full_for_kw_check.lower() in [kw.lower() for kw in STRING_KEYWORDS]:
        # If it's a keyword, return it immediately, preserving all spacing and signs.
        # This prevents it from ever being treated as an "undefined unit".
        return lead_ws + sign + leading_dollar + inner_ws + unit_full_for_kw_check	


    # ------------------------------------------------------------------ #
    # 2)  Exact unit hit (with / without inner “()” text)                #
    # ------------------------------------------------------------------ #
    unit_full    = body.strip()
    unit_noparen = re.sub(r"\([^)]*\)", "", unit_full).strip()

    if is_unit(unit_full):
        return lead_ws + sign + leading_dollar + inner_ws + unit_full
    if is_unit(unit_noparen):
        return lead_ws + sign + leading_dollar + inner_ws + unit_noparen

    # ------------------------------------------------------------------ #
    # 3)  Prefix + unit (kΩ, mA, µF …)                                   #
    # ------------------------------------------------------------------ #
    for pfx in sorted(multipliers_dict.keys(), key=len, reverse=True):
        if unit_full.startswith(pfx):
            tail = unit_full[len(pfx):]
            if is_unit(tail):
                return lead_ws + sign + leading_dollar + inner_ws + tail
            tail_np = re.sub(r"\([^)]*\)", "", tail).strip()
            if is_unit(tail_np):
                return lead_ws + sign + leading_dollar + inner_ws + tail_np

    # ------------------------------------------------------------------ #
    # 4)  Still unknown  →  surface the error (sign + $ kept)            #
    # ------------------------------------------------------------------ #
    bad_token = f"{sign}{leading_dollar}{inner_ws}{unit_full}"
    return (lead_ws + sign +
            f"Error: Undefined unit '{bad_token}' (no recognised prefix or exact match)")


# In analysis_helpers.py

# analysis_helpers.py
def resolve_compound_unit(normalized_unit: str,
                           base_units: set[str],
                           multipliers_dict: dict[str, float]) -> str:
    """
    Convert a normalized-unit string (e.g. “-$dB @ $MHz to $GHz”)
    into its absolute-unit form **without changing any spacing**.

    • All spaces before and after tokens, delimiters (“to”, “,”, “@”) and
      parentheses are preserved 1-for-1.
    • Unresolved tokens still surface as  “Error: …”.
    """

    import re

    # Delimiters that are kept verbatim (with surrounding spaces)
    delimiters = ["to", ",", "@"]

    # Your existing helper; returns pieces incl. their original spacing
    tokens = split_outside_parens_preserve(normalized_unit, delimiters)

    resolved_parts: list[str] = []

    for part in tokens:
        # ───── 1. keep delimiter blocks 100 % untouched ─────
        if part.strip().lower() in delimiters:
            resolved_parts.append(part)
            continue

        # ───── 2. capture BOTH leading + trailing whitespace ─────
        m = re.match(r'^(\s*)(.*?)(\s*)$', part, re.DOTALL)
        if not m:                               # shouldn't happen
            resolved_parts.append(part)
            continue

        lead_ws, core, trail_ws = m.groups()

        # empty (pure whitespace) → keep as-is
        if not core:
            resolved_parts.append(part)
            continue

        # ───── 3. resolve the core token (no spaces) ─────
        token_res = process_unit_token(core, base_units, multipliers_dict)

        # keep error text *and* restored spacing exactly as found
        resolved_parts.append(lead_ws + token_res + trail_ws)

    return "".join(resolved_parts)





def count_main_items(main_str: str) -> int:
    main_str = main_str.strip()
    if not main_str:
        return 0
    if " to " in main_str:
        return 1
    if "," in main_str:
        return len([s for s in main_str.split(',') if s.strip()])
    return 1


def count_conditions(cond_str: str) -> int:
    cond_str = cond_str.strip()
    if not cond_str:
        return 0
    parts = [p.strip() for p in cond_str.split(',') if p.strip()]
    return len(parts)

    # Conditions are typically comma-separated clauses.
    # Each clause might be a single value or a range.
    comma_parts = split_outside_parens(cond_str, [','])
    count = 0
    for part in comma_parts:
        part_strip = part.strip()
        if not part_strip: continue # Skip empty parts
        count += 1

    # If no commas, check if the string is non-empty -> 1 condition clause
    if count == 0 and cond_str:
        return 1

    return count


def classify_condition(cond_str: str) -> str:
    cond_str = cond_str.strip()
    if not cond_str:
        return ""
    if "," in cond_str:
        return "Multiple Conditions"
    if " to " in cond_str:
        return "Range Condition"
    return "Single Condition"


def _is_number_with_optional_id(p_str: str) -> bool:
    """
    Checks if a string is a number optionally followed ONLY by a single 
    parenthetical block.
    e.g., "10", "10 (id)", "-5.5 (test)", "1e3 (±2g)" -> True
          "10A", "10 (id) unit", "10 unit (id)", "10A (id)" -> False
    """
    if not p_str:
        return False
    
    # Regex: 
    # 1. Number part: ^[+\-±]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[eE][+\-]?\d+)?
    # 2. Optional suffix part (space + parentheses): (?:\s*\([^)]*\))?
    # 3. End of string: $
    # This pattern ensures that if there's anything after the number, it MUST be
    # a single block of parentheses, optionally preceded by space.
    # It won't match if there are unit characters between the number and the parentheses,
    # or after the parentheses.
    pattern = r"^[+\-±]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[eE][+\-]?\d+)?(?:\s*\([^)]*\))?$"
    
    if re.fullmatch(pattern, p_str):
        # Additional check: if there is a parenthetical part, ensure no letters (common unit chars)
        # are *before* an opening parenthesis if the number itself is simple.
        # Example: "10A (test)" should be False. The regex above handles this well because "10A" isn't just a number.
        # Example: "4096 (±2g)". "4096" is the number. " (±2g)" is the optional ID. This is True.
        return True
    return False

def classify_main(main_str: str) -> str:
    main_str = main_str.strip()
    if not main_str:
        return ""

    # Pattern for a simple number only (no identifier, no units)
    # Used to distinguish from Single Value if no unit characters are present AND no identifier.
    strict_number_only_pattern = r"^[+\-±]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[eE][+\-]?\d+)?$"

    if " to " in main_str:
        parts = [p.strip() for p in main_str.split(" to ", 1)]
        if len(parts) == 2:
            part1, part2 = parts[0], parts[1]
            
            # Use the helper function to check each part
            is_part1_clean_for_range_numbers = _is_number_with_optional_id(part1)
            is_part2_clean_for_range_numbers = _is_number_with_optional_id(part2)

            if is_part1_clean_for_range_numbers and is_part2_clean_for_range_numbers:
                return "Range Numbers"  # e.g., "10 to 20", "4096 (±2g) to 1024 (±8)"
            else:
                # Catches "10A to 20A", "10A (id) to 20 (id2)", "10 to 20 pF" etc.
                return "Range Value"
        else:
            # Malformed " to " string, or " to " was part of an identifier.
            # Default to "Range Value" if " to " is present but not cleanly split into two recognizable parts.
            # This could also be an "Unknown" or "Complex" type depending on requirements.
            return "Range Value" 

    # If not a range, proceed with other classifications
    if "," in main_str: # Check for comma for Multi Value (e.g., "1A, 2A")
        return "Multi Value"
    
    # Classification for single segments (no " to ", no ",")
    # Heuristic for unit presence (any letter, or specific symbols like µ, °, %)
    # This is a broad check; refine if specific non-unit letters should not trigger "Single Value".
    has_unit_chars = bool(re.search(r'[a-zA-Zµ°%]', main_str))
    is_strict_number = bool(re.fullmatch(strict_number_only_pattern, main_str))
    is_num_with_id = _is_number_with_optional_id(main_str) # This will be true for "10 (test)" and also for "10"

    if has_unit_chars:
        # Examples: "10A", "25 ppm", "HIGH", "Test (Value)"
        # If it's "10A (id)", the "A" makes it a Single Value.
        # If it's "10 (id_with_letters)", the letters in id make it Single Value.
        # This path prioritizes "Single Value" if unit-like characters are present anywhere.
        return "Single Value"
    elif is_strict_number: 
        # Purely a number, no unit characters, no parenthetical identifier. E.g., "10", "-5.5"
        return "Number"
    elif is_num_with_id: 
        # It matched the "number with optional id" pattern, but didn't have unit chars
        # and wasn't a strict_number (meaning it must have the parenthetical id).
        # e.g., "10 (id_no_letters_µ°%)", "25 (123)", "0 (0)"
        # Classify as "Number". `classify_value_type_detailed` will add "with Identifier".
        return "Number"
    else:
        # Fallback for anything else that didn't match above.
        # e.g., "---", "NA" (if not caught by unit_chars depending on alphabet), specific symbols not in unit_chars.
        # Maintaining original behavior's fallback to "Number" for minimal change,
        # though this category can be imprecise for non-numeric textual symbols.
        return "Number"


def classify_sub_value(subval: str):
    """
    Classifies a single sub-value string, which might contain main @ condition.

    Args:
        subval (str): The sub-value string (e.g., one item from a comma-split list).

    Returns:
        tuple: (classification_str, has_range_main, has_multi_main,
                has_range_cond, has_multi_cond, cond_item_count, main_item_count)
    """
    subval = str(subval).strip()
    if not subval:
         return ("Empty", False, False, False, False, 0, 0)

    # Split into main and condition parts based on '@' outside parentheses
    main_part = subval
    cond_part = ""
    at_split = split_outside_parens(subval, ['@'])
    if len(at_split) > 1:
        main_part = at_split[0].strip()
        cond_part = "@".join(at_split[1:]).strip()
    elif len(at_split) == 1:
        main_part = at_split[0].strip()

    # Classify each part
    main_class = classify_main(main_part)
    cond_class = classify_condition(cond_part)

    # Determine characteristics based on classification strings
    has_range_in_main = ("Range Value" in main_class) # Check if "Range" is part of the main classification
    has_multi_value_in_main = (main_class == "Multi Value") # Specific check for Multi Value type

    has_range_in_condition = ("Range Condition" in cond_class) # Check if "Range" is part of the condition classification
    has_multiple_conditions = (cond_class == "Multiple Conditions") # Specific check

    # Count items/conditions
    # Note: count_main_items/count_conditions operate on the text after parenthesis removal
    main_item_count = count_main_items(main_part)
    cond_item_count = count_conditions(cond_part) # Number of condition clauses

    # Combine classification strings for the final label
    if main_class and cond_class:
        classification = f"{main_class} with {cond_class}"
    elif main_class:
        classification = main_class # No condition part or condition was empty
    elif cond_class:
         # This case (condition but no main) seems unlikely for valid data
         classification = f"Condition Only: {cond_class}"
    else:
        # Neither part could be classified (e.g., input was just punctuation after cleaning?)
        classification = "Invalid/Empty Structure"

    return (classification,
            has_range_in_main,
            has_multi_value_in_main,
            has_range_in_condition,
            has_multiple_conditions,
            cond_item_count,
            main_item_count)

def classify_value_type_detailed(raw_value: str):
    """
    Process the raw value (which might contain multiple sub-values) and return:
      - final_class (e.g. "Multiple Single Value with Multiple Conditions")
      - identifiers (contents from parentheses)
      - sub_value_count (number of sub-values)
      - final_cond_item_count (uniform condition count or "Mixed")
      - has_range_in_main, has_multi_value_in_main,
        has_range_in_condition, has_multiple_conditions (bool flags)
      - final_main_item_count (uniform main item count or "Mixed")
    """
    # Extract identifiers and clean value
    found_parens = extract_identifiers_detailed(raw_value)
    identifiers = ', '.join(s.strip('()') for s in found_parens)
    clean_value = remove_parentheses_detailed(raw_value).strip()

    # Determine sub-values: if exactly one '@' treat as one sub-value; otherwise split on commas.
    at_count = clean_value.count('@')
    if at_count == 1:
        subvals = [clean_value]
    else:
        subvals = [v.strip() for v in clean_value.split(',') if v.strip()]

    sub_value_count = len(subvals)
    if sub_value_count == 0:
        return ("", identifiers, 0, 0, False, False, False, False, 0)

    # Classify each sub-value
    sub_classifications = []
    sub_range_in_main = []
    sub_multi_in_main = []
    sub_range_in_condition = []
    sub_multi_cond = []
    sub_cond_counts = []
    sub_main_item_counts = []

    for sv in subvals:
        (
            cls,
            has_range_m,
            has_multi_m,
            has_range_c,
            has_multi_c,
            cond_count,
            main_item_count
        ) = classify_sub_value(sv)

        sub_classifications.append(cls)
        sub_range_in_main.append(has_range_m)
        sub_multi_in_main.append(has_multi_m)
        sub_range_in_condition.append(has_range_c)
        sub_multi_cond.append(has_multi_c)
        sub_cond_counts.append(cond_count)
        sub_main_item_counts.append(main_item_count)

    # Determine overall classification
    if sub_value_count == 1:
        final_class = sub_classifications[0]
    else:
        unique_classes = set(sub_classifications)
        if len(unique_classes) == 1:
            final_class = "Multiple " + next(iter(unique_classes))
        else:
            final_class = "Multiple Mixed Classification"

    # Determine uniform vs. mixed counts for conditions
    unique_cond_counts = set(sub_cond_counts)
    if len(unique_cond_counts) == 1:
        final_cond_item_count = unique_cond_counts.pop()
    else:
        final_cond_item_count = "Mixed"

    # Determine uniform vs. mixed counts for main items
    unique_main_item_counts = set(sub_main_item_counts)
    if len(unique_main_item_counts) == 1:
        final_main_item_count = unique_main_item_counts.pop()
    else:
        final_main_item_count = "Mixed"

    has_range_in_main = any(sub_range_in_main)
    has_multi_value_in_main = any(sub_multi_in_main)
    has_range_in_condition = any(sub_range_in_condition)
    has_multiple_conditions = any(sub_multi_cond)

    # Existing extra check: Reassign if exactly two distinct classifications repeat equally.
    if final_class == "Multiple Mixed Classification":
        freq_map = {}
        for cls in sub_classifications:
            freq_map[cls] = freq_map.get(cls, 0) + 1
        if len(freq_map) == 2:
            keys = list(freq_map.keys())
            if freq_map[keys[0]] == freq_map[keys[1]]:
                final_class = "Multiple Repeated Pairs Classification"

    # --- NEW PATCH ---
    # Special reclassification if we got "Repeated Pairs" but suspect bad comma splitting
# --- NEW PATCH with additional check for "Number with Multiple Conditions" ---
    if final_class == "Multiple Repeated Pairs Classification" and len(sub_classifications) % 2 == 0:
        grouped_subs = []
        for i in range(0, len(sub_classifications), 2):
            pair = subvals[i] + ", " + subvals[i + 1]
            grouped_subs.append(pair)

        reclassified = [classify_sub_value(pair)[0] for pair in grouped_subs]
        
        # If all pairs classify as "Single Value with Multiple Conditions"
        if all(cls == "Single Value with Multiple Conditions" for cls in reclassified):
            final_class = "Multiple Single Value with Multiple Conditions"
            final_main_item_count = classify_sub_value(grouped_subs[0])[6]
            final_cond_item_count = classify_sub_value(grouped_subs[0])[5]
            sub_value_count = len(grouped_subs)
        # New branch: If all pairs classify as "Number with Multiple Conditions"
        elif all(cls == "Number with Multiple Conditions" for cls in reclassified):
            final_class = "Multiple Number with Multiple Conditions"
            final_main_item_count = classify_sub_value(grouped_subs[0])[6]
            final_cond_item_count = classify_sub_value(grouped_subs[0])[5]
            sub_value_count = len(grouped_subs)
    if identifiers:
        if final_class == "Number":
            final_class = "Number with Identifier"
        elif final_class == "Multiple Number":
            final_class = "Multiple Number with Identifier"

    return (
        final_class,
        identifiers,
        sub_value_count,
        final_cond_item_count,
        has_range_in_main,
        has_multi_value_in_main,
        has_range_in_condition,
        has_multiple_conditions,
        final_main_item_count
    )



def fix_exceptions(s: str) -> str:
    s = str(s)

    # OLD (too greedy – also hit mOhm, kOhm, …)
    # ohm_pattern = r'(?<![a-zA-Z])([0-9' + "".join(
    #                  re.escape(p) for p in MULTIPLIER_MAPPING.keys()
    #              ) + r'])([Oo][Hh][Mm])(?!\w)'
    # s = re.sub(ohm_pattern, r'\1 \2', s)

    # NEW – insert a space **only** when a digit is glued to “Ohm”
    #      47Ohm   ➜  47 Ohm   (✔ wanted)
    #      1kOhm   ➜  1kOhm    (✔ untouched)
    #      0.5 mOhm ➜  0.5 mOhm (✔ untouched)
    ohm_pattern = r'(?<=\d)([Oo][Hh][Mm])(?!\w)'
    s = re.sub(ohm_pattern, r' \1', s)

    # keep the rest of the clean-up exactly as before
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def replace_numbers_keep_sign_all(s: str) -> str:
    """Replaces all numbers (incl. scientific) with '$', preserving preceding sign."""
    s = str(s)
    # Regex captures optional sign ([+-]?), then digits with optional decimal/exponent.
    # It replaces the number part with '$', keeping the sign captured in group 1 (\1).
    # Handle standalone signs not followed by digits? No, regex requires digits.
    # Ensure it handles cases like ".5" correctly -> might become just "$"?
    # Pattern: (optional sign) followed by (digits OR .digits OR digits.digits) with optional exponent
    pattern = r'([+-]?)(\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?'
    return re.sub(pattern, r'\1$', s)


def replace_numbers_keep_sign_outside_parens(s: str) -> str:
    """Replaces numbers with '$' (keeping sign) only outside parentheses."""
    s = str(s) # Ensure string
    result = []
    i = 0
    depth = 0
    # Use the same number pattern as replace_numbers_keep_sign_all
    number_pattern = re.compile(r'([+-]?)(\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?')

    while i < len(s):
        char = s[i]

        if char == '(':
            depth += 1
            result.append(char)
            i += 1
        elif char == ')':
            depth = max(0, depth - 1) # Prevent negative depth
            result.append(char)
            i += 1
        elif depth == 0:
            # Outside parentheses: check for a number starting at current position
            match_ = number_pattern.match(s, i) # Match from index i
            if match_:
                sign = match_.group(1) if match_.group(1) else '' # Captured sign or empty
                result.append(sign + '$')
                i = match_.end() # Move index past the matched number
            else:
                # Not a number, just append the character
                result.append(char)
                i += 1
        else:
            # Inside parentheses, just append the character
            result.append(char)
            i += 1

    return "".join(result)
# --- NEW HELPER ---------------------------------------------------


def update_number_with_identifier(df: pd.DataFrame) -> pd.DataFrame:
    """
    In any Main-Key block that contains a Unit-row whose
        • Code has “-SN-”  *or*  “-RN-”
        • Value_processed is in parentheses “( … )”
    append “with Identifier” right after the first occurrence of the
    standalone word “Number” **or** “Numbers”, but only if that phrase
    isn’t already “Number(s) with Identifier …”.

    Columns patched (if they exist):
        Classification, DetailedValueType,
        Classification_New, DetailedValueType_New
    """
    # 1 ─ identify blocks that qualify
    trigger = (
        df["Code"].astype(str).str.contains(r"-(SN|RN)-", na=False) &
        df["Attribute"].eq("Unit") &
        df["Value_processed"].astype(str).str.match(r"^\(.*\)$")
    )
    main_keys_to_fix = df.loc[trigger, "Main Key"].unique()
    if len(main_keys_to_fix) == 0:
        return df    # nothing to change

    # 2 ─ regex: Number  *or*  Numbers, but NOT if “with Identifier” follows
    pat  = re.compile(r"\bNumber(s?)\b(?!\s+with\s+Identifier)")

    # 3 ─ target columns
    cols = ["Classification", "DetailedValueType",
            "Classification_New", "DetailedValueType_New"]

    # 4 ─ apply replacement inside each affected Main-Key block
    for key in main_keys_to_fix:
        rows = df["Main Key"] == key
        df.loc[rows, cols] = df.loc[rows, cols].applymap(
            lambda txt: pat.sub(lambda m: f"Number{m.group(1)} with Identifier", str(txt))
            if isinstance(txt, str) else txt
        )

    return df

