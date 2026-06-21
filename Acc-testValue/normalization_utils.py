import re
import pandas as pd
from mapping_utils import MULTIPLIER_MAPPING
from prefix_utils import STRING_KEYWORDS
# ====== CONFIG ======
SLASH_RANGE_UNITS = ["ppm/°C", "ppm/K"]  # Add more units if needed
DELIMITERS = ['@', ',', ';']             # These delimiters split the full value into tokens
_THOUSANDS_WHITELIST =["Cycles", "Hours","Steps"]
_THOUSANDS_SEP_RE   = re.compile(r'(?<=\d),(?=\d{3}(?:\D|$))')
# -------------------------------------------------------------------------------
# UPDATED REGEX PATTERNS
#
# For the ± pattern, we now allow for optional leading spaces:
#
#     ^(?P<leading>\s*)(?P<sign>[±])(?P<number>\d+(?:\.\d+)?)(?P<rest>.*)$
#
# This pattern does the following:
#   - (?P<leading>\s*): Captures any leading spaces.
#   - (?P<sign>[±]): Captures the "±" sign.
#   - (?P<number>\d+(?:\.\d+)?): Captures the numeric portion.
#   - (?P<rest>.*): Captures the remainder of the token exactly as-is.
#
# -------------------------------------------------------------------------------
pm_pattern = re.compile(r'^(?P<leading>\s*)(?P<sign>[±])(?P<number>\d+(?:\.\d+)?)(?P<rest>.*)$')
dash_spaced_pattern = re.compile(r'^(\d+(?:\.\d+)?)[\s]*-[\s]*(\d+(?:\.\d+)?)(\s+\S.+)$')
dash_attached_pattern = re.compile(r'^(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)(\S+)$')
slash_pattern = re.compile(r'^([±+-]?\d+(?:\.\d+)?)/\s*([±+-]?\d+(?:\.\d+)?)(\s*\S+)$')
flexible_dash_pattern = re.compile(r'''
    ^
    ([+-]?\d+(?:\.\d+)?)        # num1
    [\s]*[-–—][\s]*             # any dash with optional spaces
    ([+-]?\d+(?:\.\d+)?)        # num2
    (\s*\S.*)?                  # unit and any spacing/text that follows
    $
''', re.VERBOSE)

def is_exception(candidate: str) -> bool:
    """
    Returns True if the token matches one of the special “exception”
    formats (± range, dash range, slash range). Otherwise False.
    """
    c = candidate.strip()

    # ± pattern
    m_pm = pm_pattern.match(c)
    if m_pm:
        return True

    # dash-range patterns
    if dash_spaced_pattern.match(c):
        return True
    if dash_attached_pattern.match(c):
        return True

    # slash-range pattern
    m = slash_pattern.match(c)
    if m:
        a, b, unit = m.group(1), m.group(2), m.group(3).strip()

        # --- strip exactly ONE valid multiplier prefix, if present ---
        for p in sorted(MULTIPLIER_MAPPING.keys(), key=len, reverse=True):
            if unit.startswith(p):
                unit = unit[len(p):]
                break

        if (
            unit in SLASH_RANGE_UNITS           # e.g. "ppm/°C", "ppm/K"
            and not a.startswith('±')
            and not b.startswith('±')
        ):
            return True

    # default: not an exception
    return False

# ====== 1. Utility: Split outside parentheses ======
def split_outside_parentheses(text, delimiters):
    """
    Splits 'text' on any of the given 'delimiters', but only if the delimiter is found outside any parentheses.
    """
    result = []
    current = []
    level = 0
    i = 0
    while i < len(text):
        char = text[i]
        if char == '(':
            level += 1
        elif char == ')':
            level = max(0, level - 1)

        if level == 0:
            matched = None
            for d in delimiters:
                if text[i:i+len(d)] == d:
                    matched = d
                    break
            if matched:
                result.append(''.join(current))
                result.append(matched)
                current = []
                i += len(matched)
                continue

        current.append(char)
        i += 1

    if current:
        result.append(''.join(current))
    return result


# ====== 2. Transform subparts based on the patterns ======
# FF-main/normalization_utils.py

# ... (imports, regex definitions including the new flexible_dash_pattern) ...

def transform_subvalue(subval: str) -> tuple[str, bool, str]:
    """
    Attempts to transform a single token based on the recognized patterns.
    Returns a tuple: (updated_string, was_transformed, transformation_type).
    Pattern priority:
      1. ± pattern
      2. Dash/flexible dash pattern (output always: "val unit to val unit" with normalized spacing)
      3. Slash pattern (output: "valunit to valunit" with stripped unit, original style)
    """
    # ── preserve any original leading/trailing spaces ──────────────────────────
    pre_ws  = subval[:len(subval) - len(subval.lstrip())]
    post_ws = subval[len(subval.rstrip()):]
    core    = subval.strip()

    # -- ± Pattern (Priority 1) --
    m_pm = pm_pattern.match(core)
    if m_pm:
        leading   = m_pm.group('leading')
        number    = m_pm.group('number')
        rest      = m_pm.group('rest')
        left_val  = f'{leading}-{number}{rest}'
        right_val = f'+{number}{rest}'
        return f'{pre_ws}{left_val} to {right_val}{post_ws}', True, "±"

    # -- Flexible Dash Range Pattern (Priority 2) --
    m_flex_dash = flexible_dash_pattern.match(core)
    if m_flex_dash:
        num1_str = m_flex_dash.group(1)
        num2_str = m_flex_dash.group(2)
        raw_unit_part_from_regex = m_flex_dash.group(3)
        stripped_unit = raw_unit_part_from_regex.strip() if raw_unit_part_from_regex is not None else ""
        unit_seg = f" {stripped_unit}" if stripped_unit else ""
        return f'{pre_ws}{num1_str}{unit_seg} to {num2_str}{unit_seg}{post_ws}', True, "dash"

    # -- Slash range (Priority 3) --
    m_slash = slash_pattern.match(core)
    if m_slash:
        a_s, b_s, unit_s_raw = m_slash.group(1), m_slash.group(2), m_slash.group(3)
        unit_s = unit_s_raw.strip()
        if unit_s in SLASH_RANGE_UNITS and not a_s.startswith('±') and not b_s.startswith('±'):
            return f'{pre_ws}{a_s}{unit_s} to {b_s}{unit_s}{post_ws}', True, "slash"

    # -- No pattern matched --
    return subval, False, ""


# ====== 3. Check & transform the token (with conflict logic) ======
def handle_token(token: str) -> (str, bool, str, bool):
    """
    Processes a single token.
      - If the token contains " to ", it splits the token and then checks if
        either side is an exception (i.e., matches ±, dash, or slash patterns).
        If so, the token is flagged as a conflict and left unchanged.
      - Otherwise, the token is processed normally.
    
    Returns a tuple: (updated_token, was_transformed, transformation_type, conflict_flag)
    """
    original = token

    if " to " in original:
        parts = original.split(" to ", 1)
        left_part, right_part = parts[0], parts[1]
        if is_exception(left_part) or is_exception(right_part):
            return original, False, "", True
        else:
            new_val, was_tf, tf_type = transform_subvalue(original)
            return new_val, was_tf, tf_type, False

    new_val, was_tf, tf_type = transform_subvalue(original)
    return new_val, was_tf, tf_type, False


# ====== 4. Full value processor (chunked) ======
# In normalization_utils.py

# ... (make sure 'from prefix_utils import STRING_KEYWORDS' and other constants are at the top)

# <<< FULLY UPDATED AND CORRECTED FUNCTION >>>
def process_value_full(raw_value: str):
    """
    Splits the raw value by ALL defined delimiters, checks EACH non-delimiter chunk
    to see if it's a String Keyword, injects the number '5' if it is, and then
    reconstructs the value string. This ensures that only string keywords are modified.

    This now correctly handles keywords separated by " to " and "@".

    Returns (updated_value, exception_flag, exception_types, conflict_flag, conflict_count)
    """
    # Initialize the tracking lists for all exceptions found in the value
    all_exception_types = []

    # --- Stage 1: Pre-process for Thousands Separator (Your existing logic) ---
    # This runs once on the entire raw value string.
    value_to_process = str(raw_value) # Work with a copy
    
    parts = value_to_process.strip().rsplit(" ", 1)
    if len(parts) == 2:
        num_part, unit_part = parts
    else:
        num_part, unit_part = value_to_process, ""

    had_thousands = (
        unit_part in _THOUSANDS_WHITELIST
        and bool(_THOUSANDS_SEP_RE.search(num_part))
    )
    if had_thousands:
        value_to_process = _THOUSANDS_SEP_RE.sub("", value_to_process)
        all_exception_types.append("ThousandsSeparator")


    # --- Stage 2: Handle String Keywords and other normalizations chunk by chunk ---

    # Split the potentially comma-cleaned value into its fundamental parts, including delimiters
    # Using your original DELIMITERS list: ['@', ',', ';'] (We'll add ' to ' for this split)
    # The split_outside_parentheses needs to know about all possible separators.
    # Note: `handle_token` has its own ' to ' logic, but for keyword detection, we must split by it first.
    all_delimiters_for_split = DELIMITERS + [" to "]
    initial_tokens = split_outside_parentheses(value_to_process, all_delimiters_for_split)

    processed_tokens = []
    # These flags are for non-string exceptions
    conflict_found = False
    conflict_count = 0

    for token in initial_tokens:
        stripped_token = token.strip()

        # Check if this specific chunk is a String Keyword
        if stripped_token.lower() in [kw.lower() for kw in STRING_KEYWORDS]:
            # It's a keyword. Inject the number '5' and flag the "String" exception.
            # We preserve original spacing around the keyword by using the original token.
            leading_space = token[:len(token) - len(token.lstrip())]
            trailing_space = token[len(token.rstrip()):]
            
            # The new token becomes, e.g., " 5adj "
            new_token = f"{leading_space}5{stripped_token}{trailing_space}"

            processed_tokens.append(new_token)
            all_exception_types.append("String")

        else:
            # --- IF NOT A KEYWORD, THIS IS YOUR ORIGINAL LOGIC PATH ---

            # Check if the token is a delimiter itself, and if so, pass it through
            # This needs to be robust for multi-character delimiters like " to "
            is_delimiter = False
            for d in all_delimiters_for_split:
                if token == d:
                    is_delimiter = True
                    break
            
            if is_delimiter:
                processed_tokens.append(token)
            else:
                # If it's not a delimiter and not a string keyword, it's a normal value part.
                # Process it with your existing normalization logic for ±, slash, etc.
                updated, was_exc, exc_type, conflict = handle_token(token)
                processed_tokens.append(updated)
                if was_exc:
                    all_exception_types.append(exc_type)
                if conflict:
                    conflict_found = True
                    conflict_count += 1
                
    # Reconstruct the full string from the processed tokens
    updated_value = "".join(processed_tokens)
    
    # Determine the final exception flag based on ALL collected exception types
    exception_flag = True if all_exception_types else False

    return updated_value, exception_flag, all_exception_types, conflict_found, conflict_count



# ====== 5. Main Excel processor ======
def process_excel(input_path: str, output_path: str = None) -> pd.DataFrame:
    """
    Reads an Excel file (which must contain a 'Value' column), processes each cell using
    'process_value_full', and writes out the resulting DataFrame with the following new columns:
      - UpdatedValue  (normalized version of 'Value')
      - ExceptionFlag
      - ExceptionTypes
      - ExceptionCount
      - ExceptionTypeCount
      - ConflictFlag
      - ConflictCount
    
    In your overall workflow, you can then rename UpdatedValue to Value_Normalized.
    """
    df = pd.read_excel(input_path)
    if "Value" not in df.columns:
        raise ValueError("Input file must contain a 'Value' column.")

    updated_values = []
    exception_flags = []
    exception_types_list = []
    exception_count = []
    exception_type_count = []
    conflict_flags = []
    conflict_counts = []

    for val in df["Value"].fillna(""):
        val_str = str(val)
        (updated_val,
         found_exception,
         types,
         conflict,
         c_count) = process_value_full(val_str)
        
        updated_values.append(updated_val)
        exception_flags.append("YES" if found_exception else "NO")
        exception_types_list.append(", ".join(sorted(set(types))))
        exception_count.append(len(types))
        exception_type_count.append(len(set(types)))
        conflict_flags.append("YES" if conflict else "NO")
        conflict_counts.append(c_count)

    df["UpdatedValue"] = updated_values
    df["ExceptionFlag"] = exception_flags
    df["ExceptionTypes"] = exception_types_list
    df["ExceptionCount"] = exception_count
    df["ExceptionTypeCount"] = exception_type_count
    df["ConflictFlag"] = conflict_flags
    df["ConflictCount"] = conflict_counts

    if output_path:
        df.to_excel(output_path, index=False)

    return df


