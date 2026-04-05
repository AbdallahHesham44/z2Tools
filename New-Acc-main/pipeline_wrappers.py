import pandas as pd
import traceback
import streamlit as st
import io
import os
import re
from itertools import zip_longest
from analysis_helpers import split_outside_parens_preserve # Assuming this is the intended import
def reattach_prefixes(_, prefix_list, __, normalized_col_val):
    """
    Legacy signature:
        _(row_id, prefix_list, original_tokens, normalized_value)
    Only *prefix_list* and *normalized_value* matter for re-attachment.
    """
    prefix_csv = ",".join([p or "" for p in prefix_list])
    return reattach(prefix_csv, normalized_col_val)

   
_COMMA_SPLIT = re.compile(r',(?![^(]*\))')       # comma outside (…) only

def _split_keep_sep(s: str):
    """
    Yield (token, sep) where *sep* is the exact delimiter
    **plus all the spaces that followed it**.
    """
    last = 0
    for m in _COMMA_SPLIT.finditer(s):
        tok = s[last:m.start()]
        # delimiter itself
        sep_end = m.end()
        # swallow any spaces right after the comma
        while sep_end < len(s) and s[sep_end] == " ":
            sep_end += 1
        yield tok, s[m.start():sep_end]
        last = sep_end
    yield s[last:], "" 

from itertools import zip_longest   # still needed elsewhere

# ────────────────────────────────────────────────────────────────────────
#  New, crash-proof version
# ────────────────────────────────────────────────────────────────────────
_SEP_RE = re.compile(r'(\s*,\s*|\s*@\s*|\s+to\s+)', re.IGNORECASE)

def _split_keep_sep_custom(text: str):
    parts = _SEP_RE.split(text)
    out = []
    i = 0
    while i < len(parts):
        token = parts[i]
        sep   = parts[i + 1] if i + 1 < len(parts) else ""
        out.append((token, sep))
        i += 2
    return out

# ---------------------------------------------------------------------------
# Main: spacing-safe, crash-proof, placeholder-aware re-attachment
# ---------------------------------------------------------------------------
from itertools import islice

# ────────────────────────────────────────────────────────────────────
# Helper – split value string but *keep* the separator that follows
# every token (including its trailing spaces)
# recognised separators:   ","   "@",   "to"   (case-insensitive)
_SEP_RE = re.compile(r'(\s*,\s*|\s*@\s*|\s+to\s+)', re.IGNORECASE)

def _split_keep_sep_custom(text: str):
    """
    Return a list of (token, separator) tuples.
    Example: "$A, $A"  →  [("$A", ", "), ("$A", "")]
    """
    parts, out = _SEP_RE.split(text), []
    i = 0
    while i < len(parts):
        token = parts[i]
        sep   = parts[i + 1] if i + 1 < len(parts) else ""
        out.append((token, sep))
        i += 2
    return out
# ────────────────────────────────────────────────────────────────────


def reattach(prefix_csv: str, value_text: str) -> str:
    """
    Re-attach per-token prefixes stored in *prefix_csv* to *value_text*.

    • Handles ',', '@', and 'to' as delimiters.
    • Preserves all original spacing / punctuation.
    • Skips placeholders that are empty or whitespace-only, so no real
      prefix (e.g. “Series”) is ever discarded.
    """
    parts      = _split_keep_sep_custom(value_text)          # [(token, sep), …]
    part_iter  = iter(parts)
    combined   = []

    for raw in prefix_csv.split(','):
        if raw.strip() == "":                                # placeholder → skip
            continue
        try:
            val, sep = next(part_iter)                       # next real token
        except StopIteration:
            break                                            # more prefixes than tokens

        # keep existing spacing, or inject ONE space if none present
        prefix = raw
        if not prefix.endswith(" ") and not val.startswith(" "):
            prefix += " "

        combined.append(f"{prefix}{val}{sep}")

    # dump any remaining tokens unchanged
    combined.extend(f"{val}{sep}" for val, sep in part_iter)
    return "".join(combined)
# --- Main function logic ---

# delimiters should be defined where it's used or passed, if it's module-level here, it's fine.
delimiters = [",", " to ", "@"]

# --- Import original pipeline functions and necessary utilities ---
from detailed_pipeline import detailed_analysis_pipeline as original_detailed_analysis_pipeline_worker
from mapping_utils import read_mapping_file as read_mapping_for_detailed

# For Fixed Pipeline (ensure these are correctly imported if this file handles both wrappers)
from fixed_pipeline import (
    process_single_key as original_process_single_key_for_non_prefixed,
    read_mapping_file as read_mapping_for_fixed,
    LOCAL_BASE_UNITS as fixed_local_base_units,
    get_desired_order as original_get_desired_order,
    display_mapping as original_display_mapping,
    extract_block_texts as original_extract_block_texts,
    classify_with_regex as original_classify_helper,
    fix_exceptions as original_fix_exceptions,
    split_outside_parens as original_split_outside_parens
)
from analysis_helpers import extract_numeric_and_unit_analysis, MULTIPLIER_MAPPING

# Import our new prefix handling utilities
from prefix_utils import extract_prefix_from_string, add_prefix_to_normalized_string

# --- 1. Detailed Pipeline Wrapper (Implementing Option B cleanly) ---
# ─────────────────────────────────────────────────────────────────────
#  Wrapper: detailed pipeline with per-chunk prefix support
#  (keeps original behaviour + fixes comma-space preservation)
# ─────────────────────────────────────────────────────────────────────
def run_detailed_analysis_with_prefix_support(
        input_df: pd.DataFrame,
        mapping_file: str,
        output_file: str
    ) -> str | None:
    st.write("DEBUG (Wrapper): Detailed-wrapper (per-chunk prefix support) starting…")

    if 'Value' not in input_df.columns:
        st.error("Detailed-wrapper: Input DataFrame must contain 'Value' column.")
        return None

    # --- Helpers defined within the scope of this wrapper ---


    df_orig = input_df.copy()
    df_orig['__original_row_id__'] = range(len(df_orig))
    df_orig['Value_Original_Key_For_Merge'] = df_orig['Value']

    worker_df  = pd.DataFrame(index=df_orig.index)
    worker_df["__original_row_id__"] = df_orig["__original_row_id__"]

    extracted_chunk_prefixes_map: dict[int, list[str | None]] = {}
    extracted_original_tokens_map: dict[int, list[str]] = {}

    worker_values_list   = []
    worker_values_e_list = []

    for internal_id, row in df_orig.iterrows():
        raw_text_for_processing = str(row["Value"])
        parts_from_raw_text = split_outside_parens_preserve(raw_text_for_processing, delimiters)

        current_row_prefixes : list[str | None] = []
        current_row_original_worker_tokens : list[str] = []
        current_row_worker_parts_for_join : list[str]    = []

        for tok_from_parts in parts_from_raw_text:
            if tok_from_parts.strip() in delimiters:
                current_row_prefixes.append(None)
                current_row_original_worker_tokens.append(tok_from_parts)
                current_row_worker_parts_for_join.append(tok_from_parts)
                continue

            prefix_identified, remainder_after_prefix, has_prefix_flag = extract_prefix_from_string(tok_from_parts)
            
            if has_prefix_flag:
                leading_ws_in_tok = tok_from_parts[: len(tok_from_parts) - len(tok_from_parts.lstrip())]
                full_remainder_for_worker = leading_ws_in_tok + remainder_after_prefix

                current_row_prefixes.append(prefix_identified)
                current_row_original_worker_tokens.append(full_remainder_for_worker)
                current_row_worker_parts_for_join.append(full_remainder_for_worker)
            else:
                current_row_prefixes.append(None)
                current_row_original_worker_tokens.append(tok_from_parts)
                current_row_worker_parts_for_join.append(tok_from_parts)
        
        extracted_chunk_prefixes_map[internal_id] = current_row_prefixes
        extracted_original_tokens_map[internal_id] = current_row_original_worker_tokens
        worker_values_list.append("".join(current_row_worker_parts_for_join))
        worker_values_e_list.append(str(row.get("Value_E", raw_text_for_processing)))

    worker_df["Value"]   = worker_values_list
    worker_df["Value_E"] = worker_values_e_list

    for col in ("ExceptionFlag", "ConflictFlag", "ExceptionTypes"):
        if col in df_orig.columns:
            worker_df[col] = df_orig[col].values

    try:
        base_units, mult_map = read_mapping_for_detailed(mapping_file)
        worker_out_df = original_detailed_analysis_pipeline_worker(
            worker_df.copy(), base_units, mult_map
        )
    except Exception as exc:
        st.error(f"Detailed-wrapper: Worker crashed ➜ {exc}")
        st.error(traceback.format_exc())
        raise

    worker_cols_to_merge = [c for c in worker_out_df.columns
                            if c not in {'Value', 'Value_E', '__original_row_id__'}]
    
    df_orig['__original_row_id__'] = df_orig['__original_row_id__'].astype(worker_out_df['__original_row_id__'].dtype)

    merged_df = df_orig.merge(
        worker_out_df[['__original_row_id__'] + worker_cols_to_merge],
        on='__original_row_id__',
        how='left'
    )

    for col_to_process in ('Normalized Unit', 'Absolute Unit'):
        if col_to_process in merged_df.columns:
            merged_df[col_to_process] = merged_df.apply(
                lambda r: reattach_prefixes(
                    r['__original_row_id__'],
                    extracted_chunk_prefixes_map.get(r['__original_row_id__'], []),
                    extracted_original_tokens_map.get(r['__original_row_id__'], []),
                    r[col_to_process]
                ),
                axis=1
            )

    merged_df['Value'] = merged_df['Value_Original_Key_For_Merge']
    cols_to_drop_final = ['__original_row_id__', 'Value_Original_Key_For_Merge']
    
    for col_drop in cols_to_drop_final:
        if col_drop in merged_df.columns:
            merged_df.drop(columns=col_drop, inplace=True)

    try:
        merged_df.to_excel(output_file, index=False, engine='openpyxl')
        return output_file
    except Exception as exc:
        st.error(f"Detailed-wrapper: Failed to write “{output_file}” ➜ {exc}")
        st.error(traceback.format_exc())
        return None


def _get_fixed_pipeline_code_definitions_for_prefixed(category_name):
    """
    Parallel to fixed_pipeline.get_code_prefixes_for_category.
    Returns code structures including 'P' for prefix.
    Example structure: {"main_codes": ["M0-SV-P0", "M0-SV-V0", "M0-SV-U0"], "attributes": ["Prefix", "Value", "Unit"]}
    This is a placeholder and needs to be fully implemented for all categories.
    """
    # This needs to be a comprehensive mapping similar to the original one in fixed_pipeline.py
    # For brevity, showing a few examples:
    if category_name == "Number":
        return [{"codes": ["M0-SN-P0", "M0-SN-V0", "M0-SN-U0"], "attributes": ["Prefix", "Value", "Unit"]}]
    elif category_name == "Single Value":
        return [{"codes": ["M0-SV-P0", "M0-SV-V0", "M0-SV-U0"], "attributes": ["Prefix", "Value", "Unit"]}]
    elif category_name == "Range Value": # For range, each part (min/max) could have a prefix conceptually
        return [
            {"codes": ["M0-RV-Pn0", "M0-RV-Vn0", "M0-RV-Un0"], "attributes": ["Prefix", "Value", "Unit"]}, # Min
            {"codes": ["M0-RV-Px0", "M0-RV-Vx0", "M0-RV-Ux0"], "attributes": ["Prefix", "Value", "Unit"]}  # Max
        ]
    # ... other categories ...
    elif category_name == "Number with Single Condition":
        return [
            {"codes": ["M0-SN-P0", "M0-SN-V0", "M0-SN-U0"], "attributes": ["Prefix", "Value", "Unit"]}, # Main
            {"codes": ["M0-SC-P0", "M0-SC-V0", "M0-SC-U0"], "attributes": ["Prefix", "Value", "Unit"]}  # Condition
        ]
    # Fallback for unmapped categories
    st.warning(f"Prefixed Fixed Pipeline: Code definitions not found for category '{category_name}'. Using UNK codes.")
    return [{"codes": [f"M0-UNK-P0", f"M0-UNK-V0", f"M0-UNK-U0"], "attributes": ["Prefix", "Value", "Unit"]}]


def _fill_mapping_for_single_prefixed_part(prefix_val_unit_tuple, block_code_def):
    """
    Helper to fill mapping for one (prefix, value, unit) part.
    block_code_def example: {"codes": ["M0-SV-P0", "M0-SV-V0", "M0-SV-U0"], "attributes": ["Prefix", "Value", "Unit"]}
    """
    (p_str, v_str, u_str) = prefix_val_unit_tuple # (prefix_for_this_block, value_for_this_block, unit_for_this_block)
    result = {}
    codes = block_code_def.get("codes", [])
    attributes = block_code_def.get("attributes", [])

    for i, attr in enumerate(attributes):
        if i < len(codes):
            code_to_use = codes[i]
            if attr == "Prefix":
                result[code_to_use] = {"value": p_str if pd.notna(p_str) and p_str is not None else ""}
            elif attr == "Value":
                result[code_to_use] = {"value": v_str if pd.notna(v_str) and v_str is not None else ""}
            elif attr == "Unit":
                result[code_to_use] = {"value": u_str if pd.notna(u_str) and u_str is not None else ""}
    return result


def _generate_mapping_for_prefixed_parts(list_of_pvu_tuples, category_name_of_value_part):
    """
    Generates the code map for a list of (prefix, value, unit) tuples.
    """
    static_block_definitions = _get_fixed_pipeline_code_definitions_for_prefixed(category_name_of_value_part)
    mapping = {}

    # Assign static blocks
    for i, p_v_u_tuple in enumerate(list_of_pvu_tuples):
        if i < len(static_block_definitions):
            block_def = static_block_definitions[i]
            mapping.update(_fill_mapping_for_single_prefixed_part(p_v_u_tuple, block_def))
        else:
            # Potentially handle dynamic parts (MV-, MC-)
            # This requires defining how dynamic prefix codes are generated (e.g., M0-MV-P1, M0-MV-V1, M0-MV-U1)
            # For now, focus on static parts based on category definition.
            # Determine if it's MV or MC based on category_name_of_value_part
            dynamic_prefix_base = None
            if "Multiple Conditions" in category_name_of_value_part: dynamic_prefix_base = "MC-"
            elif category_name_of_value_part.startswith("Multi Value"): dynamic_prefix_base = "MV-"

            if dynamic_prefix_base:
                counter = i - len(static_block_definitions) + 1 # Start counter for dynamic parts
                dyn_block_def = {
                    "codes": [f"M0-{dynamic_prefix_base}P{counter}", f"M0-{dynamic_prefix_base}V{counter}", f"M0-{dynamic_prefix_base}U{counter}"],
                    "attributes": ["Prefix", "Value", "Unit"]
                }
                mapping.update(_fill_mapping_for_single_prefixed_part(p_v_u_tuple, dyn_block_def))
            else:
                st.warning(f"Prefixed Fixed Pipeline: Ran out of static block definitions for category '{category_name_of_value_part}' "
                           f"and no dynamic logic matched for part: {p_v_u_tuple}")
                # Add some error codes for unmapped parts
                err_idx = i - len(static_block_definitions)
                mapping[f"M0-ERR-P{err_idx}"] = {"value": p_v_u_tuple[0] or "ERR_P_UNMAPPED"}
                mapping[f"M0-ERR-V{err_idx}"] = {"value": p_v_u_tuple[1] or "ERR_V_UNMAPPED"}
                mapping[f"M0-ERR-U{err_idx}"] = {"value": p_v_u_tuple[2] or "ERR_U_UNMAPPED"}

    return mapping

def _get_desired_order_with_prefixes():
    """
    Generates a desired order list that includes 'P' codes.
    This should be comprehensive.
    """
    # Placeholder - This function needs careful implementation
    # Example for a few codes:
    new_order = [
        "M0-SN-P0", "M0-SN-V0", "M0-SN-U0",
        "M0-SV-P0", "M0-SV-V0", "M0-SV-U0",
        # For Range Value
        "M0-RV-Pn0", "M0-RV-Vn0", "M0-RV-Un0", "M0-RV-Px0", "M0-RV-Vx0", "M0-RV-Ux0",
        # For Multi Value (dynamic, up to a limit)
        "M0-MV-P1", "M0-MV-V1", "M0-MV-U1", "M0-MV-P2", "M0-MV-V2", "M0-MV-U2",
        # For Single Condition
        "M0-SC-P0", "M0-SC-V0", "M0-SC-U0",
        # For Range Condition
        "M0-RC-Pn0", "M0-RC-Vn0", "M0-RC-Un0", "M0-RC-Px0", "M0-RC-Vx0", "M0-RC-Ux0",
        # For Multi Condition (dynamic)
        "M0-MC-P1", "M0-MC-V1", "M0-MC-U1", "M0-MC-P2", "M0-MC-V2", "M0-MC-U2",
    ]
    # Append original error/unknown codes
    original_err_codes = [c for c in original_get_desired_order() if "ERR-" in c or "UNK-" in c]
    new_order.extend(original_err_codes)
    return list(dict.fromkeys(new_order)) # Ensure unique

def _display_mapping_with_prefixes(mapping_dict, desired_order_with_p_codes, category, main_key, overall_prefix_for_chunk):
    """
    Parallel to fixed_pipeline.display_mapping, but handles 'Prefix' attribute.
    Adds 'OverallChunkPrefix' to each row.
    """
    output_rows = []
    processed_codes = set()

    def get_attribute_from_pvu_code(code): # P, V, or U
        if "-P" in code: return "Prefix"
        if "-U" in code: return "Unit"
        if "-V" in code: return "Value"
        return "Value" # Fallback

    for code in desired_order_with_p_codes:
        if code in mapping_dict:
            attr = get_attribute_from_pvu_code(code)
            value_data = mapping_dict[code]
            val_str = value_data.get("value", "")
            if pd.isna(val_str): val_str = ""
            
            row = {
                "Main Key": main_key,
                "Category": category, # Category of the value part (after prefix removed)
                "Attribute": attr,
                "Code": code,
                "Value": str(val_str),
                "OverallChunkPrefix": overall_prefix_for_chunk if overall_prefix_for_chunk else ""
            }
            output_rows.append(row)
            processed_codes.add(code)

    # Add any codes from mapping_dict not in desired_order (e.g., dynamic codes beyond typical count)
    for code in sorted([c for c in mapping_dict if c not in processed_codes]):
        attr = get_attribute_from_pvu_code(code)
        value_data = mapping_dict[code]
        val_str = value_data.get("value", "")
        if pd.isna(val_str): val_str = ""
        row = {
            "Main Key": main_key, "Category": category, "Attribute": attr,
            "Code": code, "Value": str(val_str),
            "OverallChunkPrefix": overall_prefix_for_chunk if overall_prefix_for_chunk else ""
        }
        output_rows.append(row)
        
    return output_rows


def process_single_key_with_prefix_support(
    main_key_original_value: str,
    base_units,               # from fixed_pipeline context
    multipliers_dict          # from fixed_pipeline context
):
    """
    New function parallel to fixed_pipeline.process_single_key.
    Handles values where parts (chunks) might be prefixed.
    """
    main_key_cleaned_for_struct_class = original_fix_exceptions(str(main_key_original_value).strip())

    try:
        # --- Pass mapping args into the classifier! ---
      (category_for_struct, _, num_sub_values, _, _, _, _, _, _) = original_classify_helper(main_key_cleaned_for_struct_class)
 


    except Exception as e:
        st.error(f"Prefixed Fixed Pipeline: Error during initial classification of '{main_key_cleaned_for_struct_class}': {e}")
        category_for_struct = "Classification Error"

    if category_for_struct == "Classification Error":
        err_map = {"ERR-CLASS-PFX": {"value": "Initial classification failed for prefix handling."}}
        return original_display_mapping(err_map, original_get_desired_order(), category_for_struct, main_key_original_value)

    # … rest of function unchanged …


    all_output_rows_for_main_key = []
    is_top_level_multiple_type = category_for_struct.startswith("Multiple ")

    # Chunking: Always split by comma first if it's a "Multiple" type, otherwise treat as single chunk.
    # The string used for chunking should be the one that has had fix_exceptions applied.
    chunks_to_process_individually = []
    if is_top_level_multiple_type:
        # Use original_split_outside_parens for comma splitting.
        # The input to this split should be the string as it would be before block extraction.
        raw_comma_chunks = original_split_outside_parens(main_key_cleaned_for_struct_class, [','])
        chunks_to_process_individually = [chk.strip() for chk in raw_comma_chunks if chk.strip()]
        
        # Re-apply grouping logic if it's a special "Multiple ... with Multiple Conditions" type
        # This grouping logic from original_process_single_key might need to be here if chunks are re-combined.
        # For now, assume each comma-separated part is a distinct "chunk" for prefix processing.
        if not chunks_to_process_individually: # If splitting resulted in no valid chunks
            err_map = {"ERR-CHUNK-PFX": {"value": "No processable comma-chunks found for 'Multiple' type."}}
            return original_display_mapping(err_map, original_get_desired_order(), category_for_struct, main_key_original_value)
    else:
        chunks_to_process_individually = [main_key_cleaned_for_struct_class]


    for chunk_index, individual_chunk_text in enumerate(chunks_to_process_individually):
        # For each chunk, identify its own prefix and the remaining value part
        this_chunk_prefix, value_part_of_this_chunk, _ = extract_prefix_from_string(individual_chunk_text)

        # Now, classify the 'value_part_of_this_chunk' to understand its structure (e.g., Single, Range, with Condition)
        # This classification will drive how 'value_part_of_this_chunk' is broken into blocks.
        try:
            (category_of_value_part, _, _, _, _, _, _, _, _) = original_classify_helper(value_part_of_this_chunk)
             
           
            if not category_of_value_part or category_of_value_part in ["Empty", "Invalid/Empty Structure"]:
                 # If value_part_of_this_chunk is empty (e.g. input was just "parallel"), treat as unknown.
                category_of_value_part = "Unknown Value Part" if value_part_of_this_chunk else "Empty Value Part"
        except Exception as e:
            st.error(f"Prefixed Fixed Pipeline: Error classifying value part '{value_part_of_this_chunk}': {e}")
            category_of_value_part = "Value Part Classification Error"

        list_of_pvu_tuples_for_this_chunk = [] # List of (prefix, value, unit) for blocks in this chunk
        
        if category_of_value_part not in ["Unknown Value Part", "Value Part Classification Error", "Empty Value Part"]:
            # Extract logical blocks from 'value_part_of_this_chunk' (e.g., main, condition; or min, max)
            # original_extract_block_texts takes the value string and its category.
            blocks_within_value_part = original_extract_block_texts(value_part_of_this_chunk, category_of_value_part)

            for block_idx, block_text in enumerate(blocks_within_value_part):
                # The `this_chunk_prefix` applies to the first "main" conceptual part of `value_part_of_this_chunk`.
                # If `value_part_of_this_chunk` is "10A @ 5V", and `this_chunk_prefix` was "parallel",
                # then "parallel" applies to "10A". "5V" does not get "parallel".
                # This requires understanding if `block_text` is a main part or a condition part.
                # A simple heuristic: prefix applies if block_idx is 0 (first block). More complex for ranges.
                prefix_for_this_block = this_chunk_prefix if block_idx == 0 else None # Simplified logic

                # Parse this block_text (which is already prefix-stripped) into value and unit
                # Use analysis_helpers.extract_numeric_and_unit_analysis for robust parsing
                num_val, mult_s, base_u, _, err_flag, orig_rest_str = extract_numeric_and_unit_analysis(
                    block_text, base_units, multipliers_dict # multipliers_dict is crucial here
                )
                
                val_str_for_block, unit_str_for_block = "", ""
                # ----------------------------------------------------
                # Build VALUE  (numeric only, no multiplier, trim .0)
                # Build UNIT   (multiplier + base unit + any extras)
                # ----------------------------------------------------
                if not err_flag and num_val is not None:
                    # 1️⃣ numeric part
                    if float(num_val).is_integer():
                        val_numeric = str(int(num_val))
                    else:
                        val_numeric = f"{num_val}".rstrip("0").rstrip(".")

                    # 2️⃣ attach multiplier to VALUE (no space)
                    if mult_s and mult_s != "1":
                        val_str_for_block = val_numeric + mult_s
                    else:
                        val_str_for_block = val_numeric

                    # 3️⃣ UNIT = base unit only (multiplier stripped)
                    unit_pieces = []
                    if base_u:
                        unit_pieces.append(base_u)

                    # any residual annotation (e.g. "(typ)")
                    extra = (orig_rest_str or "").strip()
                    
                    # Skip if the “extra” is just the base unit
                    # or the multiplier + base unit we already captured (e.g. "mA")
                    dup_candidates = {base_u, f"{mult_s}{base_u}".strip()}
                    if extra and extra not in unit_pieces and extra not in dup_candidates:
                        unit_pieces.append(extra)

                    unit_str_for_block = "".join(unit_pieces)
                else:
                    # fallback when parse fails
                    val_str_for_block = block_text.strip()
                    unit_str_for_block = ""

                
                list_of_pvu_tuples_for_this_chunk.append(
                    (prefix_for_this_block, val_str_for_block.strip(), unit_str_for_block.strip())
                )
        else: # category_of_value_part indicates an issue or empty
             list_of_pvu_tuples_for_this_chunk.append(
                 (this_chunk_prefix, value_part_of_this_chunk, "") # Treat whole value_part as value
             )

        # Generate mapping for this chunk's parsed parts (prefix, value, unit)
        code_map_for_chunk = _generate_mapping_for_prefixed_parts(
            list_of_pvu_tuples_for_this_chunk, category_of_value_part
        )

        # Apply M<n>- or M0- prefix to the codes
        # M_code_prefix = f"M{chunk_index + 1}-" if is_top_level_multiple_type else "M0-"
        # The above logic has an off-by-one potential if num_sub_values is used for M index
        # The original fixed_pipeline used chunk_index for M-prefix.
        
        # Determine the M-prefix (M0- for single chunk main_keys, M<idx>- for multi-chunk main_keys)
        actual_m_prefix = f"M{chunk_index + 1}-" if is_top_level_multiple_type and len(chunks_to_process_individually) > 1 else "M0-"
        
        final_code_map_with_m_prefix = {}
        for code, val_info in code_map_for_chunk.items():
            if code.startswith("M0-"): # Generated codes from _generate_mapping... use M0-
                final_code_map_with_m_prefix[actual_m_prefix + code[3:]] = val_info
            else: # Error codes, etc.
                final_code_map_with_m_prefix[code] = val_info
        
        # Get desired order for these M-prefixed codes
        desired_order_for_m_prefixed_codes = []
        base_desired_order_p_codes = _get_desired_order_with_prefixes() # This has M0-P...
        for c_pattern in base_desired_order_p_codes:
            if c_pattern.startswith("M0-"):
                desired_order_for_m_prefixed_codes.append(actual_m_prefix + c_pattern[3:])
            else: # Keep ERR, UNK codes as is
                desired_order_for_m_prefixed_codes.append(c_pattern)
        
        # Create output rows for this chunk
        # The `category` for display_mapping should be `category_of_value_part`
        # `overall_prefix_for_chunk` is `this_chunk_prefix`
        chunk_output_rows = _display_mapping_with_prefixes(
            final_code_map_with_m_prefix,
            desired_order_for_m_prefixed_codes,
            category_of_value_part, # Category of the value after its prefix was stripped
            main_key_original_value, # The original full Main Key for this row
            this_chunk_prefix # The prefix identified for this specific chunk
        )
        all_output_rows_for_main_key.extend(chunk_output_rows)

    return all_output_rows_for_main_key


def run_fixed_pipeline_with_prefix_support(file_bytes: bytes, mapping_file_path: str):
    """
    Wrapper for the fixed pipeline to handle prefixes.
    It decides whether to call the original process_single_key or the new
    process_single_key_with_prefix_support based on prefix detection.
    """
    st.write("DEBUG: Running Fixed Pipeline with Prefix Support...")
    all_processed_rows = []
    try:
        base_units_from_file, multipliers_map_from_file = read_mapping_for_fixed(mapping_file_path)
        # original fixed_pipeline uses fixed_local_base_units and MULTIPLIER_MAPPING from analysis_helpers.
        # We need to pass the same to both original_process_single_key and our new one.
        # The MULTIPLIER_MAPPING from analysis_helpers is used by fixed_pipeline.
        current_multiplier_mapping = MULTIPLIER_MAPPING # from analysis_helpers, imported here
        
        combined_base_units = fixed_local_base_units.union(base_units_from_file)

        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        for sheet_name in xls.sheet_names:
            st.write(f"DEBUG: Prefixed Fixed Pipeline processing sheet: '{sheet_name}'")
            sheet_df = pd.read_excel(xls, sheet_name=sheet_name)
            if 'Value' not in sheet_df.columns:
                st.warning(f"Sheet '{sheet_name}' skipped (Fixed Prefix): Missing 'Value' column.")
                continue
            sheet_df['Value'] = sheet_df['Value'].fillna('').astype(str)

            for _, row_series in sheet_df.iterrows():
                main_key = row_series.get('Value', '').strip()
                if not main_key:
                    continue

                # --- Prefix Detection to Route Processing ---
                # Check if the main_key or any of its comma-separated parts start with a known prefix.
                # This is a simple check; more sophisticated logic might be needed if prefixes interact with " to " or "@".
                # For now, if any part appears prefixed, use the new handler.
                is_value_potentially_prefixed = False
                # Clean exceptions (like ±) before checking for prefixes for routing,
                # so "±parallel 5A" is not mistaken. Prefixes are typically words.
                # However, extract_prefix_from_string handles arbitrary text.
                # Let's check on the raw main_key for now.
                
                # Split by comma (outside parens) to check parts, as per user: "if any part of the value has the prefix"
                # Use original_split_outside_parens for consistency.
                # And original_fix_exceptions before splitting for stable parsing
                temp_cleaned_main_key = original_fix_exceptions(main_key)
                potential_chunks = original_split_outside_parens(temp_cleaned_main_key, [','])
                if not potential_chunks: # If split results in nothing (e.g. main_key was just ",")
                    potential_chunks = [temp_cleaned_main_key] # process the key as one chunk

                for chunk_part in potential_chunks:
                    _, _, identified_prefix_in_chunk = extract_prefix_from_string(chunk_part.strip())
                    if identified_prefix_in_chunk:
                        is_value_potentially_prefixed = True
                        break
                
                result_rows_for_key = []
                if is_value_potentially_prefixed:
                    # Use the new prefix-aware processor
                    result_rows_for_key = process_single_key_with_prefix_support(
                        main_key, # Pass original main_key
                        combined_base_units,
                        current_multiplier_mapping # Pass the globally agreed multiplier map
                    )
                else:
                    # Use the original processor for non-prefixed values
                    result_rows_for_key = original_process_single_key_for_non_prefixed(
                        main_key,
                        combined_base_units,
                        current_multiplier_mapping # Pass the globally agreed multiplier map
                    )
                
                # Add original row data to the processed results
                original_row_data_dict = row_series.to_dict()
                for processed_dict_row in result_rows_for_key:
                    final_row_data = original_row_data_dict.copy()
                    final_row_data.update(processed_dict_row)
                    final_row_data["Sheet"] = sheet_name
                    all_processed_rows.append(final_row_data)

        if not all_processed_rows:
            st.warning("Prefixed Fixed Pipeline: No output rows generated after processing all sheets.")
            return pd.DataFrame() # Return empty DataFrame if nothing was processed
            
        return pd.DataFrame(all_processed_rows)

    except FileNotFoundError:
        st.error(f"Prefixed Fixed Pipeline Error: Mapping file not found at '{mapping_file_path}'.")
        return None
    except ValueError as e: # From read_mapping_file
        st.error(f"Prefixed Fixed Pipeline Error reading mapping file: {e}.")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred in run_fixed_pipeline_with_prefix_support: {e}")
        st.error(traceback.format_exc())
        return None
