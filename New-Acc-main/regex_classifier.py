# --- START OF FILE regex_classifier.py ---

import pandas as pd
import re
from collections import Counter
from typing import List
import streamlit as st   # still used for debug/info
import traceback
from prefix_utils import extract_prefix_from_string
from analysis_helpers import extract_numeric_and_unit_analysis
from prefix_utils import STRING_KEYWORDS
# Global storage
classification_patterns: list[tuple[re.Pattern, str]] = []
multipliers_list: list[str] = []
units_list:       list[str] = []
NUM_REGEX = r"[+-]?\d+(?:\.\d+)?"

UNITS_ALLOW_SPACE: set[str] = {
    "Sec", "AWG", "Cycles", "Hrs", "Ohm"
}

# Separators that must be surrounded by ONE space:
#   123 to 456     (space–to–space)
#   1, 2, 3        (comma, then ONE space, NO space before comma)
#   5 V @ 1 MHz    (space–@–space)
SEPARATORS_REQUIRE_SPACE = {"to", ",", "@"}
# -------------------------------------------------------------------
#  STRICT separator-spacing QA
# -------------------------------------------------------------------
def _strip_multiplier_once(unit_tok: str, multipliers: list[str]) -> str:
    for p in sorted(multipliers, key=len, reverse=True):
        if unit_tok.startswith(p):
            return unit_tok[len(p):]
    return unit_tok


def detect_unit_spacing_violations_regex(raw: str,
                                         multipliers: list[str],
                                         allowed_space_units: set[str] = UNITS_ALLOW_SPACE
                                        ) -> list[str]:
    """
    Flags unit-tokens that violate spacing rules in *any* classifier pattern.

    ─────────────────────────────────────────────────────────────
      • units in *allowed_space_units*     →  exactly ONE space
      • all other units                    →  NO space
      • special base unit  'Min'
            – if it belongs to a condition (after local '@')
                                              →  NO space
            – otherwise                       →  exactly ONE space
    ─────────────────────────────────────────────────────────────
    Works block-wise:   main-part [, main-part …] , condition-part @ …, …
    """
    import re

    s = str(raw)
    offenders: list[str] = []
    token_re = re.compile(r'([+\-]?\d+(?:\.\d+)?)(\s*)([a-zA-Zµ]+)')

    depth = 0               # parenthesis depth
    in_condition = False    # reset at each top-level comma

    i = 0
    while i < len(s):
        ch = s[i]

        # ── update state ─────────────────────────────────────────────
        if ch == '(':
            depth += 1
            i += 1
            continue
        if ch == ')' and depth > 0:
            depth -= 1
            i += 1
            continue

        # top-level comma → new block, condition flag resets
        if ch == ',' and depth == 0:
            in_condition = False
            i += 1
            continue

        # '@' outside parentheses → everything after is condition side
        if ch == '@' and depth == 0:
            in_condition = True
            i += 1
            continue

        # ── try to match a number+unit token starting at i ───────────
        m = token_re.match(s, i)
        if m:
            sep      = m.group(2)           # spaces between number & unit
            unit_tok = m.group(3)
            base     = _strip_multiplier_once(unit_tok, multipliers)

            # ----- special rule for 'Min' ---------------------------
            if base == "Min":
                if in_condition:
                    if sep != "":           # must be zero spaces
                        offenders.append(unit_tok)
                else:
                    if sep != " ":          # must be exactly one space
                        offenders.append(unit_tok)
            else:
                # ----- standard rules ------------------------------
                needs_one = base in allowed_space_units
                if needs_one:
                    if sep != " ":
                        offenders.append(unit_tok)
                else:                       # space forbidden
                    if sep != "":
                        offenders.append(unit_tok)

            i = m.end()  # advance past the token
            continue

        # default advance
        i += 1

    return offenders


def detect_separator_spacing_violations(raw: str) -> list[str]:
    """
    Returns any of {"to", ",", "@"} that break the strict rule:
      • "to" / "@" : exactly ONE space before **and** after
      • ","        : NO space before, ONE space after
    """
    s = str(raw)
    bad = set()

    # --- helper lambdas ---
    one_space_before = lambda idx: idx >= 1 and s[idx-1] == " " and (idx < 2 or s[idx-2] != " ")
    one_space_after  = lambda idx: idx+1 < len(s) and s[idx+1] == " " and (idx+2 >= len(s) or s[idx+2] != " ")

    # ---------- "to" ----------
    for m in re.finditer(r'\bto\b', s):
        i = m.start()          # index of 't'
        j = m.end()-1          # index of 'o'
        if not (one_space_before(i) and one_space_after(j)):
            bad.add("to")

    # ---------- "@" -----------
    for m in re.finditer(r'@', s):
        i = m.start()
        if not (one_space_before(i) and one_space_after(i)):
            bad.add("@")

    # ---------- "," -----------
    for m in re.finditer(r',', s):
        i = m.start()
        before_ok = i == 0 or s[i-1] != " "
        # count spaces after
        k = i + 1
        spaces = 0
        while k < len(s) and s[k] == " ":
            spaces += 1
            k += 1
        after_ok = spaces == 1
        if not (before_ok and after_ok):
            bad.add(",")

    return sorted(bad)


# —————————————————————————————————————————————————————
# 1) Load mapping and split into multipliers, units, identifiers
# —————————————————————————————————————————————————————
def load_mapping(mapping_file: str):
    df = pd.read_excel(mapping_file)
    bases = df["Base Unit Symbol"].dropna().astype(str)

    identifiers = sorted(
        [s for s in bases if s.startswith("(") and s.endswith(")")],
        key=len, reverse=True
    )
    units = sorted(
        [s for s in bases if not (s.startswith("(") and s.endswith(")"))],
        key=len, reverse=True
    )
    multipliers = sorted(
        df["Multiplier Symbol"].dropna().astype(str).unique(),
        key=len, reverse=True
    )

    return multipliers, units, identifiers

# —————————————————————————————————————————————————————
# 2) Build UNIT and IDENT regex fragments
# —————————————————————————————————————————————————————
def build_regexes(mapping_file: str):
    global multipliers_list, units_list
    multipliers, units, identifiers = load_mapping(mapping_file)

    # ————————————————————————————————————————————————
    #  Add “parenthesis‐stripped” variants of each unit so that
    #  after we strip (…) in detect(), we still match.
    # ————————————————————————————————————————————————
    import re as _re  # avoid shadowing
    stripped = []
    for u in units:
        s = _re.sub(r"\([^)]*\)", "", u).strip()
        if s and s != u:
            stripped.append(s)
    # merge & dedupe, sorted longest‐first (to avoid partial‐match hijacks)
    all_mapped_units = sorted(set(units) | set(stripped), key=len, reverse=True)
    all_symbols_as_units = sorted(set(all_mapped_units) | set(STRING_KEYWORDS), key=len, reverse=True)


    units = all_symbols_as_units


    multipliers_list, units_list = multipliers, units

    mult_re  = f"(?:{'|'.join(re.escape(m) for m in multipliers)})?"
    unit_re  = f"(?:{'|'.join(re.escape(u) for u in units)})"
    power_re = "(?:²|³)?"

    UNIT     = f"{mult_re}{unit_re}{power_re}"

    # ── identifier tokens pulled from mapping ───────────────────────────
    id_inner = [re.escape(s[1:-1]) for s in identifiers]       # strip the ( )
    id_list  = "|".join(id_inner)                              # join with "|"

    # ── numeric-identifier fragment — e.g.  ±6, 2, 3.3pF ────────────────
    numeric_ident = rf"[±+\-]?{NUM_REGEX}(?:\s*{UNIT})?"

    # ── final IDENT: either a known token OR a numeric identifier ───────
    if id_list:                                                # we have tokens
        IDENT = rf"(?:{id_list}|{numeric_ident})"
    else:                                                      # none in mapping
        IDENT = numeric_ident

    return UNIT, IDENT

# —————————————————————————————————————————————————————
# 3) Compile all templates once
# —————————————————————————————————————————————————————
def initialize_regex_classifier(mapping_file_path: str):
    """Loads mapping and builds regex patterns once."""
    global classification_patterns

    if classification_patterns:
        return  # already done

    st.write(f"DEBUG: initializing regex classifier from {mapping_file_path}")
    try:
        UNIT, IDENT = build_regexes(mapping_file_path)

        TEMPLATES = [
            (rf"^{NUM_REGEX}\s*\(\s*{IDENT}(?:\s*,\s*{IDENT})*\s*\)$", "Number with Identifier"),
            # (rf"^{NUM_REGEX}\s*\(\s*[±+\-]?{NUM_REGEX}(?:\s*{UNIT})?\s*\)$","Number with Identifier"),
            (rf"^{NUM_REGEX}$",                                    "Number"),
            (rf"^{NUM_REGEX}\s*to\s*{NUM_REGEX}\s*{UNIT}$", "Range Number&value"),
            (rf"^{NUM_REGEX}\s*{UNIT}(?:\s*\([^)]*\))?$", "Single Value"),
            (rf"^{NUM_REGEX}\s*{UNIT}\s*@\s*{NUM_REGEX}\s*{UNIT}$","Single Value with Single Condition"),
            (rf"^{NUM_REGEX}\s*{UNIT}\s*@\s*{NUM_REGEX}\s*{UNIT}\s+to\s+{NUM_REGEX}\s*{UNIT}$",               "Single Value with Range Condition"),
            (rf"^{NUM_REGEX}\s*@\s*{NUM_REGEX}\s*{UNIT}$",         "Number with Single Condition"),
            (rf"^{NUM_REGEX}\s+to\s+{NUM_REGEX}$",                 "Range Numbers"),
            (rf"^{NUM_REGEX}\s*@\s*{NUM_REGEX}\s*{UNIT}(?:\s*,\s*{NUM_REGEX}\s*{UNIT})+$",                    "Number with Multiple Conditions"),
            (rf"^{NUM_REGEX}\s*{UNIT}\s+to\s+{NUM_REGEX}\s*{UNIT}$","Range Value"),
            (rf"^{NUM_REGEX}\s*{UNIT}\s+to\s+{NUM_REGEX}\s*{UNIT}\s*@\s*{NUM_REGEX}\s*{UNIT}$",               "Range Value with Single Condition"),
            (rf"^{NUM_REGEX}\s*{UNIT}\s+to\s+{NUM_REGEX}\s*{UNIT}\s*@\s*{NUM_REGEX}\s*{UNIT}\s+to\s+{NUM_REGEX}\s*{UNIT}$","Range Value with Range Condition"),
            (rf"^{NUM_REGEX}\s*{UNIT}\s*@\s*{NUM_REGEX}\s*{UNIT}(?:\s*,\s*{NUM_REGEX}\s*{UNIT})+$",           "Single Value with Multiple Conditions"),
            (rf"^(?:{NUM_REGEX}(?:\s*,\s*{NUM_REGEX})*\s*{UNIT})\s*@\s*{NUM_REGEX}\s*{UNIT}$",                "Multi Value with Single Condition"),
            (rf"^{NUM_REGEX}\s*{UNIT}\s+to\s+{NUM_REGEX}\s*{UNIT}\s*@\s*{NUM_REGEX}\s*{UNIT}(?:\s*,\s*{NUM_REGEX}\s*{UNIT})+$", "Range Value with Multiple Conditions"),
            (rf"^(?:{NUM_REGEX}\s*{UNIT})(?:\s*,\s*{NUM_REGEX}\s*{UNIT})+\s*@\s*{NUM_REGEX}\s*{UNIT}(?:\s*,\s*{NUM_REGEX}\s*{UNIT})+$", "Multi Value with Multiple Conditions"),
        ]

        for pattern, label in TEMPLATES:
            p = re.compile(pattern, re.IGNORECASE)
            classification_patterns.append((p, label))

            core = pattern.lstrip("^").rstrip("$")
            multiple = re.compile(rf"^(?:{core})(?:\s*,\s*{core})+$", re.IGNORECASE)
            classification_patterns.append((multiple, f"Multiple ({label})"))

        st.write(f"DEBUG: compiled {len(classification_patterns)} patterns.")
    except Exception as e:
        st.error(f"Initialization error: {e}")
        st.error(traceback.format_exc())
        raise

# —————————————————————————————————————————————————————
# 4) Core detection
# —————————————————————————————————————————————————————
def detect_value_type(text: str) -> str:
    if not classification_patterns:
        st.error("RegexClassifier: not initialized.")
        return "Error"

    raw = str(text).strip()
    if raw == "":
        return "Empty"

    # —— Strip known prefix ——
    prefix, stripped_raw, has_prefix = extract_prefix_from_string(raw)
    to_classify = stripped_raw if has_prefix else raw

    # —— Identifier patterns against the original raw ——
    for regex, label in classification_patterns:
        if "Identifier" in label and regex.fullmatch(raw):
            return re.split(r"\s*\[", label)[0]

    # CHANGE A ————————————————————————————————
    # First, try to match *the entire string* (with parentheses removed)
    # against every compiled pattern.  If it matches, we’re done: this
    # prevents the premature “comma-split” from breaking multi-condition
    # expressions like “1860pF @ 0 V, 1 MHz”.
    txt = re.sub(r"\([^)]*\)", "", to_classify).strip()
    if txt == "":
        return "Empty"

    for regex, label in classification_patterns:
        if regex.fullmatch(txt):
            # For “Multiple ( … )” labels we still need the uniformity
            # checks that the original code performed.
            if label.startswith("Multiple ("):
                core = label[len("Multiple ("):-1]
                anchored = next((r for r, l in classification_patterns if l == core), None)
                if not anchored:
                    return re.split(r"\s*\[", label)[0]

                pat = anchored.pattern.lstrip("^").rstrip("$")
                block_re = re.compile(rf"{pat}(?=(?:\s*,\s*{pat})|$)", re.IGNORECASE)

                blocks, i, n = [], 0, len(txt)
                while i < n:
                    m = block_re.match(txt, i)
                    if not m:
                        return "Mixed Types"
                    blocks.append(m.group(0))
                    i = m.end()
                    comma = re.match(r"\s*,\s*", txt[i:])
                    if comma:
                        i += comma.end()
                    elif i < n:
                        return "Mixed Types"

                def count_commas(s: str) -> int:
                    depth = 0; cnt = 0
                    for ch in s:
                        if ch == "(":
                            depth += 1
                        elif ch == ")" and depth > 0:
                            depth -= 1
                        elif ch == "," and depth == 0:
                            cnt += 1
                    return cnt

                if len({count_commas(b) for b in blocks}) != 1:
                    return "Inconsistent Chunk Count"

            return re.split(r"\s*\[", label)[0]
    # ————————————————————————————————————————————

    # —— Handle comma-separated lists at the top level (fallback) ——
    parts = re.split(r'\s*,\s*(?![^(]*\))', to_classify)
    if len(parts) > 1:
        cleaned_parts = []
        for p in parts:
            _, rest, _ = extract_prefix_from_string(p.strip())
            cleaned_parts.append(rest.strip() or p.strip())

        part_cats = [detect_value_type(p) for p in cleaned_parts]
        first = part_cats[0]
        if all(pc == first for pc in part_cats):
            return f"Multiple ({first})"
        else:
            return "Mixed Types"

    return "Unknown"

# —————————————————————————————————————————————————————
# 5) Helpers for DetailedValueType
# (unchanged …)
# —————————————————————————————————————————————————————
def get_blocks(raw: str) -> list[str]:
    """
    Return a list of the top-level comma-separated blocks in *raw*,
    after stripping any prefix from each chunk so the regex patterns
    recognise them.  Used by build_detailed() to compute the [m][c] xN
    counters.
    """
    # 1) remove outer-level parentheses for a clean split
    core = re.sub(r"\([^)]*\)", "", raw.strip()).strip()

    # 2) split on commas that are not inside parentheses
    parts = re.split(r'\s*,\s*(?![^(]*\))', core)

    # 3) strip a known prefix from EACH chunk
    cleaned_parts = []
    for p in parts:
        _, rest, _ = extract_prefix_from_string(p.strip())
        cleaned_parts.append(rest.strip() or p.strip())

    cleaned_core = ", ".join(cleaned_parts)

    # 4) try to match “Multiple (XXX)” patterns against the cleaned text
    for regex, label in classification_patterns:
        if regex.fullmatch(cleaned_core) and label.startswith("Multiple ("):
            core_label = label[len("Multiple ("):-1]
            anchored = next(r for r, l in classification_patterns if l == core_label)
            pat = anchored.pattern.lstrip("^").rstrip("$")
            block_re = re.compile(rf"{pat}(?=(?:\s*,\s*{pat})|$)", re.IGNORECASE)

            blocks, i, n = [], 0, len(cleaned_core)
            while i < n:
                m = block_re.match(cleaned_core, i)
                if not m:
                    break
                blocks.append(m.group(0))
                i = m.end()
                comma = re.match(r"\s*,\s*", cleaned_core[i:])
                if comma:
                    i += comma.end()
                else:
                    break
            return blocks

    # fallback: can’t match → treat whole string as one block
    return [core]

# (everything below – build_detailed, get_reason, _absolute_pattern and
# run_regex_classification – is **unchanged** from your version)
# --- END OF FILE regex_classifier.py ---


def build_detailed(raw: str, cls: str) -> str:
    blocks = get_blocks(raw)
    main_counts, cond_counts = [], []

    for block in blocks:
        if "@" in block:
            main_part, cond_part = block.split("@", 1)
        else:
            main_part, cond_part = block, ""

        # override for Range Value / Condition
        if "Range Value" in cls or "Range Numbers" in cls:
            main_counts.append(1)
        else:
            main_counts.append(2 if " to " in main_part else 1)

        if "Range Condition" in cls:
            cond_counts.append(1)
        else:
            if not cond_part:
                cond_counts.append(0)
            elif " to " in cond_part:
                cond_counts.append(2)
            else:
                parts = [p for p in cond_part.split(",") if p.strip()]
                cond_counts.append(len(parts))

    main_n = main_counts[0] if all(m == main_counts[0] for m in main_counts) else "Mixed"
    cond_n = cond_counts[0] if all(c == cond_counts[0] for c in cond_counts) else "Mixed"
    sub_n  = len(blocks)

    if cls == "Multi Value with Multiple Conditions":
        block = blocks[0]
        left, right = block.split("@", 1) if "@" in block else (block, "")
        mains = [s.strip() for s in re.split(r"\s*,\s*", left)  if s.strip()]
        conds = [s.strip() for s in re.split(r"\s*,\s*", right) if s.strip()]
        main_n, cond_n = len(mains), len(conds)

    return f"{cls} [{main_n}][{cond_n}] x{sub_n}"

# —————————————————————————————————————————————————————
# 6) Reason generation
# —————————————————————————————————————————————————————
def is_known_unit(tok: str) -> bool:
    if tok.endswith("²") or tok.endswith("³"):
        base = tok[:-1]
    else:
        base = tok
    for m in multipliers_list:
        if base.startswith(m) and base[len(m):] in units_list:
            return True
    return base in units_list

def get_reason(raw: str, cls: str) -> str:
    core = re.sub(r"\([^)]*\)", "", str(raw)).strip()
    if cls == "Unknown":
        bad = []
        for m in re.finditer(rf"{NUM_REGEX}\s*(?P<unit>[^\d\s,()@+-.]+)", core):
            u = m.group("unit")
            if not is_known_unit(u):
                bad.append(u)
        return "Missing/Unknown unit(s): " + ", ".join(sorted(set(bad))) if bad else "No matching pattern"
    if cls == "Mixed Types":
        return "Different chunk structures"
    if cls == "Inconsistent Chunk Count":
        return "Different # of sub-values per chunk"
    return ""
def _absolute_pattern(block: str) -> str:
    """
    Convert a value/condition chunk to its '$ Unit ...' absolute form.
    MODIFIED: If a subpart has no unit (e.g., just a number), use '$' only.
    """
    parts = re.split(r'\s*@\s*', block)
    abs_parts = []
  #  st.write(f"DEBUG: Processing block for abs pattern: '{block}'")
    for part in parts:
        subparts = re.split(r'\s+to\s+', part)
        abs_sub = []
        for sp in subparts:
    #        st.write(f"DEBUG:  Processing subpart: '{sp}'")
            # Call the helper function
            num, mult, base_u, *_ = extract_numeric_and_unit_analysis(
                sp, set(units_list), {}  # units_list is global after initialise
            )
        #    st.write(f"DEBUG:    Helper returned base_u: {base_u}")

            # --- MODIFICATION ---
            if not base_u:  # If no base unit was found by the helper (e.g., for "100", "80")
                # Append just "$" as requested
                abs_sub.append("$") 
            else: # If a base unit WAS found by the helper
                # Append the standard "$ BaseUnit"
                abs_sub.append(f"$ {base_u}".strip()) 
         #       st.write(f"DEBUG:    Appending '$ {base_u}'") 
            # --- END MODIFICATION ---
            
        abs_parts.append(" to ".join(abs_sub))
    return " @ ".join(abs_parts)


def add_mixed_types_detail_dvt(df: pd.DataFrame,
                               *,
                               value_col: str = "Value_Normalized",
                               detail_col: str = "MixedTypesDetail",
                               flag_col: str = "MixedTypesFlag",
                               output_col: str = "MixedTypesDetail_DVT",
                               inplace: bool = True) -> pd.DataFrame:

    if not inplace:
        df = df.copy()

    # ── define the helper *inside* so it can “see” the kwargs ────────────
    def _row_func(row):
        # 1) skip rows that aren’t mixed
        if not bool(row.get(flag_col)):
            return ""                                   # blank cell

        # 2) guard against missing / empty detail strings
        details = row.get(detail_col)
        if pd.isna(details) or details == "":
            return ""

        # 3) build the annotated string
        return build_mixed_types_detail(row[value_col], details)

    # apply & return
    df[output_col] = df.apply(_row_func, axis=1)
    return df
_DVT_EXTRACT_RE = re.compile(r"\[(?P<m>[^\]]+)]\[(?P<c>[^\]]+)]\s*x(?P<s>\d+)")

def _clean_tag(tag: str) -> str:
    """Strip the 'Multiple (' … ')' wrapper →  'Range Value'."""
    tag = tag.strip()
    if tag.lower().startswith("multiple (") and tag.endswith(")"):
        tag = tag[len("Multiple ("):-1].strip()
    return tag

def build_mixed_types_detail(value: str,
                             mixed_detail: str,
                             delim: str = ",") -> str:
    """
    Example output:
      "Multiple (Range Value) [1][0] x2, Single Value [1][0] x1"
    """
    raw_parts       = [p.strip() for p in str(value).split(delim) if p.strip()]
    declared_labels = [t.strip() for t in str(mixed_detail).split(delim) if t.strip()]

    # ── 1.  Classify every chunk ──────────────────────────────────────────
    per_chunk_info: List[tuple[str, str, str]] = []
    for chunk in raw_parts:
        base_cls = detect_value_type(chunk)
        dvt      = build_detailed(chunk, base_cls)       # e.g. "Range Value [1][0] x1"
        m, c, _  = _DVT_EXTRACT_RE.search(dvt).groups()
        per_chunk_info.append((base_cls, m, c))

    # ── 2.  Count how many chunks fall under each *clean* class name ─────
    counts = Counter(cls for cls, *_ in per_chunk_info)

    # ── 3.  Grab any m / c numbers (all chunks of same class share them) ─
    mc_map = {}
    for cls, m, c in per_chunk_info:
        mc_map.setdefault(cls, (m, c))

    # ── 4.  Assemble output in the order of the user-supplied labels ─────
    fragments = []
    for tag in declared_labels:
        key = _clean_tag(tag)                 # "Range Value" from "Multiple (Range Value)"
        if key not in counts:                 # defensive: skip unknown labels
            continue
        m, c = mc_map[key]
        cnt  = counts[key]
        fragments.append(f"{tag} [{m}][{c}] x{cnt}")

    return ", ".join(fragments)
# —————————————————————————————————————————————————————
# 7) Main entry for DataFrame processing
# —————————————————————————————————————————————————————
def run_regex_classification(df: pd.DataFrame, mapping_file_path: str) -> pd.DataFrame:
    """
    Applies regex-based classification to 'Value_Normalized', adding columns:
      - Classification_New
      - Reason_New
      - DetailedValueType_New
    """
    st.write("DEBUG: running regex classification…")
    if 'Value_Normalized' not in df.columns:
        st.error("RegexClassifier: missing 'Value_Normalized'.")
        df['Classification_New']          = "Error"
        df['Reason_New']                  = ""
        df['DetailedValueType_New']       = ""
        df['Distinct_Abs_Pattern_Count']  = None
        return df

    try:
        # initialise patterns (fills multipliers_list & units_list)
        initialize_regex_classifier(mapping_file_path)

        vals = df['Value_Normalized'].astype(str).fillna("")

        df['Classification_New']    = vals.map(detect_value_type)
        df['Reason_New']            = [
            get_reason(v, c) for v, c in zip(vals, df['Classification_New'])
        ]
        df['DetailedValueType_New'] = [
            build_detailed(v, c) for v, c in zip(vals, df['Classification_New'])
        ]
        # ─── NEW COLUMNS ──────────────────────────────────────────────────────────
        # 1) Boolean flag: overall class is “Mixed Types”
# ─── SAFE helper: recursively collect the most basic categories ─────────
        comma_split_re = re.compile(r'\s*,\s*(?![^(]*\))')
        
        def _top_level_blocks(txt: str) -> list[str]:
            """Return blocks separated by top-level commas (prefix removed)."""
            core = re.sub(r"\([^)]*\)", "", txt).strip()          # ignore (…) commas
            parts = comma_split_re.split(core)
            cleaned = []
            for p in parts:
                _, rest, _ = extract_prefix_from_string(p.strip())
                cleaned.append(rest.strip() or p.strip())
            return cleaned
        
        
        # ----------------------------------------------------------------------
        def _collect_base_types_in_value_order(txt: str,
                                               *,
                                               depth: int = 0,
                                               max_depth: int = 10) -> list[str]:
            """
            Walk the value left→right, gather base categories.  If a category
            appears more than once, represent it once as  "Multiple (<category>)".
            """
            counts: dict[str, int] = {}
            ordered: list[str] = []
        
            def _walk(s: str, d: int):
                if d > max_depth:
                    return
                cat = detect_value_type(s)
                if cat == "Mixed Types":
                    for blk in _top_level_blocks(s):           # preserves source order
                        _walk(blk, d + 1)
                else:
                    if cat.startswith("Multiple (") and cat.endswith(")"):
                        cat = cat[len("Multiple ("):-1]        # strip wrapper
                    if cat not in counts:                      # first time we see it
                        ordered.append(cat)
                    counts[cat] = counts.get(cat, 0) + 1
        
            _walk(txt, depth)
        
            # build final list in first-appearance order
            return [
                f"Multiple ({cat})" if counts[cat] > 1 else cat
                for cat in ordered
            ]
        # ----------------------------------------------------------------------
        
        # 1) Boolean flag
        df['MixedTypesFlag'] = df['Classification_New'] == "Mixed Types"
        
        # 2) Detail column (source order, duplicates collapsed)
        df['MixedTypesDetail'] = df.apply(
            lambda r: ", ".join(_collect_base_types_in_value_order(r['Value_Normalized']))
                      if r['MixedTypesFlag'] else None,
            axis=1
        )
# --
#-------------------------------------------------------------

                # --- PATCH: fix Number→Number with Identifier when parentheses indicate an ID ---
    
# --- PATCH: fix Number→Number with Identifier when parentheses indicate an ID (modified to pre-filter) ---
        identifier_pattern = re.compile(r'\(\s*[A-Za-z]\w*\s*\)')

        # Ensure 'Classification_New' is string type for .str accessor
        df['Classification_New'] = df['Classification_New'].astype(str)
        df['DetailedValueType_New'] = df['DetailedValueType_New'].astype(str) # Ensure DVT is also string

        # MODIFIED MASK:
        # Condition 1: Classification_New contains "(Number)" (e.g., "Multiple (Number)")
        # Condition 2: Classification_New is exactly "Number"
        classification_mask = (
            (
                df['Classification_New'].str.contains(r'\(Number\)', regex=True, na=False) | # Condition 1
                (df['Classification_New'] == "Number")                                      # Condition 2
            ) &
            ~df['Classification_New'].str.contains('Identifier', regex=False, na=False) # Ensure it's not already "Identifier"
        )

        for i in df.index[classification_mask]:
            raw_value = df.at[i, 'Value_Normalized']

            if identifier_pattern.search(str(raw_value)):
                current_cls = df.at[i, 'Classification_New']
                current_det = df.at[i, 'DetailedValueType_New']

                # MODIFIED REPLACEMENT LOGIC
                if current_cls == "Number":
                    df.at[i, 'Classification_New'] = "Number with Identifier"
                    # Assuming DVT should also be updated similarly if it was just "Number [...]"
                    if current_det.startswith("Number ["):
                         df.at[i, 'DetailedValueType_New'] = "Number with Identifier" + current_det[len("Number"):]
                    else: # Fallback if DVT format is unexpected
                         df.at[i, 'DetailedValueType_New'] = current_det.replace('Number', 'Number with Identifier', 1)

                elif '(Number)' in current_cls: # It contained "(Number)"
                    df.at[i, 'Classification_New'] = str(current_cls).replace('(Number)', '(Number with Identifier)')
                    df.at[i, 'DetailedValueType_New'] = str(current_det).replace('(Number)', '(Number with Identifier)')
        # -------------------------------------------------------------
        # -------------------------------------------------------------

        # ---------- NEW: distinct absolute-pattern counter ----------
        abs_counts = []
        for raw, cls in zip(vals, df['Classification_New']):
            if cls.startswith("Multiple (") and cls not in (
                "Multiple Mixed", "Multiple Repeated Pairs"
            ):
                blocks = get_blocks(raw)           # analysis_helpers helper
                patterns = {_absolute_pattern(b) for b in blocks}
                abs_counts.append(len(patterns))
            else:
                abs_counts.append(None)            # not applicable
        df['Distinct_Abs_Pattern_Count'] = abs_counts
        # -------------------------------------------------------------
        # -------------------------------------------------------------

        # ------------------------------------------------------------
        #  1) unit-spacing QA
        # ------------------------------------------------------------
        unit_flag  = []
        unit_which = []
        for txt in df['Value_Normalized']:
            u = detect_unit_spacing_violations_regex(
                txt, multipliers_list, UNITS_ALLOW_SPACE
            )
            unit_flag.append(bool(u))
            unit_which.append(", ".join(u))
        df["UnitSpacingViolationFlag"]  = unit_flag
        df["UnitSpacingViolationUnits"] = unit_which

        # ------------------------------------------------------------
        #  2) separator-spacing QA
        # ------------------------------------------------------------
        sep_flag  = []
        sep_which = []
        for txt in df['Value_Normalized']:
            bad_seps = detect_separator_spacing_violations(txt)
            sep_flag.append(bool(bad_seps))
            sep_which.append(", ".join(bad_seps))
        df["SepSpacingViolationFlag"]   = sep_flag
        df["SepSpacingViolationTokens"] = sep_which
        st.write("DEBUG: done.")
    except Exception as e:
        st.error(f"RegexClassifier: processing error: {e}")
        st.error(traceback.format_exc())
        df['Classification_New']          = "Error"
        df['Reason_New']                  = str(e)
        df['DetailedValueType_New']       = ""
        df['Distinct_Abs_Pattern_Count']  = None

    return df

def _collect_base_types_in_value_order(txt: str,
                                       *,
                                       depth: int = 0,
                                       max_depth: int = 10) -> list[str]:
    """
    Walk the value left→right, gather the basic categories that appear.
    If a category occurs >1 time, return it once as “Multiple (<category>)”.
    """
    comma_split_re = re.compile(r'\s*,\s*(?![^(]*\))')

    def _top_level_blocks(s: str) -> list[str]:
        core = re.sub(r"\([^)]*\)", "", s).strip()
        parts = comma_split_re.split(core)
        cleaned = []
        for p in parts:
            _, rest, _ = extract_prefix_from_string(p.strip())
            cleaned.append(rest.strip() or p.strip())
        return cleaned

    counts, ordered = {}, []

    def _walk(s: str, d: int):
        if d > max_depth:
            return
        cat = detect_value_type(s)
        if cat == "Mixed Types":
            for blk in _top_level_blocks(s):
                _walk(blk, d + 1)
        else:
            # unwrap “Multiple (…)” so we count the inner class
            if cat.startswith("Multiple (") and cat.endswith(")"):
                cat = cat[len("Multiple ("):-1]
            if cat not in counts:
                ordered.append(cat)
            counts[cat] = counts.get(cat, 0) + 1

    _walk(str(txt), depth)
    return [
        f"Multiple ({c})" if counts[c] > 1 else c
        for c in ordered
    ]
# --- END OF FILE regex_classifier.py ---
