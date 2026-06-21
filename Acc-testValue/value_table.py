#!/usr/bin/env python3
# value_table.py

import re
import pandas as pd

# ---------------------------------------------------------------------------
# Multiplier mapping for normalization
# ---------------------------------------------------------------------------
MULTIPLIER_MAPPING = {
    'k': 1e3,   'M': 1e6,   'G': 1e9,   'T': 1e12,
    'P': 1e15,  'E': 1e18,  'c': 1e-2,  'm': 1e-3,
    'µ': 1e-6,  'n': 1e-9,  'p': 1e-12, 'f': 1e-15,
    'a': 1e-18, 'z': 1e-21, 'y': 1e-24
}
import re
from decimal import Decimal, InvalidOperation

_SCI_RE = re.compile(
    r'^([+\-±]?\d+(?:\.\d+)?(?:[eE][+\-]?\d+)?)(?:\s*)([a-zA-Zµ]+)?$'
)

def _plain_decimal(num_txt: str) -> str:
    """Expand possible scientific notation to a plain-decimal string; else unchanged."""
    try:
        d = Decimal(num_txt)
        s = format(d.normalize(), 'f')  # drop exponent
        if s.startswith('.'):   s = '0' + s
        if s.startswith('-.'):  s = s.replace('-.', '-0.', 1)
        return s
    except InvalidOperation:
        return num_txt

def parse_value_mul_str(val_mul_str: str) -> dict:
    raw = str(val_mul_str).strip()

    # --- 1) detect & strip comma groupings for numeric parsing ---
    has_thousand = bool(re.search(r"\d,\d", raw))   # e.g. "1,234"
    clean = raw.replace(",", "")                    # numeric parsing uses this

    # --- 2) split numeric part and metric prefix (k, M, µ…) ---
    # CHANGED: accept scientific notation too
    m = _SCI_RE.match(clean)
    if m:
        numeric_str, prefix = m.group(1), m.group(2) or ""
    else:                                           # fallback – no match
        numeric_str, prefix = "", ""

    # --- 3) flag for literal space between number and prefix ---
    has_space = bool(re.search(r'\d\s+[a-zA-Zµ]', raw))

    # --- 4) build value-pattern, reinserting commas where they appeared ---
    def build_pattern(num_str: str) -> str:
        unsigned = num_str.lstrip('±+-')

        # NEW: if sci-notation present, expand only for mask construction
        if 'e' in unsigned.lower():
            unsigned = _plain_decimal(unsigned)

        if '.' in unsigned:
            int_part, frac_part = unsigned.split('.', 1)
        else:
            int_part, frac_part = unsigned, ''

        # recreate comma groups
        if has_thousand and len(int_part) > 3:
            groups = []
            rem = len(int_part) % 3
            if rem:
                groups.append('x' * rem)
            for i in range(rem, len(int_part), 3):
                groups.append('x' * 3)
            int_core = ','.join(groups)
        else:
            int_core = 'x' * len(int_part)

        frac_core = '.' + 'y' * len(frac_part) if frac_part else ''
        core = int_core + frac_core
        return f"+{core}, -{core}" if num_str.startswith('±') else core

    value_pattern = build_pattern(numeric_str)

    # --- 5) normalized numeric value (applies metric prefix) ---
    # UNCHANGED: still uses float + .12g formatting
    try:
        base_val = float(numeric_str.replace('±', ''))
        factor = MULTIPLIER_MAPPING.get(prefix, 1)
        normalized_str = f"{base_val * factor:.12g}"
    except ValueError:
        normalized_str = ""

    # --- 6) split integer / fraction columns ---
    plain = numeric_str.lstrip('±+-')
    # NEW: expand sci-notation so parts are populated
    if 'e' in plain.lower():
        plain = _plain_decimal(plain)

    if '.' in plain:
        int_part, frac_part = plain.split('.', 1)
        fraction_digits = str(len(frac_part))
        fraction_part = '.' + frac_part
    else:
        int_part, fraction_part, fraction_digits = plain, '', ''

    return {
        'numeric_str':      numeric_str,       # e.g. "400000" or "1e-3"
        'prefix':           prefix,            # e.g. "k"
        'has_space':        has_space,         # 'Value Tk' → True
        'value_pattern':    value_pattern,     # e.g. "xxx,xxx"
        'normalized_str':   normalized_str,    # same formatting as before
        'integer_part':     int_part,
        'fraction_part':    fraction_part,
        'fraction_digits':  fraction_digits
    }

    
def compute_multiplication(v):
    p = parse_value_mul_str(v)['prefix']
    # if no prefix, return "1"
    return p if p else '1'

# Exposed helpers for DataFrame.apply()
def compute_value__(v):           return parse_value_mul_str(v)['numeric_str']
#def compute_multiplication(v):    return parse_value_mul_str(v)['prefix']
def compute_space_separator(v):   return 'y' if parse_value_mul_str(v)['has_space'] else 'n'
def compute_value_pattern(v):     return parse_value_mul_str(v)['value_pattern']
def compute_normalized_value__(v):return parse_value_mul_str(v)['normalized_str']
def compute_integar_value(v):     return parse_value_mul_str(v)['integer_part']
def compute_fraction_value(v):    return parse_value_mul_str(v)['fraction_part']
def compute_fraction_digits(v):   return parse_value_mul_str(v)['fraction_digits']
