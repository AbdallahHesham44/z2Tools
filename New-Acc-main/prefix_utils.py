# File: prefix_utils.py
import re
import pandas as pd # Only for type hinting if necessary, avoid heavy logic here

# Define a list of known prefixes (case-insensitive)
# This list needs to be comprehensive and might come from configuration or mapping in a real scenario.
KNOWN_PREFIXES = [
    "parallel", "series", "Class" # Example multi-word prefixes
]
STRING_KEYWORDS = [
    "adj"  # Add more keywords here
]
# Add more prefixes here as needed, e.g.
# KNOWN_PREFIXES.extend(["another_prefix", "yet_another"])
# put this near the bottom of prefix_utils.py

# prefix_utils.py  ───── add near the bottom ─────────────────────────
def strip_prefix_once(txt: str) -> str:
    """
    Returns the string with ONE leading textual prefix removed.
    Re-uses extract_prefix_from_string so we stay in sync with KNOWN_PREFIXES.
    """
    _, rest, found = extract_prefix_from_string(txt)
    return rest if found else txt



# Compile a regex for these prefixes (word boundary, case-insensitive)
# Sort by length descending to match longer prefixes first (e.g., "common mode" before "common")
_sorted_known_prefixes = sorted(KNOWN_PREFIXES, key=len, reverse=True)
PREFIX_REGEX_PARTS = [re.escape(p) for p in _sorted_known_prefixes]

# MODIFICATION START
if PREFIX_REGEX_PARTS: # Only compile if there are prefixes
    # Match prefix at the beginning of the string, followed by ZERO OR MORE spaces, then capture the rest.
    # The prefix itself is captured in group 1, rest in group 2.
    # (?i) is now placed at the very start of the string passed to re.compile
    # Changed \s+ to \s* to allow for no space between prefix and value.
    pattern_string = r"(?i)^(" + "|".join(PREFIX_REGEX_PARTS) + r")\s*(.+)$" # MODIFIED HERE: \s+ to \s*
    PREFIX_PATTERN = re.compile(pattern_string)
else:
    PREFIX_PATTERN = None
# MODIFICATION END

def extract_prefix_from_string(text_chunk: str) -> tuple[str | None, str, bool]:
    """
    Checks if a string chunk starts with a known prefix.
    Args:
        text_chunk (str): The string to check.
    Returns:
        tuple[str | None, str, bool]: (identified_prefix, remaining_string_after_prefix, was_prefixed_flag)
                                      If no prefix, returns (None, original_string_chunk, False)
    """
    if not isinstance(text_chunk, str):
        return None, str(text_chunk), False

    if PREFIX_PATTERN is None: # Handle case where no prefixes were defined
        return None, text_chunk.strip(), False
        
    stripped_chunk = text_chunk.strip() # Process the stripped version for matching
    match = PREFIX_PATTERN.match(stripped_chunk)
    if match:
        prefix = match.group(1) 
        # The remainder should be stripped to remove any leading/trailing spaces
        # that were part of the original chunk but not part of the actual value.
        remaining = match.group(2).strip() 
        return prefix, remaining, True
    return None, stripped_chunk, False # Return the stripped_chunk if no prefix

def add_prefix_to_normalized_string(prefix: str | None, normalized_string: str) -> str:
    """
    Adds the prefix back to a normalized string if a prefix exists and the string is not empty.
    Ensures no double-prefixing if the string somehow already starts with it (case-insensitive).
    """
    if not isinstance(normalized_string, str):
        normalized_string = str(normalized_string)

    if prefix and normalized_string:
        if not normalized_string.lower().startswith(prefix.lower() + " "):
            return f"{prefix} {normalized_string}"
    return normalized_string
