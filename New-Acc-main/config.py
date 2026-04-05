# config.py

# -----------------------------
# GitHub Configuration
# -----------------------------
# You can either set these via Streamlit's secrets or override them here.
GITHUB_OWNER = None          # e.g., "your-github-username"
GITHUB_REPO = None           # e.g., "your-repo-name"
GITHUB_TOKEN = None          # e.g., "your-personal-access-token"
GITHUB_FILE_PATH = "mapping.xlsx"  # Path to the mapping file in your repo

# -----------------------------
# Pipeline File Names
# -----------------------------
MAPPING_FILE_LOCAL = "mapping.xlsx"
PROCESSED_OUTPUT_FILE = "processed_output.xlsx"
USER_INPUT_FILE = "user_input.xlsx"
CLASSIFIED_OUTPUT_FILE = "classified_output.xlsx"
QA_NEW_FILE = "QA_new.xlsx"
FINAL_COMBINED_FILE = "final_combined.xlsx"

# -----------------------------
# Unit Mapping & Base Units
# -----------------------------
# Multiplier mapping for unit prefixes.
MULTIPLIER_MAPPING = {
    'y': 1e-24, 'z': 1e-21, 'a': 1e-18, 'f': 1e-15,
    'p': 1e-12, 'n': 1e-9,  'µ': 1e-6,  'm': 1e-3,
    'c': 1e-2,  'd': 1e-1,  'da': 1e1,  'h': 1e2,
    'k': 1e3,   'M': 1e6,   'G': 1e9,   'T': 1e12,
    'P': 1e15,  'E': 1e18,  'Z': 1e21,  'Y': 1e24
}

# Initial base units set; this will be updated from the mapping file.
LOCAL_BASE_UNITS = set()
