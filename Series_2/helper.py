import pandas as pd
from difflib import SequenceMatcher

def load_file(file_path):
    """Load CSV or Excel file into DataFrame."""
    if file_path.endswith(".csv"):
        return pd.read_csv(file_path)
    return pd.read_excel(file_path)

def similarity_ratio(a, b):
    """Return similarity ratio as percentage with 2 decimal places."""
    return round(SequenceMatcher(None, a, b).ratio() * 100, 2)

def normalize_series(series):
    """Normalize series name for comparison."""
    if pd.isna(series):
        return ""
    return str(series).strip().upper()
