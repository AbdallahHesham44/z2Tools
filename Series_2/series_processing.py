import pandas as pd
import re

# =========================
# Example helper functions
# =========================
def similarity_ratio(a, b):
    # Simple placeholder (replace with your real logic)
    a, b = str(a).lower(), str(b).lower()
    matches = sum(c1 == c2 for c1, c2 in zip(a, b))
    return matches / max(len(a), len(b))

def normalize_series(series_str):
    # Example: remove spaces, standardize case
    return str(series_str).strip().upper()

def check_series_match(requested_series, candidate_series, rules_df):
    # Example matching logic based on rules
    requested_series = normalize_series(requested_series)
    candidate_series = normalize_series(candidate_series)

    for _, rule in rules_df.iterrows():
        pattern = str(rule.get("Pattern", ""))
        replacement = str(rule.get("Replacement", ""))
        candidate_series = re.sub(pattern, replacement, candidate_series)

    return requested_series == candidate_series

# =========================
# Main comparison function
# =========================
def compare_requested_series(master_file, input_file, rules_file, top_n=1):
    df_master = pd.read_excel(master_file)
    df_input = pd.read_excel(input_file)
    df_rules = pd.read_excel(rules_file)

    results = []
    for _, row in df_input.iterrows():
        requested = row.get("RequestedSeries", "")
        best_match = None
        best_score = -1

        for _, master_row in df_master.iterrows():
            score = similarity_ratio(requested, master_row.get("Series", ""))
            if score > best_score:
                best_score = score
                best_match = master_row

        if best_match is not None:
            matched_data = best_match.to_dict()
            matched_data["MatchScore"] = best_score
            matched_data["RequestedSeries"] = requested
            results.append(matched_data)

    return pd.DataFrame(results)
