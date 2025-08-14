import pandas as pd
from helper import load_file, similarity_ratio, normalize_series

def compare_requested_series(master_path, comparison_path, rules_path=None, top_n=2):
    df_master = load_file(master_path)
    df_comparison = load_file(comparison_path)
    if rules_path:
        df_rules = load_file(rules_path)
    else:
        df_rules = None

    results = []
    for idx, row in df_comparison.iterrows():
        comparison_series = normalize_series(row["Series"])
        matches = []
        for _, master_row in df_master.iterrows():
            master_series = normalize_series(master_row["Series"])
            ratio = similarity_ratio(comparison_series, master_series)
            matches.append((master_series, ratio))
        matches.sort(key=lambda x: x[1], reverse=True)
        top_matches = matches[:top_n]
        results.append({"Requested": comparison_series, "Matches": top_matches})

    return pd.DataFrame(results)

def update_master_series(update_path, master_path):
    df_update = load_file(update_path)
    df_master = load_file(master_path)
    df_master = pd.concat([df_master, df_update]).drop_duplicates().reset_index(drop=True)
    df_master.to_excel(master_path, index=False)

def delete_from_master_series(delete_path, master_path):
    df_delete = load_file(delete_path)
    df_master = load_file(master_path)
    df_master = df_master[~df_master["Series"].isin(df_delete["Series"])]
    df_master.to_excel(master_path, index=False)

def update_series_rules(update_path, rules_path):
    df_update = load_file(update_path)
    df_rules = load_file(rules_path)
    df_rules = pd.concat([df_rules, df_update]).drop_duplicates().reset_index(drop=True)
    df_rules.to_excel(rules_path, index=False)

def delete_from_series_rules(delete_path, rules_path):
    df_delete = load_file(delete_path)
    df_rules = load_file(rules_path)
    df_rules = df_rules[~df_rules["Series"].isin(df_delete["Series"])]
    df_rules.to_excel(rules_path, index=False)
