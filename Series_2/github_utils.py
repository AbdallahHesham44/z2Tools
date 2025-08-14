import pandas as pd
import requests
from io import BytesIO
from github import Github
import base64

# Load Excel file from GitHub repo
def load_excel_from_github(repo_name, file_path, token):
    g = Github(token)
    repo = g.get_repo(repo_name)
    contents = repo.get_contents(file_path)
    file_data = base64.b64decode(contents.content)
    return pd.read_excel(BytesIO(file_data))

# Save Excel file to GitHub repo
def save_excel_to_github(repo_name, file_path, df, commit_message, token):
    g = Github(token)
    repo = g.get_repo(repo_name)

    # Save DataFrame to bytes
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    content_bytes = output.read()

    # Get file if exists
    try:
        contents = repo.get_contents(file_path)
        repo.update_file(file_path, commit_message, content_bytes, contents.sha)
    except:
        repo.create_file(file_path, commit_message, content_bytes)

# Append or delete rows in MasterSeriesHistory.xlsx or SeriesRules.xlsx
def modify_master_or_rules(df_existing, df_changes, action):
    key_cols = ["VariantID", "ManufacturerName", "Category", "Family"]

    if action == "append":
        for _, row in df_changes.iterrows():
            key = tuple(row[col] for col in key_cols)
            mask = (df_existing[key_cols] == pd.Series(key, index=key_cols)).all(axis=1)

            if mask.any():
                df_existing.loc[mask, "RequestedSeries"] = row["RequestedSeries"]
            else:
                df_existing = pd.concat([df_existing, pd.DataFrame([row])], ignore_index=True)

    elif action == "delete":
        for _, row in df_changes.iterrows():
            key = tuple(row[col] for col in key_cols)
            mask = (df_existing[key_cols] == pd.Series(key, index=key_cols)).all(axis=1)
            df_existing = df_existing[~mask]

    return df_existing
