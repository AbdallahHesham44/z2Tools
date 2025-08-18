import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
from collections import defaultdict
import re
import requests
import io
import base64
from datetime import datetime

# Configuration
MASTER_URL = "https://raw.githubusercontent.com/AbdallahHesham44/Series_2/refs/heads/main/SampleMasterSeriesHistory.xlsx"
RULES_URL = "https://raw.githubusercontent.com/AbdallahHesham44/Series_2/refs/heads/main/SampleSeriesRules.xlsx"

# GitHub API Configuration
GITHUB_API_BASE = "https://api.github.com"
REPO_OWNER = "AbdallahHesham44"
REPO_NAME = "Series_2"
BRANCH = "main"
MASTER_FILE_PATH = "SampleMasterSeriesHistory.xlsx"
RULES_FILE_PATH = "SampleSeriesRules.xlsx"

def update_github_file(dataframe, file_path, commit_message, github_token):
    """
    Update a file in GitHub repository with new dataframe content.
    
    Args:
        dataframe: pandas DataFrame to upload
        file_path: path to file in repository (e.g., "SampleMasterSeriesHistory.xlsx")
        commit_message: commit message for the update
        github_token: GitHub personal access token
    
    Returns:
        tuple: (success: bool, message: str)
    """
    try:
        # Convert dataframe to Excel bytes
        output = io.BytesIO()
        dataframe.to_excel(output, index=False, engine='openpyxl')
        file_content = output.getvalue()
        
        # Encode content to base64
        content_base64 = base64.b64encode(file_content).decode('utf-8')
        
        # GitHub API headers
        headers = {
            'Authorization': f'token {github_token}',
            'Accept': 'application/vnd.github.v3+json',
            'Content-Type': 'application/json'
        }
        
        # Get current file SHA (required for updates)
        get_url = f"{GITHUB_API_BASE}/repos/{REPO_OWNER}/{REPO_NAME}/contents/{file_path}"
        get_response = requests.get(get_url, headers=headers)
        
        if get_response.status_code == 200:
            current_sha = get_response.json()['sha']
        else:
            return False, f"Failed to get current file SHA: {get_response.status_code}"
        
        # Prepare update payload
        update_data = {
            'message': commit_message,
            'content': content_base64,
            'sha': current_sha,
            'branch': BRANCH
        }
        
        # Update file via GitHub API
        put_response = requests.put(get_url, json=update_data, headers=headers)
        
        if put_response.status_code == 200:
            return True, "File updated successfully on GitHub!"
        else:
            return False, f"Failed to update file: {put_response.status_code} - {put_response.text}"
            
    except Exception as e:
        return False, f"Error updating GitHub file: {str(e)}"

def load_file_from_github(url):
    """Load Excel file from GitHub URL."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        return pd.read_excel(io.BytesIO(response.content), engine='openpyxl')
    except Exception as e:
        st.error(f"Error loading file from {url}: {e}")
        return None

def similarity_ratio(a, b):
    """Return similarity ratio as percentage with 2 decimal places."""
    return round(SequenceMatcher(None, a, b).ratio() * 100, 2)

def normalize_series(series_str):
    """Remove common separators and convert to lowercase for comparison."""
    normalized = re.sub(r'[,\-\s/]', '', str(series_str))
    return normalized.lower()

def check_series_match(requested_series, existing_series):
    """
    Check different types of matches between requested and existing series.
    Returns: (match_type, comment, similarity_score)
    """
    req_str = str(requested_series)
    exist_str = str(existing_series)

    # Skip matching against empty values (but allow dash matching)
    if exist_str in ['', 'nan', 'None']:
        return "no_match", "NotFound", 0.0

    # Special case: if requested is "-", check if existing series contains "-"
    if req_str == '-':
        if '-' in exist_str:
            return "contain", "FoundWithDiff(contain)", similarity_ratio(req_str, exist_str)
        else:
            return "no_match", "NotFound", 0.0

    # Skip matching against dash when requested is NOT dash
    if exist_str == '-':
        return "no_match", "NotFound", 0.0

    # 1. Exact match (case sensitive)
    if req_str == exist_str:
        return "exact", "insertedBeforeExact", 100.0

    # 2. Case sensitive exact match (same content, different case)
    if req_str.lower() == exist_str.lower():
        return "case_sensitive", "Exact(caseSensitive)", 100.0

    # 3. Normalize both strings (remove separators and make lowercase)
    req_normalized = normalize_series(req_str)
    exist_normalized = normalize_series(exist_str)

    # 4. Check if normalized versions are the same (NanAlphaCase)
    if req_normalized == exist_normalized and req_str != exist_str:
        return "nan_alpha", "NanAlphaCase", 100.0

    # 5. Check containment (bidirectional - either can contain the other)
    if exist_normalized in req_normalized and exist_normalized != req_normalized:
        return "contain", "FoundWithDiff(contain)", similarity_ratio(req_str, exist_str)
    elif req_normalized in exist_normalized and req_normalized != exist_normalized:
        return "contain", "FoundWithDiff(contain)", similarity_ratio(req_str, exist_str)

    # 6. Regular similarity check
    sim_score = similarity_ratio(req_str, exist_str)
    if sim_score >= 60:
        return "similar", f"similar_{sim_score}%", sim_score

    return "no_match", "NotFound", sim_score

def find_major_contain_series(input_series, lookup_data, key):
    """
    Find the most frequently used series that contains the input series.
    Returns the series name with highest count among containing series.
    """
    if key not in lookup_data:
        return None

    # Get all series for this key and their counts
    series_counts = {}
    for series, major_id in lookup_data[key]:
        # Convert MajorID percentage to count approximation
        percentage = float(major_id.replace('%', ''))
        if series in series_counts:
            series_counts[series] = max(series_counts[series], percentage)
        else:
            series_counts[series] = percentage

    # Find series that contain the input series
    containing_series = []
    input_str = str(input_series).strip()

    for series, count_percentage in series_counts.items():
        # Skip invalid series
        if series in ['', 'nan', 'None']:
            continue

        series_str = str(series).strip()

        # Special handling for dash input
        if input_str == '-':
            if '-' in series_str:
                containing_series.append((series_str, count_percentage))
        else:
            # Skip dash series when input is not dash
            if series_str == '-':
                continue

            # Check containment (case insensitive)
            if input_str.lower() in series_str.lower():
                containing_series.append((series_str, count_percentage))

    # If no containing series found, return None
    if not containing_series:
        return None

    # Sort by count percentage (highest first)
    containing_series.sort(key=lambda x: x[1], reverse=True)

    # Return the most frequent containing series
    best_series = containing_series[0][0]
    return f"Major_contain({best_series})"

def calculate_major_id(df):
    """Calculate MajorID percentages for series usage."""
    required_cols = ['VariantID', 'ManufacturerName', 'Category', 'Family', 'DataSheetURL', 'RequestedSeries']
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"Missing columns: {missing_cols}")
        return None

    # Step 1: Create key
    df['key'] = df[['ManufacturerName', 'Category', 'Family']].astype(str).agg('|'.join, axis=1)

    # Step 2: Count how often each RequestedSeries appears per key
    group_counts = df.groupby(['key', 'RequestedSeries']).size().reset_index(name='count')

    # Step 3: Count total rows per key
    total_counts = df.groupby('key').size().reset_index(name='total')

    # Step 4: Merge and calculate percentage
    merged = pd.merge(df, group_counts, on=['key', 'RequestedSeries'], how='left')
    merged = pd.merge(merged, total_counts, on='key', how='left')
    merged['MajorID'] = ((merged['count'] / merged['total']) * 100).round(2).astype(str) + '%'

    # Drop helper columns
    merged.drop(columns=['count', 'total'], inplace=True)

    return merged

def apply_rules_with_special_case(rules_df, master_df):
    """Apply business rules to the master dataframe."""
    # Validate required columns in rules file
    for col in ["ManufacturerName", "Category", "Family", "Rule"]:
        if col not in rules_df.columns:
            st.error(f"Rules file missing column: {col}")
            return master_df

    # Special case for ManufacturerName == "88xx"
    special_case_rules = rules_df[rules_df["ManufacturerName"] == "88xx"]
    normal_rules = rules_df[rules_df["ManufacturerName"] != "88xx"]

    # 1️⃣ Normal matching: create key and merge
    normal_rules["key"] = normal_rules["ManufacturerName"] + "|" + normal_rules["Category"] + "|" + normal_rules["Family"]
    master_df["key"] = master_df["ManufacturerName"] + "|" + master_df["Category"] + "|" + master_df["Family"]

    master_df = master_df.merge(normal_rules[["key", "Rule"]], on="key", how="left", suffixes=("", "_rule"))

    # 2️⃣ Special case: match only on Category + Family
    for _, row in special_case_rules.iterrows():
        mask = (
            (master_df["ManufacturerName"] == "88xx") &
            (master_df["Category"] == row["Category"]) &
            (master_df["Family"] == row["Family"])
        )
        master_df.loc[mask, "Rule"] = row["Rule"]

    # Drop helper key column
    if "key" in master_df.columns:
        master_df = master_df.drop(columns=["key"])

    return master_df

def compare_series_logic(master_df, comparison_df, rules_df=None, top_n=1):
    """Main series comparison logic."""
    
    # Calculate MajorID for master data
    df1_with_majorid = calculate_major_id(master_df.copy())
    if df1_with_majorid is None:
        return None

    # Validate required columns
    required_cols_1 = ['ManufacturerName', 'Category', 'Family', 'RequestedSeries', 'MajorID']
    required_cols_2 = ['ManufacturerName', 'Category', 'Family', 'RequestedSeries']

    if not all(col in df1_with_majorid.columns for col in required_cols_1):
        st.error(f"Master file missing columns: {set(required_cols_1) - set(df1_with_majorid.columns)}")
        return None
    if not all(col in comparison_df.columns for col in required_cols_2):
        st.error(f"Comparison file missing columns: {set(required_cols_2) - set(comparison_df.columns)}")
        return None

    # Create matching keys
    df1_with_majorid['key'] = df1_with_majorid[['ManufacturerName', 'Category', 'Family']].astype(str).agg('|'.join, axis=1)
    comparison_df['key'] = comparison_df[['ManufacturerName', 'Category', 'Family']].astype(str).agg('|'.join, axis=1)

    # Lookup dict: key -> [(RequestedSeries, MajorID)]
    lookup = defaultdict(list)
    for _, row in df1_with_majorid.iterrows():
        lookup[row['key']].append((str(row['RequestedSeries']), row['MajorID']))

    # Store results for all rows
    result_rows = []

    for _, row in comparison_df.iterrows():
        key = row['key']
        series = str(row['RequestedSeries'])

        # Most used series for the key
        if key in lookup:
            temp_df = pd.DataFrame(lookup[key], columns=['Series', 'MajorID'])
            # Filter out invalid values, but keep dashes if requested series is dash
            if series == '-':
                temp_df = temp_df[~temp_df['Series'].isin(['', 'nan', 'None'])]
            else:
                temp_df = temp_df[~temp_df['Series'].isin(['-', '', 'nan', 'None'])]

            if not temp_df.empty:
                # Convert MajorID percentage to numeric for sorting
                temp_df['MajorID_numeric'] = temp_df['MajorID'].str.replace('%', '').astype(float)
                temp_df = temp_df.groupby('Series', as_index=False)['MajorID_numeric'].max()
                sorted_list = temp_df.sort_values('MajorID_numeric', ascending=False)
                top_values = [f"{row.Series}({row.MajorID_numeric}%)" for _, row in sorted_list.head(top_n).iterrows()]
                most_used_series = " | ".join(top_values)
            else:
                most_used_series = None
        else:
            most_used_series = None

        # Calculate major_Sim
        major_sim = find_major_contain_series(series, lookup, key)

        if key not in lookup:
            # No matches found
            new_row = row.copy()
            new_row['comments'] = 'NotFound'
            new_row['MajorID'] = None
            new_row['FoundSeries'] = None
            new_row['MostUsedSeries'] = most_used_series
            new_row['similar_percentage'] = None
            new_row['AllSimilarAbove85'] = None
            new_row['major_Sim'] = major_sim
            result_rows.append(new_row)
            continue

        # Find all matches using enhanced matching logic
        matches = []
        match_priority = {
            "exact": 1,
            "case_sensitive": 2,
            "nan_alpha": 3,
            "contain": 4,
            "similar": 5,
            "no_match": 6
        }

        for existing_series, maj_id in lookup[key]:
            match_type, comment, score = check_series_match(series, existing_series)

            if match_type == "exact":
                matches.append({
                    'type': match_type,
                    'comment': comment,
                    'score': score,
                    'series': existing_series,
                    'major_id': maj_id,
                    'priority': match_priority[match_type]
                })
                break

            elif match_type != "no_match":
                matches.append({
                    'type': match_type,
                    'comment': comment,
                    'score': score,
                    'series': existing_series,
                    'major_id': maj_id,
                    'priority': match_priority[match_type]
                })

        # Calculate all matches above 85% similarity
        if series == '-':
            valid_series = [s for s, _ in lookup[key] if s not in ['', 'nan', 'None']]
            sim_scores = [(s, similarity_ratio(series, s)) for s in valid_series if '-' in s]
        else:
            valid_series = [s for s, _ in lookup[key] if s not in ['-', '', 'nan', 'None']]
            sim_scores = [(s, similarity_ratio(series, s)) for s in valid_series]

        sim_scores.sort(key=lambda x: x[1], reverse=True)
        above_85 = [f"{s}({score}%)" for s, score in sim_scores if score >= 85]
        all_similar_above_85_str = ", ".join(above_85) if above_85 else None

        # Create the new row
        new_row = row.copy()
        new_row['MostUsedSeries'] = most_used_series
        new_row['AllSimilarAbove85'] = all_similar_above_85_str
        new_row['major_Sim'] = major_sim

        if not matches:
            new_row['comments'] = 'NotFound'
            new_row['MajorID'] = None
            new_row['FoundSeries'] = None
            new_row['similar_percentage'] = None
        else:
            # Sort matches by priority, then by score
            matches.sort(key=lambda x: (x['priority'], -x['score']))
            best_match = matches[0]

            # Special case for dash requests
            if series == '-':
                dash_matches = [m for m in matches if '-' in m['series']]
                if not dash_matches:
                    new_row['comments'] = 'NotFound'
                    new_row['MajorID'] = None
                    new_row['FoundSeries'] = None
                    new_row['similar_percentage'] = None
                else:
                    best_dash_match = dash_matches[0]
                    new_row['comments'] = best_dash_match['comment']
                    new_row['MajorID'] = best_dash_match['major_id']
                    new_row['FoundSeries'] = best_dash_match['series']
                    new_row['similar_percentage'] = best_dash_match['score']
            else:
                new_row['comments'] = best_match['comment']
                new_row['MajorID'] = best_match['major_id']
                new_row['FoundSeries'] = best_match['series']
                new_row['similar_percentage'] = best_match['score']

        result_rows.append(new_row)

    # Create final dataframe
    df_final = pd.DataFrame(result_rows)

    # Apply business rules if provided
    if rules_df is not None:
        df_final = apply_rules_with_special_case(rules_df, df_final)

    return df_final

def update_master_series_logic(update_df, master_df):
    """Update master series logic."""
    # Validate required columns
    required_cols = ['VariantID', 'ManufacturerName', 'Category', 'Family', 'RequestedSeries']
    missing_cols = [col for col in required_cols if col not in update_df.columns]

    if missing_cols:
        st.error(f"Update file missing columns: {missing_cols}")
        return None, 0, 0

    # Create composite key for both dataframes
    update_df['composite_key'] = (update_df['VariantID'].astype(str) + '|' +
                                update_df['ManufacturerName'].astype(str) + '|' +
                                update_df['Category'].astype(str) + '|' +
                                update_df['Family'].astype(str))

    master_df['composite_key'] = (master_df['VariantID'].astype(str) + '|' +
                                master_df['ManufacturerName'].astype(str) + '|' +
                                master_df['Category'].astype(str) + '|' +
                                master_df['Family'].astype(str))

    updated_count = 0
    appended_count = 0
    original_len = len(master_df)

    # Process each row in update file
    for _, update_row in update_df.iterrows():
        key = update_row['composite_key']
        mask = master_df['composite_key'] == key

        if mask.any():
            # Key exists - update RequestedSeries
            master_df.loc[mask, 'RequestedSeries'] = update_row['RequestedSeries']
            updated_count += 1
        else:
            # Key doesn't exist - append new row
            new_row = update_row.drop('composite_key').to_dict()
            # Ensure all required columns from master are present
            for col in master_df.columns:
                if col not in new_row and col != 'composite_key':
                    new_row[col] = None

            master_df = pd.concat([master_df, pd.DataFrame([new_row])], ignore_index=True)
            appended_count += 1

    # Remove composite_key column
    master_df = master_df.drop('composite_key', axis=1)

    return master_df, updated_count, appended_count

def delete_from_master_series_logic(delete_df, master_df):
    """Delete from master series logic."""
    # Validate required columns
    required_cols = ['VariantID', 'ManufacturerName', 'Category', 'Family']
    missing_cols = [col for col in required_cols if col not in delete_df.columns]

    if missing_cols:
        st.error(f"Delete file missing columns: {missing_cols}")
        return None, 0

    # Create composite key for both dataframes
    delete_df['composite_key'] = (delete_df['VariantID'].astype(str) + '|' +
                                delete_df['ManufacturerName'].astype(str) + '|' +
                                delete_df['Category'].astype(str) + '|' +
                                delete_df['Family'].astype(str))

    master_df['composite_key'] = (master_df['VariantID'].astype(str) + '|' +
                                master_df['ManufacturerName'].astype(str) + '|' +
                                master_df['Category'].astype(str) + '|' +
                                master_df['Family'].astype(str))

    # Get keys to delete
    keys_to_delete = set(delete_df['composite_key'].tolist())
    original_count = len(master_df)

    # Filter out rows with keys to delete
    master_df = master_df[~master_df['composite_key'].isin(keys_to_delete)]
    deleted_count = original_count - len(master_df)

    # Remove composite_key column
    master_df = master_df.drop('composite_key', axis=1)

    return master_df, deleted_count

def update_series_rules_logic(update_df, rules_df):
    """Update series rules logic."""
    # Validate required columns
    required_cols = ['ManufacturerName', 'Category', 'Family', 'Rule']
    missing_cols = [col for col in required_cols if col not in update_df.columns]

    if missing_cols:
        st.error(f"Update file missing columns: {missing_cols}")
        return None, 0, 0

    # Create composite key for both dataframes
    update_df['composite_key'] = (update_df['ManufacturerName'].astype(str) + '|' +
                                update_df['Category'].astype(str) + '|' +
                                update_df['Family'].astype(str))

    rules_df['composite_key'] = (rules_df['ManufacturerName'].astype(str) + '|' +
                               rules_df['Category'].astype(str) + '|' +
                               rules_df['Family'].astype(str))

    updated_count = 0
    appended_count = 0

    # Process each row in update file
    for _, update_row in update_df.iterrows():
        key = update_row['composite_key']
        mask = rules_df['composite_key'] == key

        if mask.any():
            # Key exists - update Rule
            rules_df.loc[mask, 'Rule'] = update_row['Rule']
            updated_count += 1
        else:
            # Key doesn't exist - append new row
            new_row = update_row.drop('composite_key').to_dict()
            # Ensure all required columns from rules are present
            for col in rules_df.columns:
                if col not in new_row and col != 'composite_key':
                    new_row[col] = None

            rules_df = pd.concat([rules_df, pd.DataFrame([new_row])], ignore_index=True)
            appended_count += 1

    # Remove composite_key column
    rules_df = rules_df.drop('composite_key', axis=1)

    return rules_df, updated_count, appended_count

def delete_from_series_rules_logic(delete_df, rules_df):
    """Delete from series rules logic."""
    # Validate required columns
    required_cols = ['ManufacturerName', 'Category', 'Family']
    missing_cols = [col for col in required_cols if col not in delete_df.columns]

    if missing_cols:
        st.error(f"Delete file missing columns: {missing_cols}")
        return None, 0

    # Create composite key for both dataframes
    delete_df['composite_key'] = (delete_df['ManufacturerName'].astype(str) + '|' +
                                delete_df['Category'].astype(str) + '|' +
                                delete_df['Family'].astype(str))

    rules_df['composite_key'] = (rules_df['ManufacturerName'].astype(str) + '|' +
                               rules_df['Category'].astype(str) + '|' +
                               rules_df['Family'].astype(str))

    # Get keys to delete
    keys_to_delete = set(delete_df['composite_key'].tolist())
    original_count = len(rules_df)

    # Filter out rows with keys to delete
    rules_df = rules_df[~rules_df['composite_key'].isin(keys_to_delete)]
    deleted_count = original_count - len(rules_df)

    # Remove composite_key column
    rules_df = rules_df.drop('composite_key', axis=1)

    return rules_df, deleted_count

# Streamlit App
def main():
    st.set_page_config(
        page_title="Enhanced Series Comparison Tool",
        page_icon="🔍",
        layout="wide"
    )
    
    st.title("🔍 Enhanced Series Comparison Tool")
    st.markdown("Compare series data with CRUD operations support")
    
    # GitHub Token Input
    st.sidebar.header("🔑 GitHub Configuration")
    github_token = st.sidebar.text_input(
        "GitHub Personal Access Token",
        type="password",
        help="Required for updating files on GitHub. Get one from GitHub Settings > Developer settings > Personal access tokens"
    )
    
    if github_token:
        st.sidebar.success("✅ Token provided")
    else:
        st.sidebar.warning("⚠️ Token required for GitHub updates")

    # Load master files from GitHub
    with st.spinner("Loading master files from GitHub..."):
        master_df = load_file_from_github(MASTER_URL)
        rules_df = load_file_from_github(RULES_URL)

    if master_df is None or rules_df is None:
        st.error("Failed to load master files from GitHub")
        st.stop()

    st.success(f"✅ Loaded Master Series History: {len(master_df)} rows")
    st.success(f"✅ Loaded Series Rules: {len(rules_df)} rows")

    # Operation selection
    st.header("📋 Choose Operation")
    operation = st.selectbox(
        "Select what you want to do:",
        [
            "1 - Update MasterSeriesHistory.xlsx",
            "2 - Delete from MasterSeriesHistory.xlsx", 
            "3 - Update SampleSeriesRules.xlsx",
            "4 - Delete from SampleSeriesRules.xlsx",
            "5 - Compare Series"
        ]
    )

    # File upload section
    st.header("📁 Upload File")
    uploaded_file = st.file_uploader(
        "Upload your Excel file",
        type=['xlsx', 'xls'],
        help="Upload the file containing data for your selected operation"
    )

    if uploaded_file is not None:
        try:
            uploaded_df = pd.read_excel(uploaded_file, engine='openpyxl')
            st.success(f"✅ File uploaded successfully: {len(uploaded_df)} rows")
            
            # Show preview of uploaded data
            st.subheader("📊 Data Preview")
            st.dataframe(uploaded_df.head(), use_container_width=True)

            # Execute operation based on selection
            if st.button("🚀 Execute Operation", type="primary"):
                with st.spinner("Processing..."):
                    if operation.startswith("1"):
                        # Update Master Series History
                        st.subheader("📝 Updating Master Series History")
                        updated_master, updated_count, appended_count = update_master_series_logic(uploaded_df, master_df.copy())
                        
                        if updated_master is not None:
                            st.success(f"✅ Update completed!")
                            st.info(f"📊 Updated existing rows: {updated_count}")
                            st.info(f"📊 Appended new rows: {appended_count}")
                            st.info(f"📊 Total rows: {len(updated_master)}")
                            
                            # Create download buttons and GitHub update
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                # Download button
                                output = io.BytesIO()
                                updated_master.to_excel(output, index=False, engine='openpyxl')
                                st.download_button(
                                    label="💾 Download Updated Master Series",
                                    data=output.getvalue(),
                                    file_name=f"Updated_MasterSeriesHistory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            
                            with col2:
                                # GitHub update button
                                if github_token and st.button("🚀 Update GitHub Repository", key="update_master_github"):
                                    with st.spinner("Updating GitHub repository..."):
                                        success, message = update_github_file(
                                            updated_master, 
                                            MASTER_FILE_PATH,
                                            f"Update MasterSeriesHistory - {updated_count} updated, {appended_count} new rows",
                                            github_token
                                        )
                                        
                                        if success:
                                            st.success(f"🎉 {message}")
                                        else:
                                            st.error(f"❌ {message}")
                                elif not github_token:
                                    st.warning("GitHub token required for repository updates")

                    elif operation.startswith("2"):
                        # Delete from Master Series History
                        st.subheader("🗑️ Deleting from Master Series History")
                        st.warning("⚠️ This will permanently delete matching rows!")
                        
                        confirm = st.checkbox("I confirm I want to delete these rows")
                        if confirm:
                            updated_master, deleted_count = delete_from_master_series_logic(uploaded_df, master_df.copy())
                            
                            if updated_master is not None:
                                st.success(f"✅ Deletion completed!")
                                st.info(f"🗑️ Deleted rows: {deleted_count}")
                                st.info(f"📊 Remaining rows: {len(updated_master)}")
                                
                                # Create download buttons and GitHub update
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    # Download button
                                    output = io.BytesIO()
                                    updated_master.to_excel(output, index=False, engine='openpyxl')
                                    st.download_button(
                                        label="💾 Download Updated Master Series",
                                        data=output.getvalue(),
                                        file_name=f"Deleted_MasterSeriesHistory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                
                                with col2:
                                    # GitHub update button
                                    if github_token and st.button("🚀 Update GitHub Repository", key="delete_master_github"):
                                        with st.spinner("Updating GitHub repository..."):
                                            success, message = update_github_file(
                                                updated_master, 
                                                MASTER_FILE_PATH,
                                                f"Delete from MasterSeriesHistory - {deleted_count} rows deleted",
                                                github_token
                                            )
                                            
                                            if success:
                                                st.success(f"🎉 {message}")
                                            else:
                                                st.error(f"❌ {message}")
                                    elif not github_token:
                                        st.warning("GitHub token required for repository updates")
                        else:
                            st.warning("Please confirm deletion to proceed")

                    elif operation.startswith("3"):
                        # Update Series Rules
                        st.subheader("📝 Updating Series Rules")
                        updated_rules, updated_count, appended_count = update_series_rules_logic(uploaded_df, rules_df.copy())
                        
                        if updated_rules is not None:
                            st.success(f"✅ Update completed!")
                            st.info(f"📊 Updated existing rules: {updated_count}")
                            st.info(f"📊 Appended new rules: {appended_count}")
                            st.info(f"📊 Total rules: {len(updated_rules)}")
                            
                            # Create download buttons and GitHub update
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                # Download button
                                output = io.BytesIO()
                                updated_rules.to_excel(output, index=False, engine='openpyxl')
                                st.download_button(
                                    label="💾 Download Updated Series Rules",
                                    data=output.getvalue(),
                                    file_name=f"Updated_SeriesRules_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            
                            with col2:
                                # GitHub update button
                                if github_token and st.button("🚀 Update GitHub Repository", key="update_rules_github"):
                                    with st.spinner("Updating GitHub repository..."):
                                        success, message = update_github_file(
                                            updated_rules, 
                                            RULES_FILE_PATH,
                                            f"Update SeriesRules - {updated_count} updated, {appended_count} new rules",
                                            github_token
                                        )
                                        
                                        if success:
                                            st.success(f"🎉 {message}")
                                        else:
                                            st.error(f"❌ {message}")
                                elif not github_token:
                                    st.warning("GitHub token required for repository updates")

                    elif operation.startswith("4"):
                        # Delete from Series Rules
                        st.subheader("🗑️ Deleting from Series Rules")
                        st.warning("⚠️ This will permanently delete matching rules!")
                        
                        confirm = st.checkbox("I confirm I want to delete these rules")
                        if confirm:
                            updated_rules, deleted_count = delete_from_series_rules_logic(uploaded_df, rules_df.copy())
                            
                            if updated_rules is not None:
                                st.success(f"✅ Deletion completed!")
                                st.info(f"🗑️ Deleted rules: {deleted_count}")
                                st.info(f"📊 Remaining rules: {len(updated_rules)}")
                                
                                # Create download buttons and GitHub update
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    # Download button
                                    output = io.BytesIO()
                                    updated_rules.to_excel(output, index=False, engine='openpyxl')
                                    st.download_button(
                                        label="💾 Download Updated Series Rules",
                                        data=output.getvalue(),
                                        file_name=f"Deleted_SeriesRules_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                
                                with col2:
                                    # GitHub update button
                                    if github_token and st.button("🚀 Update GitHub Repository", key="delete_rules_github"):
                                        with st.spinner("Updating GitHub repository..."):
                                            success, message = update_github_file(
                                                updated_rules, 
                                                RULES_FILE_PATH,
                                                f"Delete from SeriesRules - {deleted_count} rules deleted",
                                                github_token
                                            )
                                            
                                            if success:
                                                st.success(f"🎉 {message}")
                                            else:
                                                st.error(f"❌ {message}")
                                    elif not github_token:
                                        st.warning("GitHub token required for repository updates")
                        else:
                            st.warning("Please confirm deletion to proceed")

                    elif operation.startswith("5"):
                        # Compare Series
                        st.subheader("🔍 Comparing Series")
                        
                        # Options for comparison
                        col1, col2 = st.columns(2)
                        with col1:
                            use_rules = st.checkbox("Apply business rules", value=True)
                        with col2:
                            top_n = st.number_input("Top N most used series", min_value=1, max_value=10, value=2)
                        
                        # Perform comparison
                        rules_to_use = rules_df if use_rules else None
                        result_df = compare_series_logic(master_df.copy(), uploaded_df, rules_to_use, top_n)
                        
                        if result_df is not None:
                            st.success(f"✅ Series comparison completed!")
                            st.info(f"📊 Processed rows: {len(result_df)}")
                            
                            # Show match summary
                            st.subheader("📊 Match Summary")
                            comment_counts = result_df['comments'].value_counts()
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write("**Comment Distribution:**")
                                for comment, count in comment_counts.items():
                                    st.write(f"• {comment}: {count}")
                            
                            with col2:
                                if 'major_Sim' in result_df.columns:
                                    major_sim_counts = result_df['major_Sim'].value_counts().head(5)
                                    st.write("**Top 5 Major_Sim Values:**")
                                    for major_sim, count in major_sim_counts.items():
                                        display_text = str(major_sim)[:50] + "..." if len(str(major_sim)) > 50 else str(major_sim)
                                        st.write(f"• {display_text}: {count}")
                            
                            # Show results preview
                            st.subheader("📋 Results Preview")
                            st.dataframe(result_df.head(10), use_container_width=True)
                            
                            # Download button
                            output = io.BytesIO()
                            result_df.to_excel(output, index=False, engine='openpyxl')
                            st.download_button(
                                label="💾 Download Comparison Results",
                                data=output.getvalue(),
                                file_name=f"SeriesComparison_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            # Show detailed results in expandable section
                            with st.expander("🔍 View Full Results"):
                                st.dataframe(result_df, use_container_width=True)

        except Exception as e:
            st.error(f"Error processing file: {e}")

    # Information section
    st.header("ℹ️ Required File Formats")
    
    with st.expander("📋 Required Columns for Each Operation"):
        st.markdown("""
        **1. Update MasterSeriesHistory.xlsx:**
        - VariantID
        - ManufacturerName
        - Category
        - Family
        - RequestedSeries
        
        **2. Delete from MasterSeriesHistory.xlsx:**
        - VariantID
        - ManufacturerName
        - Category
        - Family
        
        **3. Update SampleSeriesRules.xlsx:**
        - ManufacturerName
        - Category
        - Family
        - Rule
        
        **4. Delete from SampleSeriesRules.xlsx:**
        - ManufacturerName
        - Category
        - Family
        
        **5. Compare Series:**
        - ManufacturerName
        - Category
        - Family
        - RequestedSeries
        """)

    with st.expander("🔍 How Comparison Works"):
        st.markdown("""
        The comparison tool uses several matching strategies in order of priority:
        
        1. **Exact Match** (100%): Perfect case-sensitive match
        2. **Case Sensitive Match** (100%): Same content, different case
        3. **Normalized Match** (100%): Same after removing separators and case
        4. **Containment Match**: One series contains the other
        5. **Similarity Match**: ≥60% similarity using SequenceMatcher
        
        **Output Columns:**
        - `comments`: Type of match found
        - `MajorID`: Usage percentage in master data
        - `FoundSeries`: The matching series from master data
        - `MostUsedSeries`: Top N most frequently used series for this key
        - `AllSimilarAbove85`: All series with ≥85% similarity
        - `major_Sim`: Most frequent series that contains the requested series
        - `similar_percentage`: Similarity score
        """)

    with st.expander("🔧 GitHub Integration Setup"):
        st.markdown("""
        **To enable GitHub updates, you need a Personal Access Token:**
        
        1. Go to GitHub Settings → Developer settings → Personal access tokens → Tokens (classic)
        2. Click "Generate new token (classic)"
        3. Select scopes: `repo` (Full control of private repositories)
        4. Copy the token and paste it in the sidebar
        
        **Security Note:** Your token is only used during this session and is not stored.
        """)

    # Footer
    st.markdown("---")
    st.markdown("🔧 **Enhanced Series Comparison Tool** - Built with Streamlit")

if __name__ == "__main__":
    main()
