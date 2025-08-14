import streamlit as st
import pandas as pd
import io
import requests
from difflib import SequenceMatcher
from collections import defaultdict
import re
import tempfile
import os
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Series Processing Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# GitHub repository configuration
GITHUB_REPO = "your-username/your-repo-name"  # Update with your repo details
GITHUB_BRANCH = "main"
GITHUB_TOKEN = None  # Set this if you need authentication

def get_github_file_url(filename):
    """Generate GitHub raw file URL"""
    return f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/{filename}"

def download_file_from_github(filename):
    """Download file from GitHub repository"""
    try:
        url = get_github_file_url(filename)
        headers = {}
        if GITHUB_TOKEN:
            headers['Authorization'] = f'token {GITHUB_TOKEN}'
        
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.content
    except Exception as e:
        st.error(f"Error downloading {filename}: {str(e)}")
        return None

def update_github_file(filename, content, commit_message="Update file via Streamlit"):
    """Update file in GitHub repository (requires token)"""
    if not GITHUB_TOKEN:
        st.error("GitHub token required for file updates")
        return False
    
    # This is a simplified version - in production, you'd use GitHub API
    st.info("GitHub file update functionality requires implementation with GitHub API")
    return True

# Core processing functions from your original code
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

    if exist_str in ['', 'nan', 'None']:
        return "no_match", "NotFound", 0.0

    if req_str == '-':
        if '-' in exist_str:
            return "contain", "FoundWithDiff(contain)", similarity_ratio(req_str, exist_str)
        else:
            return "no_match", "NotFound", 0.0

    if exist_str == '-':
        return "no_match", "NotFound", 0.0

    if req_str == exist_str:
        return "exact", "insertedBeforeExact", 100.0

    if req_str.lower() == exist_str.lower():
        return "case_sensitive", "Exact(caseSensitive)", 100.0

    req_normalized = normalize_series(req_str)
    exist_normalized = normalize_series(exist_str)

    if req_normalized == exist_normalized and req_str != exist_str:
        return "nan_alpha", "NanAlphaCase", 100.0

    if exist_normalized in req_normalized and exist_normalized != req_normalized:
        return "contain", "FoundWithDiff(contain)", similarity_ratio(req_str, exist_str)
    elif req_normalized in exist_normalized and req_normalized != exist_normalized:
        return "contain", "FoundWithDiff(contain)", similarity_ratio(req_str, exist_str)

    sim_score = similarity_ratio(req_str, exist_str)
    if sim_score >= 60:
        return "similar", f"similar_{sim_score}%", sim_score

    return "no_match", "NotFound", sim_score

def find_major_contain_series(input_series, lookup_data, key):
    """Find the most frequently used series that contains the input series."""
    if key not in lookup_data:
        return None
    
    series_counts = {}
    for series, major_id in lookup_data[key]:
        percentage = float(major_id.replace('%', ''))
        if series in series_counts:
            series_counts[series] = max(series_counts[series], percentage)
        else:
            series_counts[series] = percentage
    
    containing_series = []
    input_str = str(input_series).strip()
    
    for series, count_percentage in series_counts.items():
        if series in ['', 'nan', 'None']:
            continue
            
        series_str = str(series).strip()
        
        if input_str == '-':
            if '-' in series_str:
                containing_series.append((series_str, count_percentage))
        else:
            if series_str == '-':
                continue
            
            if input_str.lower() in series_str.lower():
                containing_series.append((series_str, count_percentage))
    
    if not containing_series:
        return None
    
    containing_series.sort(key=lambda x: x[1], reverse=True)
    best_series = containing_series[0][0]
    return f"Major_contain({best_series})"

def calculate_major_id(df):
    """Calculate MajorID percentages for series usage."""
    required_cols = ['VariantID', 'ManufacturerName', 'Category', 'Family', 'DataSheetURL', 'RequestedSeries']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        st.error(f"Missing columns: {missing_cols}")
        return None

    df['key'] = df[['ManufacturerName', 'Category', 'Family']].astype(str).agg('|'.join, axis=1)
    group_counts = df.groupby(['key', 'RequestedSeries']).size().reset_index(name='count')
    total_counts = df.groupby('key').size().reset_index(name='total')
    
    merged = pd.merge(df, group_counts, on=['key', 'RequestedSeries'], how='left')
    merged = pd.merge(merged, total_counts, on='key', how='left')
    merged['MajorID'] = ((merged['count'] / merged['total']) * 100).round(2).astype(str) + '%'
    
    merged.drop(columns=['count', 'total'], inplace=True)
    return merged

def apply_rules_with_special_case(rules_df, master_df):
    """Apply business rules to the master dataframe."""
    for col in ["ManufacturerName", "Category", "Family", "Rule"]:
        if col not in rules_df.columns:
            st.error(f"Rules file missing column: {col}")
            return master_df

    special_case_rules = rules_df[rules_df["ManufacturerName"] == "88xx"]
    normal_rules = rules_df[rules_df["ManufacturerName"] != "88xx"]

    normal_rules["key"] = normal_rules["ManufacturerName"] + "|" + normal_rules["Category"] + "|" + normal_rules["Family"]
    master_df["key"] = master_df["ManufacturerName"] + "|" + master_df["Category"] + "|" + master_df["Family"]

    master_df = master_df.merge(normal_rules[["key", "Rule"]], on="key", how="left", suffixes=("", "_rule"))

    for _, row in special_case_rules.iterrows():
        mask = (
            (master_df["ManufacturerName"] == "88xx") &
            (master_df["Category"] == row["Category"]) &
            (master_df["Family"] == row["Family"])
        )
        master_df.loc[mask, "Rule"] = row["Rule"]

    if "key" in master_df.columns:
        master_df = master_df.drop(columns=["key"])

    return master_df

def process_series_comparison(master_df, comparison_df, rules_df=None, top_n=1):
    """Main series processing function"""
    
    # Calculate MajorID for master file
    master_with_majorid = calculate_major_id(master_df)
    if master_with_majorid is None:
        return None

    # Create matching keys
    master_with_majorid['key'] = master_with_majorid[['ManufacturerName', 'Category', 'Family']].astype(str).agg('|'.join, axis=1)
    comparison_df['key'] = comparison_df[['ManufacturerName', 'Category', 'Family']].astype(str).agg('|'.join, axis=1)

    # Create lookup dictionary
    lookup = defaultdict(list)
    for _, row in master_with_majorid.iterrows():
        lookup[row['key']].append((str(row['RequestedSeries']), row['MajorID']))

    # Process each row in comparison file
    result_rows = []

    for _, row in comparison_df.iterrows():
        key = row['key']
        series = str(row['RequestedSeries'])

        # Calculate most used series
        if key in lookup:
            temp_df = pd.DataFrame(lookup[key], columns=['Series', 'MajorID'])
            if series == '-':
                temp_df = temp_df[~temp_df['Series'].isin(['', 'nan', 'None'])]
            else:
                temp_df = temp_df[~temp_df['Series'].isin(['-', '', 'nan', 'None'])]

            if not temp_df.empty:
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

        # Process matches
        new_row = row.copy()
        new_row['MostUsedSeries'] = most_used_series
        new_row['major_Sim'] = major_sim

        if key not in lookup:
            new_row['comments'] = 'NotFound'
            new_row['MajorID'] = None
            new_row['FoundSeries'] = None
            new_row['similar_percentage'] = None
            new_row['AllSimilarAbove85'] = None
        else:
            # Find matches
            matches = []
            match_priority = {"exact": 1, "case_sensitive": 2, "nan_alpha": 3, "contain": 4, "similar": 5, "no_match": 6}

            for existing_series, maj_id in lookup[key]:
                match_type, comment, score = check_series_match(series, existing_series)
                
                if match_type == "exact":
                    matches.append({
                        'type': match_type, 'comment': comment, 'score': score,
                        'series': existing_series, 'major_id': maj_id,
                        'priority': match_priority[match_type]
                    })
                    break
                elif match_type != "no_match":
                    matches.append({
                        'type': match_type, 'comment': comment, 'score': score,
                        'series': existing_series, 'major_id': maj_id,
                        'priority': match_priority[match_type]
                    })

            # Calculate similar above 85%
            if series == '-':
                valid_series = [s for s, _ in lookup[key] if s not in ['', 'nan', 'None']]
                sim_scores = [(s, similarity_ratio(series, s)) for s in valid_series if '-' in s]
            else:
                valid_series = [s for s, _ in lookup[key] if s not in ['-', '', 'nan', 'None']]
                sim_scores = [(s, similarity_ratio(series, s)) for s in valid_series]

            sim_scores.sort(key=lambda x: x[1], reverse=True)
            above_85 = [f"{s}({score}%)" for s, score in sim_scores if score >= 85]
            new_row['AllSimilarAbove85'] = ", ".join(above_85) if above_85 else None

            if not matches:
                new_row['comments'] = 'NotFound'
                new_row['MajorID'] = None
                new_row['FoundSeries'] = None
                new_row['similar_percentage'] = None
            else:
                matches.sort(key=lambda x: (x['priority'], -x['score']))
                best_match = matches[0]

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

    # Apply rules if provided
    if rules_df is not None:
        df_final = apply_rules_with_special_case(rules_df, df_final)

    return df_final

# Streamlit UI
def main():
    st.title("📊 Series Processing Tool")
    st.markdown("---")

    # Sidebar for navigation
    st.sidebar.title("Navigation")
    mode = st.sidebar.radio(
        "Select Mode:",
        ["📥 File Management", "🔄 Series Processing"]
    )

    if mode == "📥 File Management":
        file_management_section()
    else:
        series_processing_section()

def file_management_section():
    st.header("📥 File Management")
    
    tab1, tab2 = st.tabs(["📋 Download Templates", "📤 Upload & Update Files"])
    
    with tab1:
        st.subheader("Download File Templates")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**Input Series Template**")
            template_input = create_input_template()
            st.download_button(
                label="📋 Download TemplateInput_series.xlsx",
                data=template_input,
                file_name="TemplateInput_series.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            st.markdown("**Master Series History Template**")
            template_master = create_master_template()
            st.download_button(
                label="📋 Download TemplateMasterSeriesHistory.xlsx",
                data=template_master,
                file_name="TemplateMasterSeriesHistory.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col3:
            st.markdown("**Series Rules Template**")
            template_rules = create_rules_template()
            st.download_button(
                label="📋 Download TemplateSampleSeriesRules.xlsx",
                data=template_rules,
                file_name="TemplateSampleSeriesRules.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    with tab2:
        st.subheader("Upload & Update Repository Files")
        
        # File selection
        file_to_update = st.selectbox(
            "Select file to update:",
            ["MasterSeriesHistory.xlsx", "SampleSeriesRules.xlsx"]
        )
        
        # Operation selection
        operation = st.radio(
            "Select operation:",
            ["Append/Update", "Delete"]
        )
        
        # File upload
        uploaded_file = st.file_uploader(
            f"Upload file to {operation.lower()} {file_to_update}",
            type=['xlsx', 'xls', 'csv'],
            help="Upload a file with data to append/update or delete from the selected repository file"
        )
        
        if uploaded_file is not None:
            try:
                # Read uploaded file
                if uploaded_file.name.endswith('.csv'):
                    upload_df = pd.read_csv(uploaded_file)
                else:
                    upload_df = pd.read_excel(uploaded_file)
                
                st.success(f"✅ File uploaded successfully! ({len(upload_df)} rows)")
                st.dataframe(upload_df.head())
                
                # Download current repository file
                if st.button(f"Process {operation}"):
                    with st.spinner(f"Downloading current {file_to_update}..."):
                        current_file_content = download_file_from_github(file_to_update)
                    
                    if current_file_content:
                        # Read current file
                        current_df = pd.read_excel(io.BytesIO(current_file_content))
                        
                        # Process operation
                        if operation == "Append/Update":
                            updated_df = append_update_operation(current_df, upload_df)
                        else:
                            updated_df = delete_operation(current_df, upload_df)
                        
                        if updated_df is not None:
                            # Show preview
                            st.subheader("Preview of Updated Data")
                            st.dataframe(updated_df.head(10))
                            st.info(f"Updated file will have {len(updated_df)} rows")
                            
                            # Download updated file
                            output_buffer = io.BytesIO()
                            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                                updated_df.to_excel(writer, index=False)
                            
                            st.download_button(
                                label=f"📥 Download Updated {file_to_update}",
                                data=output_buffer.getvalue(),
                                file_name=f"Updated_{file_to_update}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            # Option to update GitHub (if token available)
                            if st.button("🚀 Update GitHub Repository"):
                                if update_github_file(file_to_update, output_buffer.getvalue()):
                                    st.success("✅ GitHub repository updated successfully!")
                                else:
                                    st.error("❌ Failed to update GitHub repository")
                    
            except Exception as e:
                st.error(f"Error processing file: {str(e)}")

def series_processing_section():
    st.header("🔄 Series Processing")
    
    # Input file upload
    st.subheader("📤 Upload Input File")
    input_file = st.file_uploader(
        "Upload input file for series processing",
        type=['xlsx', 'xls', 'csv'],
        help="Upload the file containing series data to be processed"
    )
    
    if input_file is not None:
        try:
            # Read input file
            if input_file.name.endswith('.csv'):
                input_df = pd.read_csv(input_file)
            else:
                input_df = pd.read_excel(input_file)
            
            st.success(f"✅ Input file loaded successfully! ({len(input_df)} rows)")
            
            # Show preview
            with st.expander("📋 Preview Input Data"):
                st.dataframe(input_df.head())
            
            # Processing options
            st.subheader("⚙️ Processing Options")
            
            col1, col2 = st.columns(2)
            with col1:
                top_n = st.number_input("Top N most used series", min_value=1, max_value=5, value=2)
            
            with col2:
                use_rules = st.checkbox("Apply business rules", value=True)
            
            # Process button
            if st.button("🚀 Process Series Data"):
                with st.spinner("Processing series data..."):
                    
                    # Download required files from GitHub
                    st.info("📥 Downloading files from repository...")
                    
                    master_content = download_file_from_github("MasterSeriesHistory.xlsx")
                    rules_content = download_file_from_github("SampleSeriesRules.xlsx") if use_rules else None
                    
                    if master_content:
                        # Read master file
                        master_df = pd.read_excel(io.BytesIO(master_content))
                        rules_df = None
                        
                        if use_rules and rules_content:
                            rules_df = pd.read_excel(io.BytesIO(rules_content))
                        
                        # Process series
                        result_df = process_series_comparison(master_df, input_df, rules_df, top_n)
                        
                        if result_df is not None:
                            st.success("✅ Processing completed successfully!")
                            
                            # Show results summary
                            st.subheader("📊 Results Summary")
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.markdown("**Match Summary**")
                                comment_counts = result_df['comments'].value_counts()
                                st.dataframe(comment_counts)
                            
                            with col2:
                                if 'major_Sim' in result_df.columns:
                                    st.markdown("**Major_Sim Summary**")
                                    major_sim_counts = result_df['major_Sim'].value_counts().head(5)
                                    st.dataframe(major_sim_counts)
                            
                            # Show results preview
                            with st.expander("📋 Results Preview"):
                                st.dataframe(result_df.head(10))
                            
                            # Download results
                            output_buffer = io.BytesIO()
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            filename = f"SeriesOutput_{timestamp}.xlsx"
                            
                            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                                result_df.to_excel(writer, index=False)
                            
                            st.download_button(
                                label="📥 Download Results",
                                data=output_buffer.getvalue(),
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.error("❌ Error processing series data")
                    else:
                        st.error("❌ Failed to download master file from repository")
        
        except Exception as e:
            st.error(f"Error reading input file: {str(e)}")

def create_input_template():
    """Create template for input series file"""
    template_data = {
        'ManufacturerName': ['Example Corp', 'Another Corp'],
        'Category': ['Category1', 'Category2'],
        'Family': ['Family1', 'Family2'],
        'RequestedSeries': ['Series-A', 'Series-B']
    }
    
    df = pd.DataFrame(template_data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    return output.getvalue()

def create_master_template():
    """Create template for master series history file"""
    template_data = {
        'VariantID': ['VAR001', 'VAR002'],
        'ManufacturerName': ['Example Corp', 'Another Corp'],
        'Category': ['Category1', 'Category2'],
        'Family': ['Family1', 'Family2'],
        'DataSheetURL': ['http://example.com/sheet1', 'http://example.com/sheet2'],
        'RequestedSeries': ['Series-A', 'Series-B']
    }
    
    df = pd.DataFrame(template_data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    return output.getvalue()

def create_rules_template():
    """Create template for series rules file"""
    template_data = {
        'ManufacturerName': ['Example Corp', '88xx'],
        'Category': ['Category1', 'Category2'],
        'Family': ['Family1', 'Family2'],
        'Rule': ['Rule1', 'Rule2']
    }
    
    df = pd.DataFrame(template_data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    return output.getvalue()

def append_update_operation(current_df, upload_df):
    """Perform append/update operation"""
    try:
        # Create keys for both dataframes
        key_columns = ['VariantID', 'ManufacturerName', 'Category', 'Family']
        
        # Check if required columns exist
        missing_cols = [col for col in key_columns if col not in upload_df.columns]
        if missing_cols:
            st.error(f"Upload file missing columns: {missing_cols}")
            return None
        
        # Create composite keys
        current_df['key'] = current_df[key_columns].astype(str).agg('|'.join, axis=1)
        upload_df['key'] = upload_df[key_columns].astype(str).agg('|'.join, axis=1)
        
        # Update existing records
        for _, row in upload_df.iterrows():
            mask = current_df['key'] == row['key']
            if mask.any():
                # Update RequestedSeries for existing records
                current_df.loc[mask, 'RequestedSeries'] = row['RequestedSeries']
            else:
                # Append new records
                new_row = row.copy()
                new_row = new_row.drop('key')  # Remove the key column
                current_df = pd.concat([current_df, new_row.to_frame().T], ignore_index=True)
        
        # Clean up
        current_df = current_df.drop('key', axis=1)
        
        st.success(f"✅ Operation completed. {len(upload_df)} records processed.")
        return current_df
        
    except Exception as e:
        st.error(f"Error in append/update operation: {str(e)}")
        return None

def delete_operation(current_df, upload_df):
    """Perform delete operation"""
    try:
        # Create keys for both dataframes
        key_columns = ['VariantID', 'ManufacturerName', 'Category', 'Family']
        
        # Check if required columns exist
        missing_cols = [col for col in key_columns if col not in upload_df.columns]
        if missing_cols:
            st.error(f"Upload file missing columns: {missing_cols}")
            return None
        
        # Create composite keys
        current_df['key'] = current_df[key_columns].astype(str).agg('|'.join, axis=1)
        upload_df['key'] = upload_df[key_columns].astype(str).agg('|'.join, axis=1)
        
        # Get keys to delete
        keys_to_delete = set(upload_df['key'].tolist())
        
        # Filter out records to delete
        initial_count = len(current_df)
        current_df = current_df[~current_df['key'].isin(keys_to_delete)]
        deleted_count = initial_count - len(current_df)
        
        # Clean up
        current_df = current_df.drop('key', axis=1)
        
        st.success(f"✅ Delete operation completed. {deleted_count} records removed.")
        return current_df
        
    except Exception as e:
        st.error(f"Error in delete operation: {str(e)}")
        return None

if __name__ == "__main__":
    main()
