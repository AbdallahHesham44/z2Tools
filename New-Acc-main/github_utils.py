#############################################
# MODULE: GITHUB UTILITIES
# Purpose: Handles interactions with the GitHub API
#          for downloading and updating the mapping file.
#############################################

import streamlit as st
import pandas as pd
import base64
import requests
import os
from io import BytesIO # Keep needed imports, even if also in other files

# --- download_mapping_file_from_github() --- (No changes needed)
def download_mapping_file_from_github() -> pd.DataFrame:
    """
    Downloads 'mapping.xlsx' from the GitHub repo specified in Streamlit secrets,
    returns a DataFrame parsed from that file.
    """
    st.write("DEBUG: Downloading mapping.xlsx from GitHub...")
    github_token = st.secrets["github"]["token"]
    owner = st.secrets["github"]["owner"]
    repo = st.secrets["github"]["repo"]
    file_path = st.secrets["github"]["file_path"]  # e.g. "mapping.xlsx"

    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{file_path}"
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        content_json = response.json()
        encoded_content = content_json["content"]
        decoded_bytes = base64.b64decode(encoded_content)

        local_file = "mapping.xlsx" # Define local file name for saving
        with open(local_file, "wb") as f:
            f.write(decoded_bytes)

        # Now parse the local file into a DataFrame
        try:
            # Use BytesIO to read directly from bytes without saving temporarily
            # df = pd.read_excel(BytesIO(decoded_bytes))
            # Decided to keep saving locally first as the original did, simplifies flow slightly
             df = pd.read_excel(local_file)
        except Exception as e:
            st.error(f"Failed to parse downloaded mapping file: {e}")
            # Clean up local file on error?
            # if os.path.exists(local_file):
            #      os.remove(local_file)
            st.stop() # Stop execution if parsing fails

        # Remove after successful read? Keep it for pipeline steps.
        # os.remove(local_file) # Keep local file temporarily for pipeline steps
        st.write("DEBUG: Download successful. mapping_df shape:", df.shape)
        return df
    elif response.status_code == 404:
         st.error(f"Failed to download file from GitHub: File not found at '{file_path}' (owner: {owner}, repo: {repo}).")
         st.stop()
    else:
        st.error(f"Failed to download file from GitHub: {response.status_code} {response.text}")
        st.stop()

# --- update_mapping_file_on_github() --- (No changes needed)
def update_mapping_file_on_github(mapping_df: pd.DataFrame) -> bool:
    """
    Updates 'mapping.xlsx' on GitHub using a PUT request to the GitHub API.
    """
    st.write("DEBUG: Attempting to update mapping.xlsx on GitHub.")
    st.write("DEBUG: DataFrame shape before upload:", mapping_df.shape)

    github_token = st.secrets["github"]["token"]
    owner = st.secrets["github"]["owner"]
    repo = st.secrets["github"]["repo"]
    file_path = st.secrets["github"]["file_path"]

    # 1) Save DF to BytesIO buffer, then encode
    try:
        output_buffer = BytesIO()
        mapping_df.to_excel(output_buffer, index=False, engine='openpyxl')
        output_buffer.seek(0)
        content_bytes = output_buffer.read()
        encoded_content = base64.b64encode(content_bytes).decode("utf-8")
    except Exception as e:
         st.error(f"Error converting DataFrame to Excel for upload: {e}")
         return False

    # 3) Get the current file's SHA
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{file_path}"
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }
    sha = None
    try:
        current_response = requests.get(url, headers=headers)
        if current_response.status_code == 200:
            sha = current_response.json().get("sha")
            st.write("DEBUG: Current file SHA:", sha)
        elif current_response.status_code == 404:
            st.write("DEBUG: No existing file found. Creating a new one...")
        else:
            st.error(f"Failed to get current file info from GitHub: {current_response.status_code} {current_response.text}")
            return False
    except requests.exceptions.RequestException as e:
         st.error(f"Network error getting file SHA: {e}")
         return False


    # 4) Prepare data payload
    data = {
        "message": "Update mapping file via Streamlit app",
        "content": encoded_content
    }
    if sha:
        data["sha"] = sha

    # 5) PUT request to update file
    try:
        update_response = requests.put(url, headers=headers, json=data)
    except requests.exceptions.RequestException as e:
        st.error(f"Network error during file update: {e}")
        return False


    if update_response.status_code in [200, 201]: # 200 OK (update), 201 Created (new file)
        st.write("DEBUG: Update/creation successful:", update_response.status_code)
        # Force re-download next time or update session state
        if "mapping_df" in st.session_state:
            del st.session_state["mapping_df"] # Invalidate cache
        return True
    else:
        st.error(f"Failed to update file on GitHub: {update_response.status_code} {update_response.text}")
        # Provide more specific feedback if possible (e.g., rate limits, auth errors)
        if update_response.status_code == 401:
            st.error("Authentication failed. Check your GitHub token in secrets.")
        elif update_response.status_code == 409:
             st.error("Conflict detected (SHA mismatch?). Try refreshing the app and saving again.")
        return False
