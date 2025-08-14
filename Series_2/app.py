import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from github import Github
from series_processing import compare_requested_series

# ======================
# GitHub setup
# ======================
GITHUB_REPO = "AbdallahHesham44/z2Tools"
GITHUB_BRANCH = "main"
TOKEN = st.secrets.get("GITHUB_TOKEN", "")

def load_from_github(path):
    url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/{path}"
    response = requests.get(url)
    response.raise_for_status()
    return BytesIO(response.content)

def upload_to_github(path, content, message):
    if not TOKEN:
        st.error("GitHub token is missing in secrets.")
        return
    g = Github(TOKEN)
    repo = g.get_repo(GITHUB_REPO)
    try:
        file_content = repo.get_contents(path, ref=GITHUB_BRANCH)
        repo.update_file(file_content.path, message, content, file_content.sha, branch=GITHUB_BRANCH)
    except:
        repo.create_file(path, message, content, branch=GITHUB_BRANCH)

# ======================
# Streamlit UI
# ======================
st.title("📊 Series Processing Tool")

# 1️⃣ Download Template files
st.subheader("📥 Download Templates")
templates = {
    "TemplateInput_series.xlsx": "Serise/TempleteInput_series.xlsx",
    "TemplateMasterSeriesHistory.xlsx": "Serise/TempleteMasterSeriesHistory.xlsx",
    "TemplateSampleSeriesRules.xlsx": "Serise/TempleteSampleSeriesRules.xlsx"
}

for name, path in templates.items():
    btn = st.download_button(
        label=f"Download {name}",
        data=load_from_github(path).read(),
        file_name=name
    )

# 2️⃣ Select operation
st.subheader("⚙️ Operation Mode")
operation = st.radio("Choose an operation:", ["Append", "Delete"])

# 3️⃣ File upload
uploaded_input = st.file_uploader("Upload Input_series.xlsx", type=["xlsx"])

# 4️⃣ Process button
if uploaded_input:
    # Load GitHub master and rules
    master_file = load_from_github("Serise/MasterSeriesHistory.xlsx")
    rules_file = load_from_github("Serise/SampleSeriesRules.xlsx")

    st.write("Processing...")
    output_df = compare_requested_series(master_file, uploaded_input, rules_file, top_n=1)

    st.dataframe(output_df)

    # Save locally
    towrite = BytesIO()
    output_df.to_excel(towrite, index=False)
    towrite.seek(0)

    st.download_button(
        label="💾 Download Output",
        data=towrite,
        file_name="Output_Series.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 5️⃣ Option to update GitHub
    if st.checkbox("Update MasterSeriesHistory.xlsx and SeriesRules.xlsx in GitHub"):
        upload_to_github("Serise/MasterSeriesHistory.xlsx", towrite.getvalue(), f"{operation} operation update")
        st.success("GitHub files updated successfully!")
