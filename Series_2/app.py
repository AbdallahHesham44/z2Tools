# app.py
import streamlit as st
import pandas as pd
import base64
from io import BytesIO
from github import Github

st.set_page_config(page_title="Series Updater", layout="centered")

# ======================
# Your provided helpers
# ======================
def load_excel_from_github(repo_name, file_path, token):
    g = Github(token)
    repo = g.get_repo(repo_name)
    contents = repo.get_contents(file_path)
    file_data = base64.b64decode(contents.content)
    return pd.read_excel(BytesIO(file_data)), contents.sha  # also return sha for safe updates

def save_excel_to_github(repo_name, file_path, df, commit_message, token, sha=None, branch="main"):
    g = Github(token)
    repo = g.get_repo(repo_name)

    # Save DataFrame to bytes
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    content_bytes = output.read()

    try:
        if sha is None:
            contents = repo.get_contents(file_path, ref=branch)
            sha = contents.sha
        repo.update_file(file_path, commit_message, content_bytes, sha, branch=branch)
    except Exception:
        repo.create_file(file_path, commit_message, content_bytes, branch=branch)

def modify_master_or_rules(df_existing, df_changes, action):
    key_cols = ["VariantID", "ManufacturerName", "Category", "Family"]

    # Validate key columns exist
    for col in key_cols + (["RequestedSeries"] if action == "append" else []):
        if col not in df_changes.columns:
            raise ValueError(f"Changes file is missing required column: {col}")
    for col in key_cols:
        if col not in df_existing.columns:
            raise ValueError(f"Target file is missing required key column: {col}")

    if action == "append":
        # stats
        existing_keys = set(map(tuple, df_existing[key_cols].astype(str).values))
        change_keys = list(map(tuple, df_changes[key_cols].astype(str).values))
        overwrite_count = sum(1 for k in change_keys if k in existing_keys)

        # do work
        for _, row in df_changes.iterrows():
            key_vals = tuple(map(str, (row[c] for c in key_cols)))
            mask = (df_existing[key_cols].astype(str) == pd.Series(key_vals, index=key_cols)).all(axis=1)
            if mask.any():
                # overwrite RequestedSeries only
                if "RequestedSeries" not in df_existing.columns:
                    raise ValueError("Target file has no 'RequestedSeries' column to overwrite.")
                df_existing.loc[mask, "RequestedSeries"] = row["RequestedSeries"]
            else:
                df_existing = pd.concat([df_existing, pd.DataFrame([row])], ignore_index=True)

        new_existing_keys = set(map(tuple, df_existing[key_cols].astype(str).values))
        appended_count = len(new_existing_keys) - len(existing_keys)
        return df_existing, {"overwritten": overwrite_count, "appended": appended_count}

    elif action == "delete":
        # stats
        existing_keys = set(map(tuple, df_existing[key_cols].astype(str).values))
        change_keys = set(map(tuple, df_changes[key_cols].astype(str).values))
        delete_hits = len(existing_keys.intersection(change_keys))

        # do work
        idx = pd.MultiIndex.from_frame(df_existing[key_cols].astype(str))
        kill = pd.MultiIndex.from_tuples(list(change_keys), names=key_cols)
        keep_mask = ~idx.isin(kill)
        df_existing = df_existing.loc[keep_mask].reset_index(drop=True)
        return df_existing, {"deleted": delete_hits, "ignored": len(change_keys) - delete_hits}

    else:
        raise ValueError("action must be 'append' or 'delete'.")

# ======================
# UI
# ======================
st.title("🔧 Series GitHub Updater")

with st.expander("🔐 GitHub Settings", expanded=True):
    repo_name = st.text_input("Repository (owner/repo)", value="AbdallahHesham44/z2Tools")
    branch = st.text_input("Branch", value="main")
    token = st.text_input("Personal Access Token (repo scope)", type="password", help="Token is used only from your browser to GitHub API.")

st.markdown("### 1) Choose action")
action = st.radio("Action", ["append", "delete"], horizontal=True)

st.markdown("### 2) Choose target file(s)")
file_choices = {
    "MasterSeriesHistory.xlsx": "Serise/MasterSeriesHistory.xlsx",
    "SeriesRules.xlsx (SampleSeriesRules.xlsx)": "Serise/SampleSeriesRules.xlsx",
}
targets = st.multiselect(
    "Select one or both files to modify",
    options=list(file_choices.keys()),
    default=["MasterSeriesHistory.xlsx"]
)

st.markdown("### 3) Upload changes Excel")
changes_file = st.file_uploader("Upload .xlsx containing changes (must include VariantID, ManufacturerName, Category, Family; and RequestedSeries for append)", type=["xlsx"])

run = st.button("🚀 Apply")

if run:
    if not token:
        st.error("Please enter your GitHub token.")
        st.stop()
    if not targets:
        st.error("Please choose at least one target file.")
        st.stop()
    if not changes_file:
        st.error("Please upload the changes Excel file.")
        st.stop()

    try:
        df_changes = pd.read_excel(changes_file)
    except Exception as e:
        st.error(f"Could not read your changes file: {e}")
        st.stop()

    for label in targets:
        st.divider()
        st.subheader(f"Target: {label}")

        gh_path = file_choices[label]

        # Load current file + sha for safe update
        try:
            df_existing, sha = load_excel_from_github(repo_name, gh_path, token)
        except Exception as e:
            st.error(f"❌ Failed to download {gh_path} from {repo_name}: {e}")
            continue

        before_rows = len(df_existing)
        try:
            updated_df, stats = modify_master_or_rules(df_existing.copy(), df_changes.copy(), action)
        except Exception as e:
            st.error(f"❌ Failed to apply changes to {label}: {e}")
            continue

        after_rows = len(updated_df)

        # Show stats
        st.write("**Summary**")
        if action == "append":
            st.write(f"- Rows before: **{before_rows}**")
            st.write(f"- Rows after: **{after_rows}**")
            st.write(f"- Overwritten: **{stats.get('overwritten', 0)}**")
            st.write(f"- Appended: **{stats.get('appended', 0)}**")
        else:
            st.write(f"- Rows before: **{before_rows}**")
            st.write(f"- Rows after: **{after_rows}**")
            st.write(f"- Deleted: **{stats.get('deleted', 0)}**")
            st.write(f"- Ignored (not found): **{stats.get('ignored', 0)}**")

        # Preview
        st.write("**Preview (top 20 rows)**")
        st.dataframe(updated_df.head(20))

        # Download updated file
        out_buf = BytesIO()
        updated_df.to_excel(out_buf, index=False)
        out_buf.seek(0)
        st.download_button(
            label=f"⬇️ Download updated {label}",
            data=out_buf,
            file_name=label.replace(" (SampleSeriesRules.xlsx)", ""),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Push back to GitHub
        if st.checkbox(f"Push updated {label} to GitHub", key=f"push_{label}"):
            try:
                save_excel_to_github(
                    repo_name=repo_name,
                    file_path=gh_path,
                    df=updated_df,
                    commit_message=f"{action.title()} via Streamlit on {label}",
                    token=token,
                    sha=sha,
                    branch=branch
                )
                st.success(f"✅ Pushed {label} to GitHub.")
            except Exception as e:
                st.error(f"❌ Failed to push {label}: {e}")
