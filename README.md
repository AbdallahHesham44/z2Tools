# z2Tools

## Deploying the Streamlit App

This repository contains the Streamlit app in `New-Acc-main/app.py`.

### Required files
- `New-Acc-main/app.py`
- `requirements.txt`
- `New-Acc-main/mapping.xlsx` (optional fallback)
- Supporting modules such as `github_utils.py`, `result_combiner.py`, `mapping_utils.py`, etc.

### Requirements
The app depends on:
- `streamlit`
- `pandas`
- `openpyxl`
- `xlsxwriter`
- `requests`
- `tqdm`
- `gdown`
- `PyDrive2`

### Streamlit Cloud deployment
1. Push this repository to GitHub.
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Create a new app.
4. Select repository: `AbdallahHesham44/z2Tools`.
5. Branch: `main`.
6. Main file path: `New-Acc-main/app.py`.

### Streamlit secrets
Add these in Streamlit app secrets if you want GitHub mapping file support:

```toml
[github]
token = "<your-github-token>"
owner = "AbdallahHesham44"
repo = "z2Tools"
file_path = "New-Acc-main/mapping.xlsx"
```

If you do not want GitHub access, commit `New-Acc-main/mapping.xlsx` to use the local fallback instead.

### Notes
- Streamlit Cloud uses its own secrets settings; local `.streamlit/secrets.toml` is not used in deployment.
- If GitHub secrets are missing, the app will load `New-Acc-main/mapping.xlsx` if available.
