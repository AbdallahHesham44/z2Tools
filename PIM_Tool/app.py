```python
import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
import requests
import io

# =========================
# Utility Functions
# =========================

def make_pattern(text):
    def replacer(match):
        integer_part = match.group(1)
        decimal_part = match.group(2)
        integer_replaced = "$"
        decimal_replaced = "$" * len(decimal_part) if decimal_part else ""
        return integer_replaced + ("." + decimal_replaced if decimal_part else "")
    return re.sub(r"(\d+)(?:\.(\d+))?", replacer, text)

def process_excel(df):
    df["key"] = (
        df["Category"].astype(str) + "|" +
        df["Sub-Category"].astype(str) + "|" +
        df["Attribute Name"].astype(str)
    )
    df["Helper_pattern"] = df["Preset values"].astype(str).apply(
        lambda x: re.sub(r"\d+(\.\d+)?", "$", x)
    )
    df["pattern"] = df["Preset values"].astype(str).apply(make_pattern)
    df["count"] = df.groupby(["key", "pattern"])["pattern"].transform("count")
    return df

def sort_preset_values(df):
    def is_alpha_only(s):
        return bool(re.fullmatch(r"[A-Za-z\s]+", s))

    def sort_groups(val):
        if not isinstance(val, str):
            return val
        groups = [grp.strip() for grp in val.split(" - ")]
        sorted_groups = []
        for grp in groups:
            if ", " in grp:
                parts = [p.strip() for p in grp.split(",")]
                if all(is_alpha_only(p) for p in parts):
                    parts_sorted = sorted(parts, key=lambda x: x.lower())
                    sorted_groups.append(", ".join(parts_sorted))
                else:
                    sorted_groups.append(", ".join(parts))
            else:
                sorted_groups.append(grp.strip())
        return " - ".join(sorted_groups)

    df["sorted Preset values"] = df["Preset values"].apply(sort_groups)
    df["Was Sorted"] = df["Preset values"].astype(str).str.strip() != df["sorted Preset values"].astype(str).str.strip()
    return df

def normalize_pattern(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"\s+", "", s)
    s = s.replace("±", "+-")
    return s.lower()

def mark_patterns(df1, df2):
    df1["pattern_norm"] = df1["Helper_pattern"].apply(normalize_pattern)
    df2["Pattern_norm"] = df2["Pattern"].apply(normalize_pattern)
    patterns_set = set(df2["Pattern_norm"].unique())
    df1["IsNumber"] = df1["pattern_norm"].apply(lambda x: "Yes" if x in patterns_set else "No")
    df1.drop(columns=["pattern_norm"], inplace=True)
    return df1

# =========================
# Streamlit App
# =========================
st.title("🔍 Excel Fuzzy Matcher Tool")

uploaded_file = st.file_uploader("Upload your input Excel file", type=["xlsx"])

if uploaded_file:
    # Load uploaded file
    df_input = pd.read_excel(uploaded_file, dtype=str)

    # GitHub raw links for reference files
    github_files = {
        "pattern_file": "https://raw.githubusercontent.com/yourusername/yourrepo/main/zvaluepatternbystatus_input.xlsx",
        "preset_file": "https://raw.githubusercontent.com/yourusername/yourrepo/main/Preset_15_pattern-count_new.xlsx"
    }

    # Load GitHub files
    st.info("📥 Downloading reference files from GitHub...")
    github_data = {}
    for name, url in github_files.items():
        r = requests.get(url)
        r.raise_for_status()
        github_data[name] = pd.read_excel(io.BytesIO(r.content), dtype=str)

    st.success("✅ All files loaded successfully!")

    # Step 1: Process Excel
    st.subheader("Step 1: Pattern & Count")
    df_proc = process_excel(df_input)
    st.dataframe(df_proc.head())

    # Step 2: Smart Sort
    st.subheader("Step 2: Smart Sort")
    df_sorted = sort_preset_values(df_proc)
    st.dataframe(df_sorted.head())

    # Step 3: Mark Patterns
    st.subheader("Step 3: Mark Patterns")
    df_marked = mark_patterns(df_sorted, github_data["pattern_file"])
    st.dataframe(df_marked.head())

    # Step 4: Final Output
    st.subheader("Step 4: Save Results")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_marked.to_excel(writer, index=False, sheet_name="Processed")
    st.download_button("⬇️ Download Processed Excel", data=output.getvalue(), file_name="processed_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
```
