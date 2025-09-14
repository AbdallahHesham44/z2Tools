import streamlit as st
import pandas as pd
import re
import io
import zipfile
from difflib import SequenceMatcher
from datetime import datetime

# --------------------------
# Helper Functions
# --------------------------

def extract_pattern(part_number: str) -> str:
    """Extract numeric resistance value pattern from a part number."""
    match = re.search(r"(\d+\.?\d*[a-zA-Z]*)", part_number)
    if not match:
        return "NoMatch"

    value = match.group(1)

    # Replace digits before decimal with $
    if "." in value:
        before, after = value.split(".", 1)
        before = "$" * len(before)
        after = "".join("$" if c.isdigit() else c for c in after)
        return f"{before}.{after}"
    else:
        return "$" * len(value)


def compute_similarity(a: str, b: str) -> float:
    """Compute similarity ratio between two strings."""
    return SequenceMatcher(None, a, b).ratio()


def process_file(uploaded_file, similarity_threshold=0.8):
    """Process uploaded Excel file and return matched/unmatched DataFrames."""
    df = pd.read_excel(uploaded_file)

    if "PartNumber" not in df.columns:
        st.error("❌ Excel file must contain a 'PartNumber' column.")
        return None, None, None

    df["ExtractedPattern"] = df["PartNumber"].astype(str).apply(extract_pattern)

    matched, unmatched = [], []

    for i, row1 in df.iterrows():
        best_match, best_score = None, 0.0
        for j, row2 in df.iterrows():
            if i == j:
                continue
            score = compute_similarity(row1["ExtractedPattern"], row2["ExtractedPattern"])
            if score > best_score:
                best_match, best_score = row2, score

        if best_match is not None and best_score >= similarity_threshold:
            matched.append({
                "PartNumber": row1["PartNumber"],
                "Pattern": row1["ExtractedPattern"],
                "SuggestedValue": best_match["PartNumber"],
                "Similarity": f"{best_score:.2%}"
            })
        else:
            unmatched.append({
                "PartNumber": row1["PartNumber"],
                "Pattern": row1["ExtractedPattern"],
                "SuggestedValue": "N/A",
                "Similarity": "N/A"
            })

    return df, pd.DataFrame(matched), pd.DataFrame(unmatched)


def create_download_zip(matched_df, unmatched_df, summary_df):
    """Create downloadable ZIP with results."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as zipf:
        if matched_df is not None and not matched_df.empty:
            matched_bytes = io.BytesIO()
            matched_df.to_excel(matched_bytes, index=False)
            zipf.writestr("matched_results.xlsx", matched_bytes.getvalue())

        if unmatched_df is not None and not unmatched_df.empty:
            unmatched_bytes = io.BytesIO()
            unmatched_df.to_excel(unmatched_bytes, index=False)
            zipf.writestr("unmatched_results.xlsx", unmatched_bytes.getvalue())

        if summary_df is not None and not summary_df.empty:
            summary_bytes = io.BytesIO()
            summary_df.to_excel(summary_bytes, index=False)
            zipf.writestr("processing_summary.xlsx", summary_bytes.getvalue())

    buffer.seek(0)
    return buffer

# --------------------------
# Streamlit UI
# --------------------------

def main():
    st.set_page_config(page_title="Resistance Code Parser", layout="wide")

    tab_main, tab_template = st.tabs(["🔍 Parser Tool", "📥 Download Template"])

    # ---------------- TEMPLATE TAB ----------------
    with tab_template:
        st.subheader("📥 Download Input Template")
        template_df = pd.DataFrame({"PartNumber": ["ABC123", "XYZ456", "10.228 (5.8mm)"]})
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            template_df.to_excel(writer, index=False, sheet_name="Template")
        buffer.seek(0)
        st.download_button(
            label="⬇️ Download Template Excel",
            data=buffer,
            file_name="template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ---------------- MAIN TOOL TAB ----------------
    with tab_main:
        st.title("⚡ Enhanced Resistance Code Parser")
        st.markdown("Upload an Excel file with resistance part numbers to extract and match resistance values.")

        uploaded_file = st.file_uploader("📤 Upload Excel file", type=["xlsx"])

        if uploaded_file:
            df, matched_df, unmatched_df = process_file(uploaded_file)

            if df is not None:
                st.success("✅ File processed successfully!")

                st.subheader("📊 Extracted Patterns")
                st.dataframe(df)

                if matched_df is not None and not matched_df.empty:
                    st.subheader("✅ Matched Results")
                    st.dataframe(matched_df)

                if unmatched_df is not None and not unmatched_df.empty:
                    st.subheader("⚠️ Unmatched Results")
                    st.dataframe(unmatched_df)

                # Summary info
                summary_df = pd.DataFrame([{
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Total Records": len(df),
                    "Matched": len(matched_df),
                    "Unmatched": len(unmatched_df)
                }])

                # ZIP download
                zip_buffer = create_download_zip(matched_df, unmatched_df, summary_df)
                st.download_button(
                    label="📦 Download All Results (ZIP)",
                    data=zip_buffer,
                    file_name="resistance_parser_results.zip",
                    mime="application/zip"
                )

    # ---------------- SIDEBAR TEST ----------------
    st.sidebar.header("🧪 Test Pattern Extraction")
    test_input = st.sidebar.text_input("Enter Part Number", "10.228 (5.8mm)")
    if test_input:
        st.sidebar.write("🔹 Extracted Pattern:", extract_pattern(test_input))


if __name__ == "__main__":
    main()
