import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Excel Key Matcher", layout="wide")

st.title("🔍 Excel Key Matcher")
tab1, tab2 = st.tabs(["🔍 Matcher", "📥 Templates"])

with tab1:
    # =========================
    # Functions
    # =========================
    def normalize_key(value):
        return re.sub(r"[^A-Za-z0-9]", "", str(value)).upper()
    
    def process_files(file1, file2):
        df1 = pd.read_excel(file1, dtype=str).fillna("")
        df2 = pd.read_excel(file2, dtype=str).fillna("")
    
        required_file1 = ["OtherPartNumber", "PartNumber"]
        required_file2 = ["Digikey_Part_Number", "Manufacturer_Part_Number"]
    
        missing1 = [c for c in required_file1 if c not in df1.columns]
        missing2 = [c for c in required_file2 if c not in df2.columns]
    
        if missing1:
            raise ValueError(
                f"First file is missing columns: {', '.join(missing1)}"
            )
    
        if missing2:
            raise ValueError(
                f"Second file is missing columns: {', '.join(missing2)}"
            )
    
        # Exact Keys
        df1["Key"] = (
            df1["OtherPartNumber"].astype(str).str.strip()
            + "|"
            + df1["PartNumber"].astype(str).str.strip()
        )
    
        df2["Key"] = (
            df2["Digikey_Part_Number"].astype(str).str.strip()
            + "|"
            + df2["Manufacturer_Part_Number"].astype(str).str.strip()
        )
    
        # Non-Alpha Keys
        df1["Key_NonAlpha"] = df1["Key"].apply(normalize_key)
        df2["Key_NonAlpha"] = df2["Key"].apply(normalize_key)
    
        exact_keys = set(df2["Key"])
        nonalpha_keys = set(df2["Key_NonAlpha"])
    
        def get_status(row):
            if row["Key"] in exact_keys:
                return "FoundExact"
            elif row["Key_NonAlpha"] in nonalpha_keys:
                return "Exact-nonAlpha"
            else:
                return "Not Exact"
    
        df1["Status"] = df1.apply(get_status, axis=1)
    
        # Summary
        summary = df1["Status"].value_counts().to_dict()
    
        # Remove helper columns if you don't want them
        df1.drop(columns=["Key_NonAlpha"], inplace=True)
    
        return df1, summary
    
    def to_excel(df):
        output = BytesIO()
    
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Results", index=False)
    
        output.seek(0)
        return output
    
    # =========================
    # UI
    # =========================
    col1, col2 = st.columns(2)
    
    with col1:
        file1 = st.file_uploader(
            "📄 Upload First Excel File",
            type=["xlsx", "xls"],
            key="file1"
        )
    
    with col2:
        file2 = st.file_uploader(
            "📄 Upload Second Excel File",
            type=["xlsx", "xls"],
            key="file2"
        )
    
    if file1 and file2:
    
        if st.button("🚀 Start Matching", type="primary"):
    
            try:
                with st.spinner("Processing files..."):
    
                    result_df, summary = process_files(file1, file2)
    
                    st.success("✅ Processing completed")
    
                    # Metrics
                    c1, c2, c3 = st.columns(3)
    
                    c1.metric(
                        "FoundExact",
                        summary.get("FoundExact", 0)
                    )
    
                    c2.metric(
                        "Exact-nonAlpha",
                        summary.get("Exact-nonAlpha", 0)
                    )
    
                    c3.metric(
                        "Not Exact",
                        summary.get("Not Exact", 0)
                    )
    
                    st.subheader("Preview")
                    st.dataframe(
                        result_df.head(100),
                        use_container_width=True
                    )
    
                    excel_data = to_excel(result_df)
    
                    st.download_button(
                        label="⬇️ Download Result",
                        data=excel_data,
                        file_name="Matched_Result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    
            except Exception as e:
                st.error(f"❌ Error: {e}")
    with tab2:

    st.subheader("Download Templates")

    # Template 1
    template1 = pd.DataFrame(
        columns=[
            "OtherPartNumber",
            "PartNumber"
        ]
    )

    # Template 2
    template2 = pd.DataFrame(
        columns=[
            "Digikey_Part_Number",
            "Manufacturer_Part_Number"
        ]
    )

    def dataframe_to_excel(df):
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)

        output.seek(0)
        return output

    st.markdown("### First File Template")

    st.download_button(
        label="📥 Download First File Template",
        data=dataframe_to_excel(template1),
        file_name="Template_File1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("Required Columns:")
    st.code("OtherPartNumber\nPartNumber")

    st.divider()

    st.markdown("### Second File Template")

    st.download_button(
        label="📥 Download Second File Template",
        data=dataframe_to_excel(template2),
        file_name="Template_File2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("Required Columns:")
    st.code("Digikey_Part_Number\nManufacturer_Part_Number")
