import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

def normalize_pin_group(pin: str) -> str:
    """
    Normalize pin names into logical groups.
    Works for VIN, VOUT, GND, etc.
    """
    pin = str(pin).upper().strip()

    # +PREFIX / -PREFIX → PREFIX
    m = re.match(r"^[+-]?([A-Z]+)$", pin)
    if m:
        return m.group(1)

    # PREFIX + number → PREFIX
    m = re.match(r"^([A-Z]+)\d+$", pin)
    if m:
        return m.group(1)

    # Otherwise keep as-is (e.g. VINCOM, VOUTCOM, AVIN, etc.)
    return pin


def process_excel_data(df):
    """
    Process the uploaded dataframe and return processed results
    """
    # Create normalized pin group
    df["PinGroup"] = df["Pin Name"].apply(normalize_pin_group)

    # ---- EXACT GROUP (DataDefinition + Pin Name) ----
    exact_stats = (
        df.groupby(["DataDefinition", "Pin Name"])["Normalized Pin NAME"]
        .agg(["nunique", lambda x: len(set(x)) > 1])
        .reset_index()
        .rename(columns={"nunique": "CountExact", "<lambda_0>": "DiffExact"})
    )

    # ---- SIMILARITY GROUP (DataDefinition + PinGroup) ----
    sim_stats = (
        df.groupby(["DataDefinition", "PinGroup"])["Normalized Pin NAME"]
        .agg(["nunique", lambda x: len(set(x)) > 1])
        .reset_index()
        .rename(columns={"nunique": "CountSim", "<lambda_0>": "DiffSim"})
    )

    # ---- Merge back to main ----
    final_df = df.merge(exact_stats, on=["DataDefinition", "Pin Name"], how="left")
    final_df = final_df.merge(sim_stats, on=["DataDefinition", "PinGroup"], how="left")

    return final_df


def summarize_all_normalized(df):
    """
    Summarize PartsCount for all Normalized Pin NAME values within
    each (DataDefinition, PinGroup). Produces percentages for each.
    Keeps original Normalized Pin NAME spellings in output.
    """
    # Keep original and add lowercase for comparison
    df["Norm_lower"] = df["Normalized Pin NAME"].astype(str).str.lower()

    # Filter rows where CountExact != 1
    df_filtered = df[df["CountExact"] != 1]

    results = []

    # Loop by DataDefinition
    for dd, group_dd in df_filtered.groupby("DataDefinition"):
        # Loop by PinGroup
        for pg, group_pg in group_dd.groupby("PinGroup"):

            # Sum PartsCount for each lowercase Normalized Pin NAME
            norm_sums = group_pg.groupby("Norm_lower")["PartsCount"].sum()

            total = norm_sums.sum()

            for norm_lower, cnt in norm_sums.items():
                perc = (cnt / total * 100) if total > 0 else 0

                # Pick one representative original spelling (first one in group)
                original_name = (
                    group_pg.loc[group_pg["Norm_lower"] == norm_lower, "Normalized Pin NAME"]
                    .iloc[0]
                )

                results.append({
                    "DataDefinition": dd,
                    "SumCountExact": total,   # total for this group
                    "PinGroup": pg,
                    "Normalized Pin NAME": original_name,  # back to original spelling
                    "Percentage": round(perc, 2),
                    "PartsCount": cnt
                })

    summary_df = pd.DataFrame(results)
    return summary_df


def to_excel_bytes(df1, df2, sheet1_name="Processed Data", sheet2_name="Summary"):
    """
    Convert dataframes to Excel bytes for download
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name=sheet1_name, index=False)
        df2.to_excel(writer, sheet_name=sheet2_name, index=False)
    processed_data = output.getvalue()
    return processed_data


# Streamlit App
def main():
    st.set_page_config(
        page_title="Pin Analysis Tool",
        page_icon="🔌",
        layout="wide"
    )

    st.title("🔌 Pin Analysis Tool")
    st.markdown("Upload your Excel file to analyze pin data and generate summaries")

    # File upload
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload your PINOUT Excel file for analysis"
    )

    if uploaded_file is not None:
        try:
            # Read the uploaded file
            with st.spinner('Loading data...'):
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ File loaded successfully! Shape: {df.shape}")
            
            # Show original data preview
            with st.expander("📊 Preview Original Data", expanded=False):
                st.dataframe(df.head(10))
                st.write(f"**Columns:** {list(df.columns)}")
                st.write(f"**Total rows:** {len(df)}")

            # Check for required columns
            required_columns = ["Pin Name", "Normalized Pin NAME", "DataDefinition", "PartsCount"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"❌ Missing required columns: {missing_columns}")
                st.stop()

            # Process data
            with st.spinner('Processing data...'):
                processed_df = process_excel_data(df)
                summary_df = summarize_all_normalized(processed_df)

            # Create columns for results
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("📈 Processed Data")
                st.write(f"**Rows:** {len(processed_df)}")
                with st.expander("View Processed Data", expanded=False):
                    st.dataframe(processed_df)

            with col2:
                st.subheader("📊 Summary Analysis")
                st.write(f"**Rows:** {len(summary_df)}")
                with st.expander("View Summary", expanded=True):
                    st.dataframe(summary_df)

            # Statistics
            st.subheader("📋 Analysis Statistics")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Pin Groups", processed_df["PinGroup"].nunique())
            
            with col2:
                st.metric("Data Definitions", processed_df["DataDefinition"].nunique())
            
            with col3:
                st.metric("Unique Pins", processed_df["Pin Name"].nunique())
            
            with col4:
                st.metric("Summary Entries", len(summary_df))

            # Charts
            if not summary_df.empty:
                st.subheader("📊 Data Visualization")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Top 10 Pin Groups by Parts Count**")
                    top_pingroups = summary_df.groupby("PinGroup")["PartsCount"].sum().sort_values(ascending=False).head(10)
                    st.bar_chart(top_pingroups)
                
                with col2:
                    st.write("**Data Definition Distribution**")
                    dd_dist = summary_df["DataDefinition"].value_counts().head(10)
                    st.bar_chart(dd_dist)

            # Download section
            st.subheader("💾 Download Results")
            
            # Generate Excel file
            excel_data = to_excel_bytes(processed_df, summary_df)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"pin_analysis_results_{timestamp}.xlsx"
            
            st.download_button(
                label="📥 Download Excel Results",
                data=excel_data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download both processed data and summary in Excel format"
            )

            # Individual CSV downloads
            col1, col2 = st.columns(2)
            
            with col1:
                processed_csv = processed_df.to_csv(index=False)
                st.download_button(
                    label="📄 Download Processed Data (CSV)",
                    data=processed_csv,
                    file_name=f"processed_data_{timestamp}.csv",
                    mime="text/csv"
                )
            
            with col2:
                summary_csv = summary_df.to_csv(index=False)
                st.download_button(
                    label="📄 Download Summary (CSV)",
                    data=summary_csv,
                    file_name=f"summary_data_{timestamp}.csv",
                    mime="text/csv"
                )

        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.write("Please check that your file has the required columns:")
            st.write("- Pin Name")
            st.write("- Normalized Pin NAME") 
            st.write("- DataDefinition")
            st.write("- PartsCount")

    else:
        st.info("👆 Please upload an Excel file to get started")
        
        # Show sample data format
        with st.expander("📝 Expected File Format", expanded=False):
            sample_data = pd.DataFrame({
                'Pin Name': ['VIN1', 'VOUT2', 'GND', '+VCC'],
                'Normalized Pin NAME': ['VIN_NORM', 'VOUT_NORM', 'GND_NORM', 'VCC_NORM'],
                'DataDefinition': ['Type A', 'Type A', 'Type B', 'Type B'],
                'PartsCount': [10, 15, 8, 12]
            })
            st.dataframe(sample_data)


if __name__ == "__main__":
    main()
