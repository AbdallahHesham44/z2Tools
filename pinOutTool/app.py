import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

def normalize_pin_group(pin: str) -> str:
    """
    Normalize pin names into logical groups.
    Removes digits, keeps letters, '-', and '#'.
    Preserves leading and trailing '-' or '#' if they match.
    """
    # Handle empty strings or whitespace-only strings
    if not pin or not pin.strip():
        return ""

    pin = str(pin).upper().strip()

    # Preserve matching start and end characters if both '-' or both '#'
    if (pin.startswith('-') and pin.endswith('-')) or (pin.startswith('#') and pin.endswith('#')):
        core = pin[1:-1]
        core = re.sub(r'\d', '', core)  # remove digits
        core = re.sub(r'[^A-Z#-]', '', core)  # keep only A-Z, #, -
        return pin[0] + core + pin[-1]

    # Otherwise: just remove digits, keep letters + - + #
    pin = re.sub(r'\d', '', pin)
    pin = re.sub(r'[^A-Z#-]', '', pin)
    return pin

def process_excel(df):
    """Process the uploaded dataframe"""
    # Ensure 'PartsCount' is numeric
    df['PartsCount'] = pd.to_numeric(df['PartsCount'], errors='coerce')
    df = df.dropna(subset=['PartsCount'])  # Drop rows where conversion failed

    # Create PinGroup column for grouping, but keep original Pin Name intact
    df["PinGroup"] = df["Pin Name"].apply(normalize_pin_group)

    # ---- EXACT GROUP (DataDefinition + Pin Name) ----
    # Use original Pin Name column, not modified version
    exact_stats = (
        df.groupby(["DataDefinition", "Pin Name"], dropna=False)["Normalized Pin NAME"]
        .agg(["nunique", lambda x: len(set(x)) > 1])
        .reset_index()
        .rename(columns={"nunique": "CountExact", "<lambda_0>": "DiffExact"})
    )

    # ---- SIMILARITY GROUP (DataDefinition + PinGroup) ----
    sim_stats = (
        df.groupby(["DataDefinition", "PinGroup"], dropna=False)["Normalized Pin NAME"]
        .agg(["nunique", lambda x: len(set(x)) > 1])
        .reset_index()
        .rename(columns={"nunique": "CountSim", "<lambda_0>": "DiffSim"})
    )

    # ---- Merge back to main ----
    final_df = df.merge(exact_stats,
                        on=["DataDefinition", "Pin Name"],
                        how="left")
    final_df = final_df.merge(sim_stats,
                              on=["DataDefinition", "PinGroup"],
                              how="left")

    return final_df

def summarize_all_normalized(df_proc):
    """
    Summarize PartsCount for all Normalized Pin NAME values within
    each (DataDefinition, PinGroup). Produces percentages for each.
    Adds a Status column.
    """
    # Ensure 'PartsCount' is numeric for calculations
    df_proc['PartsCount'] = pd.to_numeric(df_proc['PartsCount'], errors='coerce').fillna(0)

    # For calculations, normalize case but don't modify original data
    df_proc["Norm_lower"] = df_proc["Normalized Pin NAME"].astype(str).str.upper()

    results = []

    # Loop by DataDefinition + PinGroup
    for (dd, pg), group_pg in df_proc.groupby(["DataDefinition", "PinGroup"]):
        norm_sums = group_pg.groupby("Norm_lower")["PartsCount"].sum()
        total = norm_sums.sum()
        max_norm = norm_sums.idxmax() if not norm_sums.empty else None

        for norm_lower, cnt in norm_sums.items():
            perc = (cnt / total * 100) if total > 0 else 0
            # Find the first original Normalized Pin NAME for this group
            original_name_series = group_pg.loc[group_pg["Norm_lower"] == norm_lower, "Normalized Pin NAME"]
            original_name = original_name_series.iloc[0] if not original_name_series.empty else "Unknown"

            status = "seems Ok" if norm_lower == max_norm else "Conflict in same PL | Pin name"

            results.append({
                "DataDefinition": dd,
                "SumCountExact": total,
                "PinGroup": pg,
                "Normalized Pin NAME": original_name,
                "Percentage": round(perc, 2),
                "PartsCount": cnt,
                "Status": status
            })

    summary_df = pd.DataFrame(results)
    return summary_df

def create_template():
    """Create a template Excel file for users to download"""
    template_data = {
        "DataDefinition": [
            "ConnectorA_Type1",
            "ConnectorA_Type1", 
            "ConnectorB_Type2",
            "ConnectorB_Type2",
            "ConnectorC_Type3"
        ],
        "Pin Name": [
            "VCC12",
            "GND1",
            "DATA-1",
            "CLK#2",
            "RST"
        ],
        "Normalized Pin NAME": [
            "VCC",
            "GND",
            "DATA",
            "CLK",
            "RESET"
        ],
        "PartsCount": [
            100,
            150,
            75,
            80,
            45
        ]
    }
    
    template_df = pd.DataFrame(template_data)
    return template_df

def to_excel_bytes(df, sheet_name="Sheet1"):
    """Convert dataframe to Excel bytes for download"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

# Streamlit App
def main():
    st.set_page_config(page_title="Pin Analysis Tool", page_icon="🔌", layout="wide")
    
    st.title("🔌 Pin Analysis Tool")
    st.markdown("Analyze pin names and detect conflicts in connector definitions")
    
    # Create main tabs
    tab1, tab2, tab3 = st.tabs(["📤 Upload & Process", "📋 Instructions", "📄 Template & Download"])
    
    with tab1:
        # Main content area
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.header("📤 Upload Your File")
            uploaded_file = st.file_uploader(
                "Choose an Excel file",
                type=['xlsx', 'xls'],
                help="Upload your Excel file with pin data. Use the template format from the Template tab."
            )
        
        with col2:
            st.header("⚙️ Processing Options")
            show_preview = st.checkbox("Show data preview", value=True)
            show_statistics = st.checkbox("Show processing statistics", value=True)

        if uploaded_file is not None:
            try:
                # Load the uploaded file
                df = pd.read_excel(uploaded_file, dtype=str, keep_default_na=False)
                
                # Validate required columns
                required_columns = ["DataDefinition", "Pin Name", "Normalized Pin NAME", "PartsCount"]
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    st.error(f"Missing required columns: {', '.join(missing_columns)}")
                    st.info("Please use the template file format from the Template tab.")
                    return
                
                st.success(f"✅ File uploaded successfully! Found {len(df)} rows.")
                
                if show_preview:
                    st.subheader("📊 Data Preview")
                    st.dataframe(df.head(10), use_container_width=True)
                
                # Process the data
                with st.spinner("Processing data..."):
                    processed_df = process_excel(df)
                    summary_df = summarize_all_normalized(processed_df)
                
                # Update main dataframe with status and percentage
                merged_df = processed_df.merge(
                    summary_df[["DataDefinition", "PinGroup", "Normalized Pin NAME", "Percentage", "Status"]],
                    on=["DataDefinition", "PinGroup", "Normalized Pin NAME"],
                    how="left"
                )
                merged_df["Status"] = merged_df["Status"].fillna("Case Sensitive Different")
                
                st.success("✅ Processing completed!")
                
                # Show statistics
                if show_statistics:
                    st.subheader("📈 Processing Statistics")
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Total Rows", len(merged_df))
                    with col2:
                        st.metric("Unique DataDefinitions", merged_df["DataDefinition"].nunique())
                    with col3:
                        st.metric("Unique Pin Groups", merged_df["PinGroup"].nunique())
                    with col4:
                        conflicts = (merged_df["Status"] == "Conflict in same PL | Pin name").sum()
                        st.metric("Conflicts Detected", conflicts)
                
                # Display results in tabs
                tab1a, tab2a, tab3a = st.tabs(["📋 Processed Data", "📊 Summary", "⚠️ Conflicts"])
                
                with tab1a:
                    st.subheader("Processed Data with Status")
                    st.dataframe(merged_df, use_container_width=True)
                    
                    # Download processed data
                    processed_bytes = to_excel_bytes(merged_df, "Processed_Data")
                    st.download_button(
                        label="📥 Download Processed Data",
                        data=processed_bytes,
                        file_name=f"processed_pin_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with tab2a:
                    st.subheader("Summary by Pin Groups")
                    st.dataframe(summary_df, use_container_width=True)
                    
                    # Download summary
                    summary_bytes = to_excel_bytes(summary_df, "Summary")
                    st.download_button(
                        label="📥 Download Summary",
                        data=summary_bytes,
                        file_name=f"pin_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with tab3a:
                    st.subheader("Detected Conflicts")
                    conflicts_df = merged_df[merged_df["Status"] == "Conflict in same PL | Pin name"]
                    
                    if len(conflicts_df) > 0:
                        st.warning(f"Found {len(conflicts_df)} rows with conflicts")
                        st.dataframe(conflicts_df, use_container_width=True)
                        
                        # Show conflict summary
                        st.subheader("Conflict Summary by DataDefinition")
                        conflict_summary = conflicts_df.groupby("DataDefinition").agg({
                            "Pin Name": "count",
                            "PinGroup": "nunique"
                        }).rename(columns={"Pin Name": "Conflict_Count", "PinGroup": "Affected_Groups"})
                        st.dataframe(conflict_summary, use_container_width=True)
                    else:
                        st.success("🎉 No conflicts detected!")
                
                # Additional insights
                st.subheader("📊 Data Insights")
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**Top 5 Pin Groups by Frequency:**")
                    pin_group_counts = merged_df["PinGroup"].value_counts().head()
                    for pin_group, count in pin_group_counts.items():
                        st.write(f"• {pin_group}: {count} occurrences")
                
                with col2:
                    st.markdown("**Data Definitions Overview:**")
                    dd_counts = merged_df["DataDefinition"].value_counts()
                    for dd, count in dd_counts.items():
                        st.write(f"• {dd}: {count} pins")
                        
            except Exception as e:
                st.error(f"Error processing file: {str(e)}")
                st.info("Please check that your file matches the template format from the Template tab.")
    
    with tab2:
        st.header("📋 How to Use This Tool")
        
        st.markdown("""
        ## 🎯 Purpose
        This tool analyzes pin naming conventions in connector definitions to detect conflicts and inconsistencies.
        
        ## 📝 Step-by-Step Instructions
        
        ### Step 1: Prepare Your Data
        1. Organize your pin data in Excel format (.xlsx or .xls)
        2. Ensure your file has the required columns (see Template tab)
        3. Make sure PartsCount column contains only numeric values
        
        ### Step 2: Upload and Process
        1. Go to the "Upload & Process" tab
        2. Click "Browse files" and select your Excel file
        3. Enable preview options if desired
        4. The tool will automatically validate and process your data
        
        ### Step 3: Review Results
        - **Processed Data**: View your original data with added analysis columns
        - **Summary**: See aggregated statistics by pin groups
        - **Conflicts**: Review any detected naming conflicts
        
        ### Step 4: Download Results
        - Download processed data with status indicators
        - Download summary report for analysis
        - Use timestamp-based filenames to avoid overwrites
        """)
        
        st.markdown("""
        ## 🔍 What the Tool Does
        
        ### Pin Group Normalization
        The tool groups similar pin names by:
        - Converting to uppercase
        - Removing all digits (0-9)
        - Keeping only letters (A-Z), hyphens (-), and hash symbols (#)
        - Preserving matching start/end characters
        
        ### Conflict Detection
        Identifies when multiple normalized pin names exist within the same:
        - DataDefinition (connector type)
        - PinGroup (normalized group)
        
        ### Status Assignment
        - **"seems Ok"**: Pin name is the dominant version in its group
        - **"Conflict in same PL | Pin name"**: Multiple conflicting pin names detected
        - **"Case Sensitive Different"**: Case sensitivity differences found
        """)
        
        st.markdown("""
        ## 📊 Understanding Results
        
        ### Processed Data Columns
        - **PinGroup**: Normalized grouping for similar pins
        - **CountExact**: Count of exact pin name matches
        - **DiffExact**: Boolean indicating if differences exist in exact matches
        - **CountSim**: Count of similar pin names in the group
        - **DiffSim**: Boolean indicating if differences exist in similar names
        - **Status**: Conflict status for each pin
        - **Percentage**: Percentage of parts count within the pin group
        
        ### Summary Report
        - Shows aggregated data by DataDefinition and PinGroup
        - Includes total parts count and percentage breakdown
        - Helps identify the most common pin naming patterns
        """)
        
        st.markdown("""
        ## ⚠️ Common Issues & Solutions
        
        | Issue | Solution |
        |-------|----------|
        | Missing required columns | Use the template file format |
        | Non-numeric PartsCount | Ensure all PartsCount values are numbers |
        | Empty file error | Check that your Excel file contains data rows |
        | File format error | Save as .xlsx or .xls format |
        | High conflict count | Review pin naming conventions for consistency |
        """)
    
    with tab3:
        st.header("📄 Template File & Downloads")
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.subheader("📥 Download Template")
            st.markdown("Download the Excel template to get started:")
            
            template_df = create_template()
            template_bytes = to_excel_bytes(template_df, "Template")
            
            st.download_button(
                label="📄 Download Template File",
                data=template_bytes,
                file_name="pin_analysis_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.info("💡 This template includes sample data to show the exact format required.")
        
        with col2:
            st.subheader("📋 Required Format")
            st.markdown("Your Excel file must contain these columns:")
            
            required_cols = pd.DataFrame({
                "Column Name": ["DataDefinition", "Pin Name", "Normalized Pin NAME", "PartsCount"],
                "Data Type": ["Text", "Text", "Text", "Numeric"],
                "Description": [
                    "Connector type/definition identifier",
                    "Original pin name (may contain numbers)",
                    "Normalized pin name (without numbers)",
                    "Number of parts (must be numeric)"
                ]
            })
            st.dataframe(required_cols, use_container_width=True, hide_index=True)
        
        st.subheader("📊 Template Preview")
        st.markdown("Here's what the template file contains:")
        
        template_df = create_template()
        st.dataframe(template_df, use_container_width=True, hide_index=True)
        
        st.subheader("🔄 Example Transformations")
        st.markdown("See how pin names are normalized:")
        
        examples = pd.DataFrame({
            "Original Pin Name": ["VCC12", "GND1", "DATA-1", "CLK#2", "-RST3-", "#PWR5#"],
            "Normalized Group": ["VCC", "GND", "DATA-", "CLK#", "-RST-", "#PWR#"],
            "Explanation": [
                "Removes digits, keeps letters",
                "Removes digits, keeps letters", 
                "Removes digits, preserves trailing dash",
                "Removes digits, preserves hash symbol",
                "Removes digits, preserves matching dashes",
                "Removes digits, preserves matching hash symbols"
            ]
        })
        st.dataframe(examples, use_container_width=True, hide_index=True)
        
        st.markdown("""
        ### 📝 Notes:
        - Pin names are case-insensitive for grouping
        - Original pin names are preserved exactly as entered
        - Empty or whitespace-only pin names are handled gracefully
        - Special characters like '-' and '#' are preserved strategically
        """)

if __name__ == "__main__":
    main()
