import streamlit as st
import pandas as pd
import numpy as np
import os
import time
import io
from datetime import datetime
import gc
import zipfile
import tempfile

# Page configuration
st.set_page_config(
    page_title="Excel Mask Processing Tool",
    page_icon="📊",
    layout="wide"
)

def apply_two_stage_replacement(df_row, port_key, replacement='__'):
    """
    Apply two-stage replacement logic:
    1. If 2+ occurrences: replace only 2nd occurrence and mark as MultiPortion
    2. If 1 occurrence: replace the single occurrence
    """
    original_part = str(df_row['ZPartNumber'])
    if pd.isna(original_part) or original_part == '' or original_part == 'nan':
        return original_part, '0', False

    # Find all occurrences of the port_key
    occurrences = []
    start = 0
    while True:
        pos = original_part.find(str(port_key), start)
        if pos == -1:
            break
        occurrences.append(pos)
        start = pos + 1

    if len(occurrences) >= 2:
        # Stage 1: Multiple occurrences - replace only the 2nd occurrence
        second_occurrence_index = occurrences[1]
        masked_part = (original_part[:second_occurrence_index] +
                      replacement +
                      original_part[second_occurrence_index + len(str(port_key)):])
        return masked_part, second_occurrence_index, True  # True indicates MultiPortion

    elif len(occurrences) == 1:
        # Stage 2: Single occurrence - replace it
        first_occurrence_index = occurrences[0]
        masked_part = original_part.replace(str(port_key), replacement, 1)  # Replace only first occurrence
        return masked_part, first_occurrence_index, False  # False indicates single portion

    else:
        # No occurrences found
        return original_part, '0', False

def process_excel_data(uploaded_file):
    """Process the uploaded Excel file with the mask replacement logic"""
    
    try:
        # Load Excel sheets
        with pd.ExcelFile(uploaded_file, engine='openpyxl') as excel_file:
            df_portion = pd.read_excel(excel_file, sheet_name='PC_portion')
            df_B1 = pd.read_excel(excel_file, sheet_name='B1')
            df_B2 = pd.read_excel(excel_file, sheet_name='B2')
        
        # Store dataframes in a dictionary for easier processing
        dataframes = {'B1': df_B1, 'B2': df_B2}
        
        # Ensure required columns exist for both sheets
        for sheet_name, df in dataframes.items():
            for col in ['masked_code', 'maskMatch', 'PortionStart']:
                if col not in df.columns:
                    df[col] = np.nan
                    st.info(f"Added column '{col}' to {sheet_name}")
        
        # Clean values with vectorized operations
        st.info("🧹 Cleaning data...")
        
        for col in ['PortionName', 'Family', 'PortionKey', 'FamilyName', 'ZPartNumber']:
            if col in df_portion.columns:
                df_portion[col] = df_portion[col].fillna('').astype(str).str.strip()
            for sheet_name, df in dataframes.items():
                if col in df.columns:
                    df[col] = df[col].fillna('').astype(str).str.strip()
        
        # Keep Part Mask column unchanged - only clean it without removing underscores
        for sheet_name, df in dataframes.items():
            if 'Part Mask' in df.columns:
                df['Part Mask'] = df['Part Mask'].fillna('').astype(str).str.strip()
        
        # Filter Packaging rows
        df_packaging = df_portion[df_portion['PortionName'] == 'Packaging']
        
        # Get distinct families
        family_list = df_packaging['Family'].unique().tolist()
        total_families = len(family_list)
        
        st.info(f"📋 Found {total_families} families to process for both B1 and B2 sheets")
        
        # Create progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Process each family for both sheets
        for i, family in enumerate(family_list, 1):
            status_text.text(f"🔄 Processing Family: {family} ({i}/{total_families})")
            progress_bar.progress(i / total_families)
            
            # Get PortionKeys for this family sorted by length descending
            portKeys = df_packaging[df_packaging['Family'] == family]['PortionKey'].unique().tolist()
            portKeys = [pk for pk in portKeys if pk and str(pk).strip() and str(pk) != 'nan']
            portKeys = sorted(portKeys, key=lambda x: len(str(x)), reverse=True)
            
            # Process both B1 and B2 sheets for this family
            for sheet_name, df in dataframes.items():
                # Filter rows for this family
                mask_family = df['FamilyName'] == family
                family_row_count = mask_family.sum()
                
                if family_row_count == 0:
                    continue
                
                # Apply each portKey with enhanced two-stage logic
                processed_rows = 0
                multi_portion_count = 0
                
                for portKey in portKeys:
                    if not portKey or str(portKey).strip() == '' or str(portKey) == 'nan':
                        continue
                    
                    # Mask for rows containing portKey and empty masked_code
                    mask_key = (mask_family &
                               df['ZPartNumber'].astype(str).str.contains(str(portKey), na=False) &
                               (df['masked_code'].isna() | (df['masked_code'] == '') | (df['masked_code'] == 'nan')))
                    
                    matched_rows = mask_key.sum()
                    if matched_rows > 0:
                        # Apply two-stage replacement logic
                        for row_idx in df[mask_key].index:
                            masked_code, portion_start, is_multi_portion = apply_two_stage_replacement(
                                df.loc[row_idx], portKey
                            )
                            
                            # Set the values
                            df.loc[row_idx, 'masked_code'] = masked_code
                            
                            # Set PortionStart based on the logic
                            if is_multi_portion:
                                df.loc[row_idx, 'PortionStart'] = 'MultiPortion'
                                multi_portion_count += 1
                            else:
                                df.loc[row_idx, 'PortionStart'] = portion_start
                            
                            # Set maskMatch based on comparison with Part Mask
                            part_mask_val = str(df.loc[row_idx, 'Part Mask'])
                            if str(masked_code) == part_mask_val:
                                df.loc[row_idx, 'maskMatch'] = 'maskMatch'
                            else:
                                df.loc[row_idx, 'maskMatch'] = 'NotMatch'
                        
                        processed_rows += matched_rows
                
                # Handle rows where no portion was found
                no_portion_mask = mask_family & (df['masked_code'].isna() | (df['masked_code'] == '') | (df['masked_code'] == 'nan'))
                
                if no_portion_mask.any():
                    no_portion_count = no_portion_mask.sum()
                    
                    # Set masked_code to original ZPartNumber
                    df.loc[no_portion_mask, 'masked_code'] = df.loc[no_portion_mask, 'ZPartNumber']
                    
                    # Set PortionStart to 'NoPacking'
                    df.loc[no_portion_mask, 'PortionStart'] = 'NoPacking'
                    
                    # Set maskMatch based on comparison with Part Mask
                    for row_idx in df[no_portion_mask].index:
                        part_number = str(df.loc[row_idx, 'ZPartNumber'])
                        part_mask_val = str(df.loc[row_idx, 'Part Mask'])
                        
                        if part_number == part_mask_val:
                            df.loc[row_idx, 'maskMatch'] = 'maskMatch'
                        else:
                            df.loc[row_idx, 'maskMatch'] = 'NotMatch'
            
            # Memory cleanup
            gc.collect()
        
        progress_bar.progress(1.0)
        status_text.text("✅ Processing completed successfully!")
        
        return dataframes, df_portion
        
    except Exception as e:
        st.error(f"❌ Error occurred during processing: {str(e)}")
        return None, None

def create_template_file():
    """Create a template Excel file for download"""
    
    # Create sample data for template
    pc_portion_data = {
        'PortionName': ['Packaging', 'Packaging', 'Packaging'],
        'Family': ['SampleFamily1', 'SampleFamily1', 'SampleFamily2'],
        'PortionKey': ['CF1/4CT', 'R51J', 'ABC123'],
        'FamilyName': ['SampleFamily1', 'SampleFamily1', 'SampleFamily2']
    }
    
    b1_data = {
        'ZPartNumber': ['CF1/4CT52RR51J', 'ABC123XYZ', 'DEF456GHI'],
        'FamilyName': ['SampleFamily1', 'SampleFamily2', 'SampleFamily1'],
        'Part Mask': ['CF1/4CT52____51J', 'ABC123___', 'DEF456___'],
        'masked_code': ['', '', ''],
        'maskMatch': ['', '', ''],
        'PortionStart': ['', '', '']
    }
    
    b2_data = {
        'ZPartNumber': ['CF1/4CT52RR51J', 'ABC123XYZ', 'DEF456GHI'],
        'FamilyName': ['SampleFamily1', 'SampleFamily2', 'SampleFamily1'],
        'Part Mask': ['CF1/4CT52____51J', 'ABC123___', 'DEF456___'],
        'masked_code': ['', '', ''],
        'maskMatch': ['', '', ''],
        'PortionStart': ['', '', '']
    }
    
    # Create DataFrames
    df_pc_portion = pd.DataFrame(pc_portion_data)
    df_b1 = pd.DataFrame(b1_data)
    df_b2 = pd.DataFrame(b2_data)
    
    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_pc_portion.to_excel(writer, sheet_name='PC_portion', index=False)
        df_b1.to_excel(writer, sheet_name='B1', index=False)
        df_b2.to_excel(writer, sheet_name='B2', index=False)
    
    output.seek(0)
    return output

def main():
    st.title("📊 Excel Mask Processing Tool")
    st.markdown("---")
    
    # Create tabs
    tab1, tab2 = st.tabs(["🔄 Process File", "📥 Download Template"])
    
    with tab1:
        st.header("Process Excel File")
        st.markdown("Upload your Excel file with the required sheets (PC_portion, B1, B2) to process mask replacements.")
        
        # File uploader
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="The file should contain sheets: PC_portion, B1, and B2"
        )
        
        if uploaded_file is not None:
            st.success(f"✅ File uploaded: {uploaded_file.name}")
            
            # Display file info
            file_details = {
                "Filename": uploaded_file.name,
                "File size": f"{uploaded_file.size / 1024:.2f} KB",
                "File type": uploaded_file.type
            }
            
            col1, col2 = st.columns(2)
            with col1:
                st.json(file_details)
            
            # Process button
            if st.button("🚀 Start Processing", type="primary"):
                with st.spinner("Processing file... This may take a few minutes for large files."):
                    start_time = time.time()
                    
                    # Process the file
                    processed_dataframes, df_portion = process_excel_data(uploaded_file)
                    
                    if processed_dataframes is not None:
                        processing_time = time.time() - start_time
                        st.success(f"✅ Processing completed in {processing_time:.2f} seconds")
                        
                        # Display summary statistics
                        st.markdown("### 📊 Processing Summary")
                        
                        col1, col2 = st.columns(2)
                        
                        for i, (sheet_name, df) in enumerate(processed_dataframes.items()):
                            col = col1 if i == 0 else col2
                            
                            with col:
                                st.subheader(f"Sheet: {sheet_name}")
                                
                                total_rows = len(df)
                                processed_rows = len(df[df['masked_code'].notna() & 
                                                   (df['masked_code'] != '') & 
                                                   (df['masked_code'] != 'nan')])
                                matches = len(df[df['maskMatch'] == 'maskMatch'])
                                not_matches = len(df[df['maskMatch'] == 'NotMatch'])
                                multi_portion_rows = len(df[df['PortionStart'] == 'MultiPortion'])
                                no_portion_rows = len(df[df['PortionStart'] == 'NoPacking'])
                                
                                metrics_data = {
                                    "Total Rows": total_rows,
                                    "Processed Rows": processed_rows,
                                    "Exact Matches": matches,
                                    "Non-matches": not_matches,
                                    "MultiPortion": multi_portion_rows,
                                    "NoPacking": no_portion_rows
                                }
                                
                                for metric, value in metrics_data.items():
                                    if metric == "Processed Rows":
                                        percentage = f"({value/total_rows*100:.1f}%)" if total_rows > 0 else ""
                                        st.metric(metric, f"{value} {percentage}")
                                    else:
                                        st.metric(metric, value)
                        
                        # Sample data display
                        st.markdown("### 📋 Sample Processed Data")
                        
                        for sheet_name, df in processed_dataframes.items():
                            with st.expander(f"Sample from {sheet_name}"):
                                # Show MultiPortion samples
                                multi_sample = df[df['PortionStart'] == 'MultiPortion'].head(3)
                                if len(multi_sample) > 0:
                                    st.markdown("**MultiPortion Rows:**")
                                    st.dataframe(multi_sample[['ZPartNumber', 'FamilyName', 'Part Mask', 'masked_code', 'maskMatch', 'PortionStart']])
                                
                                # Show Single Portion samples
                                single_sample = df[pd.to_numeric(df['PortionStart'], errors='coerce').notna()].head(3)
                                if len(single_sample) > 0:
                                    st.markdown("**Single Portion Rows:**")
                                    st.dataframe(single_sample[['ZPartNumber', 'FamilyName', 'Part Mask', 'masked_code', 'maskMatch', 'PortionStart']])
                        
                        # Download processed file
                        st.markdown("### 💾 Download Processed File")
                        
                        # Create processed file
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_portion.to_excel(writer, sheet_name='PC_portion', index=False)
                            processed_dataframes['B1'].to_excel(writer, sheet_name='B1', index=False)
                            processed_dataframes['B2'].to_excel(writer, sheet_name='B2', index=False)
                        
                        output.seek(0)
                        
                        st.download_button(
                            label="📥 Download Processed File",
                            data=output.getvalue(),
                            file_name=f"processed_{uploaded_file.name.rsplit('.', 1)[0]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
    
    with tab2:
        st.header("Download Template File")
        st.markdown("Download a template Excel file to understand the required format and structure.")
        
        st.markdown("""
        ### 📋 Template Structure
        
        The template file contains three sheets:
        
        1. **PC_portion**: Contains portion definitions
           - `PortionName`: Should be 'Packaging' for mask processing
           - `Family`: Family identifier
           - `PortionKey`: The key to be replaced in part numbers
           - `FamilyName`: Family name for matching
        
        2. **B1**: First data sheet with part numbers to be processed
           - `ZPartNumber`: Original part number
           - `FamilyName`: Family name for matching
           - `Part Mask`: Expected masked result
           - `masked_code`: Will be populated after processing
           - `maskMatch`: Will show match status
           - `PortionStart`: Will show portion start position
        
        3. **B2**: Second data sheet (same structure as B1)
        """)
        
        # Create and offer template download
        template_file = create_template_file()
        
        st.download_button(
            label="📥 Download Template File",
            data=template_file.getvalue(),
            file_name=f"excel_mask_template_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.markdown("### ℹ️ Instructions")
        st.info("""
        1. Download the template file above
        2. Fill in your data following the same structure
        3. Make sure you have all three required sheets: PC_portion, B1, B2
        4. Upload your completed file in the "Process File" tab
        5. The tool will process the mask replacements according to your two-stage logic:
           - If 2+ occurrences of a portion key: replace only the 2nd occurrence
           - If 1 occurrence: replace that single occurrence
           - If 0 occurrences: mark as 'NoPacking'
        """)

if __name__ == "__main__":
    main()
