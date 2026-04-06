# --- START OF FILE app.py ---

#############################################
# STREAMLIT SCRIPT - UNIFIED PIPELINE APP
# Main application file for UI and pipeline orchestration.
#############################################

import streamlit as st
import pandas as pd
import os
import gc
from io import BytesIO
import io # Make sure io is imported
import traceback # Import traceback for better error logging

# Import functions from custom modules
from github_utils import download_mapping_file_from_github, update_mapping_file_on_github
from mapping_utils import save_mapping_to_disk
# Import the NEW preprocessor function
from preprocessor import preprocess_input_file
# Import the existing pipeline functions
from fixed_pipeline import process_fixed_pipeline_bytes
from pipeline_wrappers import run_fixed_pipeline_with_prefix_support, run_detailed_analysis_with_prefix_support
from detailed_pipeline import detailed_analysis
from result_combiner import combine_results

#############################################
# HIDE GITHUB ICON & OTHER ELEMENTS
#############################################
hide_button = """
    <style>
    [data-testid="stBaseButton-header"] {
        display: none;
    }
    </style>
    """
st.markdown(hide_button, unsafe_allow_html=True)


#############################################
# STREAMLIT APP UI & LOGIC
#############################################

st.title("ACC Project - Unified Pipeline")

# Define mapping file name constant
MAPPING_FILENAME = "mapping.xlsx"

# --- Initialize Mapping File ---
# (Keep this section as is)
if "mapping_df" not in st.session_state or st.session_state.get("mapping_df") is None:
    st.write("DEBUG: mapping_df not in session state or is None. Attempting load...")
    mapping_data = None

    github_config = st.secrets.get("github", {}) if hasattr(st, "secrets") else {}
    github_config_valid = all(
        github_config.get(key) for key in ["token", "owner", "repo", "file_path"]
    )

    if github_config_valid:
        st.write("DEBUG: GitHub secrets appear configured, attempting download...")
        try:
            mapping_data = download_mapping_file_from_github()
            if mapping_data is not None:
                st.session_state["mapping_df"] = mapping_data
                save_mapping_to_disk(st.session_state["mapping_df"], MAPPING_FILENAME)
                st.success("Mapping file downloaded and saved locally from GitHub.")
        except Exception as e:
            st.warning(f"GitHub download failed: {e}")

    if mapping_data is None:
        st.write("DEBUG: Falling back to local mapping file.")
        try:
            if os.path.exists(MAPPING_FILENAME):
                st.session_state["mapping_df"] = pd.read_excel(MAPPING_FILENAME)
                st.warning("Loaded mapping file from local mapping.xlsx.")
            else:
                if not github_config_valid:
                    st.error("GitHub secrets are not configured. Please add valid GitHub secrets or provide a local mapping.xlsx file.")
                else:
                    st.error("Failed to load mapping file from GitHub and local fallback mapping.xlsx does not exist.")
                st.session_state["mapping_df"] = None
                st.stop()
        except Exception as e2:
            st.error(f"Error loading local mapping file: {e2}")
            st.session_state["mapping_df"] = None
            st.stop()

# --- Validate Mapping File (after potential download) ---
# (Keep this section as is)
if st.session_state.get("mapping_df") is None:
    st.error("Mapping data could not be loaded. Please check GitHub configuration/connection and refresh.")
    if st.button("Retry Download Mapping"):
        if "mapping_df" in st.session_state: del st.session_state["mapping_df"]
        st.rerun()
    st.stop()
else:
    required_cols = {"Base Unit Symbol", "Multiplier Symbol"}
    current_mapping_df = st.session_state["mapping_df"]
    if not required_cols.issubset(current_mapping_df.columns):
        st.error(f"Mapping file from GitHub (or session state) is missing required columns: {required_cols}. Please fix the file via 'Manage Units' or directly on GitHub.")
        mapping_valid_for_processing = False
    else:
        mapping_valid_for_processing = True


# --- Main App Navigation ---
operation = st.selectbox("Select Operation", ["Get Pattern", "Manage Units"])

############################
# OPERATION: GET PATTERN
############################
if operation == "Get Pattern":
    st.header("Get Pattern")

    # Re-check mapping file validity specifically for this operation
    # (Keep this check as is)
    if not mapping_valid_for_processing:
         st.error("Cannot run 'Get Pattern' because the mapping file is missing required columns. Please check/fix it via 'Manage Units' or on GitHub.")
         st.stop()
    if not os.path.exists(MAPPING_FILENAME):
        st.error(f"Local mapping file '{MAPPING_FILENAME}' not found. Please ensure it was downloaded/saved (try refreshing or managing units).")
        if st.session_state.get("mapping_df") is not None:
            save_mapping_to_disk(st.session_state["mapping_df"], MAPPING_FILENAME)
            st.rerun()
        st.stop()


    st.write("Upload an Excel file containing a 'Value' column for processing.")

    # --- File Uploader ---
    input_file = st.file_uploader("Upload Input Excel File", type=["xlsx"], key="pattern_uploader")

    if input_file:
        # Define output filenames
        user_input_filename_temp = "user_input_temp_original.xlsx" # Temp file for ORIGINAL uploaded data
        preprocessed_input_filename_temp = "user_input_temp_preprocessed.xlsx" # Temp file for PREPROCESSED data (optional, for debugging or if needed)
        analysis_output_filename = "detailed_analysis_output.xlsx"
        final_output_filename = "final_combined_output.xlsx"

        # --- Progress Bar and Status ---
        progress_bar = st.progress(0)
        status_text = st.empty()
        status_text.info("Starting processing...")

        try:
            # --- Read Input File Bytes ---
            status_text.info("Reading input file...")
            progress_bar.progress(5)
            input_file_bytes = input_file.read()

            # --- Save ORIGINAL input bytes locally (needed for preprocess_input_file) ---
            try:
                with open(user_input_filename_temp, "wb") as f:
                    f.write(input_file_bytes)
                st.write(f"DEBUG: Saved original uploaded data to {user_input_filename_temp}")
            except Exception as e:
                 st.error(f"Error saving uploaded file temporarily: {e}")
                 st.stop()

            # --- STEP 0: PREPROCESSING ---
            status_text.info("Preprocessing input data (normalization)...")
            progress_bar.progress(10)
            preprocessed_df = None
            preprocessed_bytes_for_fixed_pipeline = None
            try:
                # Call the preprocessor using the temporary file path of the ORIGINAL data
                # preprocess_input_file reads the file, processes 'Value', adds columns,
                # renames Value->Value_E, UpdatedValue->Value_Normalized, and Value=Value_Normalized
                preprocessed_df = preprocess_input_file(user_input_filename_temp)
                st.write(f"DEBUG: Preprocessing complete. DataFrame shape: {preprocessed_df.shape}")
                st.write(f"DEBUG: Preprocessed columns: {preprocessed_df.columns.tolist()}") # Show columns after preprocessing

                # Check if 'Value' column still exists (it should be the normalized one)
                if 'Value' not in preprocessed_df.columns:
                     st.error("Preprocessing Error: 'Value' column lost after preprocessing. Check preprocessor.py and normalization_utils.py.")
                     # Clean up temp file
                     if os.path.exists(user_input_filename_temp): os.remove(user_input_filename_temp)
                     st.stop()
                if 'Value_E' not in preprocessed_df.columns:
                     st.warning("Preprocessing Warning: 'Value_E' (original value) column not found after preprocessing.") # Should exist
                if 'Value_Normalized' not in preprocessed_df.columns:
                     st.warning("Preprocessing Warning: 'Value_Normalized' column not found after preprocessing.") # Should exist


                # Prepare bytes of the *preprocessed* data for the fixed pipeline
                # Convert the preprocessed DataFrame to Excel bytes in memory
                output_buffer = io.BytesIO()
                preprocessed_df.to_excel(output_buffer, index=False, engine='openpyxl')
                preprocessed_bytes_for_fixed_pipeline = output_buffer.getvalue()
                st.write(f"DEBUG: Prepared preprocessed bytes for fixed pipeline (Size: {len(preprocessed_bytes_for_fixed_pipeline)} bytes)")

                # The preprocessed DataFrame is now the input for detailed analysis
                # No need to read from disk again, just use preprocessed_df
                input_df_for_analysis = preprocessed_df

            except ValueError as ve: # Catch specific ValueError from preprocessor/normalization utils
                st.error(f"Preprocessing Error: {ve}")
                if os.path.exists(user_input_filename_temp): os.remove(user_input_filename_temp) # Clean up
                st.stop()
            except Exception as e:
                st.error(f"An error occurred during preprocessing: {e}")
                st.error(traceback.format_exc())
                if os.path.exists(user_input_filename_temp): os.remove(user_input_filename_temp) # Clean up
                st.stop()
            progress_bar.progress(20) # Update progress after preprocessing


            # --- Step 1: Ensure mapping is on disk ---
            status_text.info("Verifying local mapping file...")
            progress_bar.progress(25) # Adjusted progress
            if not os.path.exists(MAPPING_FILENAME):
                 st.error(f"Critical Error: Local mapping file '{MAPPING_FILENAME}' disappeared.")
                 # Clean up temp file
                 if os.path.exists(user_input_filename_temp): os.remove(user_input_filename_temp)
                 st.stop()


            # --- Step 2: Run Fixed Processing Pipeline (using PREPROCESSED data) ---
            status_text.info("Running fixed processing pipeline (using preprocessed data)...")
            progress_bar.progress(30) # Adjusted progress
            # Pass the *preprocessed* bytes
            processed_df_fixed = run_fixed_pipeline_with_prefix_support(preprocessed_bytes_for_fixed_pipeline, MAPPING_FILENAME)
            #processed_df_fixed = process_fixed_pipeline_bytes(preprocessed_bytes_for_fixed_pipeline, MAPPING_FILENAME)

            if processed_df_fixed is None: # Check for critical failure
                st.error("Fixed Processing Pipeline failed critically. Aborting.")
                # Clean up temp file
                if os.path.exists(user_input_filename_temp): os.remove(user_input_filename_temp)
                st.stop()
            if processed_df_fixed.empty:
                st.warning("Fixed Processing Pipeline did not generate any output rows. Check preprocessed 'Value' column content and mapping.")
                # Let combine_results handle the empty processed_df.
            else:
                 st.write(f"DEBUG: Fixed processing pipeline complete. Output shape: {processed_df_fixed.shape}")
                 # Check if Value_E etc. were carried through
                 st.write(f"DEBUG: Columns in fixed pipeline output: {processed_df_fixed.columns.tolist()}")
            progress_bar.progress(50) # Adjusted progress


            # --- Step 3: Run Detailed Analysis Pipeline (using PREPROCESSED data) ---
            status_text.info("Running detailed analysis pipeline (using preprocessed data)...")
            progress_bar.progress(60) # Adjusted progress
            # Pass the *preprocessed* DataFrame directly
            analysis_result_path = run_detailed_analysis_with_prefix_support( # NEW CALL
                input_df=input_df_for_analysis,
                mapping_file=MAPPING_FILENAME,
                output_file=analysis_output_filename
            )
 # analysis_result_path = detailed_analysis(input_df=input_df_for_analysis, # Pass the DataFrame from preprocessingmapping_file=MAPPING_FILENAME,output_file=analysis_output_filename # Saves result to this file )

            if analysis_result_path is None:
                st.error("Detailed Analysis Pipeline failed. Aborting.")
                # Clean up temp files
                if os.path.exists(user_input_filename_temp): os.remove(user_input_filename_temp)
                st.stop()
            st.write(f"DEBUG: Detailed analysis complete. Output saved to {analysis_result_path}")
            progress_bar.progress(80) # Adjusted progress


            # --- Step 4: Combine Results ---
            status_text.info("Combining fixed and detailed results...")
            progress_bar.progress(90)
            # Pass the DataFrame from fixed pipeline (should include Value_E, Value_Normalized etc.)
            # and the path to the analysis results file (based on preprocessed 'Value')
            final_result_path = combine_results(
                processed_df=processed_df_fixed,
                analysis_file=analysis_output_filename,
                output_file=final_output_filename
            )

            if final_result_path is None:
                st.error("Combining results failed. Aborting.")
                # Clean up temp files
                if os.path.exists(user_input_filename_temp): os.remove(user_input_filename_temp)
                if os.path.exists(analysis_output_filename): os.remove(analysis_output_filename)
                st.stop()

            progress_bar.progress(100)
            status_text.success("Processing Complete!")


            # --- Step 5: Offer Download ---
            try:
                with open(final_result_path, "rb") as fp:
                    final_bytes = fp.read()
                st.download_button(
                    label=f"Download Combined Results ({os.path.basename(final_result_path)})",
                    data=final_bytes,
                    file_name=os.path.basename(final_result_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_final"
                )
            except Exception as e:
                st.error(f"Error preparing download link for final results: {e}")

            # --- Clean up temporary files ---
            # Make sure to remove the original temp file too
            cleanup_files = [user_input_filename_temp, analysis_output_filename] # Keep final_output_filename for download
            for f in cleanup_files:
                try:
                    if os.path.exists(f):
                         os.remove(f)
                         st.write(f"DEBUG: Removed temporary file {f}")
                except Exception as e:
                     st.warning(f"Could not remove temporary file {f}: {e}")


        except Exception as e:
            status_text.error(f"An error occurred during the 'Get Pattern' process: {e}")
            st.error(traceback.format_exc()) # Show detailed error
            # Attempt cleanup on error
            cleanup_files = [user_input_filename_temp, analysis_output_filename, final_output_filename]
            for f in cleanup_files:
                try:
                    if os.path.exists(f): os.remove(f)
                except Exception as clean_e:
                     st.warning(f"Could not remove temp file {f} during error cleanup: {clean_e}")


############################
# OPERATION: MANAGE UNITS
############################
elif operation == "Manage Units":
    # (Keep this section as is)
    st.header("Manage Units (GitHub mapping file)")

    # Display current mapping from session state
    st.subheader("Current Mapping File (from GitHub / Session State)")
    current_mapping_df = st.session_state.get("mapping_df")

    if current_mapping_df is not None:
         st.dataframe(current_mapping_df)
         if st.button("Refresh Mapping from GitHub"):
              if "mapping_df" in st.session_state: del st.session_state["mapping_df"]
              st.rerun()
    else:
         st.warning("Mapping data not currently loaded in session state.")
         if st.button("Retry Download from GitHub"):
              if "mapping_df" in st.session_state: del st.session_state["mapping_df"]
              st.rerun()

    # --- Add New Unit ---
    # (Keep this section as is)
    st.subheader("Add New Base Unit")
    with st.form("add_unit_form"):
        new_unit = st.text_input("Enter new Base Unit Symbol").strip()
        submitted_new = st.form_submit_button("Add Unit to Local Session")

    if submitted_new and new_unit:
        if current_mapping_df is not None:
            if new_unit in current_mapping_df["Base Unit Symbol"].astype(str).values:
                 st.warning(f"Unit '{new_unit}' already exists in the session mapping.")
            else:
                new_row_data = {"Base Unit Symbol": new_unit, "Multiplier Symbol": None}
                for col in current_mapping_df.columns:
                    if col not in new_row_data:
                         new_row_data[col] = None
                new_row_df = pd.DataFrame([new_row_data], columns=current_mapping_df.columns)
                st.session_state["mapping_df"] = pd.concat(
                    [current_mapping_df, new_row_df], ignore_index=True)
                st.success(f"New unit '{new_unit}' added to the current session. Save to GitHub to persist.")
                st.rerun()
        else:
            st.error("Mapping data not available in session state to add unit.")
    elif submitted_new:
        st.error("Base Unit Symbol cannot be empty.")


    # --- Delete Unit ---
    # (Keep this section as is)
    st.subheader("Delete Base Unit")
    if current_mapping_df is not None and not current_mapping_df.empty:
        try:
             existing_units = sorted(current_mapping_df["Base Unit Symbol"].dropna().astype(str).unique())
             existing_units = [unit for unit in existing_units if unit]
        except KeyError:
             st.error("Cannot retrieve units: 'Base Unit Symbol' column not found.")
             existing_units = []

        if existing_units:
            unit_to_delete = st.selectbox(
                "Select a unit to delete from local session",
                options=["--Select--"] + existing_units, key="delete_unit_select")
            if st.button("Delete Selected Unit from Local Session"):
                if unit_to_delete != "--Select--":
                    original_shape = st.session_state["mapping_df"].shape
                    st.session_state["mapping_df"] = st.session_state["mapping_df"][
                        st.session_state["mapping_df"]["Base Unit Symbol"] != unit_to_delete
                    ].reset_index(drop=True)
                    new_shape = st.session_state["mapping_df"].shape
                    st.success(f"Unit '{unit_to_delete}' deleted from the current session. (Rows before: {original_shape[0]}, after: {new_shape[0]}). Save to GitHub to persist.")
                    st.rerun()
                else:
                    st.warning("Please select a unit to delete.")
        else:
            st.info("No base units found in the current mapping session to delete.")
    elif current_mapping_df is not None and current_mapping_df.empty:
         st.info("Mapping data in session is currently empty.")
    else:
        st.info("Mapping data is not loaded.")

    # --- Persist Changes ---
    # (Keep this section as is)
    st.subheader("Persist Changes")
    st.warning("Changes made using the forms above only affect the current browser session.")

    if current_mapping_df is not None:
        try:
            output_buffer = BytesIO()
            current_mapping_df.to_excel(output_buffer, index=False, engine='openpyxl')
            output_buffer.seek(0)
            st.download_button(
                label="Download Current Mapping File (Local Session)",
                data=output_buffer,
                file_name="session_mapping.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_session_mapping"
            )
        except Exception as e:
             st.error(f"Error creating download link for session mapping: {e}")

    if st.button("Save Current Session Mapping to GitHub"):
        if current_mapping_df is not None:
            st.info("Attempting to save current session mapping to GitHub...")
            required_cols = {"Base Unit Symbol", "Multiplier Symbol"}
            if not required_cols.issubset(current_mapping_df.columns):
                 st.error(f"Cannot save to GitHub: Mapping file is missing required columns: {required_cols}. Please add them back or refresh from GitHub.")
            else:
                success = update_mapping_file_on_github(current_mapping_df)
                if success:
                    st.success("Mapping file updated on GitHub! Changes will be reflected after the app fully refreshes or restarts.")
                else:
                    st.error("Failed to update mapping file on GitHub. Check console/logs for details (e.g., permissions, token validity, conflicts).")
        else:
            st.error("No mapping data found in the current session to save.")

# --- END OF FILE app.py ---
