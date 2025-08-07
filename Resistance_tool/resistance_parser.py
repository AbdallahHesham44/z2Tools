# streamlit_app.py
import streamlit as st
import pandas as pd
import os
import time
import shutil
from datetime import datetime
from resistance_parser import ResistanceParser  # We'll modularize your class into another file

# Title
st.title("🔍 Resistance Code Parser Tool")
st.markdown("Upload an Excel file with a `PartNumber` column and optional `Value` column.")

# Upload input Excel file
uploaded_file = st.file_uploader("📁 Upload Excel File", type=["xlsx"])

# User parameters
batch_size = st.number_input("🔢 Batch Size", min_value=100, max_value=10000, step=100, value=1000)
checkpoint_interval = st.number_input("💾 Checkpoint Interval", min_value=1000, max_value=50000, step=1000, value=5000)

# Output folder
output_dir = "streamlit_output"
os.makedirs(output_dir, exist_ok=True)

if uploaded_file is not None:
    # Save uploaded file
    input_file_path = os.path.join(output_dir, uploaded_file.name)
    with open(input_file_path, "wb") as f:
        f.write(uploaded_file.read())

    st.success("✅ File uploaded successfully.")

    # Run parsing on button click
    if st.button("🚀 Start Processing"):
        with st.spinner("⏳ Processing in progress..."):

            parser = ResistanceParser(
                input_file=input_file_path,
                output_dir=output_dir,
                batch_size=batch_size,
                checkpoint_interval=checkpoint_interval
            )

            # Replace print with progress reporting
            original_print = print

            def streamlit_print(*args, **kwargs):
                st.text(" ".join(str(arg) for arg in args))

            try:
                print = streamlit_print
                parser.run()
                print = original_print
            except Exception as e:
                print = original_print
                st.error(f"❌ Error: {e}")
                raise

            st.success("✅ Done!")

        # Show download links
        all_file = os.path.join(output_dir, "extracted_resistances_all.xlsx")
        best_file = os.path.join(output_dir, "extracted_resistances.xlsx")

        if os.path.exists(best_file):
            st.download_button("📥 Download Best Matches", data=open(best_file, "rb").read(), file_name="best_matches.xlsx")

        if os.path.exists(all_file):
            st.download_button("📥 Download All Matches", data=open(all_file, "rb").read(), file_name="all_matches.xlsx")
