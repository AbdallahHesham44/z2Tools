# streamlit_app.py
import streamlit as st
import pandas as pd
import tempfile
import os
from CapacitorValueMatcher import CapacitorValueMatcher
from datetime import datetime

st.set_page_config(page_title="Capacitor Value Matcher", layout="wide")
st.title("⚙️ Capacitor Value Matcher Tool")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

batch_size = st.number_input("Batch Size", value=10000, step=1000)
num_threads = st.slider("Number of Threads", 1, 10, value=4)
checkpoint_interval = st.number_input("Checkpoint Interval", value=5000, step=1000)

output_dir = tempfile.mkdtemp()

if uploaded_file is not None:
    st.success("✅ File uploaded successfully")

    if st.button("🚀 Run Matcher"):
        progress_bar = st.progress(0, text="Starting...")
        status_placeholder = st.empty()

        # Save uploaded file
        temp_input_file = os.path.join(output_dir, f"input_{datetime.now().timestamp()}.xlsx")
        with open(temp_input_file, "wb") as f:
            f.write(uploaded_file.read())

        # Define a callback function to update progress bar
        def update_progress(progress, total, batch_num):
            percentage = progress / total if total else 0
            progress_bar.progress(percentage, text=f"Processing batch {batch_num} - {int(percentage * 100)}%")
            status_placeholder.info(f"Processed {progress} of {total} rows")

        # Run matcher
        with st.spinner("Processing..."):
            matcher = CapacitorValueMatcher(
                input_file_path=temp_input_file,
                output_dir=output_dir,
                batch_size=batch_size,
                num_threads=num_threads,
                checkpoint_interval=checkpoint_interval,
                progress_callback=update_progress  # <== pass the callback
            )
            matcher.process_file()

        st.success("🎉 Matching completed!")
        st.balloons()

        # Offer downloads
        matched_file = os.path.join(output_dir, "MatchedOutput.xlsx")
        unmatched_file = os.path.join(output_dir, "notMatchedOutput.xlsx")

        if os.path.exists(matched_file):
            with open(matched_file, "rb") as f:
                st.download_button("📥 Download Matched Results", f, file_name="MatchedOutput.xlsx")

        if os.path.exists(unmatched_file):
            with open(unmatched_file, "rb") as f:
                st.download_button("📥 Download Unmatched Results", f, file_name="notMatchedOutput.xlsx")
