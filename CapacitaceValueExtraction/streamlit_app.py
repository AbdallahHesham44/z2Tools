# streamlit_app.py
import streamlit as st
import pandas as pd
import tempfile
import os
from CapacitorValueMatcher import CapacitorValueMatcher
from datetime import datetime
import gc
import threading
from pathlib import Path
import time

# Configure page with memory optimization
st.set_page_config(
    page_title="Capacitor Value Matcher", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("⚙️ Capacitor Value Matcher Tool")
st.warning("🧠 **Low Memory Mode** - Optimized for 1GB RAM environment")

# Initialize session state
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
if 'output_files' not in st.session_state:
    st.session_state.output_files = {}
if 'processing_stats' not in st.session_state:
    st.session_state.processing_stats = {}
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = None

# Ultra-conservative settings for 1GB RAM
with st.sidebar:
    st.header("⚙️ Memory-Optimized Settings")
    
    # Much smaller batch sizes for 1GB RAM
    batch_size = st.number_input(
        "Batch Size", 
        value=500,  # Very small for 1GB RAM
        min_value=100,
        max_value=2000,
        step=100,
        help="CRITICAL: Keep small for 1GB RAM! 500-1000 max recommended"
    )
    
    # Single thread only for memory conservation
    num_threads = st.selectbox(
        "Number of Threads", 
        [1],  # Force single thread for memory conservation
        help="Fixed at 1 thread to conserve memory"
    )
    
    checkpoint_interval = st.number_input(
        "Checkpoint Interval", 
        value=500,  # Very frequent checkpoints
        min_value=100,
        max_value=1000,
        step=100,
        help="Frequent saves to prevent data loss"
    )
    
    # Memory monitoring
    st.subheader("💾 Memory Tips")
    st.info("For 500K rows with 1GB RAM:\n- Use batch size: 300-500\n- Process in multiple smaller files if possible\n- Close other browser tabs")

# File upload section
st.header("📁 Upload File")

# File size warning
st.error("⚠️ **IMPORTANT for 1GB RAM**: Files over 50MB may fail. Consider splitting large files into smaller chunks.")

uploaded_file = st.file_uploader(
    "Upload Excel File", 
    type=["xlsx", "xls"],
    help="For 1GB RAM: Recommended max file size is 50MB"
)

# File info and memory check
if uploaded_file is not None:
    # Check file size immediately without loading into memory
    file_size_mb = len(uploaded_file.getvalue()) / (1024 * 1024)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📄 File Name", uploaded_file.name)
    with col2:
        st.metric("💾 File Size", f"{file_size_mb:.2f} MB")
    with col3:
        if file_size_mb > 100:
            st.error("❌ Too large for 1GB RAM")
        elif file_size_mb > 50:
            st.warning("⚠️ May cause memory issues")
        else:
            st.success("✅ Size OK for 1GB RAM")
    
    # Memory usage warning
    if file_size_mb > 50:
        st.error("🚨 **FILE TOO LARGE FOR 1GB RAM**")
        st.markdown("""
        **Options:**
        1. Split your file into smaller chunks (< 50MB each)
        2. Process on a machine with more RAM
        3. Use a local installation instead of Streamlit Cloud
        """)
        st.stop()
    
    # Estimate with conservative memory usage
    estimated_rows = int(file_size_mb * 8000)  # Conservative estimate
    estimated_batches = max(1, estimated_rows // batch_size)
    estimated_time = max(1, estimated_batches // 10)  # Very conservative
    
    st.info(f"📊 Est. rows: ~{estimated_rows:,} | Est. batches: {estimated_batches} | Est. time: ~{estimated_time} min")
    
    # Memory usage prediction
    estimated_memory_mb = (batch_size * 0.001) + (file_size_mb * 2)  # Very rough estimate
    if estimated_memory_mb > 800:  # Leave 200MB headroom
        st.error(f"⚠️ Estimated memory usage: {estimated_memory_mb:.0f}MB - Too high for 1GB RAM!")
        st.error("Reduce batch size or use smaller file")
        st.stop()
    
    # Processing section
    if st.button("🚀 Start Processing (Low Memory Mode)", type="primary"):
        # Force garbage collection before starting
        gc.collect()
        
        # Reset session state
        st.session_state.processing_complete = False
        st.session_state.output_files = {}
        st.session_state.processing_stats = {}
        
        # Create temporary directory
        output_dir = tempfile.mkdtemp(prefix="cap_match_")
        st.session_state.temp_dir = output_dir
        
        # Progress indicators
        progress_bar = st.progress(0, text="Initializing...")
        status_placeholder = st.empty()
        memory_placeholder = st.empty()
        
        try:
            # Save uploaded file with minimal memory usage
            temp_input_file = os.path.join(output_dir, f"input_{int(time.time())}.xlsx")
            
            # Write file in very small chunks to minimize memory
            with open(temp_input_file, "wb") as f:
                chunk_size = 1024  # Very small chunks for 1GB RAM
                file_data = uploaded_file.getvalue()
                for i in range(0, len(file_data), chunk_size):
                    f.write(file_data[i:i + chunk_size])
                    if i % (chunk_size * 100) == 0:  # Every 100 chunks
                        gc.collect()  # Aggressive garbage collection
            
            # Clear the uploaded file from memory immediately
            del file_data
            uploaded_file = None
            gc.collect()
            
            # Memory-aware progress callback
            progress_lock = threading.Lock()
            
            # Use a mutable object to share state between functions
            gc_state = {'last_gc_time': time.time()}
            
            def update_progress(progress, total, batch_num, matched_count=0, unmatched_count=0):
                with progress_lock:
                    try:
                        percentage = min(progress / total, 1.0) if total > 0 else 0
                        
                        progress_bar.progress(
                            percentage, 
                            text=f"Batch {batch_num}/{estimated_batches} - {int(percentage * 100)}%"
                        )
                        
                        status_placeholder.info(
                            f"📊 Processed: {progress:,} / {total:,} rows (Batch size: {batch_size})"
                        )
                        
                        memory_placeholder.success(f"🧠 Memory optimized - Single thread, {batch_size} batch size")
                        
                        # Force garbage collection every 5 batches or 30 seconds
                        current_time = time.time()
                        if batch_num % 5 == 0 or (current_time - gc_state['last_gc_time']) > 30:
                            gc.collect()
                            gc_state['last_gc_time'] = current_time
                            
                    except Exception as e:
                        st.error(f"Progress update error: {e}")
            
            # Initialize matcher with ultra-conservative settings
            with st.spinner("🔄 Processing with memory optimization..."):
                matcher = CapacitorValueMatcher(
                    input_file_path=temp_input_file,
                    output_dir=output_dir,
                    batch_size=batch_size,
                    num_threads=1,  # Force single thread
                    checkpoint_interval=checkpoint_interval,
                    progress_callback=update_progress
                )
                
                # Process the file
                result = matcher.process_file()
                
                # Final garbage collection
                gc.collect()
                
                # Update progress to 100%
                progress_bar.progress(1.0, text="✅ Processing complete!")
                
            # Handle output files with memory conservation
            matched_file = os.path.join(output_dir, "MatchedOutput.xlsx")
            unmatched_file = os.path.join(output_dir, "notMatchedOutput.xlsx")
            
            # Instead of loading files into memory, just store file paths
            st.session_state.output_files = {}
            
            if os.path.exists(matched_file):
                st.session_state.output_files['matched_path'] = matched_file
                matched_size = os.path.getsize(matched_file) / (1024 * 1024)
                st.session_state.processing_stats['matched_size'] = matched_size
            
            if os.path.exists(unmatched_file):
                st.session_state.output_files['unmatched_path'] = unmatched_file
                unmatched_size = os.path.getsize(unmatched_file) / (1024 * 1024)
                st.session_state.processing_stats['unmatched_size'] = unmatched_size
            
            st.session_state.processing_complete = True
            
            # Clean up input file immediately
            try:
                os.remove(temp_input_file)
            except:
                pass
            
            # Final garbage collection
            gc.collect()
            
            st.success("🎉 Processing completed successfully!")
            st.balloons()
            
        except MemoryError:
            st.error("❌ **OUT OF MEMORY!** Try:")
            st.error("- Reduce batch size to 100-300")
            st.error("- Use smaller input file")
            st.error("- Split file into multiple parts")
            
        except Exception as e:
            st.error(f"❌ Error during processing: {str(e)}")
            
            # Clean up on error
            try:
                if 'temp_input_file' in locals() and os.path.exists(temp_input_file):
                    os.remove(temp_input_file)
            except:
                pass

# Download section - load files only when downloading
if st.session_state.processing_complete and st.session_state.output_files:
    st.header("📥 Download Results")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if 'matched_path' in st.session_state.output_files:
            matched_path = st.session_state.output_files['matched_path']
            matched_size = st.session_state.processing_stats.get('matched_size', 0)
            
            if os.path.exists(matched_path):
                # Load file only when download button is clicked
                try:
                    with open(matched_path, "rb") as f:
                        matched_data = f.read()
                    
                    st.download_button(
                        label="📄 Download Matched Results",
                        data=matched_data,
                        file_name=f"MatchedOutput_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help=f"File size: {matched_size:.2f} MB"
                    )
                    
                    # Clear from memory immediately
                    del matched_data
                    gc.collect()
                    
                except Exception as e:
                    st.error(f"Error loading matched file: {e}")
            else:
                st.error("Matched file not found")
        else:
            st.info("No matched results available")
    
    with col2:
        if 'unmatched_path' in st.session_state.output_files:
            unmatched_path = st.session_state.output_files['unmatched_path']
            unmatched_size = st.session_state.processing_stats.get('unmatched_size', 0)
            
            if os.path.exists(unmatched_path):
                try:
                    with open(unmatched_path, "rb") as f:
                        unmatched_data = f.read()
                    
                    st.download_button(
                        label="📄 Download Unmatched Results",
                        data=unmatched_data,
                        file_name=f"NotMatchedOutput_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help=f"File size: {unmatched_size:.2f} MB"
                    )
                    
                    # Clear from memory immediately
                    del unmatched_data
                    gc.collect()
                    
                except Exception as e:
                    st.error(f"Error loading unmatched file: {e}")
            else:
                st.error("Unmatched file not found")
        else:
            st.info("No unmatched results available")

# Cleanup button
if st.session_state.temp_dir and os.path.exists(st.session_state.temp_dir):
    if st.button("🧹 Clean Up Temporary Files"):
        try:
            import shutil
            shutil.rmtree(st.session_state.temp_dir)
            st.session_state.temp_dir = None
            st.session_state.processing_complete = False
            st.session_state.output_files = {}
            gc.collect()
            st.success("✅ Temporary files cleaned up")
            st.rerun()
        except Exception as e:
            st.error(f"Error cleaning up: {e}")

# Critical tips for 1GB RAM
with st.expander("🚨 CRITICAL: 1GB RAM Optimization Guide"):
    st.markdown("""
    **FOR 500K ROWS WITH 1GB RAM:**
    
    🔥 **CRITICAL SETTINGS:**
    - Batch size: 300-500 MAX
    - Threads: 1 (forced)
    - Checkpoint: 500
    - File size: < 50MB
    
    🚨 **MEMORY MANAGEMENT:**
    - Close ALL other browser tabs
    - Don't run other applications
    - Process during low-traffic hours
    - Consider splitting large files
    
    ⚡ **IF IT STILL FAILS:**
    1. Reduce batch size to 100-200
    2. Split your 500K file into 5x 100K files
    3. Process each file separately
    4. Combine results manually
    
    🎯 **REALISTIC EXPECTATIONS:**
    - 500K rows might take 2-4 hours
    - Memory errors are still possible
    - Consider upgrading RAM or using desktop version
    """)

# Footer
st.markdown("---")
st.markdown("🧠 **1GB RAM Optimized Version** - Maximum memory conservation")

# Force garbage collection on every run
gc.collect()
