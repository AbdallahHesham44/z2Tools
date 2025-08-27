import streamlit as st
import pandas as pd
import re
import time
import json
import os
from datetime import datetime, timedelta
from pathlib import Path
import gc
import warnings
import tempfile
import zipfile
import io
warnings.filterwarnings('ignore')

# Set page config
st.set_page_config(
    page_title="Resistance Code Parser",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
.stProgress > div > div > div > div {
    background-color: #ff6b35;
}
.success-box {
    padding: 1rem;
    border-radius: 0.5rem;
    background-color: #d4edda;
    border: 1px solid #c3e6cb;
    color: #155724;
}
.warning-box {
    padding: 1rem;
    border-radius: 0.5rem;
    background-color: #fff3cd;
    border: 1px solid #ffeaa7;
    color: #856404;
}
.error-box {
    padding: 1rem;
    border-radius: 0.5rem;
    background-color: #f8d7da;
    border: 1px solid #f5c6cb;
    color: #721c24;
}
.metric-container {
    background-color: #f0f2f6;
    padding: 1rem;
    border-radius: 0.5rem;
    margin: 0.5rem 0;
}
</style>
""", unsafe_allow_html=True)

class StreamlitResistanceParser:
    def __init__(self):
        """Initialize the parser with configuration"""
        # Define character multiplier rules
        self.rule1_multipliers = {
            'J': 1e-1,   # 10^-1
            'K': 1e-2,   # 10^-2
            'L': 1e-3,   # 10^-3
            'M': 1e-4,   # 10^-4
            'N': 1e-5,   # 10^-5
            'P': 1e-6    # 10^-6
        }

        self.rule2_no_multiplier = ['A', 'R', 'W', 'D', 'G', 'J', 'M']
        self.rule2_1k_multiplier = ['B', 'U', 'Y', 'E', 'H', 'K', 'N']
        self.rule2_1m_multiplier = ['C', 'V', 'Z', 'F', 'T', 'L', 'P']

    def parse_r_decimal_patterns(self, code):
        """Parse R-decimal patterns where R acts as decimal point"""
        results = []

        # Pattern 1: R followed by digits (R047, R100, R22)
        for i in range(len(code) - 1):
            match = re.match(r'^R(\d{1,4})$', code[i:i+5] if i+5 <= len(code) else code[i:])
            if match:
                digits = match.group(1)
                if len(digits) == 1:
                    value = float(f"0.0{digits}")
                elif len(digits) == 2:
                    value = float(f"0.{digits}")
                elif len(digits) == 3:
                    value = float(f"0.{digits}")
                else:
                    value = float(f"0.{digits}")

                pattern = f"R{digits}"
                results.append({
                    'pattern': pattern,
                    'type': 'r-decimal',
                    'rule': 'R-Decimal-Leading',
                    'value': value,
                    'unit': 'Ohm',
                    'position': i,
                    'base_digits': f"0.{digits}",
                    'multiplier_char': 'R',
                    'multiplier_value': 1
                })

        # Pattern 2: digits + R + digits (47R0, 4R7, 100R5)
        for i in range(len(code) - 2):
            for length in range(3, min(6, len(code) - i + 1)):
                substring = code[i:i+length]
                match = re.match(r'^(\d{1,3})R(\d{1,3})$', substring)
                if match:
                    before, after = match.groups()
                    value = float(f"{before}.{after}")

                    results.append({
                        'pattern': substring,
                        'type': 'r-decimal',
                        'rule': 'R-Decimal-Middle',
                        'value': value,
                        'unit': 'Ohm',
                        'position': i,
                        'base_digits': f"{before}.{after}",
                        'multiplier_char': 'R',
                        'multiplier_value': 1
                    })

        # Pattern 3: digits + R (trailing R)
        for i in range(len(code) - 1):
            for length in range(2, min(5, len(code) - i + 1)):
                substring = code[i:i+length]
                match = re.match(r'^(\d{1,3})R$', substring)
                if match:
                    digits = match.group(1)
                    value = float(f"{digits}.0")

                    results.append({
                        'pattern': substring,
                        'type': 'r-decimal',
                        'rule': 'R-Decimal-Trailing',
                        'value': value,
                        'unit': 'Ohm',
                        'position': i,
                        'base_digits': f"{digits}.0",
                        'multiplier_char': 'R',
                        'multiplier_value': 1
                    })

        return results

    def parse_4digit_rule1(self, code):
        """Rule 1: 4-digit pattern with single character (decimal multiplier)"""
        results = []

        for i in range(len(code) - 3):
            substring = code[i:i+4]

            patterns = [
                (r'^(\d{3})([A-Z])$', lambda d, c: (float(d), c, d)),
                (r'^([A-Z])(\d{3})$', lambda c, d: (float(d), c, d)),
                (r'^(\d{2})([A-Z])(\d{1})$', lambda b, c, a: (float(f"{b}.{a}"), c, f"{b}.{a}")),
                (r'^(\d{1})([A-Z])(\d{2})$', lambda b, c, a: (float(f"{b}.{a}"), c, f"{b}.{a}"))
            ]

            for pattern, processor in patterns:
                match = re.match(pattern, substring)
                if match:
                    if len(match.groups()) == 2:
                        base_value, char, base_digits = processor(match.group(1), match.group(2))
                    else:
                        base_value, char, base_digits = processor(match.group(1), match.group(2), match.group(3))

                    if char in self.rule1_multipliers:
                        multiplier = self.rule1_multipliers[char]
                        value = base_value * multiplier

                        results.append({
                            'pattern': substring,
                            'type': '4-digit-rule1',
                            'rule': f'Rule1-{char}',
                            'value': value,
                            'unit': 'Ohm',
                            'position': i,
                            'base_digits': str(base_digits),
                            'multiplier_char': char,
                            'multiplier_value': multiplier
                        })

        return results

    def parse_4digit_rule2(self, code):
        """Rule 2: 4-character code with character replacing decimal point"""
        results = []

        patterns = [
            r'^(\d{2})([A-Z])(\d)$',    # DDLD
            r'^(\d)([A-Z])(\d{2})$',    # DLDD
            r'^([A-Z])(\d{3})$',        # LDDD
            r'^(\d{3})([A-Z])$',        # DDDL
        ]

        for i in range(len(code) - 3):
            substring = code[i:i+4]
            for pattern in patterns:
                match = re.match(pattern, substring)
                if match:
                    if pattern == r'^(\d{3})([A-Z])$':
                        before, char = match.groups()
                        base_value = f"{before}.0"
                    elif pattern == r'^([A-Z])(\d{3})$':
                        char, after = match.groups()
                        base_value = f"0.{after}"
                    else:
                        before, char, after = match.groups()
                        base_value = f"{before}.{after}"

                    if char in self.rule2_no_multiplier:
                        multiplier = 1
                        rule_type = "no-multiplier"
                    elif char in self.rule2_1k_multiplier:
                        multiplier = 1e3
                        rule_type = "1k-multiplier"
                    elif char in self.rule2_1m_multiplier:
                        multiplier = 1e6
                        rule_type = "1m-multiplier"
                    else:
                        continue

                    value = float(base_value) * multiplier
                    results.append({
                        'pattern': substring,
                        'type': '4-digit-rule2',
                        'rule': f'Rule2-{char}-{rule_type}',
                        'value': value,
                        'unit': 'Ohm',
                        'position': i,
                        'base_digits': base_value,
                        'multiplier_char': char,
                        'multiplier_value': multiplier
                    })

        return results

    def parse_traditional_patterns(self, code):
        """Parse traditional 3-digit and 4-digit numeric patterns"""
        results = []

        # 3-digit patterns
        for i in range(len(code) - 2):
            substring = code[i:i+3]
            if substring.isdigit():
                try:
                    base = int(substring[:2])
                    multiplier = int(substring[2])
                    if multiplier <= 6:
                        value = base * (10 ** multiplier)
                        results.append({
                            'pattern': substring,
                            'type': '3-digit-traditional',
                            'rule': 'Traditional-3digit',
                            'value': value,
                            'unit': 'Ohm',
                            'position': i,
                            'base_digits': str(base),
                            'multiplier_char': str(multiplier),
                            'multiplier_value': 10 ** multiplier
                        })
                except:
                    pass

        # 4-digit patterns
        for i in range(len(code) - 3):
            substring = code[i:i+4]
            if substring.isdigit():
                try:
                    base = int(substring[:3])
                    multiplier = int(substring[3])
                    if multiplier <= 6:
                        value = base * (10 ** multiplier)
                        results.append({
                            'pattern': substring,
                            'type': '4-digit-traditional',
                            'rule': 'Traditional-4digit',
                            'value': value,
                            'unit': 'Ohm',
                            'position': i,
                            'base_digits': str(base),
                            'multiplier_char': str(multiplier),
                            'multiplier_value': 10 ** multiplier
                        })
                except:
                    pass

        return results

    def parse_all_resistance_codes_enhanced(self, code):
        """Extract ALL possible resistance codes using all rules"""
        code = str(code).upper().strip()
        if not code or code == 'NAN':
            return []

        results = []

        # Apply all parsing rules
        results.extend(self.parse_r_decimal_patterns(code))
        results.extend(self.parse_4digit_rule1(code))
        results.extend(self.parse_4digit_rule2(code))
        results.extend(self.parse_traditional_patterns(code))

        # Remove duplicates
        seen = set()
        unique_results = []
        for result in results:
            key = (result['pattern'], result['type'], result['rule'])
            if key not in seen:
                seen.add(key)
                unique_results.append(result)

        # Sort by position
        unique_results.sort(key=lambda x: (x['position'], x['pattern']))
        return unique_results

    def convert_to_ohm(self, value_str):
        """Convert string like '1.02 MOhm' or '5.6 KOhm' to float in Ohms"""
        if not isinstance(value_str, str) or pd.isna(value_str):
            return None

        value_str = str(value_str).replace(" ", "").upper()
        match = re.match(r'^([0-9]*\.?[0-9]+)([A-Z]+)$', value_str)
        if not match:
            return None

        val, unit = match.groups()
        try:
            val = float(val)
        except:
            return None

        unit_multipliers = {
            'OHM': 1,
            'KOHM': 1e3,
            'MOHM': 1e6,
            'LOHM': 1e-2
        }

        return val * unit_multipliers.get(unit, 1)

    def find_best_match(self, parsed_results, target_ohm):
        """Find the best matching resistance value from parsed results"""
        if not parsed_results or pd.isna(target_ohm):
            return None

        best_match = None
        min_difference = float('inf')

        for result in parsed_results:
            try:
                parsed_value = float(result['value'])
                difference = abs(parsed_value - target_ohm)
                relative_diff = difference / max(target_ohm, parsed_value) if max(target_ohm, parsed_value) > 0 else float('inf')

                if difference < 1e-10:
                    return result

                if relative_diff < 0.001 and difference < min_difference:
                    min_difference = difference
                    best_match = result

            except (ValueError, TypeError):
                continue

        return best_match

    def process_dataframe(self, df, progress_bar, status_text):
        """Process the entire dataframe"""
        match_rows = []
        no_match_rows = []
        total_rows = len(df)

        for idx, row in df.iterrows():
            # Update progress
            progress = (idx + 1) / total_rows
            progress_bar.progress(progress)
            status_text.text(f"Processing row {idx + 1:,} of {total_rows:,} ({progress*100:.1f}%)")

            part = str(row['PartNumber'])
            original_value = str(row.get('Value', '')).strip()
            converted_val = self.convert_to_ohm(original_value)

            parsed = self.parse_all_resistance_codes_enhanced(part)

            base_row = {
                "PartNumber": part,
                "CompanyName": row.get("CompanyName", ""),
                "ProductLine": row.get("ProductLine", ""),
                "FeatureName": row.get("FeatureName", ""),
                "OriginalValue": original_value,
                "OriginalOhm": converted_val,
            }

            if not parsed:
                no_match_row = base_row.copy()
                no_match_row.update({
                    "ParsedPattern": "No Pattern Found",
                    "ParsedType": "none",
                    "ParsedRule": "none",
                    "ParsedValue": None,
                    "ParsedUnit": None,
                    "Position": -1,
                    "BaseDigits": "",
                    "MultiplierChar": "",
                    "MultiplierValue": None,
                    "MatchStatus": "no_pattern"
                })
                no_match_rows.append(no_match_row)
                continue

            best_match = self.find_best_match(parsed, converted_val)

            if best_match:
                match_row = base_row.copy()
                match_row.update({
                    "ParsedPattern": best_match['pattern'],
                    "ParsedType": best_match['type'],
                    "ParsedRule": best_match['rule'],
                    "ParsedValue": best_match['value'],
                    "ParsedUnit": best_match['unit'],
                    "Position": best_match['position'],
                    "BaseDigits": best_match['base_digits'],
                    "MultiplierChar": best_match['multiplier_char'],
                    "MultiplierValue": best_match['multiplier_value'],
                    "MatchStatus": "matched",
                    "MatchDifference": abs(best_match['value'] - converted_val) if converted_val else None
                })
                match_rows.append(match_row)
            else:
                best_parsed = parsed[0]
                no_match_row = base_row.copy()
                no_match_row.update({
                    "ParsedPattern": best_parsed['pattern'],
                    "ParsedType": best_parsed['type'],
                    "ParsedRule": best_parsed['rule'],
                    "ParsedValue": best_parsed['value'],
                    "ParsedUnit": best_parsed['unit'],
                    "Position": best_parsed['position'],
                    "BaseDigits": best_parsed['base_digits'],
                    "MultiplierChar": best_parsed['multiplier_char'],
                    "MultiplierValue": best_parsed['multiplier_value'],
                    "MatchStatus": "no_match",
                    "ParsedAlternatives": len(parsed)
                })
                no_match_rows.append(no_match_row)

        return match_rows, no_match_rows

def create_downloadable_excel(df, filename):
    """Create a downloadable Excel file"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Results')
    output.seek(0)
    return output.getvalue()

def main():
    st.title("⚡ Resistance Code Parser")
    st.markdown("Upload an Excel file to parse resistance codes and separate matches from non-matches")

    # Initialize parser
    if 'parser' not in st.session_state:
        st.session_state.parser = StreamlitResistanceParser()

    # Sidebar configuration
    st.sidebar.header("🔧 Configuration")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose Excel file",
        type=['xlsx', 'xls'],
        help="Upload an Excel file containing PartNumber and Value columns"
    )

    if uploaded_file is not None:
        try:
            # Load data
            with st.spinner("Loading data..."):
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ Loaded {len(df):,} rows from {uploaded_file.name}")
            
            # Show data preview
            st.subheader("📊 Data Preview")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Show column info
            st.subheader("📋 Column Information")
            col_info = pd.DataFrame({
                'Column': df.columns,
                'Type': df.dtypes,
                'Non-Null Count': df.count(),
                'Sample Values': [df[col].dropna().head(3).tolist() for col in df.columns]
            })
            st.dataframe(col_info, use_container_width=True)

            # Validation
            required_columns = ['PartNumber']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"❌ Missing required columns: {missing_columns}")
                st.stop()

            # Test specific part number
            st.sidebar.subheader("🧪 Test Single Part")
            test_part = st.sidebar.text_input("Enter part number to test:", value="MCR100JZHJSR047")
            test_value = st.sidebar.text_input("Enter target value:", value="47 mOhm")
            
            if st.sidebar.button("Test Part"):
                st.sidebar.write("**Test Results:**")
                parsed = st.session_state.parser.parse_all_resistance_codes_enhanced(test_part)
                target_ohm = st.session_state.parser.convert_to_ohm(test_value)
                
                if parsed:
                    best_match = st.session_state.parser.find_best_match(parsed, target_ohm)
                    for result in parsed:
                        st.sidebar.write(f"- {result['pattern']}: {result['value']} Ω")
                    if best_match:
                        st.sidebar.success(f"✅ Match: {best_match['pattern']}")
                    else:
                        st.sidebar.warning("❌ No good match")
                else:
                    st.sidebar.warning("No patterns found")

            # Processing section
            st.subheader("🚀 Process Data")
            
            if st.button("Start Processing", type="primary", use_container_width=True):
                start_time = time.time()
                
                # Progress tracking
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Results containers
                results_container = st.container()
                
                try:
                    with st.spinner("Processing resistance codes..."):
                        match_rows, no_match_rows = st.session_state.parser.process_dataframe(
                            df, progress_bar, status_text
                        )
                    
                    processing_time = time.time() - start_time
                    
                    # Clear progress indicators
                    progress_bar.empty()
                    status_text.empty()
                    
                    # Show results
                    with results_container:
                        st.success(f"✅ Processing completed in {processing_time:.1f} seconds!")
                        
                        # Statistics
                        total_processed = len(match_rows) + len(no_match_rows)
                        match_rate = (len(match_rows) / total_processed * 100) if total_processed > 0 else 0
                        
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Total Processed", f"{total_processed:,}")
                        with col2:
                            st.metric("Matched", f"{len(match_rows):,}")
                        with col3:
                            st.metric("No Match", f"{len(no_match_rows):,}")
                        with col4:
                            st.metric("Match Rate", f"{match_rate:.1f}%")
                        
                        # Tabs for results
                        tab1, tab2, tab3 = st.tabs(["📊 Summary", "✅ Matches", "❌ No Matches"])
                        
                        with tab1:
                            st.subheader("Processing Summary")
                            
                            # Match status distribution
                            if match_rows or no_match_rows:
                                status_data = {
                                    'Status': ['Matched', 'No Match'],
                                    'Count': [len(match_rows), len(no_match_rows)]
                                }
                                st.bar_chart(pd.DataFrame(status_data).set_index('Status'))
                            
                            # Rule distribution for matches
                            if match_rows:
                                match_df = pd.DataFrame(match_rows)
                                rule_counts = match_df['ParsedRule'].value_counts()
                                st.subheader("Match Rules Distribution")
                                st.bar_chart(rule_counts)
                        
                        with tab2:
                            if match_rows:
                                match_df = pd.DataFrame(match_rows)
                                st.subheader(f"✅ Matched Results ({len(match_rows):,} rows)")
                                st.dataframe(match_df.head(100), use_container_width=True)
                                
                                # Download button
                                excel_data = create_downloadable_excel(match_df, "matched_results.xlsx")
                                st.download_button(
                                    label="📥 Download Matched Results",
                                    data=excel_data,
                                    file_name=f"matched_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            else:
                                st.info("No matched results found.")
                        
                        with tab3:
                            if no_match_rows:
                                no_match_df = pd.DataFrame(no_match_rows)
                                st.subheader(f"❌ No Match Results ({len(no_match_rows):,} rows)")
                                st.dataframe(no_match_df.head(100), use_container_width=True)
                                
                                # Download button
                                excel_data = create_downloadable_excel(no_match_df, "no_match_results.xlsx")
                                st.download_button(
                                    label="📥 Download No Match Results",
                                    data=excel_data,
                                    file_name=f"no_match_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            else:
                                st.info("No unmatched results found.")
                        
                        # Download all results as ZIP
                        if match_rows or no_match_rows:
                            st.subheader("📦 Download All Results")
                            
                            # Create ZIP file
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                if match_rows:
                                    match_excel = create_downloadable_excel(pd.DataFrame(match_rows), "matched_results.xlsx")
                                    zip_file.writestr("matched_results.xlsx", match_excel)
                                
                                if no_match_rows:
                                    no_match_excel = create_downloadable_excel(pd.DataFrame(no_match_rows), "no_match_results.xlsx")
                                    zip_file.writestr("no_match_results.xlsx", no_match_excel)
                                
                                # Add summary
                                summary = {
                                    'processing_completed': datetime.now().isoformat(),
                                    'total_processed_rows': total_processed,
                                    'matched_count': len(match_rows),
                                    'no_match_count': len(no_match_rows),
                                    'match_rate_percent': match_rate,
                                    'processing_time_seconds': processing_time
                                }
                                zip_file.writestr("summary.json", json.dumps(summary, indent=2))
                            
                            zip_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Download Complete Results Package (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name=f"resistance_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                mime="application/zip"
                            )
                
                except Exception as e:
                    st.error(f"❌ Error during processing: {str(e)}")
                    st.exception(e)

        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")
            st.exception(e)
    
    else:
        # Show instructions and examples
        st.info("👆 Upload an Excel file to get started")
        
        st.subheader("📖 How it Works")
        
        st.markdown("""
        This tool parses resistance codes from part numbers using multiple rules:
        
        **🔍 Supported Patterns:**
        - **R-Decimal**: `R047` → 0.047Ω, `47R2` → 47.2Ω
        - **4-Digit Rule 1**: `123K` → 123 × 10⁻² = 1.23Ω  
        - **4-Digit Rule 2**: `47B2` → 47.2Ω, `6H11` → 6.11KΩ
        - **Traditional**: `472` → 47 × 10² = 4700Ω
        
        **📊 Output Files:**
        - **Matched**: Parts where parsed value matches target value
        - **No Match**: Parts with no pattern or no good match
        """)
        
        st.subheader("📋 Required Columns")
        st.markdown("""
        Your Excel file should contain:
        - **PartNumber** (required): The part number to parse
        - **Value** (optional): Target resistance value for matching
        - Other columns will be preserved in output
        """)
        
        st.subheader("🧪 Example Patterns")
        examples_df = pd.DataFrame([
            {"Part Number": "MCR100JZHJSR047", "Expected": "0.047 Ω", "Rule": "R-Decimal"},
            {"Part Number": "PA1206FRE470R012Z", "Expected": "0.012 Ω", "Rule": "R-Decimal"}, 
            {"Part Number": "CHP1206L75R0JNT", "Expected": "75.0 Ω", "Rule": "R-Decimal"},
            {"Part Number": "TEST123K456", "Expected": "1.23 Ω", "Rule": "4-Digit Rule 1"},
            {"Part Number": "ABC47B2DEF", "Expected": "47.2 Ω", "Rule": "4-Digit Rule 2"},
            {"Part Number": "XYZ4701GHI", "Expected": "4700 Ω", "Rule": "Traditional"}
        ])
        st.dataframe(examples_df, use_container_width=True)

if __name__ == "__main__":
    main()
