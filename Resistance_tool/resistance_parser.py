import streamlit as st
import pandas as pd
import re
import time
import json
import os
import tempfile
import zipfile
from datetime import datetime, timedelta
from io import BytesIO
import gc
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="Enhanced Resistance Code Parser",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

class EnhancedResistanceParser:
    def __init__(self, batch_size=1000, checkpoint_interval=5000):
        """Initialize the parser with configuration"""
        self.batch_size = batch_size
        self.checkpoint_interval = checkpoint_interval
        
        # Progress tracking
        self.start_time = None
        self.processed_count = 0
        self.total_rows = 0
        self.matched_results = []
        self.no_match_results = []

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

        # Multiplier letters with their values
        self.multiplier_letters = {
            'K': 1e3,    # Kilo - 1,000
            'M': 1e6,    # Mega - 1,000,000
            'G': 1e9,    # Giga - 1,000,000,000
            'T': 1e12    # Tera - 1,000,000,000,000
        }

    def parse_multiplier_decimal_patterns(self, code):
        """Parse patterns where K, M, G, T act as decimal points with multipliers"""
        results = []

        # Pattern 1: digits + multiplier (82K, 47M)
        for i in range(len(code) - 1):
            for length in range(2, min(5, len(code) - i + 1)):
                substring = code[i:i+length]
                match = re.match(r'^(\d{1,3})([KMG])$', substring)
                if match:
                    digits, multiplier_char = match.groups()
                    base_value = float(digits)
                    multiplier = self.multiplier_letters[multiplier_char]
                    value = base_value * multiplier

                    results.append({
                        'pattern': substring,
                        'type': 'multiplier-decimal',
                        'rule': f'Multiplier-Trailing-{multiplier_char}',
                        'value': value,
                        'unit': 'Ohm',
                        'position': i,
                        'base_digits': digits,
                        'multiplier_char': multiplier_char,
                        'multiplier_value': multiplier
                    })

        # Pattern 2: digits + multiplier + digits (3K3, 2M20, 47K5)
        for i in range(len(code) - 2):
            for length in range(3, min(6, len(code) - i + 1)):
                substring = code[i:i+length]
                match = re.match(r'^(\d{1,3})([KMG])(\d{1,3})$', substring)
                if match:
                    before, multiplier_char, after = match.groups()
                    base_value = float(f"{before}.{after}")
                    multiplier = self.multiplier_letters[multiplier_char]
                    value = base_value * multiplier

                    results.append({
                        'pattern': substring,
                        'type': 'multiplier-decimal',
                        'rule': f'Multiplier-Decimal-{multiplier_char}',
                        'value': value,
                        'unit': 'Ohm',
                        'position': i,
                        'base_digits': f"{before}.{after}",
                        'multiplier_char': multiplier_char,
                        'multiplier_value': multiplier
                    })

        return results

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

        return results

    def parse_4digit_rule1(self, code):
        """Rule 1: 4-digit pattern with single character (decimal multiplier)"""
        results = []

        for i in range(len(code) - 3):
            substring = code[i:i+4]

            # Pattern: 3 digits + 1 character
            match = re.match(r'^(\d{3})([A-Z])$', substring)
            if match:
                digits, char = match.groups()
                if char in self.rule1_multipliers:
                    base_value = float(digits)
                    multiplier = self.rule1_multipliers[char]
                    value = base_value * multiplier

                    results.append({
                        'pattern': substring,
                        'type': '4-digit-rule1',
                        'rule': f'Rule1-{char}',
                        'value': value,
                        'unit': 'Ohm',
                        'position': i,
                        'base_digits': digits,
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

                    # Multiplier mapping
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
        results.extend(self.parse_multiplier_decimal_patterns(code))
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

        value_str = str(value_str).replace(" ", "").replace("Ω", "Ohm")
        match = re.match(r'^([0-9]*\.?[0-9]+)([a-zA-Z]+)$', value_str)
        if not match:
            return None

        val_str, unit = match.groups()

        try:
            val = float(val_str)
        except ValueError:
            return None

        unit_lower = unit.lower()

        unit_multipliers = {
            'ohm': 1,
            'kohm': 1e3,
            'mohm': 1e6,
            'lohm': 1e-2,
        }

        # Handle case sensitivity for Mega vs milli
        if unit == 'MOhm':
            multiplier = 1e6
        elif unit == 'mOhm':
            multiplier = 1e-3
        else:
            multiplier = unit_multipliers.get(unit_lower, None)

        if multiplier is None:
            return None

        return val * multiplier

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

                # Exact match gets priority
                if difference < 1e-10:
                    return result

                # Or very close match (within 0.1% relative difference)
                if relative_diff < 0.001 and difference < min_difference:
                    min_difference = difference
                    best_match = result

            except (ValueError, TypeError):
                continue

        return best_match

    def process_dataframe(self, df, progress_callback=None):
        """Process the entire dataframe and separate into match/no-match"""
        self.matched_results = []
        self.no_match_results = []
        self.total_rows = len(df)
        self.start_time = time.time()

        for idx, row in df.iterrows():
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
                # No patterns found
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
                self.no_match_results.append(no_match_row)
                continue

            # Try to find best match
            best_match = self.find_best_match(parsed, converted_val)

            if best_match:
                # Found a match
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
                self.matched_results.append(match_row)
            else:
                # No good match found
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
                self.no_match_results.append(no_match_row)

            # Update progress
            if progress_callback and idx % 100 == 0:
                progress_callback(idx + 1, self.total_rows)

        return self.matched_results, self.no_match_results


def main():
    st.title("⚡ Enhanced Resistance Code Parser")
    st.markdown("Upload an Excel file with resistance part numbers to extract and match resistance values")

    # Sidebar configuration
    st.sidebar.header("Configuration")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload an Excel file containing part numbers and resistance values"
    )

    # Processing settings
    st.sidebar.subheader("Processing Settings")
    batch_size = st.sidebar.slider("Batch Size", 100, 10000, 1000, 100)
    
    # Pattern testing section
    st.sidebar.subheader("Pattern Testing")
    test_part = st.sidebar.text_input("Test Part Number", "AF0603FR-0782KL")
    test_value = st.sidebar.text_input("Expected Value", "82 Kohm")
    
    if st.sidebar.button("Test Pattern"):
        parser = EnhancedResistanceParser()
        results = parser.parse_all_resistance_codes_enhanced(test_part)
        target_ohm = parser.convert_to_ohm(test_value)
        
        st.sidebar.markdown("**Test Results:**")
        if results:
            best_match = parser.find_best_match(results, target_ohm)
            for result in results:
                st.sidebar.write(f"- {result['pattern']}: {result['value']:.2f} Ohm")
            if best_match:
                st.sidebar.success(f"Best match: {best_match['pattern']} -> {best_match['value']:.2f} Ohm")
            else:
                st.sidebar.warning("No good match found")
        else:
            st.sidebar.error("No patterns found")

    # Main content area
    if uploaded_file is not None:
        try:
            # Load data
            with st.spinner("Loading file..."):
                df = pd.read_excel(uploaded_file)
            
            st.success(f"File loaded successfully! {len(df):,} rows found")
            
            # Display file info
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", f"{len(df):,}")
            with col2:
                st.metric("Columns", len(df.columns))
            with col3:
                if 'PartNumber' in df.columns:
                    st.metric("Part Numbers", f"{df['PartNumber'].notna().sum():,}")
                else:
                    st.error("PartNumber column not found!")

            # Show data preview
            st.subheader("Data Preview")
            st.dataframe(df.head(10), use_container_width=True)

            # Processing section
            st.subheader("Processing")
            
            if st.button("Start Processing", type="primary"):
                if 'PartNumber' not in df.columns:
                    st.error("Required column 'PartNumber' not found in the file!")
                    return

                # Initialize parser
                parser = EnhancedResistanceParser(batch_size=batch_size)
                
                # Progress tracking
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                def update_progress(processed, total):
                    progress = processed / total
                    progress_bar.progress(progress)
                    status_text.text(f"Processing: {processed:,}/{total:,} rows ({progress*100:.1f}%)")

                # Process data
                start_time = time.time()
                with st.spinner("Processing resistance codes..."):
                    matched_results, no_match_results = parser.process_dataframe(df, update_progress)

                processing_time = time.time() - start_time
                
                # Display results
                st.success(f"Processing completed in {processing_time:.2f} seconds!")
                
                # Statistics
                total_processed = len(matched_results) + len(no_match_results)
                match_rate = (len(matched_results) / total_processed * 100) if total_processed > 0 else 0
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Processed", f"{total_processed:,}")
                with col2:
                    st.metric("Matched", f"{len(matched_results):,}")
                with col3:
                    st.metric("No Match", f"{len(no_match_results):,}")
                with col4:
                    st.metric("Match Rate", f"{match_rate:.1f}%")

                # Results tabs
                tab1, tab2, tab3 = st.tabs(["Matched Results", "No Match Results", "Download Files"])
                
                with tab1:
                    if matched_results:
                        df_matched = pd.DataFrame(matched_results)
                        st.subheader(f"Matched Results ({len(matched_results):,} rows)")
                        st.dataframe(df_matched, use_container_width=True)
                        
                        # Download button for matched results
                        buffer = BytesIO()
                        df_matched.to_excel(buffer, index=False)
                        st.download_button(
                            label="Download Matched Results (Excel)",
                            data=buffer.getvalue(),
                            file_name="extracted_resistances_MATCH.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.info("No matched results found")

                with tab2:
                    if no_match_results:
                        df_no_match = pd.DataFrame(no_match_results)
                        st.subheader(f"No Match Results ({len(no_match_results):,} rows)")
                        st.dataframe(df_no_match, use_container_width=True)
                        
                        # Download button for no-match results
                        buffer = BytesIO()
                        df_no_match.to_excel(buffer, index=False)
                        st.download_button(
                            label="Download No Match Results (Excel)",
                            data=buffer.getvalue(),
                            file_name="extracted_resistances_NotMATCH.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.info("No unmatched results found")

                with tab3:
                    st.subheader("Download All Results")
                    
                    # Create combined results
                    all_results = matched_results + no_match_results
                    if all_results:
                        df_all = pd.DataFrame(all_results)
                        
                        # Single combined file
                        buffer_all = BytesIO()
                        df_all.to_excel(buffer_all, index=False)
                        st.download_button(
                            label="Download All Results (Excel)",
                            data=buffer_all.getvalue(),
                            file_name="extracted_resistances_ALL.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        # ZIP file with separate files
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            if matched_results:
                                matched_buffer = BytesIO()
                                pd.DataFrame(matched_results).to_excel(matched_buffer, index=False)
                                zip_file.writestr("extracted_resistances_MATCH.xlsx", matched_buffer.getvalue())
                            
                            if no_match_results:
                                no_match_buffer = BytesIO()
                                pd.DataFrame(no_match_results).to_excel(no_match_buffer, index=False)
                                zip_file.writestr("extracted_resistances_NotMATCH.xlsx", no_match_buffer.getvalue())
                            
                            # Add summary
                            summary = {
                                'processing_completed': datetime.now().isoformat(),
                                'total_processed_rows': total_processed,
                                'matched_count': len(matched_results),
                                'no_match_count': len(no_match_results),
                                'match_rate_percent': match_rate
                            }
                            zip_file.writestr("processing_summary.json", json.dumps(summary, indent=2))
                        
                        st.download_button(
                            label="Download ZIP Package (All Files)",
                            data=zip_buffer.getvalue(),
                            file_name="resistance_parser_results.zip",
                            mime="application/zip"
                        )

        except Exception as e:
            st.error(f"Error loading file: {str(e)}")
            st.info("Please ensure the file is a valid Excel file with the required columns")

    else:
        # Instructions when no file is uploaded
        st.info("Please upload an Excel file to begin processing")
        
        st.subheader("Expected File Format")
        st.markdown("""
        Your Excel file should contain the following columns:
        - **PartNumber** (required): Part numbers containing resistance codes
        - **Value** (optional): Expected resistance values (e.g., "82 Kohm", "47 mOhm")
        - **CompanyName** (optional): Company information
        - **ProductLine** (optional): Product line information
        - **FeatureName** (optional): Feature information
        """)

        st.subheader("Supported Patterns")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **Multiplier Decimal Patterns:**
            - `82K` → 82,000 Ohm
            - `3K3` → 3,300 Ohm
            - `2M20` → 2,200,000 Ohm
            - `47K5` → 47,500 Ohm
            
            **R-Decimal Patterns:**
            - `R047` → 0.047 Ohm
            - `47R0` → 47.0 Ohm
            - `4R7` → 4.7 Ohm
            """)
        
        with col2:
            st.markdown("""
            **Traditional Patterns:**
            - `472` → 4,700 Ohm (47 × 10²)
            - `1003` → 100,000 Ohm (100 × 10³)
            
            **4-Digit Rules:**
            - Rule 1: Character multipliers (J, K, L, M, N, P)
            - Rule 2: Character as decimal point replacement
            """)

    # Footer
    st.markdown("---")
    st.markdown("**Enhanced Resistance Code Parser** - Extracts resistance values from part numbers using multiple parsing rules")


if __name__ == "__main__":
    main()
