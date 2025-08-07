
import pandas as pd
import re
import time
import json
import os
from datetime import datetime, timedelta
import gc
import warnings
warnings.filterwarnings('ignore')

class ResistanceParser:
    def __init__(self, input_file, output_dir, batch_size=1000, checkpoint_interval=5000):
        self.input_file = input_file
        self.output_dir = output_dir
        self.batch_size = batch_size
        self.checkpoint_interval = checkpoint_interval

        os.makedirs(output_dir, exist_ok=True)
        self.checkpoint_file = os.path.join(output_dir, "checkpoint.json")
        self.output_all = os.path.join(output_dir, "extracted_resistances_all.xlsx")
        self.output_best = os.path.join(output_dir, "extracted_resistances.xlsx")
        self.temp_dir = os.path.join(output_dir, "temp_batches")
        os.makedirs(self.temp_dir, exist_ok=True)
        self.start_time = None
        self.processed_count = 0
        self.total_rows = 0

    def save_checkpoint(self, processed_count, batch_files):
        checkpoint_data = {
            'processed_count': processed_count,
            'batch_files': batch_files,
            'timestamp': datetime.now().isoformat(),
            'total_rows': self.total_rows
        }
        with open(self.checkpoint_file, 'w') as f:
            json.dump(checkpoint_data, f, indent=2)
        print(f"💾 Checkpoint saved at row {processed_count}")

    def load_checkpoint(self):
        if os.path.exists(self.checkpoint_file):
            with open(self.checkpoint_file, 'r') as f:
                checkpoint_data = json.load(f)
            print(f"📂 Resuming from checkpoint at row {checkpoint_data['processed_count']}")
            return checkpoint_data
        return None

    def estimate_time_remaining(self, processed, total, elapsed_time):
        if processed == 0:
            return "Calculating..."
        rate = processed / elapsed_time
        remaining_rows = total - processed
        remaining_seconds = remaining_rows / rate
        return str(timedelta(seconds=int(remaining_seconds)))

    def parse_all_resistance_codes_overlapping(self, code):
        code = str(code).upper()
        results = []
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
                            'type': '3-digit',
                            'value': value,
                            'unit': 'Ohm',
                            'position': i
                        })
                except:
                    pass

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
                            'type': '4-digit',
                            'value': value,
                            'unit': 'Ohm',
                            'position': i
                        })
                except:
                    pass

        patterns_3char = [
            (r'^([RKML])(\d{2})$', 'letter-first'),
            (r'^(\d{2})([RKML])$', 'letter-last'),
            (r'^(\d{1})([RKML])(\d{1})$', 'letter-middle')
        ]

        for i in range(len(code) - 2):
            substring = code[i:i+3]
            for pattern, pattern_type in patterns_3char:
                match = re.match(pattern, substring)
                if match:
                    try:
                        if pattern_type == 'letter-first':
                            letter, digits = match.groups()
                            numeric = float(digits)
                        elif pattern_type == 'letter-last':
                            digits, letter = match.groups()
                            numeric = float(digits)
                        else:
                            before, letter, after = match.groups()
                            numeric = float(f"{before}.{after}")

                        if letter == 'L':
                            value, unit = numeric * 1e-2, "Ohm"
                        elif letter == 'R':
                            value, unit = numeric, "Ohm"
                        elif letter == 'K':
                            value, unit = numeric * 1e3, "Ohm"
                        elif letter == 'M':
                            value, unit = numeric * 1e6, "Ohm"
                        else:
                            continue

                        results.append({
                            'pattern': substring,
                            'type': '3-char-letter',
                            'value': value,
                            'unit': unit,
                            'position': i
                        })
                    except:
                        pass

        patterns_4char = [
            (r'^([RKML])(\d{3})$', 'letter-first'),
            (r'^(\d{2})([RKML])(\d{1})$', 'letter-middle'),
            (r'^(\d{3})([RKML])$', 'letter-last'),
            (r'^(\d+)([RKML])(\d+)$', 'mixed-pattern')
        ]

        for i in range(len(code) - 3):
            substring = code[i:i+4]
            for pattern, pattern_type in patterns_4char:
                match = re.match(pattern, substring)
                if match and len(substring) == 4:
                    try:
                        if pattern_type == 'letter-first':
                            letter, digits = match.groups()
                            numeric = float(digits)
                        elif pattern_type == 'letter-last':
                            digits, letter = match.groups()
                            numeric = float(digits)
                        else:
                            before, letter, after = match.groups()
                            numeric = float(f"{before}.{after}")

                        if letter == 'L':
                            value, unit = numeric * 1e-2, "Ohm"
                        elif letter == 'R':
                            value, unit = numeric, "Ohm"
                        elif letter == 'K':
                            value, unit = numeric * 1e3, "Ohm"
                        elif letter == 'M':
                            value, unit = numeric * 1e6, "Ohm"
                        else:
                            continue

                        results.append({
                            'pattern': substring,
                            'type': '4-char-letter',
                            'value': value,
                            'unit': unit,
                            'position': i
                        })
                    except:
                        pass

        seen = set()
        unique_results = []
        for result in results:
            key = (result['pattern'], result['type'])
            if key not in seen:
                seen.add(key)
                unique_results.append(result)
        return unique_results

    def convert_to_ohm(self, value_str):
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

    def process_batch(self, batch_df, batch_num):
        rows = []
        for _, row in batch_df.iterrows():
            part = str(row['PartNumber'])
            original_value = str(row.get('Value', '')).strip()
            converted_val = self.convert_to_ohm(original_value)
            parsed = self.parse_all_resistance_codes_overlapping(part)
            for result in parsed:
                rows.append({
                    "PartNumber": part,
                    "CompanyName": row.get("CompanyName", ""),
                    "ProductLine": row.get("ProductLine", ""),
                    "FeatureName": row.get("FeatureName", ""),
                    "OriginalValue": original_value,
                    "OriginalOhm": converted_val,
                    "ParsedPattern": result['pattern'],
                    "ParsedType": result['type'],
                    "ParsedValue": result['value'],
                    "ParsedUnit": result['unit'],
                    "Position": result['position']
                })

        batch_df_processed = pd.DataFrame(rows)
        batch_file = os.path.join(self.temp_dir, f"batch_{batch_num:06d}.parquet")
        batch_df_processed.to_parquet(batch_file, index=False)
        del batch_df_processed, rows
        gc.collect()
        return batch_file

    def combine_batches(self, batch_files):
        print("\n🔄 Combining batches into final files...")
        combined_dfs = []
        for batch_file in batch_files:
            if os.path.exists(batch_file):
                df_batch = pd.read_parquet(batch_file)
                combined_dfs.append(df_batch)
        if not combined_dfs:
            print("❌ No batch files found!")
            return
        df_all = pd.concat(combined_dfs, ignore_index=True)
        print("💾 Saving all patterns...")
        df_all.to_excel(self.output_all, index=False)
        print("🎯 Filtering best matches...")
        df_final = self.filter_best_matches(df_all)
        df_final.to_excel(self.output_best, index=False)
        self.cleanup_temp_files(batch_files)
        print(f"✅ Processing complete!")
        print(f"📁 All patterns: {self.output_all}")
        print(f"📁 Best matches: {self.output_best}")

    def filter_best_matches(self, df_all):
        def reduce_group(group):
            target = group.iloc[0]['OriginalOhm']
            if pd.isna(target):
                result = group.iloc[[0]].copy()
                result['status'] = 'notMatch'
                return result
            matches = group[abs(group['ParsedValue'] - target) < 0.1]
            if not matches.empty:
                result = matches.iloc[[0]].copy()
                result['status'] = 'match'
            else:
                result = group.iloc[[0]].copy()
                result['status'] = 'notMatch'
            return result
        return df_all.groupby('PartNumber', group_keys=False).apply(reduce_group)

    def cleanup_temp_files(self, batch_files):
        print("🧹 Cleaning up temporary files...")
        for batch_file in batch_files:
            if os.path.exists(batch_file):
                os.remove(batch_file)
        try:
            os.rmdir(self.temp_dir)
        except:
            pass

    def run(self):
        print("🚀 Starting enhanced resistance code parser...")
        print(f"📂 Input file: {self.input_file}")
        print(f"📁 Output directory: {self.output_dir}")
        checkpoint_data = self.load_checkpoint()
        start_row = 0
        batch_files = []
        if checkpoint_data:
            start_row = checkpoint_data['processed_count']
            batch_files = checkpoint_data.get('batch_files', [])
            self.total_rows = checkpoint_data['total_rows']
            print(f"🔄 Resuming from row {start_row}")

        if self.total_rows == 0:
            print("📊 Analyzing input file...")
            df_info = pd.read_excel(self.input_file, nrows=0)
            with pd.ExcelFile(self.input_file) as xls:
                self.total_rows = len(pd.read_excel(xls, usecols=[0]))

        print(f"📈 Total rows to process: {self.total_rows:,}")
        print(f"📦 Batch size: {self.batch_size:,}")
        print(f"💾 Checkpoint interval: {self.checkpoint_interval:,}")

        self.start_time = time.time()
        batch_num = len(batch_files)

        try:
            with pd.ExcelFile(self.input_file) as xls:
                for chunk_start in range(start_row, self.total_rows, self.batch_size):
                    chunk_end = min(chunk_start + self.batch_size, self.total_rows)
                    batch_df = pd.read_excel(
                        xls,
                        skiprows=range(1, chunk_start + 1) if chunk_start > 0 else None,
                        nrows=chunk_end - chunk_start
                    )
                    if batch_df.empty:
                        break
                    batch_df['PartNumber'] = batch_df['PartNumber'].astype(str)
                    batch_file = self.process_batch(batch_df, batch_num)
                    batch_files.append(batch_file)
                    batch_num += 1
                    self.processed_count = chunk_end
                    elapsed_time = time.time() - self.start_time
                    progress_pct = (self.processed_count / self.total_rows) * 100
                    eta = self.estimate_time_remaining(self.processed_count, self.total_rows, elapsed_time)
                    print(f"✅ Batch {batch_num:,} completed | "
                          f"Progress: {self.processed_count:,}/{self.total_rows:,} ({progress_pct:.1f}%) | "
                          f"ETA: {eta}")
                    if self.processed_count % self.checkpoint_interval == 0:
                        self.save_checkpoint(self.processed_count, batch_files)
                    del batch_df
                    gc.collect()
                    time.sleep(0.1)

        except KeyboardInterrupt:
            print("\n⚠️ Processing interrupted by user")
            self.save_checkpoint(self.processed_count, batch_files)
            print("💾 Progress saved. You can resume later.")
            return

        except Exception as e:
            print(f"\n❌ Error during processing: {str(e)}")
            self.save_checkpoint(self.processed_count, batch_files)
            print("💾 Progress saved. Check the error and resume later.")
            raise

        self.combine_batches(batch_files)

        if os.path.exists(self.checkpoint_file):
            os.remove(self.checkpoint_file)

        total_time = time.time() - self.start_time
        print(f"\n🎉 Processing completed in {timedelta(seconds=int(total_time))}")
        print(f"📊 Processed {self.processed_count:,} rows")
        print(f"⚡ Average speed: {self.processed_count/total_time:.1f} rows/second")
