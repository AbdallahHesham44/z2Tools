# CapacitorValueMatcher.py
import pandas as pd
import numpy as np
import re
import os
import json
import threading
import pickle
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

class CapacitorValueMatcher:
    def __init__(self, input_file_path, output_dir="output", batch_size=10000, num_threads=4, checkpoint_interval=5000, progress_callback=None):
        self.input_file_path = input_file_path
        self.output_dir = output_dir
        self.batch_size = batch_size
        self.num_threads = num_threads
        self.checkpoint_interval = checkpoint_interval
        self.progress_callback = progress_callback

        os.makedirs(output_dir, exist_ok=True)

        self.checkpoint_file = os.path.join(output_dir, "checkpoint.json")
        self.matched_temp_file = os.path.join(output_dir, "matched_temp.pkl")
        self.unmatched_temp_file = os.path.join(output_dir, "unmatched_temp.pkl")

        self.processed_rows = 0
        self.matched_results = []
        self.unmatched_results = []

        self.lock = threading.Lock()

    def load_checkpoint(self):
        if os.path.exists(self.checkpoint_file):
            try:
                with open(self.checkpoint_file, 'r') as f:
                    checkpoint = json.load(f)
                self.processed_rows = checkpoint.get('processed_rows', 0)

                if os.path.exists(self.matched_temp_file):
                    with open(self.matched_temp_file, 'rb') as f:
                        self.matched_results = pickle.load(f)

                if os.path.exists(self.unmatched_temp_file):
                    with open(self.unmatched_temp_file, 'rb') as f:
                        self.unmatched_results = pickle.load(f)
                return True
            except:
                return False
        return False

    def save_checkpoint(self):
        checkpoint = {
            'processed_rows': self.processed_rows,
            'timestamp': datetime.now().isoformat(),
            'matched_count': len(self.matched_results),
            'unmatched_count': len(self.unmatched_results)
        }
        with open(self.checkpoint_file, 'w') as f:
            json.dump(checkpoint, f, indent=2)
        with open(self.matched_temp_file, 'wb') as f:
            pickle.dump(self.matched_results, f)
        with open(self.unmatched_temp_file, 'wb') as f:
            pickle.dump(self.unmatched_results, f)

    def extract_patterns(self, part_number):
        if pd.isna(part_number) or not isinstance(part_number, str):
            return []
        patterns = []
        for i in range(len(part_number) - 2):
            substring = part_number[i:i+3]
            if substring.isdigit():
                patterns.append(substring)
        for i in range(len(part_number) - 3):
            substring = part_number[i:i+4]
            if substring.isdigit():
                patterns.append(substring)
        patterns.extend(re.findall(r'\d*R\d*', part_number))
        patterns.extend(re.findall(r'R\d+', part_number))
        return list(set(patterns))

    def calculate_values(self, pattern):
        calculated_values = []
        if 'R' in pattern:
            if pattern.startswith('R'):
                numeric_part = pattern[1:]
                if numeric_part:
                    decimal_value = float('0.' + numeric_part)
                    calculated_values.append(decimal_value)
            else:
                parts = pattern.split('R')
                if len(parts) == 2:
                    before_r = parts[0] if parts[0] else '0'
                    after_r = parts[1] if parts[1] else '0'
                    decimal_value = float(before_r + '.' + after_r)
                    calculated_values.append(decimal_value)
        elif pattern.isdigit() and len(pattern) in [3, 4]:
            digits = pattern
            if len(digits) >= 2:
                first_digits = int(digits[:-1])
                last_digit = int(digits[-1])
                calculated_values.append(first_digits * (10 ** last_digit))
                if last_digit == 7:
                    calculated_values.append(first_digits * (10 ** -1))
                    calculated_values.append(first_digits * (10 ** -3))
                elif last_digit == 8:
                    calculated_values.append(first_digits * (10 ** -2))
                elif last_digit == 9:
                    calculated_values.append(first_digits * (10 ** -1))
                    calculated_values.append(first_digits * (10 ** -3))
                first_digit = int(digits[0])
                remaining_digits = int(digits[1:])
                calculated_values.append(remaining_digits * (10 ** first_digit))
        return calculated_values

    def convert_to_pf(self, value, unit):
        if pd.isna(value) or value == 0:
            return 0
        unit = str(unit).lower().strip()
        conversion_factors = {
            'pf': 1,
            'nf': 1000,
            'uf': 1000000,
            'µf': 1000000,
            'mf': 1000000000,
            'f': 1000000000000
        }
        return value * conversion_factors.get(unit, 1)

    def parse_value_column(self, value_str):
        if pd.isna(value_str):
            return 0, 'pf'
        value_str = str(value_str).strip()
        numeric_match = re.search(r'[\d.]+', value_str)
        if not numeric_match:
            return 0, 'pf'
        numeric_value = float(numeric_match.group())
        unit_match = re.search(r'[a-zA-Zµ]+', value_str)
        unit = unit_match.group().lower() if unit_match else 'pf'
        return numeric_value, unit

    def generate_unit_variants(self, pf_value):
        variants = [(pf_value, 'pf')]
        if pf_value != 0:
            uf_value = pf_value / 1000000
            variants.append((uf_value, 'uf'))
        return variants

    def process_single_row(self, row_data):
        row_index, row = row_data
        part_number = str(row.get('PartNumber', ''))
        target_value_str = row.get('Value', '')
        patterns = self.extract_patterns(part_number)
        if not patterns:
            result_row = row.copy()
            result_row['ExValue'] = 'No patterns found'
            result_row['Status'] = 'no_match'
            return ('unmatched', result_row)
        all_calculated_values = []
        for pattern in patterns:
            all_calculated_values.extend(self.calculate_values(pattern))
        if not all_calculated_values:
            result_row = row.copy()
            result_row['ExValue'] = 'No values calculated'
            result_row['Status'] = 'no_match'
            return ('unmatched', result_row)
        target_numeric, target_unit = self.parse_value_column(target_value_str)
        target_pf = self.convert_to_pf(target_numeric, target_unit)
        for calc_value in all_calculated_values:
            for variant_value, variant_unit in self.generate_unit_variants(calc_value):
                variant_pf = self.convert_to_pf(variant_value, variant_unit)
                if abs(variant_pf - target_pf) < max(1e-6, target_pf * 1e-10):
                    result_row = row.copy()
                    result_row['ExValue'] = f"{variant_value} {variant_unit}"
                    result_row['Status'] = 'match'
                    return ('matched', result_row)
        result_row = row.copy()
        result_row['ExValue'] = f"Calc:{[f'{v:.6g}' for v in all_calculated_values[:5]]} vs Target:{target_pf:.6g}pF"
        result_row['Status'] = 'no_match'
        return ('unmatched', result_row)

    def process_batch(self, batch_df):
        batch_matched = []
        batch_unmatched = []
        row_data = [(idx, row) for idx, row in batch_df.iterrows()]
        with ThreadPoolExecutor(max_workers=self.num_threads) as executor:
            futures = {executor.submit(self.process_single_row, rd): rd for rd in row_data}
            for future in as_completed(futures):
                try:
                    result_type, result_row = future.result()
                    if result_type == 'matched':
                        batch_matched.append(result_row)
                    else:
                        batch_unmatched.append(result_row)
                except Exception as e:
                    batch_unmatched.append({'Error': str(e), 'ExValue': 'Error', 'Status': 'error'})
        return batch_matched, batch_unmatched

    def combine_batch_files(self):
        matched_files = [f for f in os.listdir(self.output_dir) if f.startswith('matched_batch_')]
        if matched_files:
            matched_dfs = [pd.read_excel(os.path.join(self.output_dir, f)) for f in matched_files]
            final_matched = pd.concat(matched_dfs, ignore_index=True)
            final_matched.to_excel(os.path.join(self.output_dir, "MatchedOutput.xlsx"), index=False)
            for f in matched_files:
                os.remove(os.path.join(self.output_dir, f))

        unmatched_files = [f for f in os.listdir(self.output_dir) if f.startswith('unmatched_batch_')]
        if unmatched_files:
            unmatched_dfs = [pd.read_excel(os.path.join(self.output_dir, f)) for f in unmatched_files]
            final_unmatched = pd.concat(unmatched_dfs, ignore_index=True)
            final_unmatched.to_excel(os.path.join(self.output_dir, "notMatchedOutput.xlsx"), index=False)
            for f in unmatched_files:
                os.remove(os.path.join(self.output_dir, f))

    def process_file(self):
        self.load_checkpoint()
        df = pd.read_excel(self.input_file_path)
        total_rows = len(df)
        if self.processed_rows > 0:
            df = df.iloc[self.processed_rows:].reset_index(drop=True)
        batch_num = self.processed_rows // self.batch_size
        for start in range(0, len(df), self.batch_size):
            end = min(start + self.batch_size, len(df))
            batch_df = df.iloc[start:end].copy()
            batch_matched, batch_unmatched = self.process_batch(batch_df)
            if batch_matched:
                pd.DataFrame(batch_matched).to_excel(os.path.join(self.output_dir, f"matched_batch_{batch_num}.xlsx"), index=False)
            if batch_unmatched:
                pd.DataFrame(batch_unmatched).to_excel(os.path.join(self.output_dir, f"unmatched_batch_{batch_num}.xlsx"), index=False)
            with self.lock:
                self.processed_rows += len(batch_df)
                self.matched_results.extend(batch_matched)
                self.unmatched_results.extend(batch_unmatched)
            if (self.processed_rows) % self.checkpoint_interval == 0:
                self.save_checkpoint()
            if self.progress_callback:
                self.progress_callback(self.processed_rows, total_rows, batch_num + 1)
            batch_num += 1
        self.combine_batch_files()
        for temp_file in [self.checkpoint_file, self.matched_temp_file, self.unmatched_temp_file]:
            if os.path.exists(temp_file):
                os.remove(temp_file)
