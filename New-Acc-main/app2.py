#!/usr/bin/env python3

"""

run_pipeline.py  –  End-to-end driver for the ACC Unified Pipeline

(Colab-friendly version that uses the *new* prefix-aware modules)



Steps:

  1. Ensure mapping.xlsx (local or via GitHub API download)

  2. Pre-process the input (robust CSV/XLSX handling)

  3. Run the *fixed* processing pipeline (prefix-aware wrapper)

  4. Run the *detailed* analysis pipeline (prefix-aware wrapper)

  5. Combine fixed + detailed outputs into final report

"""



#───────────────────────────────────────────────────────────────────────────

# 0  Colab-only pip installs  (harmless elsewhere)

#───────────────────────────────────────────────────────────────────────────

try:

    import google.colab  # noqa: F401

    IN_COLAB = True

except ImportError:

    IN_COLAB = False



if IN_COLAB:

    # Quiet install to avoid extra output in Colab

    # %pip install -q pandas openpyxl streamlit pdfplumber requests pyngrok
    # %pip install -q --ignore-installed pandas openpyxl streamlit pdfplumber requests pyngrok
    %pip install -q --ignore-installed blinker
    %pip install -q "pandas==2.2.2" "requests==2.32.4" "tornado==6.5.1" openpyxl streamlit pdfplumber pyngrok



#───────────────────────────────────────────────────────────────────────────

# 1  Standard libs

#───────────────────────────────────────────────────────────────────────────

import os

import sys

import argparse

import pandas as pd



#───────────────────────────────────────────────────────────────────────────

# 2  Colab upload hook

#───────────────────────────────────────────────────────────────────────────

if IN_COLAB:

    from google.colab import files

    print("🔔  Colab detected – please upload your input file (CSV or XLSX)…")

    uploaded = files.upload()

    if not uploaded:

        sys.exit("No file uploaded; exiting.")

    # Fake argv so argparse sees only the uploaded filename

    sys.argv = [sys.argv[0], next(iter(uploaded.keys()))]



#───────────────────────────────────────────────────────────────────────────

# 3  Optional GitHub token (only needed for API download of mapping.xlsx)

#───────────────────────────────────────────────────────────────────────────

os.environ.setdefault("GITHUB_TOKEN", "")

os.environ.setdefault("GITHUB_OWNER", "Hima9791")

os.environ.setdefault("GITHUB_REPO",  "map")

os.environ.setdefault("GITHUB_FILE_PATH", "mapping.xlsx")



#───────────────────────────────────────────────────────────────────────────

# 4  Project-module imports  (all new wrappers included)

#───────────────────────────────────────────────────────────────────────────

from mapping_utils      import read_mapping_file, save_mapping_to_disk

from github_utils       import download_mapping_file_from_github

from preprocessor       import preprocess_input_file



# NEW prefix-aware wrappers

from pipeline_wrappers  import (

    run_fixed_pipeline_with_prefix_support,

    run_detailed_analysis_with_prefix_support

)



from result_combiner    import combine_results



#───────────────────────────────────────────────────────────────────────────

# 5  Helper: ensure mapping.xlsx exists (local or GitHub)

#───────────────────────────────────────────────────────────────────────────

def ensure_mapping(path: str, allow_github: bool) -> None:

    """Guarantee mapping.xlsx is present at *path* (download if allowed)."""

    if os.path.exists(path):

        print(f"[✔] Found mapping file at '{path}'.")

        return

    if not allow_github:

        sys.exit(f"[✖] mapping file '{path}' not found. "

                 f"Run with --use-github to download it.")

    print("[ ] mapping.xlsx not found – downloading via GitHub API…")

    df_map = download_mapping_file_from_github()

    save_mapping_to_disk(df_map, path)

    print(f"[✔] mapping.xlsx saved to '{path}'.")



#───────────────────────────────────────────────────────────────────────────

# 6  Main driver

#───────────────────────────────────────────────────────────────────────────

def main() -> None:

    # ── CLI args ──────────────────────────────────────────────────

    ap = argparse.ArgumentParser(description="ACC Unified Pipeline runner")

    ap.add_argument("input_excel", help="Input file (CSV or XLSX with 'Value' column)")

    ap.add_argument("--mapping", default="mapping.xlsx",

                    help="Local mapping file name (default: mapping.xlsx)")

    ap.add_argument("--use-github", action="store_true",

                    help="If mapping.xlsx missing, download via GitHub API")

    ap.add_argument("--workdir", default=".",

                    help="Working directory for outputs")

    args, _ = ap.parse_known_args()



    # ── Convert CSV→XLSX if needed ───────────────────────────────

    input_path = args.input_excel

    if input_path.lower().endswith(".csv"):

        print(f"[ ] Detected CSV '{input_path}' – converting to Excel…")

        df_csv = pd.read_csv(input_path)

        input_path = input_path.rsplit(".", 1)[0] + "_converted.xlsx"

        df_csv.to_excel(input_path, index=False, engine="openpyxl")

        print(f"[✔] Saved converted file '{input_path}'")



    # ── Working directory + mapping ───────────────────────────────

    wd = args.workdir.rstrip("/")

    os.makedirs(wd, exist_ok=True)

    mapping_path = os.path.join(wd, args.mapping)

    ensure_mapping(mapping_path, args.use_github)



    #───────────────────────────────────────────────────────────────

    # 1  Pre-processing

    #───────────────────────────────────────────────────────────────

    print("[ ] Pre-processing input…")

    try:

        df_pre = preprocess_input_file(input_path)

    except Exception as e:

        sys.exit(f"[✖] Pre-processing failed: {e}")

    pre_xlsx = os.path.join(wd, "preprocessed.xlsx")

    df_pre.to_excel(pre_xlsx, index=False, engine="openpyxl")

    print(f"[✔] Preprocessed data → {pre_xlsx}")



    #───────────────────────────────────────────────────────────────

    # 2  Fixed pipeline  (prefix-aware)

    #───────────────────────────────────────────────────────────────

    print("[ ] Running fixed pipeline (prefix support)…")

    with open(pre_xlsx, "rb") as fh:

        df_fixed = run_fixed_pipeline_with_prefix_support(fh.read(), mapping_path)

    fixed_xlsx = os.path.join(wd, "fixed_output.xlsx")

    df_fixed.to_excel(fixed_xlsx, index=False, engine="openpyxl")

    print(f"[✔] Fixed output → {fixed_xlsx}")



    #───────────────────────────────────────────────────────────────

    # 3  Detailed analysis  (prefix-aware)

    #───────────────────────────────────────────────────────────────

    print("[ ] Running detailed analysis (prefix support)…")

    detailed_xlsx = os.path.join(wd, "detailed_analysis.xlsx")

    ok = run_detailed_analysis_with_prefix_support(

        input_df=df_pre,

        mapping_file=mapping_path,

        output_file=detailed_xlsx

    )

    if ok is None:

        sys.exit("[✖] Detailed analysis failed.")

    print(f"[✔] Detailed output → {detailed_xlsx}")



    #───────────────────────────────────────────────────────────────

    # 4  Combine fixed + detailed

    #───────────────────────────────────────────────────────────────

    print("[ ] Combining results…")

    final_xlsx = os.path.join(wd, "final_combined.xlsx")

    combined_ok = combine_results(

        processed_df=df_fixed,

        analysis_file=detailed_xlsx,

        output_file=final_xlsx

    )

    if combined_ok is None:

        sys.exit("[✖] Combining results failed.")

    print(f"[✔] Final combined report → {final_xlsx}")



    #───────────────────────────────────────────────────────────────

    # 5  Summary

    #───────────────────────────────────────────────────────────────

    print("\n🎉  Pipeline completed successfully!")

    for label, path in [

        ("Preprocessed",      pre_xlsx),

        ("Fixed pipeline",    fixed_xlsx),

        ("Detailed analysis", detailed_xlsx),

        ("Final report",      final_xlsx)

    ]:

        print(f"   • {label}: {path}")



#───────────────────────────────────────────────────────────────────────────

# 7  Entrypoint

#───────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    main()
