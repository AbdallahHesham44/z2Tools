import streamlit as st
import pandas as pd
import numpy as np
import io
import time
from datetime import datetime
import gc

st.set_page_config(page_title="Excel Masking Tool", layout="wide")
st.title("🔒 Excel Part Number Masking Tool")

uploaded_file = st.file_uploader("📤 Upload Excel file", type=["xlsx"])

# Replacement function
def apply_two_stage_replacement(df_row, port_key, replacement="__"):
    original_part = str(df_row["ZPartNumber"])
    if pd.isna(original_part) or original_part in ["", "nan"]:
        return original_part, "0", False

    occurrences = []
    start = 0
    while True:
        pos = original_part.find(str(port_key), start)
        if pos == -1:
            break
        occurrences.append(pos)
        start = pos + 1

    if len(occurrences) >= 2:
        second_occurrence_index = occurrences[1]
        masked_part = (
            original_part[:second_occurrence_index]
            + replacement
            + original_part[second_occurrence_index + len(str(port_key)) :]
        )
        return masked_part, second_occurrence_index, True

    elif len(occurrences) == 1:
        first_occurrence_index = occurrences[0]
        masked_part = original_part.replace(str(port_key), replacement, 1)
        return masked_part, first_occurrence_index, False

    else:
        return original_part, "0", False


if uploaded_file:
    try:
        start_time = time.time()

        with pd.ExcelFile(uploaded_file, engine="openpyxl") as excel_file:
            df_portion = pd.read_excel(excel_file, sheet_name="PC_portion")
            df_B1 = pd.read_excel(excel_file, sheet_name="B1")
            df_B2 = pd.read_excel(excel_file, sheet_name="B2")

        dataframes = {"B1": df_B1, "B2": df_B2}

        for sheet_name, df in dataframes.items():
            for col in ["masked_code", "maskMatch", "PortionStart"]:
                if col not in df.columns:
                    df[col] = np.nan

        # Clean values
        for col in ["PortionName", "Family", "PortionKey", "FamilyName", "ZPartNumber"]:
            if col in df_portion.columns:
                df_portion[col] = df_portion[col].fillna("").astype(str).str.strip()
            for df in dataframes.values():
                if col in df.columns:
                    df[col] = df[col].fillna("").astype(str).str.strip()

        for df in dataframes.values():
            if "Part Mask" in df.columns:
                df["Part Mask"] = df["Part Mask"].fillna("").astype(str).str.strip()

        df_packaging = df_portion[df_portion["PortionName"] == "Packaging"]
        family_list = df_packaging["Family"].unique().tolist()
        total_families = len(family_list)

        progress_bar = st.progress(0)
        log_box = st.empty()

        for i, family in enumerate(family_list, 1):
            portKeys = (
                df_packaging[df_packaging["Family"] == family]["PortionKey"]
                .dropna()
                .unique()
                .tolist()
            )
            portKeys = sorted(portKeys, key=lambda x: len(str(x)), reverse=True)

            for sheet_name, df in dataframes.items():
                mask_family = df["FamilyName"] == family

                for portKey in portKeys:
                    if not portKey or str(portKey).strip() == "" or str(portKey) == "nan":
                        continue

                    mask_key = (
                        mask_family
                        & df["ZPartNumber"].astype(str).str.contains(str(portKey), na=False)
                        & (
                            df["masked_code"].isna()
                            | (df["masked_code"] == "")
                            | (df["masked_code"] == "nan")
                        )
                    )

                    for row_idx in df[mask_key].index:
                        masked_code, portion_start, is_multi_portion = (
                            apply_two_stage_replacement(df.loc[row_idx], portKey)
                        )

                        df.loc[row_idx, "masked_code"] = masked_code
                        df.loc[row_idx, "PortionStart"] = (
                            "MultiPortion" if is_multi_portion else portion_start
                        )

                        part_mask_val = str(df.loc[row_idx, "Part Mask"])
                        df.loc[row_idx, "maskMatch"] = (
                            "maskMatch"
                            if str(masked_code) == part_mask_val
                            else "NotMatch"
                        )

                # Handle no portion rows
                no_portion_mask = mask_family & (
                    df["masked_code"].isna()
                    | (df["masked_code"] == "")
                    | (df["masked_code"] == "nan")
                )
                df.loc[no_portion_mask, "masked_code"] = df.loc[
                    no_portion_mask, "ZPartNumber"
                ]
                df.loc[no_portion_mask, "PortionStart"] = "NoPacking"
                for row_idx in df[no_portion_mask].index:
                    if df.loc[row_idx, "ZPartNumber"] == df.loc[row_idx, "Part Mask"]:
                        df.loc[row_idx, "maskMatch"] = "maskMatch"
                    else:
                        df.loc[row_idx, "maskMatch"] = "NotMatch"

            progress_bar.progress(i / total_families)
            log_box.text(f"✅ Processed {i}/{total_families} families")

            gc.collect()

        # Save results to downloadable Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_portion.to_excel(writer, sheet_name="PC_portion", index=False)
            dataframes["B1"].to_excel(writer, sheet_name="B1", index=False)
            dataframes["B2"].to_excel(writer, sheet_name="B2", index=False)
        output.seek(0)

        st.success("🎉 Processing complete!")
        st.download_button(
            label="💾 Download Processed Excel",
            data=output,
            file_name=f"processed_mask_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Show preview
        st.subheader("📋 Preview of Processed Data (B1)")
        st.dataframe(dataframes["B1"].head(20))

    except Exception as e:
        st.error(f"❌ Error: {e}")
