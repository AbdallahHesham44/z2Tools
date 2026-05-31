import io
import re
from urllib.parse import urlparse, unquote

import pandas as pd
import streamlit as st


GENERIC_PDF_RE = re.compile(r"^pdf\s*file[_-]?\d+\.pdf$|^pdffile[_-]?\d+\.pdf$", re.IGNORECASE)


def extract_filename(value: object) -> str:
    """Return the PDF/file name from a Local URL cell."""
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if not text:
        return ""

    parsed = urlparse(text)
    path = parsed.path if parsed.path else text
    filename = path.rstrip("/").split("/")[-1]
    return unquote(filename).strip()


def is_generic_pdf_filename(filename: str) -> bool:
    """Detect strange generic names like PdfFile212651.pdf or PdfFile_232283.pdf."""
    if not filename:
        return False

    compact = filename.replace(" ", "")
    return bool(GENERIC_PDF_RE.match(compact))


def flag_reason(local_url: object) -> str:
    filename = extract_filename(local_url)
    if is_generic_pdf_filename(filename):
        return "FLAG: generic PDF filename"
    return ""


def process_workbook(uploaded_file) -> tuple[dict[str, pd.DataFrame], dict[str, pd.DataFrame], int]:
    """
    Reads all sheets. For every sheet that contains Local URL, inserts URL Flag
    immediately beside it. Returns original sheets, processed sheets, flag count.
    """
    original_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    processed_sheets = {}
    total_flags = 0

    for sheet_name, df in original_sheets.items():
        processed = df.copy()

        if "Local URL" in processed.columns:
            local_url_position = processed.columns.get_loc("Local URL")

            # Remove old flag column if user re-uploads an already flagged file.
            if "URL Flag" in processed.columns:
                processed = processed.drop(columns=["URL Flag"])
                local_url_position = processed.columns.get_loc("Local URL")

            flags = processed["Local URL"].apply(flag_reason)
            total_flags += int((flags != "").sum())

            processed.insert(local_url_position + 1, "URL Flag", flags)

        processed_sheets[sheet_name] = processed

    return original_sheets, processed_sheets, total_flags


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            # Excel sheet names are limited to 31 chars.
            safe_sheet_name = str(sheet_name)[:31]
            df.to_excel(writer, index=False, sheet_name=safe_sheet_name)

            worksheet = writer.sheets[safe_sheet_name]
            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, len(value))
                worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 55)

    return output.getvalue()


st.set_page_config(page_title="Local URL Datasheet Flagger", page_icon="📄", layout="wide")

st.title("Local URL Datasheet Flagger")
st.write(
    "Upload an Excel file. The app finds the `Local URL` column, detects generic PDF filenames "
    "such as `PdfFile212651.pdf` or `PdfFile_232283.pdf`, adds `URL Flag` beside it, "
    "and lets you download the updated workbook."
)

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        original_sheets, processed_sheets, total_flags = process_workbook(uploaded_file)

        st.success(f"Done. Found {total_flags} strange Local URL value(s).")

        sheet_names_with_local_url = [
            name for name, df in processed_sheets.items() if "Local URL" in df.columns
        ]

        if not sheet_names_with_local_url:
            st.warning("No column named `Local URL` was found in this workbook.")
        else:
            selected_sheet = st.selectbox("Preview sheet", sheet_names_with_local_url)
            preview_df = processed_sheets[selected_sheet]

            local_url_position = preview_df.columns.get_loc("Local URL")
            preview_columns = preview_df.columns[
                max(0, local_url_position - 3): min(len(preview_df.columns), local_url_position + 5)
            ]
            st.dataframe(preview_df.loc[:, preview_columns], use_container_width=True)

            flagged_rows = preview_df[preview_df["URL Flag"] != ""]
            if not flagged_rows.empty:
                st.subheader("Flagged rows")
                st.dataframe(flagged_rows.loc[:, preview_columns], use_container_width=True)

        excel_bytes = to_excel_bytes(processed_sheets)
        output_name = uploaded_file.name.rsplit(".", 1)[0] + "_flagged.xlsx"

        st.download_button(
            label="Download flagged Excel file",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as exc:
        st.error(f"Could not process the file: {exc}")
