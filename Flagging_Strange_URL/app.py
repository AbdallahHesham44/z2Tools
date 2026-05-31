import io
import re
from urllib.parse import urlparse, unquote

import pandas as pd
import requests
import streamlit as st
from pypdf import PdfReader

GENERIC_PDF_RE = re.compile(
    r"^pdf\s*file[_-]?\d+\.pdf$|^pdffile[_-]?\d+\.pdf$",
    re.IGNORECASE,
)

GENERATED_COLUMNS = [
    "URL Flag",
    "FamilyName Found",
    "FamilyName Match Source",
    "FamilyName Check Status",
    "Same Family Local Count",
    "Same Family Locals",
    "Same Local Family Count",
    "Families Sharing Same Local",
    "Local Family Relation Status",
    "FamilyName Found in Local URL",
]


def extract_filename(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    parsed = urlparse(text)
    path = parsed.path if parsed.path else text
    return unquote(path.rstrip("/").split("/")[-1]).strip()


def clean_display(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = unquote(str(value)).lower().strip()
    return re.sub(r"[^a-z0-9]+", "", text)


def normalize_key(value: object) -> str:
    return normalize_text(value)


def is_generic_pdf_filename(filename: str) -> bool:
    if not filename:
        return False
    return bool(GENERIC_PDF_RE.match(filename.replace(" ", "")))


def flag_reason(local_url: object) -> str:
    filename = extract_filename(local_url)
    if is_generic_pdf_filename(filename):
        return "FLAG: generic PDF filename"
    return ""


@st.cache_data(show_spinner=False, ttl=3600)
def download_and_extract_pdf_text(url: str, timeout: int = 30) -> tuple[str, str]:
    """Return (text, status). Status is used in the output sheet for troubleshooting."""
    if not url or not str(url).lower().startswith(("http://", "https://")):
        return "", "Skipped: not a URL"

    try:
        response = requests.get(str(url), timeout=timeout)
        response.raise_for_status()
    except Exception as exc:
        return "", f"Download error: {exc}"

    content_type = response.headers.get("content-type", "").lower()
    if "pdf" not in content_type and not str(url).lower().split("?")[0].endswith(".pdf"):
        return "", f"Skipped: not PDF ({content_type or 'unknown content type'})"

    try:
        reader = PdfReader(io.BytesIO(response.content))
        pages_text = []
        for page in reader.pages:
            pages_text.append(page.extract_text() or "")
        text = "\n".join(pages_text)
        if not text.strip():
            return "", "PDF text empty or scanned image"
        return text, "PDF text extracted"
    except Exception as exc:
        return "", f"PDF parse error: {exc}"


def family_match_result(family_name: object, local_url: object, search_pdf_text: bool) -> tuple[str, str, str]:
    """Return (Yes/No, source, status)."""
    family_normalized = normalize_text(family_name)
    if not family_normalized:
        return "No", "", "Missing FamilyName"

    url_normalized = normalize_text(local_url)
    if url_normalized and family_normalized in url_normalized:
        return "Yes", "Local URL", "Matched URL text"

    if not search_pdf_text:
        return "No", "", "Not found in URL text"

    pdf_text, status = download_and_extract_pdf_text(str(local_url).strip())
    pdf_normalized = normalize_text(pdf_text)
    if pdf_normalized and family_normalized in pdf_normalized:
        return "Yes", "PDF content", status

    return "No", "", status if status else "Not found"


def unique_join(values: list[str]) -> str:
    seen = set()
    output = []
    for value in values:
        text = clean_display(value)
        if not text:
            continue
        key = text.lower()
        if key not in seen:
            seen.add(key)
            output.append(text)
    return "\n".join(output)


def add_family_local_relations(processed: pd.DataFrame) -> pd.DataFrame:
    """Add relation columns between FamilyName and Local URL.

    Same Family Local Count / Same Family Locals:
        For each FamilyName, show all unique Local URLs used by that family.

    Same Local Family Count / Families Sharing Same Local:
        For each Local URL, show all unique FamilyName values using that same local.
        This helps detect when one PDF is reused across different families.
    """
    if "FamilyName" not in processed.columns or "Local URL" not in processed.columns:
        return processed

    df = processed.copy()
    family_keys = df["FamilyName"].apply(normalize_key)
    local_keys = df["Local URL"].apply(normalize_key)

    family_to_locals: dict[str, list[str]] = {}
    local_to_families: dict[str, list[str]] = {}

    for idx, row in df.iterrows():
        family_key = family_keys.loc[idx]
        local_key = local_keys.loc[idx]
        family_display = clean_display(row["FamilyName"])
        local_display = clean_display(row["Local URL"])

        if family_key and local_display:
            family_to_locals.setdefault(family_key, []).append(local_display)
        if local_key and family_display:
            local_to_families.setdefault(local_key, []).append(family_display)

    same_family_local_count = []
    same_family_locals = []
    same_local_family_count = []
    families_sharing_same_local = []
    relation_status = []

    for idx, row in df.iterrows():
        family_key = family_keys.loc[idx]
        local_key = local_keys.loc[idx]

        family_locals_text = unique_join(family_to_locals.get(family_key, []))
        local_families_text = unique_join(local_to_families.get(local_key, []))

        family_local_count = len([x for x in family_locals_text.split("\n") if x.strip()])
        local_family_count = len([x for x in local_families_text.split("\n") if x.strip()])

        same_family_local_count.append(family_local_count)
        same_family_locals.append(family_locals_text)
        same_local_family_count.append(local_family_count)
        families_sharing_same_local.append(local_families_text)

        if not family_key:
            relation_status.append("Missing FamilyName")
        elif not local_key:
            relation_status.append("Missing Local URL")
        elif local_family_count > 1:
            relation_status.append("CHECK: same Local URL is used by multiple families")
        elif family_local_count > 1:
            relation_status.append("Same family has multiple Local URLs")
        else:
            relation_status.append("Single Local URL for this family")

    insert_after = "FamilyName Check Status" if "FamilyName Check Status" in df.columns else "Local URL"
    insert_pos = df.columns.get_loc(insert_after) + 1
    relation_cols = pd.DataFrame(
        {
            "Same Family Local Count": same_family_local_count,
            "Same Family Locals": same_family_locals,
            "Same Local Family Count": same_local_family_count,
            "Families Sharing Same Local": families_sharing_same_local,
            "Local Family Relation Status": relation_status,
        }
    )

    for offset, col in enumerate(relation_cols.columns):
        if col in df.columns:
            df = df.drop(columns=[col])
        df.insert(insert_pos + offset, col, relation_cols[col])

    return df


def process_workbook(uploaded_file, search_pdf_text: bool, add_relations: bool) -> tuple[dict[str, pd.DataFrame], int, int, int]:
    sheets = pd.read_excel(uploaded_file, sheet_name=None)
    processed_sheets = {}
    total_url_flags = 0
    total_family_yes = 0
    total_shared_local_issues = 0

    for sheet_name, df in sheets.items():
        processed = df.copy()
        drop_cols = [col for col in GENERATED_COLUMNS if col in processed.columns]
        if drop_cols:
            processed = processed.drop(columns=drop_cols)

        if "Local URL" in processed.columns:
            local_pos = processed.columns.get_loc("Local URL")
            flags = processed["Local URL"].apply(flag_reason)
            total_url_flags += int((flags != "").sum())
            processed.insert(local_pos + 1, "URL Flag", flags)

            if "FamilyName" in processed.columns:
                results = processed.apply(
                    lambda row: family_match_result(row["FamilyName"], row["Local URL"], search_pdf_text),
                    axis=1,
                    result_type="expand",
                )
                results.columns = ["FamilyName Found", "FamilyName Match Source", "FamilyName Check Status"]
                total_family_yes += int((results["FamilyName Found"] == "Yes").sum())
                insert_pos = processed.columns.get_loc("URL Flag") + 1
                for offset, col in enumerate(results.columns):
                    processed.insert(insert_pos + offset, col, results[col])

                if add_relations:
                    processed = add_family_local_relations(processed)
                    if "Local Family Relation Status" in processed.columns:
                        total_shared_local_issues += int(
                            processed["Local Family Relation Status"].eq(
                                "CHECK: same Local URL is used by multiple families"
                            ).sum()
                        )

        processed_sheets[sheet_name] = processed

    return processed_sheets, total_url_flags, total_family_yes, total_shared_local_issues


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        used_sheet_names = set()
        for sheet_name, df in sheets.items():
            base_name = str(sheet_name)[:31] or "Sheet"
            safe_sheet_name = base_name
            counter = 1
            while safe_sheet_name in used_sheet_names:
                suffix = f"_{counter}"
                safe_sheet_name = f"{base_name[:31 - len(suffix)]}{suffix}"
                counter += 1
            used_sheet_names.add(safe_sheet_name)

            df.to_excel(writer, index=False, sheet_name=safe_sheet_name)
            worksheet = writer.sheets[safe_sheet_name]
            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, len(value))
                worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 70)
            worksheet.freeze_panes = "A2"
    return output.getvalue()


st.set_page_config(page_title="Local URL + Family Relation Checker", page_icon="📄", layout="wide")
st.title("Local URL + PDF FamilyName + Relation Checker")
st.write(
    "Upload an Excel file. The app flags generic PDF filenames, checks whether `FamilyName` "
    "exists in the URL/PDF content, and maps relations between Local URLs and FamilyName values."
)

search_pdf_text = st.checkbox(
    "Also download each PDF and search inside PDF content",
    value=True,
    help="Use this when the PDF file name is generic, for example PdfFile_180482.pdf. This is slower but more accurate.",
)

add_relations = st.checkbox(
    "Add relation columns between FamilyName and Local URL",
    value=True,
    help="Groups all Local URLs used by the same FamilyName and all FamilyName values sharing the same Local URL.",
)

with st.expander("How the relation check works"):
    st.write(
        "For every row, the app groups by `FamilyName` and lists all unique `Local URL` values used by that family. "
        "It also groups by `Local URL` and lists all unique families using the same local. "
        "If one Local URL is used by multiple families, the row is marked as `CHECK: same Local URL is used by multiple families`."
    )

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        with st.spinner("Processing workbook..."):
            processed_sheets, total_url_flags, total_family_yes, total_shared_local_issues = process_workbook(
                uploaded_file,
                search_pdf_text,
                add_relations,
            )

        st.success(
            f"Done. Found {total_url_flags} generic Local URL filename(s). "
            f"Found FamilyName in {total_family_yes} row(s). "
            f"Found {total_shared_local_issues} row(s) where one Local URL is shared by multiple families."
        )

        sheet_names_with_local_url = [name for name, df in processed_sheets.items() if "Local URL" in df.columns]
        if not sheet_names_with_local_url:
            st.warning("No column named `Local URL` was found in this workbook.")
        else:
            selected_sheet = st.selectbox("Preview sheet", sheet_names_with_local_url)
            preview_df = processed_sheets[selected_sheet]
            local_pos = preview_df.columns.get_loc("Local URL")
            preview_cols = preview_df.columns[max(0, local_pos - 4): min(len(preview_df.columns), local_pos + 14)]
            st.dataframe(preview_df.loc[:, preview_cols], use_container_width=True)

            if "Local Family Relation Status" in preview_df.columns:
                st.subheader("Rows needing relation check")
                flagged_relation_df = preview_df[
                    preview_df["Local Family Relation Status"].str.startswith("CHECK", na=False)
                ]
                st.dataframe(flagged_relation_df.loc[:, preview_cols], use_container_width=True)

        excel_bytes = to_excel_bytes(processed_sheets)
        output_name = uploaded_file.name.rsplit(".", 1)[0] + "_family_relation_checked.xlsx"
        st.download_button(
            label="Download checked Excel file",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as exc:
        st.error(f"Could not process the file: {exc}")
