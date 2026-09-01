import streamlit as st
import pandas as pd
import requests
import pdfplumber
import io
import re
import urllib3

# Handle SSL certificate verification issue
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
st.set_page_config(page_title="PN & Family Validator", layout="wide")

st.title("📄 PartNumber & Family Validator")

# =========================
# HELPER: EXTRACT TEXT FROM PDF
# =========================
def extract_text_from_pdf(pdf_bytes):
    pdf_text = ""
    pdf_text_lower = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pdf_text += text + "\n"
                    pdf_text_lower += text.lower() + " "
        return pdf_text, pdf_text_lower
    except:
        return "", ""

# =========================
# HELPER: SEARCH PART NUMBER
# =========================
def search_part_number(part_number, pdf_text_lower):

    part_str = str(part_number).strip()

    # 1. Exact match
    if part_str.lower() in pdf_text_lower:
        return True, part_str

    # # 2. Remove suffix after '-' or '/'
    # if "-" in part_str or "/" in part_str:
    #     stripped = re.split(r'[-/]', part_str)[0]
    #     if stripped and stripped.lower() in pdf_text_lower:
    #         return True, stripped

    # # 3. Remove last 3 characters
    # if len(part_str) > 3:
    #     stripped_3 = part_str[:-3]
    #     if stripped_3.lower() in pdf_text_lower:
    #         return True, stripped_3

    return False, None

# =========================
# SEARCH PDF
# =========================
def search_pdf(pdf_url, part_number, family):

    try:
        response = requests.get(pdf_url, timeout=20,verify=False)
        response.raise_for_status()
    except:
        return {
            "PartNumber": "LinkIssue",
            "PartNumber_Matched": None,
            "Family": False,
        }

    pdf_text, pdf_text_lower = extract_text_from_pdf(response.content)

    if not pdf_text_lower:
        return {
            "PartNumber": "LinkIssue",
            "PartNumber_Matched": None,
            "Family": False,
        }

    part_found, part_matched = search_part_number(part_number, pdf_text_lower)
    family_found = str(family).lower() in pdf_text_lower

    return {
        "PartNumber": part_found,
        "PartNumber_Matched": part_matched,
        "Family": family_found,
    }
# =========================
# DOWNLOAD TEMPLATE
# =========================
def generate_template():
    template_df = pd.DataFrame(columns=[
        "PartNumber",
        "SupplierName",
        "Family",
        "Datasheet"
    ])

    output = io.BytesIO()
    template_df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    return output

st.subheader("📥 Download Excel Template")

template_file = generate_template()

st.download_button(
    label="Download Template Excel",
    data=template_file,
    file_name="PN_Family_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()

# =========================
# FILE UPLOAD
# =========================
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    required_cols = ["PartNumber", "SupplierName", "Family", "Datasheet"]
    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        st.error(f"Missing columns: {missing}")
        st.stop()

    st.success("File loaded successfully")

    if st.button("Start Validation"):

        statuses = []
        reasons = []
        part_matched_details = []

        progress = st.progress(0)
        total_rows = len(df)

        for idx, row in df.iterrows():

            result = search_pdf(
                pdf_url=row["Datasheet"],
                part_number=row["PartNumber"],
                family=row["Family"],
            )

            reason_parts = []
            failed = False

            # PartNumber
            if result["PartNumber"] == "LinkIssue":
                reason_parts.append("Link Issue")
                part_matched_details.append(None)
                failed = True

            elif result["PartNumber"]:
                matched = result["PartNumber_Matched"]
                original = str(row["PartNumber"]).strip()

                if matched.lower() == original.lower():
                    reason_parts.append("Found PartNumber (exact)")
                else:
                    reason_parts.append(
                        f"Found PartNumber (fallback: '{matched}')"
                    )

                part_matched_details.append(matched)

            else:
                reason_parts.append("Not found PartNumber")
                part_matched_details.append(None)
                failed = True

            # Family
            if result["Family"]:
                reason_parts.append("Found Family")
            else:
                reason_parts.append("Not found Family")
                failed = True

            statuses.append("FAIL" if failed else "PASS")
            reasons.append(" | ".join(reason_parts))

            progress.progress((idx + 1) / total_rows)

        df["status"] = statuses
        df["reason"] = reasons
        df["PartNumber_Matched"] = part_matched_details

        st.success("Validation Completed")

        # Convert to Excel in memory
        output = io.BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.download_button(
            label="Download Result Excel",
            data=output,
            file_name="validation_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
