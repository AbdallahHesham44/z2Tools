import streamlit as st
import pandas as pd
import requests
import pdfplumber
import io
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from fuzzywuzzy import fuzz

# =========================
# ⚙️ CONFIG
# =========================
MAX_WORKERS = 10
REQUEST_TIMEOUT = 15
MAX_RETRIES = 2

# =========================
# 📄 TEMPLATE
# =========================
def generate_template():
    df = pd.DataFrame({
        "PartNumber": ["ABC123", "XYZ789"],
        "SupplierName": ["Supplier A", "Supplier B"],
        "Family": ["Zener Diode", "Capacitor"],
        "Datasheet": [
            "https://example.com/datasheet1.pdf",
            "https://example.com/datasheet2.pdf"
        ]
    })
    output = io.BytesIO()
    df.to_excel(output, index=False)
    return output.getvalue()

# =========================
# 🧠 HELPERS
# =========================
def normalize(text):
    return re.sub(r'[\s\-\_]', '', str(text).lower())

def download_pdf(url):
    for _ in range(MAX_RETRIES):
        try:
            r = requests.get(url, timeout=REQUEST_TIMEOUT)
            if r.status_code == 200 and len(r.content) > 500:
                return r.content
        except:
            time.sleep(1)
    return None

def extract_text(content, search_terms=None):
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for page in pdf.pages[:5]:
                txt = page.extract_text()
                if txt:
                    lower = txt.lower()

                    # 🚀 early stop
                    if search_terms:
                        for term in search_terms:
                            if term and str(term).lower() in lower:
                                return lower

            return lower
    except:
        return ""

def match_part(part, text):
    part_clean = normalize(part)
    text_clean = normalize(text)

    if part_clean in text_clean:
        return True

    for i in range(2, 5):
        base = part_clean[:-i]
        if len(base) > 4 and base in text_clean:
            return True

    return False

def match_family(family, text):
    family = str(family).lower()

    if family in text:
        return True

    for line in text.split("\n"):
        if fuzz.partial_ratio(family, line.lower()) > 85:
            return True

    return False

def process_row(row):
    url = row["Datasheet"]
    part = row["PartNumber"]
    fam  = row["Family"]

    if not isinstance(url, str) or not url.startswith("http"):
        return "FAIL", "INVALID_URL"

    pdf = download_pdf(url)
    if not pdf:
        return "FAIL", "DOWNLOAD_FAIL"

    text = extract_text(pdf, [part, fam])
    if not text:
        return "FAIL", "EMPTY_PDF"

    part_ok = match_part(part, text)
    fam_ok  = match_family(fam, text)

    if part_ok and fam_ok:
        return "PASS", "OK"
    elif not part_ok:
        return "FAIL", "NO_PART"
    else:
        return "FAIL", "NO_FAMILY"

# =========================
# 🎨 UI
# =========================
st.set_page_config(page_title="PDF Validator", layout="wide")

st.title("🚀 PDF Part Number & Family Validator")

# =========================
# 📥 TEMPLATE DOWNLOAD
# =========================
st.subheader("📥 Download Template")

st.download_button(
    label="📄 Download Excel Template",
    data=generate_template(),
    file_name="template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.info("Required columns: PartNumber, SupplierName, Family, Datasheet")

# =========================
# 📂 UPLOAD FILE
# =========================
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.write("### 📊 Preview")
    st.dataframe(df.head())

    # =========================
    # ✅ VALIDATE COLUMNS
    # =========================
    required_cols = ["PartNumber", "SupplierName", "Family", "Datasheet"]

    if not all(col in df.columns for col in required_cols):
        st.error(f"❌ Missing columns! Required: {required_cols}")
        st.stop()

    # =========================
    # ▶️ RUN BUTTON
    # =========================
    if st.button("▶️ Run Validation"):
        progress = st.progress(0)
        status_text = st.empty()

        results = []
        total = len(df)
        done = 0

        start = time.time()

        # =========================
        # ⚡ PARALLEL PROCESSING
        # =========================
        with ThreadPoolExecutor(MAX_WORKERS) as executor:
            futures = {executor.submit(process_row, row): i for i, row in df.iterrows()}

            for future in as_completed(futures):
                idx = futures[future]

                try:
                    status, reason = future.result()
                except Exception as e:
                    status, reason = "FAIL", str(e)

                results.append((idx, status, reason))

                done += 1
                progress.progress(done / total)
                status_text.text(f"Processed {done}/{total}")

        # =========================
        # 📊 ADD RESULTS
        # =========================
        for idx, status, reason in results:
            df.at[idx, "status"] = status
            df.at[idx, "reason"] = reason

        st.success("✅ Validation Completed")

        # =========================
        # 📊 SUMMARY
        # =========================
        st.write("### 📊 Results Summary")
        st.write(df["status"].value_counts())

        # =========================
        # 📋 TABLE
        # =========================
        st.dataframe(df)

        # =========================
        # 📥 DOWNLOAD OUTPUT
        # =========================
        output = io.BytesIO()
        df.to_excel(output, index=False)

        st.download_button(
            label="📥 Download Results",
            data=output.getvalue(),
            file_name="validated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
