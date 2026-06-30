# ============================
#  Streamlit PDF Family Checker (Robust Mode)
#  Ready to run with:  streamlit run streamlit_pdf_checker.py
#  Dependencies: streamlit pandas requests PyMuPDF openpyxl tqdm
# ============================

import os
import re
import json
import time
import pickle
import shutil
import difflib
import tempfile
import logging
import threading
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO

import fitz  # PyMuPDF
import pandas as pd
import requests
import streamlit as st

# ─── CONFIG ──────────────────────────────────────────────────────────────
MAX_WORKERS        = 10          # default #threads; expose in sidebar slider
PROCESSING_VERSION = "v1.3"      # bump whenever logic changes (forces re-cache)
HEADER_CHARS       = 150         # how many chars from page top     (header)
FOOTER_CHARS       = 150         # how many chars from page bottom  (footer)
FIRST20_RATIO      = 0.20        # % of first page extracted for quick search
TOP50_RATIO        = 0.50        # % of first page for quick preview

# ─── LOGGING ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ─── SMALL UTILS ─────────────────────────────────────────────────────────

def clean_text(s: str | int | float | None):
    """Remove control characters; keep None / non-str unchanged."""
    if isinstance(s, str):
        return re.sub(r"[\x00-\x1F\x7F-\x9F]", "", s)
    return s

def calculate_similarity(a: str, b: str) -> float:
    """Percent similarity (0-100) using SequenceMatcher."""
    try:
        return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio() * 100
    except Exception:
        return 0.0

def search_in_metadata(metadata: dict, family_name: str):
    """Return (status, similarity_str) for metadata search."""
    if not metadata:
        return "Not Found", ""
    for v in metadata.values():
        if isinstance(v, str) and family_name.lower() in v.lower():
            return "Found in Metadata", ""
    # else max similarity
    max_sim = max(
        (calculate_similarity(family_name, str(v)) for v in metadata.values() if isinstance(v, str)),
        default=0.0,
    )
    return "Not Found", f"{max_sim:.2f}%" if max_sim else ""

def search_in_text(section_text: str, family_name: str):
    """Return (found:bool, similarity:float)."""
    if family_name.lower() in section_text.lower():
        return True, 100.0
    return False, calculate_similarity(family_name, section_text)

def extract_header_footer(text: str):
    return text[:HEADER_CHARS], text[-FOOTER_CHARS:]

def slice_ratio(text: str, ratio: float):
    end = int(len(text) * ratio)
    return text[:end]

# ─── EXTRA HELPERS FOR STATUS/PRI ────────────────────────────────────────

def whitespace_ratio(s: str | None) -> float:
    """Fraction of whitespace characters in a string (0..1). Empty → 1.0."""
    s = "" if s is None else str(s)
    if not s:
        return 1.0
    total = len(s)
    ws = sum(1 for ch in s if str(ch).isspace())
    return ws / total if total else 1.0

def norm_code(s: str | None) -> str:
    """Normalize typical revision/code/header tokens for robust comparison."""
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ").replace("\u200B", "")  # nbsp, zero-width
    s = re.sub(r"\b(rev(?:ision)?)[\s:.\-]*", "rev ", s, flags=re.I)
    s = re.sub(r"[^A-Za-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

# ─── SIMPLE DISK CACHE (pickle) ──────────────────────────────────────────

def _cache_path(cache_dir: str):
    return os.path.join(cache_dir, "processed_pdfs_cache.pkl")

def load_cache(cache_dir: str):
    try:
        with open(_cache_path(cache_dir), "rb") as f:
            return pickle.load(f)
    except Exception:
        return {}

def save_cache(cache_dir: str, cache: dict):
    try:
        with open(_cache_path(cache_dir), "wb") as f:
            pickle.dump(cache, f)
    except Exception as e:
        logger.warning(f"Cache save failed: {e}")

# ─── PDF HELPERS ─────────────────────────────────────────────────────────

def download_pdf(session: requests.Session, url: str, cache_dir: str, max_attempts: int = 3):
    """Return (local_path, error_msg_or_None)."""
    if not url:
        return None, "Missing URL"
    filename = os.path.basename(url.split("?", 1)[0]) or "downloaded.pdf"
    local_path = os.path.join(cache_dir, filename)
    # cached?
    if os.path.exists(local_path) and os.path.getsize(local_path) > 0:
        try:
            with fitz.open(local_path):
                return local_path, None
        except Exception:
            os.remove(local_path)
    for attempt in range(1, max_attempts + 1):
        try:
            r = session.get(url, timeout=60)
            r.raise_for_status()
            with open(local_path, "wb") as f:
                f.write(r.content)
            with fitz.open(local_path):  # verify
                return local_path, None
        except Exception as e:
            logger.warning(f"{url} (attempt {attempt}/{max_attempts}) → {e}")
            time.sleep(2)
    return None, f"Failed after {max_attempts} attempts"

def process_pdf(file_path: str, family_name: str, company_name: str,
                cache_dir: str, cache: dict, cache_lock: threading.Lock):
    key = (file_path, family_name, company_name, PROCESSING_VERSION)
    with cache_lock:
        if key in cache:
            return cache[key]
    # open doc
    try:
        doc = fitz.open(file_path)
    except Exception as e:
        res = {"Metadata": f"Error opening PDF: {e}"}
        with cache_lock:
            cache[key] = res
        return res
    # metadata
    meta = doc.metadata or {}
    meta_json = json.dumps({k: str(v) for k, v in meta.items()})
    meta_status, meta_sim = search_in_metadata(meta, family_name)
    # first page text
    try:
        first_page_text = doc.load_page(0).get_text("text")
    except Exception:
        first_page_text = ""
    header, footer = extract_header_footer(first_page_text)
    found_h, sim_h = search_in_text(header, family_name)
    found_f, sim_f = search_in_text(footer, family_name)
    if found_h or found_f:
        hf_status = ("Header" if found_h else "") + (" & " if found_h and found_f else "") + ("Footer" if found_f else "") + " Found"
        hf_sim = ""
    else:
        hf_status = "Not Found"
        hf_sim = f"{max(sim_h, sim_f):.2f}%" if max(sim_h, sim_f) else ""
    first20 = slice_ratio(first_page_text, FIRST20_RATIO)
    found_20, sim_20 = search_in_text(first20, family_name)
    first20_status = "Found" if found_20 else "Not Found"
    first20_sim = "" if found_20 else f"{sim_20:.2f}%" if sim_20 else ""
    top50 = slice_ratio(first_page_text, TOP50_RATIO)

    # Keep raw header for the whitespace ratio rule; preserve your stripped output
    res = {
        "Metadata": meta_status + (f" | Similarity: {meta_sim}" if meta_sim else ""),
        "Metadata Similarity": meta_sim,
        "Header and Footer": hf_status + (f" | Similarity: {hf_sim}" if hf_sim else ""),
        "Header and Footer Similarity": hf_sim,
        "First 20% of First Page": first20_status + (f" | Similarity: {first20_sim}" if first20_sim else ""),
        "headerContent_raw": header,               # unstripped, for ws% rule
        "headerContent": header.strip(),
        "footerContent": footer.strip(),
        "TopContent 50%": top50.strip(),
        "countOfPage": doc.page_count,
        "Similarity Family Header - with the first": "",
        "Similarity Family Header - with the last": "",
        "compare header with median": "",
        "Similarity Supplier Footer": "",
        "comment": "",
        "metadata.Content": meta_json,
        "Log": ""
    }
    doc.close()
    with cache_lock:
        cache[key] = res
    return res

# ─── ROW PROCESSOR ───────────────────────────────────────────────────────

def _row_template():
    return {
        "Metadata": "", "Metadata Similarity": "",
        "Header and Footer": "", "Header and Footer Similarity": "",
        "First 20% of First Page": "",
        "headerContent_raw": "",
        "headerContent": "", "footerContent": "",
        "TopContent 50%": "",
        "countOfPage": 0,
        "Similarity Family Header - with the first": "",
        "Similarity Family Header - with the last": "",
        "compare header with median": "",
        "Similarity Supplier Footer": "",
        "comment": "",                         # will be replaced by family-level
        "metadata.Content": "",
        "Log": "",
        "PRI": "",
        "Status": ""
    }

def process_row(session: requests.Session, row: pd.Series, cache_dir: str,
                cache: dict, cache_lock: threading.Lock):
    res = _row_template()
    fam  = str(row.get("FamilyName", "")).strip()
    comp = str(row.get("CompanyName", "")).strip()
    url  = str(row.get("Local URL", "")).strip()

    if not fam:
        res.update({k: "Missing Family Name" for k in ["Metadata", "Header and Footer", "First 20% of First Page"]})
        res["Log"] = "Missing Family Name"
        return res
    if not url:
        res.update({k: "Missing URL" for k in ["Metadata", "Header and Footer", "First 20% of First Page"]})
        res["Log"] = "Missing URL"
        return res
    pdf_path, err = download_pdf(session, url, cache_dir)
    if err:
        res.update({k: f"Download Error: {err}" for k in ["Metadata", "Header and Footer", "First 20% of First Page"]})
        res["Log"] = f"Download Error: {err}"
        return res
    res.update(process_pdf(pdf_path, fam, comp, cache_dir, cache, cache_lock))
    return res

# ─── SIMILARITY POST-PROCESS ─────────────────────────────────────────────

def compute_similarity(df: pd.DataFrame):
    grouped = df.groupby("FamilyName")
    for _, group in grouped:
        n = len(group)
        if n == 0:
            continue
        first_idx = group.index[0]
        last_idx  = group.index[-1]
        median_idx = group.index[(n - 1)//2]
        first_header = df.at[first_idx, "headerContent"]
        last_header  = df.at[last_idx,  "headerContent"]
        median_header= df.at[median_idx, "headerContent"] if n > 2 else ""
        first_url    = df.at[first_idx, "Local URL"]
        last_url     = df.at[last_idx,  "Local URL"]
        median_url   = df.at[median_idx, "Local URL"] if n > 2 else ""
        for idx in group.index:
            curr_header = df.at[idx, "headerContent"]
            curr_footer = df.at[idx, "footerContent"]
            curr_url    = df.at[idx, "Local URL"]
            comp        = str(df.at[idx, "CompanyName"]).lower()
            # header-first
            if idx == first_idx or curr_url == first_url:
                df.at[idx, "Similarity Family Header - with the first"] = "-"
            else:
                df.at[idx, "Similarity Family Header - with the first"] = f"{calculate_similarity(curr_header, first_header):.2f}%"
            # header-last
            if idx == last_idx or curr_url == last_url:
                df.at[idx, "Similarity Family Header - with the last"] = "-"
            else:
                df.at[idx, "Similarity Family Header - with the last"] = f"{calculate_similarity(curr_header, last_header):.2f}%"
            # header-median
            if n < 3 or idx == median_idx or curr_url == median_url:
                df.at[idx, "compare header with median"] = "-" if n >= 3 else "N/A"
            else:
                df.at[idx, "compare header with median"] = f"{calculate_similarity(curr_header, median_header):.2f}%"
            # footer vs supplier
            footer_lower = str(curr_footer).lower()
            sim_footer = 100.0 if comp and comp in footer_lower else calculate_similarity(comp, curr_footer)
            df.at[idx, "Similarity Supplier Footer"] = f"{sim_footer:.2f}%"
            # (row-level comment was here; we’ll replace with family-level later)
    return df

# ─── STATUS/PRI ASSIGNMENT (GROUP-AWARE) ─────────────────────────────────

def assign_dynamic_status(df: pd.DataFrame, family_col: str = "FamilyName", sim_threshold: float = 0.65) -> pd.DataFrame:
    # Precompute row-level signals
    log_s  = df.get("Log", pd.Series([""] * len(df))).astype(str)
    meta_s = df.get("Metadata", pd.Series([""] * len(df))).astype(str)

    df["__download_err"] = log_s.str.startswith("Download Error:")
    df["__open_err"]     = meta_s.str.startswith("Error opening PDF")

    # Use raw header if present; per-row fallback to stripped header when raw is empty
    if "headerContent_raw" in df.columns:
        header_raw = df["headerContent_raw"].astype(str)
        header_src = header_raw.where(header_raw.str.len() > 0, df["headerContent"].astype(str))
    else:
        header_src = df["headerContent"].astype(str)

    df["__ws80"]   = header_src.apply(whitespace_ratio).ge(0.80)
    df["__usable"] = ~(df["__download_err"] | df["__open_err"] | df["__ws80"])

    # Normalized code tokens from header for comparison
    df["__code_norm"] = header_src.apply(norm_code)

    pri_out    = [""] * len(df)
    status_out = [""] * len(df)

    # group by family
    for fam, idxs in df.groupby(family_col).groups.items():
        idxs = list(idxs)
        U = int(df.loc[idxs, "__usable"].sum())

        def set_row(i, pri, status):
            pri_out[i] = pri
            status_out[i] = status

        # A) No usable PDFs
        if U == 0:
            for i in idxs:
                if df.at[i, "__download_err"]:
                    set_row(i, "LOW", "LOW PRI, download error")
                elif df.at[i, "__open_err"]:
                    set_row(i, "LOW", "LOW PRI, PDFs not Open")
                elif df.at[i, "__ws80"]:
                    set_row(i, "LOW", "LOW PRI as the First characters are spaces")
                else:
                    set_row(i, "LOW", "LOW PRI, no usable headers")
            continue

        # B) Exactly one usable PDF → can't compare
        if U == 1:
            for i in idxs:
                if df.at[i, "__download_err"]:
                    set_row(i, "LOW", "LOW PRI, download error")
                elif df.at[i, "__open_err"]:
                    set_row(i, "LOW", "LOW PRI, PDFs not Open")
                elif df.at[i, "__ws80"]:
                    set_row(i, "LOW", "LOW PRI as the First characters are spaces")
                elif df.at[i, "__usable"]:
                    set_row(i, "LOW", "LOW PRI, only one usable PDF (can’t compare)")
                else:
                    set_row(i, "LOW", "LOW PRI, no usable headers")
            continue

        # C) Two or more usable → compare among usable
        usable_idxs = [i for i in idxs if df.at[i, "__usable"]]
        codes       = [df.at[i, "__code_norm"] for i in usable_idxs]
        non_empty   = [c for c in codes if c]
        baseline    = Counter(non_empty).most_common(1)[0][0] if non_empty else ""

        def similar(a: str, b: str) -> bool:
            if not a and not b:
                return True
            if not a or not b:
                return False
            A, B = set(a.split()), set(b.split())
            inter = len(A & B)
            union = len(A | B) or 1
            return (inter / union) >= sim_threshold

        for i in idxs:
            if df.at[i, "__download_err"]:
                set_row(i, "LOW", "LOW PRI, download error")
            elif df.at[i, "__open_err"]:
                set_row(i, "LOW", "LOW PRI, PDFs not Open")
            elif df.at[i, "__ws80"]:
                set_row(i, "LOW", "LOW PRI as the First characters are spaces")
            elif df.at[i, "__usable"]:
                code_i = df.at[i, "__code_norm"]
                if similar(code_i, baseline):
                    set_row(i, "", "Seems to be ok")
                else:
                    set_row(i, "FIRST", "First PRI to check.")
            else:
                set_row(i, "LOW", "LOW PRI, no usable headers")

    df["PRI"]    = pri_out
    df["Status"] = status_out

    # Optional: per-row whitespace ratio (for debugging/filters)
    df["headerWhitespaceRatio"] = header_src.apply(whitespace_ratio).round(3)

    # cleanup temp columns
    df.drop(columns=["__download_err","__open_err","__ws80","__usable","__code_norm"], inplace=True, errors="ignore")
    return df

# ─── FAMILY-LEVEL COMMENTS & SUMMARY ─────────────────────────────────────

CATS = [
    ("First PRI to check.", "FIRST"),
    ("LOW PRI, download error", "LOW_DOWNLOAD"),
    ("LOW PRI, PDFs not Open", "LOW_OPEN"),
    ("LOW PRI as the First characters are spaces", "LOW_SPACES"),
    ("LOW PRI, only one usable PDF (can’t compare)", "LOW_ONE_USABLE"),
    ("LOW PRI, no usable headers", "LOW_NO_USABLE"),
    ("Seems to be ok", "OK"),
]

def build_family_comments_and_summary(df: pd.DataFrame, family_col: str = "FamilyName") -> tuple[pd.DataFrame, pd.DataFrame]:
    fam_comment_map = {}
    summary_rows = []

    # Make sure we have a stable row number to reference in comments
    if "RowNum" not in df.columns:
        df.insert(0, "RowNum", range(1, len(df) + 1))

    if "FamilyStatus" not in df.columns:
        df["FamilyStatus"] = ""

    for fam, g in df.groupby(family_col, sort=False):
        idxs = g.index.tolist()
        total = len(g)

        # Re-derive usability for transparency
        download_err = g["Log"].astype(str).str.startswith("Download Error:")
        open_err     = g["Metadata"].astype(str).str.startswith("Error opening PDF")
        ws80         = g.get("headerWhitespaceRatio", pd.Series([0]*len(g), index=g.index)).ge(0.80)
        usable       = ~(download_err | open_err | ws80)
        U = int(usable.sum())

        # Baseline code (normalize headers from usable rows)
        header_src = g["headerContent_raw"].fillna("").astype(str) if "headerContent_raw" in g.columns else g["headerContent"].fillna("").astype(str)
        header_src = header_src.where(header_src.str.len() > 0, g["headerContent"].fillna("").astype(str))
        codes = header_src.apply(norm_code)
        non_empty = codes[usable & codes.astype(bool)]
        baseline = Counter(non_empty).most_common(1)[0][0] if len(non_empty) > 0 else ""
        baseline_disp = baseline if baseline else "—"

        parts = [f"Family '{fam}': usable={U}/{total}; baseline={baseline_disp}"]

        cat_counts = {}
        details_by_cat = {}
        family_status_values = []

        for status_text, label in CATS:
            sel = g[g["Status"] == status_text]
            cat_counts[label] = len(sel)
            if len(sel) > 0:
                where_list = [f"#{int(r.RowNum)}|{r['Local URL']}" for _, r in sel.iterrows()]
                details_by_cat[label] = "; ".join(where_list)
                family_status_values.append(status_text)
                parts.append(f"{label}={len(sel)} [{details_by_cat[label]}]")
            else:
                details_by_cat[label] = ""

        family_status = "; ".join(family_status_values) if family_status_values else "Seems to be ok"
        parts.insert(1, f"Family status: {family_status}")

        comment_str = "\n".join(parts)
        for i in idxs:
            fam_comment_map[i] = comment_str

        df.loc[idxs, "FamilyStatus"] = family_status

        # summary record
        rec = {
            "FamilyName": fam,
            "Total": total,
            "Usable": U,
            "BaselineCode": baseline_disp,
        }
        for _, label in CATS:
            rec[label] = cat_counts.get(label, 0)
            rec[f"{label}_where"] = details_by_cat.get(label, "")
        rec["FamilyStatus"] = family_status
        summary_rows.append(rec)

    # apply family comment to every row; also mirror it into 'comment'
    df["FamilyComment"] = df.index.map(fam_comment_map)
    df["comment"] = df["FamilyComment"]

    summary_df = pd.DataFrame(summary_rows)
    return df, summary_df

# ─── MAIN PROCESSOR ──────────────────────────────────────────────────────

def run_checker(input_excel_bytes: BytesIO, progress_cb=None, max_workers: int = MAX_WORKERS) -> BytesIO:
    start = time.time()
    cache_dir = tempfile.mkdtemp(prefix="pdf_cache_")
    cache     = load_cache(cache_dir)
    cache_lock = threading.Lock()

    df = pd.read_excel(input_excel_bytes)
    required = {"FamilyName", "CompanyName", "Local URL"}
    if not required.issubset(df.columns):
        raise ValueError(f"Excel must have columns: {required}")

    # ensure columns exist
    # add_cols = _row_template().keys()
    # for c in add_cols:
    #     if c not in df.columns:
    #         df[c] = ""
    for c in add_cols:
        if c not in df.columns:
            if c == "countOfPage":
                df[c] = pd.Series(dtype="Int64")
            else:
                df[c] = pd.Series(dtype="object")
    if "RowNum" not in df.columns:
        df.insert(0, "RowNum", range(1, len(df) + 1))

    session = requests.Session()
    total = len(df)
    completed = 0
    futures = {}
    with ThreadPoolExecutor(max_workers=max_workers) as exe:
        for idx, row in df.iterrows():
            futures[exe.submit(process_row, session, row, cache_dir, cache, cache_lock)] = idx
        for fut in as_completed(futures):
            idx = futures[fut]
            res = fut.result()
            for k, v in res.items():
                df.at[idx, k] = v
            completed += 1
            if progress_cb:
                progress_cb(completed / total)

    df = compute_similarity(df)
    df = assign_dynamic_status(df, family_col="FamilyName", sim_threshold=0.65)

    # Build family-level comments & a summary sheet
    df, family_summary = build_family_comments_and_summary(df, family_col="FamilyName")

    # final clean
    for c in df.columns:
        df[c] = df[c].apply(clean_text)

    # save to bytes (two sheets)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Rows")
        family_summary.to_excel(writer, index=False, sheet_name="Family_Summary")
    output.seek(0)

    save_cache(cache_dir, cache)
    shutil.rmtree(cache_dir, ignore_errors=True)
    logger.info("Processing took %.1fs", time.time() - start)
    return output

# ─── STREAMLIT UI ────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="PDF Family Checker", layout="wide")
    st.title("📄 PDF Family Checker – Robust Mode")

    st.sidebar.header("Settings")
    max_workers = st.sidebar.slider("Parallel workers", 1, 20, MAX_WORKERS)
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], accept_multiple_files=False)

    if uploaded_file is not None:
        st.write("Excel loaded. Rows:", len(pd.read_excel(uploaded_file)))
        if st.button("🚀 Start Processing"):
            progress_bar = st.progress(0.0)
            try:
                out_bytes = run_checker(uploaded_file, progress_cb=progress_bar.progress, max_workers=max_workers)
                st.success("Processing complete! Click to download.")
                st.download_button(
                    label="📥 Download Checked Excel",
                    data=out_bytes,
                    file_name="checked_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Processing failed: {e}")

    st.markdown("---")
    st.markdown("**How to use:** Upload an Excel file containing the columns `FamilyName`, `CompanyName`, and `Local URL`. The app downloads each PDF, extracts headers/footers & metadata, then reports whether the family name appears.")

if __name__ == "__main__":
    main()
