import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="PL Feature Inserter", layout="wide")

st.title("📌 Product Line Feature Inserter")
st.markdown(
    "Add feature names to all Product Lines after the last **Parametric** feature without duplication."
)

# =========================
# REQUIRED COLUMNS
# =========================
required_columns = [
    "FunctionName",
    "ProductName",
    "ProductKey",
    "FeatureName",
    "FeatureKey",
    "ProductFeatureName",
    "DisplayOrder",
    "Comment",
    "ProductFeatureDefinition",
    "IsPlFilter",
    "IsTab",
    "Status"
]

# =========================
# FILE UPLOAD
# =========================
uploaded_file = st.file_uploader(
    "📂 Upload Excel File",
    type=["xlsx"]
)

# =========================
# FEATURES INPUT
# =========================
st.subheader("🧩 Features To Add")

feature_text = st.text_area(
    "Enter one feature per line",
    value="Minimum Storage Temperature\nMaximum Storage Temperature",
    height=150
)

features_to_add = [
    x.strip()
    for x in feature_text.splitlines()
    if x.strip()
]

# =========================
# PROCESS
# =========================
if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.subheader("📋 Preview")
    st.dataframe(df.head())

    missing_cols = [c for c in required_columns if c not in df.columns]

    if missing_cols:
        st.error(f"❌ Missing columns: {missing_cols}")
        st.stop()

    if st.button("🚀 Process File"):

        all_groups = []

        progress = st.progress(0)

        grouped = list(
            df.groupby(["ProductName", "ProductKey"], sort=False)
        )

        total_groups = len(grouped)

        for idx, ((product_name, product_key), group) in enumerate(grouped):

            group = group.copy()

            existing_features = set(
                group["FeatureName"].astype(str)
            )

            parametric_rows = group[
                group["FunctionName"] == "Parametric"
            ]

            # Skip if no Parametric rows
            if parametric_rows.empty:
                all_groups.append(group)
                continue

            last_parametric_idx = parametric_rows.index[-1]

            base_row = group.loc[last_parametric_idx].copy()

            new_rows = []

            for feature in features_to_add:

                # Prevent duplication inside same PL
                if feature in existing_features:
                    continue

                new_row = base_row.copy()

                new_row["FeatureName"] = feature
                new_row["ProductFeatureName"] = feature

                # Optional cleanup
                new_row["Comment"] = None
                new_row["ProductFeatureDefinition"] = None

                new_rows.append(new_row)

            before = group.loc[:last_parametric_idx]
            after = group.loc[last_parametric_idx + 1:]

            updated_group = pd.concat(
                [
                    before,
                    pd.DataFrame(new_rows),
                    after
                ],
                ignore_index=False
            )

            all_groups.append(updated_group)

            progress.progress((idx + 1) / total_groups)

        final_df = pd.concat(all_groups, ignore_index=True)

        st.success("✅ Processing Completed")

        st.subheader("📊 Final Preview")
        st.dataframe(final_df.head(20))

        # =========================
        # EXPORT
        # =========================
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False)

        output.seek(0)

        st.download_button(
            label="⬇️ Download Updated Excel File",
            data=output,
            file_name="updated_product_lines.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
