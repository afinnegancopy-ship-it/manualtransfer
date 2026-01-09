import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Manual Transfer Processor", layout="wide")
st.title("Brown Thomas – Manual Transfer File Processor")

uploaded_file = st.file_uploader(
    "Upload XLS file (template + 2 data sheets)",
    type=["xls"]
)

def normalize(col):
    return (
        str(col)
        .lower()
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
    )

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    TEMPLATE_SHEET = "brownthomas_new_template"

    if TEMPLATE_SHEET not in xls.sheet_names:
        st.error("Sheet 'brownthomas_new_template' not found.")
        st.stop()

    # -----------------------------
    # READ TEMPLATE (KEEP ORIGINAL COLUMNS)
    # -----------------------------
    template_df = pd.read_excel(xls, sheet_name=TEMPLATE_SHEET)
    template_columns = template_df.columns.tolist()

    template_norm = template_df.copy()
    template_norm.columns = [normalize(c) for c in template_norm.columns]

    if "ppid" not in template_norm.columns:
        st.error("Template must contain a PPID column.")
        st.stop()

    # -----------------------------
    # READ SOURCE SHEETS
    # -----------------------------
    data_frames = []

    for sheet in xls.sheet_names:
        if sheet == TEMPLATE_SHEET:
            continue

        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = [normalize(c) for c in df.columns]

        if "pim parent id" in df.columns:
            df = df.rename(columns={"pim parent id": "ppid"})

        if "ppid" not in df.columns:
            continue

        data_frames.append(df)

    if not data_frames:
        st.error("No valid source sheets with PPID found.")
        st.stop()

    source_df = pd.concat(data_frames, ignore_index=True)
    source_df = source_df.drop_duplicates(subset=["ppid"])

    # -----------------------------
    # COLUMN MAPPING (NORMALIZED)
    # -----------------------------
    column_map = {
        "SKU": "retek id",
        "BARCODE": "barcode",
        "DESCRIPTION": "retek item description",
        "COLOUR": "diff 1 description",
        "SIZE": "uk size concat",
        "PRODUCT TYPE": "product type uda",
        "DIVISION": "division name",
        "BRAND": "brand",
        "DEPARTMENT": "department name",
        "DEPARTMENT NUMBER": "department number",
        "DIVISION NUMBER": "division number",
        "STORE 301 ALLOCATION": "store 301 allocation",
        "STORE 401 ALLOCATION": "store 401 allocation",
        "ITEM STORE FLAG": "item store flag",
        "VPN PARENT": "vpn parent",
    }

    # -----------------------------
    # MERGE USING NORMALIZED PPID
    # -----------------------------
    merged = template_norm.merge(source_df, on="ppid", how="left")

    # -----------------------------
    # FILL TEMPLATE DATA
    # -----------------------------
    for template_col, source_col in column_map.items():
        template_col_norm = normalize(template_col)
        if template_col_norm in merged.columns and source_col in merged.columns:
            template_norm[template_col_norm] = merged[source_col]

    # Barcode cleanup
    if "barcode" in template_norm.columns:
        template_norm["barcode"] = (
            template_norm["barcode"]
            .astype(str)
            .str.replace(r"\.0+$", "", regex=True)
        )

    # -----------------------------
    # RESTORE ORIGINAL COLUMN NAMES
    # -----------------------------
    template_norm.columns = template_columns
    final_df = template_norm

    # -----------------------------
    # EXPORT
    # -----------------------------
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    output_file = f"Processed Manual Transfer File - {timestamp}.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name=TEMPLATE_SHEET)

    st.success("File processed successfully.")

    with open(output_file, "rb") as f:
        st.download_button(
            "Download Processed Manual Transfer File",
            data=f,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
