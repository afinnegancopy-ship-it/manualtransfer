import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Brown Thomas Manual Transfer Processor", layout="wide")
st.title("Brown Thomas Manual Transfer File Processor")

uploaded_file = st.file_uploader(
    "Upload XLS file (template + 2 data sheets)",
    type=["xls"]
)

def normalize(col):
    return (
        col.lower()
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
    )

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)

    TEMPLATE_SHEET = "brownthomas_new_template"

    template_df = pd.read_excel(xls, sheet_name=TEMPLATE_SHEET)
    template_df.columns = [normalize(c) for c in template_df.columns]

    if "ppid" not in template_df.columns:
        st.error("Template must contain PPID column.")
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

    data_df = pd.concat(data_frames, ignore_index=True)
    data_df = data_df.drop_duplicates(subset=["ppid"])

    # -----------------------------
    # COLUMN MAP (NORMALIZED)
    # -----------------------------
    column_map = {
        "sku": "retek id",
        "barcode": "barcode",
        "description": "retek item description",
        "colour": "diff 1 description",
        "size": "uk size concat",
        "product type": "product type uda",
        "division": "division name",
        "brand": "brand",
        "department": "department name",
        "department number": "department number",
        "division number": "division number",
        "store 301 allocation": "store 301 allocation",
        "store 401 allocation": "store 401 allocation",
        "item store flag": "item store flag",
        "vpn parent": "vpn parent"
    }

    # -----------------------------
    # MERGE
    # -----------------------------
    merged_df = template_df.merge(
        data_df,
        on="ppid",
        how="left"
    )

    # -----------------------------
    # FILL TEMPLATE COLUMNS
    # -----------------------------
    for template_col, source_col in column_map.items():
        if template_col in merged_df.columns and source_col in merged_df.columns:
            merged_df[template_col] = merged_df[source_col]

    # Barcode cleanup
    if "barcode" in merged_df.columns:
        merged_df["barcode"] = (
            merged_df["barcode"]
            .astype(str)
            .str.replace(r"\.0+$", "", regex=True)
        )

    # Restore original column casing/order
    final_df = merged_df[template_df.columns]

    # -----------------------------
    # EXPORT
    # -----------------------------
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    output_filename = f"Processed Manual Transfer File - {timestamp}.xlsx"

    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name=TEMPLATE_SHEET)

    st.success("File processed successfully!")

    with open(output_filename, "rb") as f:
        st.download_button(
            "Download Processed Manual Transfer File",
            f,
            output_filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
