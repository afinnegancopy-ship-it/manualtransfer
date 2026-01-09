import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Manual Transfer Processor", layout="wide")
st.title("Brown Thomas – Manual Transfer File Processor")

uploaded_file = st.file_uploader(
    "Upload XLS file (template + 2 data sheets)",
    type=["xls"]
)

def normalize_col(col):
    return (
        str(col)
        .lower()
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
    )

def normalize_ppid(val):
    if pd.isna(val):
        return ""
    val = str(val).strip()
    val = val.replace(".0", "")
    return val

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    TEMPLATE_SHEET = "brownthomas_new_template"

    if TEMPLATE_SHEET not in xls.sheet_names:
        st.error("Sheet 'brownthomas_new_template' not found.")
        st.stop()

    # -----------------------------
    # TEMPLATE
    # -----------------------------
    template_df = pd.read_excel(xls, sheet_name=TEMPLATE_SHEET, dtype=str)
    template_columns = template_df.columns.tolist()

    template_df.columns = [normalize_col(c) for c in template_df.columns]

    if "ppid" not in template_df.columns:
        st.error("Template must contain PPID column.")
        st.stop()

    template_df["ppid"] = template_df["ppid"].apply(normalize_ppid)

    # -----------------------------
    # SOURCE SHEETS
    # -----------------------------
    source_frames = []

    for sheet in xls.sheet_names:
        if sheet == TEMPLATE_SHEET:
            continue

        df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
        df.columns = [normalize_col(c) for c in df.columns]

        if "pim parent id" in df.columns:
            df = df.rename(columns={"pim parent id": "ppid"})

        if "ppid" not in df.columns:
            continue

        df["ppid"] = df["ppid"].apply(normalize_ppid)
        source_frames.append(df)

    if not source_frames:
        st.error("No source sheets with Pim Parent ID found.")
        st.stop()

    source_df = pd.concat(source_frames, ignore_index=True)
    source_df = source_df.drop_duplicates(subset=["ppid"])

    # -----------------------------
    # COLUMN MAP
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
        "vpn parent": "vpn parent",
    }

    # -----------------------------
    # MERGE
    # -----------------------------
    merged = template_df.merge(
        source_df,
        on="ppid",
        how="left",
        indicator=True
    )

    # DEBUG (remove later)
    st.write("PPID match summary:")
    st.write(merged["_merge"].value_counts())

    # -----------------------------
    # FILL DATA
    # -----------------------------
    for template_col, source_col in column_map.items():
        if template_col in merged.columns and source_col in merged.columns:
            template_df[template_col] = merged[source_col]

    # Barcode cleanup
    if "barcode" in template_df.columns:
        template_df["barcode"] = (
            template_df["barcode"]
            .astype(str)
            .str.replace(r"\.0+$", "", regex=True)
        )

    # Restore original headers
    template_df.columns = template_columns
    final_df = template_df

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
            f,
            output_file,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
