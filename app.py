import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(
    page_title="Brown Thomas Manual Transfer Processor",
    layout="wide"
)

st.title("Brown Thomas Manual Transfer File Processor")

uploaded_file = st.file_uploader(
    "Upload XLS file (template + 2 data sheets)",
    type=["xls"]
)

if uploaded_file:

    # Load workbook
    xls = pd.ExcelFile(uploaded_file)

    # Validate template sheet
    TEMPLATE_SHEET = "brownthomas_new_template"
    if TEMPLATE_SHEET not in xls.sheet_names:
        st.error(f"Sheet '{TEMPLATE_SHEET}' not found.")
        st.stop()

    # Read template
    template_df = pd.read_excel(xls, sheet_name=TEMPLATE_SHEET)
    template_df.columns = template_df.columns.str.strip()

    if "PPID" not in template_df.columns:
        st.error("Template sheet must contain a 'PPID' column.")
        st.stop()

    # -----------------------------
    # READ & CLEAN SOURCE SHEETS
    # -----------------------------
    data_frames = []

    for sheet in xls.sheet_names:
        if sheet == TEMPLATE_SHEET:
            continue

        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = df.columns.str.strip()

        # Map Pim Parent ID → PPID
        if "Pim Parent ID" in df.columns:
            df = df.rename(columns={"Pim Parent ID": "PPID"})

        if "PPID" not in df.columns:
            st.warning(f"Sheet '{sheet}' skipped — no PPID/Pim Parent ID column.")
            continue

        data_frames.append(df)

    if not data_frames:
        st.error("No valid source sheets with PPID found.")
        st.stop()

    # Combine and deduplicate
    data_df = pd.concat(data_frames, ignore_index=True)
    data_df = data_df.drop_duplicates(subset=["PPID"])

    # -----------------------------
    # COLUMN MAPPING
    # -----------------------------
    column_map = {
        "SKU": "Retek ID",
        "BARCODE": "Barcode",
        "DESCRIPTION": "Retek Item Description",
        "COLOUR": "Diff 1 Description",
        "SIZE": "UK Size Concat",
        "PRODUCT TYPE": "Product Type UDA",
        "DIVISION": "Division Name",
        "BRAND": "Brand",
        "DEPARTMENT": "Department Name",
        "DEPARTMENT NUMBER": "Department Number",
        "DIVISION NUMBER": "Division Number",
        "STORE 301 ALLOCATION": "Store 301 Allocation",
        "STORE 401 ALLOCATION": "Store 401 Allocation",
        "ITEM STORE FLAG": "Item Store Flag",
        "VPN PARENT": "VPN Parent"
    }

    # -----------------------------
    # MERGE USING PPID
    # -----------------------------
    merged_df = template_df.merge(
        data_df,
        on="PPID",
        how="left",
        suffixes=("", "_src")
    )

    # Fill template columns
    for template_col, source_col in column_map.items():
        if source_col in merged_df.columns and template_col in merged_df.columns:
            merged_df[template_col] = merged_df[source_col]

    # -----------------------------
    # BARCODE CLEANUP
    # -----------------------------
    if "BARCODE" in merged_df.columns:
        merged_df["BARCODE"] = (
            merged_df["BARCODE"]
            .astype(str)
            .str.replace(r"\.0+$", "", regex=True)
        )

    # Keep only template structure
    final_df = merged_df[template_df.columns]

    # -----------------------------
    # EXPORT FILE
    # -----------------------------
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    output_filename = f"Processed Manual Transfer File - {timestamp}.xls"

    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        final_df.to_excel(
            writer,
            index=False,
            sheet_name=TEMPLATE_SHEET
        )

    st.success("File processed successfully!")

    with open(output_filename, "rb") as f:
        st.download_button(
            label="Download Processed Manual Transfer File",
            data=f,
            file_name=output_filename,
            mime="application/vnd.ms-excel"
        )
