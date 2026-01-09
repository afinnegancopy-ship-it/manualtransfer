import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Brown Thomas Manual Transfer Processor", layout="wide")

st.title("Brown Thomas Manual Transfer File Processor")

uploaded_file = st.file_uploader(
    "Upload XLS file (template + 2 data sheets)",
    type=["xls"]
)

if uploaded_file:
    # Load all sheets
    xls = pd.ExcelFile(uploaded_file)

    if "brownthomas_new_template" not in xls.sheet_names:
        st.error("Sheet 'brownthomas_new_template' not found.")
        st.stop()

    # Read template
    template_df = pd.read_excel(xls, sheet_name="brownthomas_new_template")

    # Read other sheets
    data_sheets = [
        pd.read_excel(xls, sheet_name=s)
        for s in xls.sheet_names
        if s != "brownthomas_new_template"
    ]

    # Combine the two data sheets
    data_df = pd.concat(data_sheets, ignore_index=True)

    # Remove duplicate PPIDs in source data
    data_df = data_df.drop_duplicates(subset=["PPID"])

    # Mapping between template columns and source headers
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

    # Merge template with source data using PPID
    merged_df = template_df.merge(
        data_df,
        on="PPID",
        how="left",
        suffixes=("", "_src")
    )

    # Fill template columns from source columns
    for template_col, source_col in column_map.items():
        if source_col in merged_df.columns:
            merged_df[template_col] = merged_df[source_col]

    # Barcode cleanup: convert to number & remove decimals
    if "BARCODE" in merged_df.columns:
        merged_df["BARCODE"] = (
            pd.to_numeric(merged_df["BARCODE"], errors="coerce")
            .astype("Int64")
        )

    # Keep only original template columns
    final_df = merged_df[template_df.columns]

    # Generate output filename
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    output_filename = f"Processed Manual Transfer File - {timestamp}.xls"

    # Save file
    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="brownthomas_new_template")

    st.success("File processed successfully!")

    with open(output_filename, "rb") as f:
        st.download_button(
            label="Download Processed Manual Transfer File",
            data=f,
            file_name=output_filename,
            mime="application/vnd.ms-excel"
        )
