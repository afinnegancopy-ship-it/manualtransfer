import streamlit as st
import pandas as pd
from datetime import datetime
import io
import xlwt

st.set_page_config(page_title="Brown Thomas File Processor", layout="wide")

st.title("Brown Thomas Manual Transfer File Processor")
st.markdown("---")

# File uploader
uploaded_file = st.file_uploader("Upload your XLS file", type=['xls', 'xlsx'])

if uploaded_file is not None:
    try:
        # Read all sheets from the Excel file
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        st.info(f"Found {len(sheet_names)} sheets: {', '.join(sheet_names)}")
        
        # Let user select the source sheets
        st.subheader("Sheet Selection")
        
        # Default to sheets that are not the template
        default_sheets = [s for s in sheet_names if 'template' not in s.lower()]
        
        source_sheets = st.multiselect(
            "Select the Source Sheets (containing PPID/Pim Parent ID data):",
            sheet_names,
            default=default_sheets if default_sheets else sheet_names
        )
        
        if st.button("Process File", type="primary"):
            with st.spinner("Processing..."):
                # Read source sheets
                source_dfs = []
                for sheet in source_sheets:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet)
                    source_dfs.append(df)
                    st.write(f"**{sheet}** - {len(df)} rows")
                
                # Combine all source data
                all_source = pd.concat(source_dfs, ignore_index=True)
                
                # Column mapping: Template column -> Source column name
                column_mapping = {
                    'PPID': 'Pim Parent ID',
                    'SKU': 'Retek ID',
                    'BARCODE': 'Barcode',
                    'DESCRIPTION': 'Retek Item Description',
                    'COLOUR': 'Diff 1 Description',
                    'SIZE': 'UK Size Concat',
                    'PRODUCT TYPE': 'Product Type UDA',
                    'DIVISION': 'Division Name',
                    'BRAND': 'Brand',
                    'DEPARTMENT': 'Department Name',
                    'DEPARTMENT NUMBER': 'Department Number',
                    'DIVISION NUMBER': 'Division Number',
                    'STORE 301 ALLOCATION': 'Store 301 Allocation',
                    'STORE 401 ALLOCATION': 'Store 401 Allocation',
                    'ITEM STORE FLAG': 'Item Store Flag',
                    'VPN PARENT': 'VPN Parent'
                }
                
                # Find Pim Parent ID column
                ppid_col = None
                for col in all_source.columns:
                    if 'pim parent id' in col.lower():
                        ppid_col = col
                        break
                
                if ppid_col is None:
                    st.error("Could not find 'Pim Parent ID' column in source sheets!")
                else:
                    st.success(f"Found Pim Parent ID column: '{ppid_col}'")
                    
                    # Get unique PPIDs
                    unique_ppids = all_source[ppid_col].dropna().unique()
                    st.info(f"Found {len(unique_ppids)} unique PPIDs")
                    
                    # Create output
                    output_data = []
                    
                    for ppid in unique_ppids:
                        matching_rows = all_source[all_source[ppid_col] == ppid]
                        
                        row_data = {'PPID': ppid}
                        
                        for template_col, source_col in column_mapping.items():
                            if template_col == 'PPID':
                                continue
                            
                            # Find source column (case-insensitive)
                            source_col_actual = None
                            for col in all_source.columns:
                                if col.lower() == source_col.lower():
                                    source_col_actual = col
                                    break
                            
                            if source_col_actual:
                                values = matching_rows[source_col_actual].dropna()
                                if len(values) > 0:
                                    value = values.iloc[0]
                                    
                                    # BARCODE - convert to integer (remove decimals)
                                    if template_col == 'BARCODE' and pd.notna(value):
                                        try:
                                            value = int(float(value))
                                        except (ValueError, TypeError):
                                            pass
                                    
                                    row_data[template_col] = value
                                else:
                                    row_data[template_col] = None
                            else:
                                row_data[template_col] = None
                        
                        output_data.append(row_data)
                    
                    # Create output dataframe with correct column order
                    output_columns = ['PPID', 'SKU', 'BARCODE', 'DESCRIPTION', 'COLOUR', 'SIZE', 
                                     'PRODUCT TYPE', 'DIVISION', 'BRAND', 'DEPARTMENT', 
                                     'DEPARTMENT NUMBER', 'DIVISION NUMBER', 'STORE 301 ALLOCATION', 
                                     'STORE 401 ALLOCATION', 'ITEM STORE FLAG', 'VPN PARENT']
                    
                    output_df = pd.DataFrame(output_data, columns=output_columns)
                    
                    st.success(f"✅ Created {len(output_df)} rows in output file")
                    
                    # Display preview of processed data
                    st.subheader("Preview of Processed Data")
                    st.dataframe(output_df)
                    
                    # Generate output filename with timestamp
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"Processed_Manual_Transfer_File_{timestamp}.xls"
                    
                    # Create .xls file using xlwt
                    workbook = xlwt.Workbook()
                    sheet = workbook.add_sheet('Processed Data')
                    
                    # Write headers
                    for col_idx, col_name in enumerate(output_columns):
                        sheet.write(0, col_idx, col_name)
                    
                    # Write data
                    for row_idx, row in output_df.iterrows():
                        for col_idx, col_name in enumerate(output_columns):
                            value = row[col_name]
                            if pd.isna(value):
                                sheet.write(row_idx + 1, col_idx, '')
                            else:
                                sheet.write(row_idx + 1, col_idx, value)
                    
                    # Save to buffer
                    output = io.BytesIO()
                    workbook.save(output)
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Download Processed File",
                        data=output,
                        file_name=output_filename,
                        mime="application/vnd.ms-excel"
                    )
                    
                    st.markdown("---")
                    st.subheader("Processing Summary")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Unique PPIDs", len(unique_ppids))
                    with col2:
                        st.metric("Output Rows", len(output_df))
                    with col3:
                        st.metric("Source Sheets Used", len(source_sheets))
                    
                    # Show column mapping status
                    st.subheader("Column Mapping Status")
                    mapping_status = []
                    for template_col, source_col in column_mapping.items():
                        found = any(col.lower() == source_col.lower() for col in all_source.columns)
                        status = "✅ Found" if found else "❌ Not Found"
                        mapping_status.append({
                            'Template Column': template_col,
                            'Source Column': source_col,
                            'Status': status
                        })
                    st.dataframe(pd.DataFrame(mapping_status))
                        
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.exception(e)

else:
    st.info("👆 Please upload an XLS/XLSX file to begin processing")
    
    st.markdown("---")
    st.subheader("Column Mapping Reference")
    st.markdown("""
    | Template Column | Source Column |
    |-----------------|---------------|
    | PPID | Pim Parent ID |
    | SKU | Retek ID |
    | BARCODE | Barcode (converted to integer) |
    | DESCRIPTION | Retek Item Description |
    | COLOUR | Diff 1 Description |
    | SIZE | UK Size Concat |
    | PRODUCT TYPE | Product Type UDA |
    | DIVISION | Division Name |
    | BRAND | Brand |
    | DEPARTMENT | Department Name |
    | DEPARTMENT NUMBER | Department Number |
    | DIVISION NUMBER | Division Number |
    | STORE 301 ALLOCATION | Store 301 Allocation |
    | STORE 401 ALLOCATION | Store 401 Allocation |
    | ITEM STORE FLAG | Item Store Flag |
    | VPN PARENT | VPN Parent |
    """)
