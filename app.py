import streamlit as st
import pandas as pd
from datetime import datetime
import io

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
        
        # Let user select the template sheet and source sheets
        st.subheader("Sheet Selection")
        
        template_sheet = st.selectbox(
            "Select the Template Sheet (brownthomas_new_template):",
            sheet_names,
            index=0
        )
        
        source_sheets = st.multiselect(
            "Select the Source Sheets (containing PPID data):",
            [s for s in sheet_names if s != template_sheet],
            default=[s for s in sheet_names if s != template_sheet]
        )
        
        if st.button("Process File", type="primary"):
            with st.spinner("Processing..."):
                # Read the template sheet
                template_df = pd.read_excel(uploaded_file, sheet_name=template_sheet)
                
                st.write("Template columns found:", template_df.columns.tolist())
                
                # Combine all source sheets into one dataframe
                source_dfs = []
                for sheet in source_sheets:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet)
                    source_dfs.append(df)
                    st.write(f"Source sheet '{sheet}' columns:", df.columns.tolist())
                
                # Combine source data and remove duplicates based on Pim Parent ID
                combined_source = pd.concat(source_dfs, ignore_index=True)
                
                # Column mapping from template to source
                column_mapping = {
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
                
                # Find the PPID column in template (case-insensitive search)
                ppid_col_template = None
                for col in template_df.columns:
                    if 'PPID' in col.upper():
                        ppid_col_template = col
                        break
                
                # Find the Pim Parent ID column in source (case-insensitive search)
                ppid_col_source = None
                for col in combined_source.columns:
                    if 'pim parent id' in col.lower():
                        ppid_col_source = col
                        break
                
                if ppid_col_template is None:
                    st.error("Could not find PPID column in template sheet!")
                elif ppid_col_source is None:
                    st.error("Could not find 'Pim Parent ID' column in source sheets!")
                else:
                    st.success(f"Found PPID column in template: '{ppid_col_template}'")
                    st.success(f"Found Pim Parent ID column in source: '{ppid_col_source}'")
                    
                    # Remove duplicate PPIDs from source, keeping first occurrence
                    combined_source_unique = combined_source.drop_duplicates(subset=[ppid_col_source], keep='first')
                    
                    st.info(f"Total source rows: {len(combined_source)}, After removing duplicates: {len(combined_source_unique)}")
                    
                    # Create a lookup dictionary from source data
                    source_lookup = combined_source_unique.set_index(ppid_col_source)
                    
                    # Process each row in template
                    processed_count = 0
                    not_found_count = 0
                    
                    for idx, row in template_df.iterrows():
                        ppid = row[ppid_col_template]
                        
                        # Skip if PPID is empty
                        if pd.isna(ppid):
                            continue
                        
                        # Look up the PPID in source data
                        if ppid in source_lookup.index:
                            source_row = source_lookup.loc[ppid]
                            
                            # Map each column
                            for template_col, source_col in column_mapping.items():
                                # Find matching column in template (case-insensitive)
                                template_col_actual = None
                                for col in template_df.columns:
                                    if col.upper() == template_col.upper():
                                        template_col_actual = col
                                        break
                                
                                # Find matching column in source (case-insensitive)
                                source_col_actual = None
                                for col in source_lookup.columns:
                                    if col.lower() == source_col.lower():
                                        source_col_actual = col
                                        break
                                
                                if template_col_actual and source_col_actual:
                                    value = source_row[source_col_actual]
                                    
                                    # Special handling for BARCODE - convert to integer
                                    if template_col.upper() == 'BARCODE' and pd.notna(value):
                                        try:
                                            value = int(float(value))
                                        except (ValueError, TypeError):
                                            pass
                                    
                                    template_df.at[idx, template_col_actual] = value
                            
                            processed_count += 1
                        else:
                            not_found_count += 1
                    
                    st.success(f"✅ Processed {processed_count} rows successfully")
                    if not_found_count > 0:
                        st.warning(f"⚠️ {not_found_count} PPIDs not found in source sheets")
                    
                    # Display preview of processed data
                    st.subheader("Preview of Processed Data")
                    st.dataframe(template_df.head(20))
                    
                    # Generate output filename with timestamp
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"Processed_Manual_Transfer_File_{timestamp}.xlsx"
                    
                    # Create download button
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        template_df.to_excel(writer, index=False, sheet_name='Processed Data')
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Download Processed File",
                        data=output,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.markdown("---")
                    st.subheader("Processing Summary")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Template Rows", len(template_df))
                    with col2:
                        st.metric("Successfully Matched", processed_count)
                    with col3:
                        st.metric("Not Found", not_found_count)
                        
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
