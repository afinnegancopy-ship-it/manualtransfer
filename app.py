import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Brown Thomas Manual Transfer Processor", layout="wide")

st.title("Brown Thomas Manual Transfer File Processor")
st.markdown("---")

# Column mapping from source sheets to template
COLUMN_MAPPING = {
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

def load_excel_file(uploaded_file):
    """Load all sheets from the uploaded Excel file"""
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheets = {}
        for sheet_name in xls.sheet_names:
            sheets[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
        return sheets, xls.sheet_names
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None, None

def clean_barcode(barcode_value):
    """Convert barcode to number and remove decimal digits"""
    try:
        if pd.isna(barcode_value):
            return None
        # Convert to float first, then to int to remove decimals
        return int(float(barcode_value))
    except (ValueError, TypeError):
        return barcode_value

def build_lookup_dict(source_sheets, template_sheet_name):
    """Build a lookup dictionary from all source sheets (excluding template) keyed by PPID"""
    lookup = {}
    
    for sheet_name, df in source_sheets.items():
        if sheet_name == template_sheet_name:
            continue
            
        # Find the PPID column (case-insensitive search)
        ppid_col = None
        for col in df.columns:
            if str(col).upper().strip() == 'PPID':
                ppid_col = col
                break
        
        if ppid_col is None:
            st.warning(f"Sheet '{sheet_name}' does not have a PPID column. Skipping...")
            continue
        
        st.info(f"Processing sheet: '{sheet_name}' with {len(df)} rows")
        
        # Create column mapping for this sheet (case-insensitive)
        col_map = {}
        for col in df.columns:
            col_upper = str(col).upper().strip()
            col_map[col_upper] = col
        
        # Process each row
        for idx, row in df.iterrows():
            ppid = row[ppid_col]
            if pd.isna(ppid):
                continue
            
            # Convert PPID to string for consistent matching
            ppid_key = str(ppid).strip()
            
            # Only add if not already in lookup (avoid duplicates)
            if ppid_key not in lookup:
                row_data = {}
                
                for template_col, source_col in COLUMN_MAPPING.items():
                    source_col_upper = source_col.upper().strip()
                    if source_col_upper in col_map:
                        actual_col = col_map[source_col_upper]
                        value = row[actual_col]
                        
                        # Special handling for barcode
                        if template_col == 'BARCODE':
                            value = clean_barcode(value)
                        
                        row_data[template_col] = value
                    else:
                        row_data[template_col] = None
                
                lookup[ppid_key] = row_data
    
    return lookup

def process_template(template_df, lookup):
    """Process the template by filling in data from lookup"""
    # Find PPID column in template (case-insensitive)
    ppid_col = None
    for col in template_df.columns:
        if str(col).upper().strip() == 'PPID':
            ppid_col = col
            break
    
    if ppid_col is None:
        st.error("Template sheet does not have a PPID column!")
        return None
    
    # Create a copy of the template
    result_df = template_df.copy()
    
    # Create column mapping for template (case-insensitive)
    template_col_map = {}
    for col in result_df.columns:
        col_upper = str(col).upper().strip()
        template_col_map[col_upper] = col
    
    # Track statistics
    matched = 0
    unmatched = 0
    
    # Process each row in the template
    for idx, row in result_df.iterrows():
        ppid = row[ppid_col]
        if pd.isna(ppid):
            continue
        
        ppid_key = str(ppid).strip()
        
        if ppid_key in lookup:
            matched += 1
            row_data = lookup[ppid_key]
            
            for template_col, value in row_data.items():
                if template_col in template_col_map:
                    actual_col = template_col_map[template_col]
                    if value is not None:
                        result_df.at[idx, actual_col] = value
        else:
            unmatched += 1
    
    st.success(f"✅ Matched {matched} PPIDs")
    if unmatched > 0:
        st.warning(f"⚠️ {unmatched} PPIDs not found in source sheets")
    
    return result_df

def main():
    st.markdown("""
    ### Instructions:
    1. Upload your XLS file containing the 'brownthomas_new_template' sheet and source data sheets
    2. Select the template sheet (brownthomas_new_template)
    3. The script will map data from source sheets to the template based on PPID
    4. Download the processed file
    """)
    
    uploaded_file = st.file_uploader("Upload your Excel file (.xls or .xlsx)", type=['xls', 'xlsx'])
    
    if uploaded_file is not None:
        # Load all sheets
        sheets, sheet_names = load_excel_file(uploaded_file)
        
        if sheets is not None:
            st.success(f"✅ Loaded {len(sheet_names)} sheets: {', '.join(sheet_names)}")
            
            # Let user select the template sheet
            template_sheet = st.selectbox(
                "Select the template sheet (brownthomas_new_template):",
                options=sheet_names,
                index=0 if 'brownthomas_new_template' not in sheet_names else sheet_names.index('brownthomas_new_template') if 'brownthomas_new_template' in sheet_names else 0
            )
            
            # Show preview of template
            st.subheader("Template Preview")
            st.dataframe(sheets[template_sheet].head(10))
            
            # Show source sheets info
            st.subheader("Source Sheets")
            source_sheets = [s for s in sheet_names if s != template_sheet]
            for sheet in source_sheets:
                with st.expander(f"Sheet: {sheet} ({len(sheets[sheet])} rows)"):
                    st.write("Columns:", list(sheets[sheet].columns))
                    st.dataframe(sheets[sheet].head(5))
            
            if st.button("🚀 Process File", type="primary"):
                with st.spinner("Processing..."):
                    # Build lookup from source sheets
                    st.subheader("Building Lookup Dictionary...")
                    lookup = build_lookup_dict(sheets, template_sheet)
                    st.info(f"Found {len(lookup)} unique PPIDs in source sheets")
                    
                    # Process template
                    st.subheader("Processing Template...")
                    result_df = process_template(sheets[template_sheet], lookup)
                    
                    if result_df is not None:
                        st.subheader("Processed Result Preview")
                        st.dataframe(result_df.head(20))
                        
                        # Generate filename with date and time
                        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                        output_filename = f"Processed Manual Transfer File_{timestamp}.xlsx"
                        
                        # Create download button
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False, sheet_name='Processed Data')
                        output.seek(0)
                        
                        st.download_button(
                            label="📥 Download Processed File",
                            data=output,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.success(f"✅ Processing complete! Click above to download: {output_filename}")

if __name__ == "__main__":
    main()
