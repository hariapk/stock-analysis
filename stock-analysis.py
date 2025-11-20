import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import re # For robust string manipulation

# --- Configuration and Title ---
st.set_page_config(
    page_title="Stock Data Processor",
    layout="centered",
    initial_sidebar_state="auto"
)

st.title("📈 Streamlit Stock Analysis Flagging Tool")
st.markdown("Upload your Excel file to calculate industry averages, apply custom flags, sort, and download the formatted result.")

# --- File Uploader ---
uploaded_file = st.file_uploader(
    "1. Choose an Excel file (`.xlsx`)",
    type=['xlsx'],
    help="The file must contain columns like 'Industry', 'PE1', 'Market Cap (mil)', 'EG1', and 'EG2'. We will format PE1, PE2, PEG1, PEG2 to one decimal place, EG1 and EG2 as percentages (0 decimals), and EPS0, EPS1, EPS2 to two decimal places."
)

def find_robust_column(df_columns, required_key):
    """
    Finds the actual column name in the DataFrame based on a required key,
    handling casing, spacing, and punctuation variations (like parentheses).
    Returns the found column name or None.
    """
    # Function to clean a string: remove spaces, lowercase, remove punctuation (like parentheses)
    def clean_string(s):
        s = s.lower().replace(' ', '')
        # Remove common punctuation symbols for robustness
        s = re.sub(r'[()\[\]\{\}\.\-\_]', '', s)
        return s

    # 1. Standardize the required key for search
    standard_key = clean_string(required_key)
    
    # 2. Map cleaned column names to their original names
    standardized_columns = {clean_string(col): col for col in df_columns}
    
    # 3. Check for exact match of the standardized key
    if standard_key in standardized_columns:
        return standardized_columns[standard_key]

    # 4. Handle common abbreviations (e.g., matching 'marketcapmil')
    if 'marketcapmil' in standardized_columns:
        return standardized_columns['marketcapmil']

    # 5. Fallback fuzzy search 
    for std_col, original_col in standardized_columns.items():
        if standard_key in std_col or std_col in standard_key:
             return original_col
    
    return None

@st.cache_data
def process_excel_data(uploaded_file):
    """Reads, processes, calculates flags, sorts data, and returns a processed Pandas DataFrame."""
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

    # --- 1. Data Preparation and Robust Column Mapping ---
    
    # Define the 5 core columns needed for calculation
    required_cols = {
        'industry': None,
        'pe1': None,
        'market cap (mil)': None, 
        'eg1': None,
        'eg2': None,
    }
    
    # Map the required keys to the actual column names in the DataFrame
    for key in required_cols.keys():
        actual_col = find_robust_column(df.columns, key)
        if actual_col:
            required_cols[key] = actual_col
        else:
            st.error(f"FATAL ERROR: Could not find a matching column for the required field: **'{key}'**.")
            st.info(f"The code requires a column for '{key}'. Please check your Excel headers: {list(df.columns)}")
            return None

    # --- Columns for Rounding ---
    
    # Identify columns for 0 decimal place rounding (EG1, EG2)
    zero_decimal_cols_keys = ['eg1', 'eg2']
    actual_zero_decimal_cols = []
    
    for key in zero_decimal_cols_keys:
        # NOTE: EG1 and EG2 are already in required_cols, so we just map them for rounding.
        # Use the actual column name found in the initial required_cols mapping
        if key in required_cols and required_cols[key]:
            actual_zero_decimal_cols.append(required_cols[key])
        
    # Identify columns for 1 decimal place rounding (PE1, PE2, PEG1, PEG2)
    one_decimal_cols_keys = ['pe1', 'pe2', 'peg1', 'peg2']
    actual_one_decimal_cols = []
    
    for key in one_decimal_cols_keys:
        actual_col = find_robust_column(df.columns, key)
        if actual_col:
            actual_one_decimal_cols.append(actual_col)
            # Ensure PE2, PEG1, PEG2 are mapped if they weren't in the 5 core (PE1 is always core)
            if key not in required_cols:
                 required_cols[key] = actual_col
            
    # Identify columns for 2 decimal place rounding (EPS0, EPS1, EPS2)
    two_decimal_cols_keys = ['eps0', 'eps1', 'eps2']
    actual_two_decimal_cols = []
    
    for key in two_decimal_cols_keys:
        actual_col = find_robust_column(df.columns, key)
        if actual_col:
            actual_two_decimal_cols.append(actual_col)
            # Add to required_cols to ensure they are handled if not already present
            required_cols[key] = actual_col 


    # Create a simplified map for easy access
    col_map = {k: v for k, v in required_cols.items()}

    # All columns that need to be numeric: core calculation columns + rounding columns
    core_calc_cols = [col_map['pe1'], col_map['market cap (mil)'], col_map['eg1'], col_map['eg2']]
    
    # Combine all numeric columns and use set to ensure unique list of actual column names
    numeric_cols_actual = list(set(core_calc_cols + actual_one_decimal_cols + actual_two_decimal_cols + actual_zero_decimal_cols)) 

    for col in numeric_cols_actual:
        # Check if the column exists before converting
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
    # Apply 0 decimal rounding to EG columns (This is needed for the number format "0%")
    for col in actual_zero_decimal_cols:
        if col in df.columns:
            # We round to 0 decimal places to get a whole number value for the percentage format
            df[col] = df[col].round(0)
            
    # Apply 1 decimal rounding to the PE/PEG columns
    for col in actual_one_decimal_cols:
        if col in df.columns:
            df[col] = df[col].round(1)
    
    # Apply 2 decimal rounding to the requested EPS columns
    for col in actual_two_decimal_cols:
        if col in df.columns:
            df[col] = df[col].round(2)


    # Drop rows where 'Industry' or 'PE1' is missing/invalid
    df.dropna(subset=[col_map['industry'], col_map['pe1']], inplace=True)
    
    # --- 2. Calculation of Industry Average PE1 ---
    industry_avg_pe = df.groupby(col_map['industry'])[col_map['pe1']].mean().reset_index()
    industry_avg_pe.rename(columns={col_map['pe1']: 'IndustryAvgPE1_Calc'}, inplace=True) # Use a temp name

    # Merge industry average back
    df = pd.merge(df, industry_avg_pe, on=col_map['industry'], how='left')

    # --- 3. Feature/Flag Creation ---
    
    # SpecialFlag: PE1 above industry average AND Market Cap 3k-10k
    df['SpecialFlag'] = (
        (df[col_map['pe1']] > df['IndustryAvgPE1_Calc']) &
        (df[col_map['market cap (mil)']] >= 3000) &
        (df[col_map['market cap (mil)']] <= 10000)
    )

    # EG2_gt_EG1 flag: True if EG2 > EG1
    df['EG2_gt_EG1'] = df[col_map['eg2']] > df[col_map['eg1']]

    # Round the calculated average to 1 decimal place (to match PE1, PE2, etc.)
    df['IndustryAvgPE1'] = df['IndustryAvgPE1_Calc'].round(1)
    df.drop(columns=['IndustryAvgPE1_Calc'], inplace=True)

    # --- 4. Sorting ---
    df_sorted = df.sort_values(
        by=[col_map['industry'], col_map['pe1']],
        ascending=[True, False]
    )
    
    # Reset index for clean export
    df_sorted.reset_index(drop=True, inplace=True)
    
    # --- 5. Column Renaming and Ordering to Match Request ---
    
    # Rename the 5 core columns to standardized names for consistency and shading function compliance
    standardized_names = {
        col_map['industry']: 'Industry',
        col_map['pe1']: 'PE1',
        col_map['market cap (mil)']: 'Market Cap (mil)',
        col_map['eg1']: 'EG1',
        col_map['eg2']: 'EG2',
    }
    
    # Preserve original names for all other mapped columns (like PE2, PEG1, EPS0, etc.)
    for key, actual_col in col_map.items():
        if key not in ['industry', 'pe1', 'market cap (mil)', 'eg1', 'eg2']:
            standardized_names[actual_col] = actual_col # Keep original name

    df_sorted.rename(columns=standardized_names, inplace=True)
    
    # Get all original columns (now potentially with standardized names for the 5 core ones)
    calculated_cols = ['IndustryAvgPE1', 'SpecialFlag', 'EG2_gt_EG1']
    
    original_non_calculated_cols = [
        col for col in df_sorted.columns
        if col not in calculated_cols
    ]
    
    # Define the final order requested by the user (all originals, then the 3 calculated ones)
    final_output_cols = original_non_calculated_cols + calculated_cols

    # Select and reorder the DataFrame, ensuring we only select columns that exist
    df_final = df_sorted[[col for col in final_output_cols if col in df_sorted.columns]]
    
    return df_final

def apply_shading_and_save(df):
    """Applies openpyxl formatting to a dataframe and saves to BytesIO."""
    output = BytesIO()
    
    # Use Pandas to export to the buffer, which creates the structure
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Processed Stocks', index=False)
    
    # Load the workbook from the same buffer to apply styles
    wb = load_workbook(output)
    ws = wb['Processed Stocks']

    # Define the light gray fill
    gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    # Identify column indices for Industry and Percentage columns
    industry_col = None
    percentage_cols = ['EG1', 'EG2']
    percentage_col_indices = {}
    
    for col_num, cell in enumerate(ws[1], start=1):
        if cell.value == "Industry": 
            industry_col = col_num
        if cell.value in percentage_cols:
             percentage_col_indices[cell.value] = col_num

    # Apply alternating shading by industry
    if industry_col:
        current_industry = None
        gray_toggle = False
        # Start from row 2 (index 2) for data
        for row in range(2, ws.max_row + 1):
            industry_value = ws.cell(row=row, column=industry_col).value
            
            # Toggle shading when industry changes
            if industry_value != current_industry:
                if industry_value is not None:
                    gray_toggle = not gray_toggle
                current_industry = industry_value
            
            # Apply fill if toggled
            if gray_toggle:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = gray_fill
    
    # Apply percentage number format to EG1 and EG2 columns
    # Excel format "0%" displays the value of 15 as 15% (i.e., it doesn't multiply by 100, just appends %)
    percentage_format = "0%"
    for col_name, col_num in percentage_col_indices.items():
        # Apply format to all data rows (starting from row 2)
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_num)
            cell.number_format = percentage_format


    # Save the modified workbook back to a new BytesIO buffer
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# --- Main App Logic ---
if uploaded_file is not None:
    st.info("File successfully uploaded. Processing data...")
    
    df_processed = process_excel_data(uploaded_file)
    
    if df_processed is not None and not df_processed.empty:
        st.success(f"Processing complete! {len(df_processed)} records analyzed.")
        
        # Display Summary
        st.subheader("Processing Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Records", len(df_processed))
        col2.metric("Special Flags (True)", df_processed['SpecialFlag'].sum())
        col3.metric("New Columns", 3)
        
        # Display the first few rows of the data
        st.subheader("Preview of Processed Data (First 10 Rows) - Check Column Order")
        st.dataframe(df_processed.head(10), use_container_width=True)

        # --- Download Button ---
        st.markdown("---")
        st.subheader("2. Download Results")
        
        # Prepare the final excel file for download
        excel_bytes = apply_shading_and_save(df_processed)
        
        st.download_button(
            label="Download stocks_ranked_with_flags.xlsx",
            data=excel_bytes,
            file_name="stocks_ranked_with_flags.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.markdown(
            """
            <div style="font-size: 0.9em; color: gray;">
                The downloaded file is sorted by Industry (ASC) and PE1 (DESC) and includes alternating row shading.
            </div>
            """, unsafe_allow_html=True
        )
    elif df_processed is not None:
        st.warning("Processed DataFrame is empty. Check if your input file contains the required data and columns.")
