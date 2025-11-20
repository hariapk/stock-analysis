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
    help="The file must contain columns like 'Industry', 'PE1', 'Market Cap in millions', 'EG1', and 'EG2'."
)

def find_robust_column(df_columns, required_key):
    """
    Finds the actual column name in the DataFrame based on a required key,
    handling casing, leading/trailing spaces, and internal differences.
    Returns the found column name or None.
    """
    # 1. Standardize the required key for search (remove spaces and lowercase)
    standard_key = required_key.lower().replace(' ', '')
    
    # 2. Standardize all actual column names to create a map (standardized -> original name)
    standardized_columns = {col.lower().replace(' ', ''): col for col in df_columns}
    
    # 3. Check for exact match of the standardized key
    if standard_key in standardized_columns:
        return standardized_columns[standard_key]

    # 4. Handle common variations (basic fuzzy search, useful if the user used e.g., 'marketcap' instead of 'marketcapinmillions')
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
    
    # IMPORTANT DEBUGGING STEP: Show the user the column names from their file
    st.warning(f"Columns found in your uploaded file (case sensitive): {list(df.columns)}")

    required_cols = {
        'industry': None,
        'pe1': None,
        'market cap in millions': None,
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
            st.info("Please verify the spelling in your Excel file and try again. The column names must be close to the required fields.")
            return None

    # Create a simplified map for easy access
    col_map = {k: v for k, v in required_cols.items()}

    # Convert essential columns to numeric using the actual column names found
    numeric_cols_actual = [col_map['pe1'], col_map['market cap in millions'], col_map['eg1'], col_map['eg2']]
    for col in numeric_cols_actual:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Drop rows where 'Industry' or 'PE1' is missing/invalid
    df.dropna(subset=[col_map['industry'], col_map['pe1']], inplace=True)
    
    # --- 2. Calculation of Industry Average PE1 ---
    industry_avg_pe = df.groupby(col_map['industry'])[col_map['pe1']].mean().reset_index()
    industry_avg_pe.rename(columns={col_map['pe1']: 'IndustryAvgPE1'}, inplace=True)

    # Merge industry average back
    df = pd.merge(df, industry_avg_pe, on=col_map['industry'], how='left')

    # --- 3. Feature/Flag Creation ---
    
    # SpecialFlag: PE1 above industry average AND Market Cap 3k-10k
    df['SpecialFlag'] = (
        (df[col_map['pe1']] > df['IndustryAvgPE1']) &
        (df[col_map['market cap in millions']] >= 3000) &
        (df[col_map['market cap in millions']] <= 10000)
    )

    # EG2_gt_EG1 flag: True if EG2 > EG1
    df['EG2_gt_EG1'] = df[col_map['eg2']] > df[col_map['eg1']]

    # Round the calculated average for cleaner display
    df['IndustryAvgPE1'] = df['IndustryAvgPE1'].round(2)

    # --- 4. Sorting ---
    df_sorted = df.sort_values(
        by=[col_map['industry'], col_map['pe1']],
        ascending=[True, False]
    )
    
    # Reset index for clean export
    df_sorted.reset_index(drop=True, inplace=True)
    
    # Rename columns back to a clean, standardized format for output
    final_cols = {
        col_map['industry']: 'Industry',
        col_map['pe1']: 'PE1',
        col_map['market cap in millions']: 'Market Cap in millions',
        col_map['eg1']: 'EG1',
        col_map['eg2']: 'EG2',
    }
    df_sorted.rename(columns=final_cols, inplace=True)
    
    return df_sorted

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

    # Identify column for Industry (now standardized as 'Industry' in the output DataFrame)
    industry_col = None
    for col_num, cell in enumerate(ws[1], start=1):
        if cell.value == "Industry": 
            industry_col = col_num
            break

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
        st.subheader("Preview of Processed Data (First 10 Rows)")
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
