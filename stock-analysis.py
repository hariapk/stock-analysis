import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO

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

@st.cache_data
def process_excel_data(uploaded_file):
    """Reads, processes, calculates flags, sorts data, and returns a processed Pandas DataFrame."""
    try:
        # Load Excel file into a Pandas DataFrame
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

    # --- 1. Data Preparation ---
    # Strip spaces from column names
    df.columns = df.columns.str.strip()

    # Convert essential columns to numeric, coercing errors to NaN
    numeric_cols = ['PE1', 'Market Cap in millions', 'EG1', 'EG2']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Drop rows where 'Industry' or 'PE1' is missing/invalid
    df.dropna(subset=['Industry', 'PE1'], inplace=True)
    
    # --- 2. Calculation of Industry Average PE1 ---
    industry_avg_pe = df.groupby('Industry')['PE1'].mean().reset_index()
    industry_avg_pe.rename(columns={'PE1': 'IndustryAvgPE1'}, inplace=True)

    # Merge industry average back
    df_merged = pd.merge(df, industry_avg_pe, on='Industry', how='left')

    # --- 3. Feature/Flag Creation ---
    
    # SpecialFlag: PE1 above industry average AND Market Cap 3k-10k
    df_merged['SpecialFlag'] = (
        (df_merged['PE1'] > df_merged['IndustryAvgPE1']) &
        (df_merged['Market Cap in millions'] >= 3000) &
        (df_merged['Market Cap in millions'] <= 10000)
    )

    # EG2_gt_EG1 flag: True if EG2 > EG1
    df_merged['EG2_gt_EG1'] = df_merged['EG2'] > df_merged['EG1']

    # Round the calculated average for cleaner display
    df_merged['IndustryAvgPE1'] = df_merged['IndustryAvgPE1'].round(2)

    # --- 4. Sorting ---
    df_sorted = df_merged.sort_values(
        by=['Industry', 'PE1'],
        ascending=[True, False]
    )
    
    # Reset index for clean export
    df_sorted.reset_index(drop=True, inplace=True)

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

    # Identify column for Industry (assuming header is in row 1)
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
                # Only toggle if the value is not None (to handle sorting edge cases)
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
    
    # Function call is outside the logic block to use the cached data across Streamlit runs
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