import streamlit as st
import pandas as pd
import os
import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# --- 1. UPDATED DATA LOADING (Reads CSV and XLSX) ---
@st.cache_data
def load_all_models():
    all_data = []
    # Look for both CSV and Excel files
    data_files = [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx'))]
    
    for file in data_files:
        try:
            # Skip templates or system files
            if any(x in file.lower() for x in ["template", "requirements", "packages"]):
                continue
            
            # Read logic based on file type
            if file.endswith('.csv'):
                df_raw = pd.read_csv(file, header=None).fillna('')
            else:
                # For Excel files, read the first sheet
                df_raw = pd.read_excel(file, header=None).fillna('')
            
            category = file.replace('.csv', '').replace('.xlsx', '').split('-')[-1].strip()
            model_col_idx = -1
            header_row = 0
            
            # Scan for the "Model" column
            for r_idx in range(min(len(df_raw), 20)):
                row_vals = [str(v).strip().lower() for v in df_raw.iloc[r_idx]]
                if 'model' in row_vals or 'bewis no' in row_vals:
                    model_col_idx = row_vals.index('model') if 'model' in row_vals else row_vals.index('bewis no')
                    header_row = r_idx
                    break
            
            if model_col_idx != -1:
                # Reload with the correct header
                if file.endswith('.csv'):
                    df = pd.read_csv(file, header=header_row).fillna('')
                else:
                    df = pd.read_excel(file, header=header_row).fillna('')
                
                df.columns = [str(c).strip() for c in df.columns]
                m_col_name = df.columns[model_col_idx]
                
                for _, row in df.iterrows():
                    m_name = str(row[m_col_name]).strip()
                    if not m_name or m_name.lower() in ['model', 'nan', 'bewis no']:
                        continue
                    
                    specs = []
                    # Check for key specs
                    for col in df.columns:
                        if any(k in col.lower() for k in ['accuracy', 'range', 'axis', 'output']):
                            if str(row[col]).strip():
                                specs.append(f"{col}: {row[col]}")
                    
                    all_data.append({
                        "Model": m_name,
                        "Category": category,
                        "Specs": "\n".join(specs)
                    })
        except:
            continue
    return pd.DataFrame(all_data)

# --- 2. THE REST OF THE APP REMAINS THE SAME ---
# (Include the safe_write function and UI code from the previous message here)
