import streamlit as st
import pandas as pd
import os
import datetime
import requests
from bs4 import BeautifulSoup
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.cell.cell import MergedCell

# --- 1. DEEP SCAN DATA LOADING ---
@st.cache_data
def load_all_models():
    all_data = []
    # Search for all CSVs. We won't restrict by name anymore to ensure we catch everything.
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]
    
    for file in csv_files:
        try:
            # Skip the template files if they were converted to CSV
            if "template" in file.lower() or "Quotation" in file: continue
            
            category = file.replace('.csv', '').split('-')[-1].strip()
            df = pd.read_csv(file, header=None).fillna('')
            
            # DEEP SCAN: Find which column and row the data actually starts
            model_col_idx = -1
            start_row = 0
            
            for r_idx in range(min(len(df), 10)): # Look at first 10 rows
                row_vals = [str(x).strip().lower() for x in df.iloc[r_idx]]
                if 'model' in row_vals or 'bewis no' in row_vals:
                    model_col_idx = row_vals.index('model') if 'model' in row_vals else row_vals.index('bewis no')
                    start_row = r_idx + 1
                    break
            
            if model_col_idx != -1:
                # Re-read with correct header
                df = pd.read_csv(file, header=start_row-1).fillna('')
                df.columns = [str(c).strip() for c in df.columns]
                m_col_name = df.columns[model_col_idx]
                
                for _, row in df.iterrows():
                    m_name = str(row[m_col_name]).strip()
                    if not m_name or m_name.lower() in ['model', 'nan', 'bewis no']: continue
                    
                    specs = []
                    # 1. Axis detection
                    axis_col = next((c for c in df.columns if "axis" in c.lower()), None)
                    if axis_col and row[axis_col]: specs.append(f"Axis: {row[axis_col]}")
                    
                    # 2. Key parameters
                    keys = ['accuracy', 'resolution', 'range', 'output']
                    for col in df.columns:
                        if any(k in col.lower() for k in keys) and row[col]:
                            specs.append(f"{col}: {row[col]}")
                    
                    all_data.append({
                        "Model": m_name,
                        "Category": category,
                        "Specs": "\n".join(specs)
                    })
        except: continue
    
    return pd.DataFrame(all_data)

# --- 2. WRITING SAFETY ---
def safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

# --- 3. UI ---
st.set_page_config(layout="wide", page_title="BWSensing Official Generator")
model_db = load_all_models()

# Sidebar
st.sidebar.header("Quote Parameters")
c_name = st.sidebar.text_input("Customer Name")
c_addr = st.sidebar.text_area("Address")
country_code = st.sidebar.text_input("Country Short Code (e.g. SA)", value="SA").upper()
rmb_rate = st.sidebar.number_input("RMB to USD Exchange", value=7.2)

today = datetime.date.today()
quote_no = f"{today.strftime('%Y%m%d')}-MC-{country_code}"

# Main App
st.title(f"Quotation: {quote_no}")

if model_db.empty:
    st.error("❌ No models found! Please ensure your CSV files (like 'Model list - Tilt.csv') are in the GitHub folder.")
else:
    st.success(f"✅ Found {len(model_db)} models in your database.")

if 'items' not in st.session_state:
    st.session_state.items = [{"model": ""}]

# Build Rows
selected_items = []
for i, row in enumerate(st.session_state.items):
    cols = st.columns([3, 1, 1])
    m_options = [""] + sorted(model_db['Model'].unique().tolist()) if not model_db.empty else [""]
    
    choice = cols[0].selectbox(f"Select Model {i+1}", m_options, key=f"sel_{i}")
    qty = cols[1].number_input("Qty", min_value=1, key=f"qty_{i}")
    rmb = cols[2].number_input("RMB Price", key=f"rmb_{i}")
    
    if choice:
        data = model_db[model_db['Model'] == choice].iloc[0]
        selected_items.append({
            "model": choice, "desc": data['Category'], "qty": qty,
            "remark": data['Specs'], "usd": round(rmb/rmb_rate, 2)
        })

if st.button("➕ Add Another Product"):
    st.session_state.items.append({"model": ""})
    st.rerun()

# --- 4. EXCEL GENERATION ---
if st.button("🚀 Create Official Quote"):
    t_path = 'template.xlsx' if os.path.exists('template.xlsx') else 'template.xls'
    if not os.path.exists(t_path):
        st.error("Error: Please upload 'template.xlsx' to GitHub.")
    else:
        wb = load_workbook(t_path)
        ws = wb.active
        
        # Fill Template
        safe_write(ws, 4, 9, today.strftime("%B, %dth. %Y")) # I4
        safe_write(ws, 6, 9, quote_no)                      # I6
        safe_write(ws, 10, 2, c_name)                        # B10
        safe_write(ws, 12, 2, c_addr)                        # B12

        for idx, item in enumerate(selected_items):
            r = 17 + idx
            safe_write(ws, r, 1, item['desc'])
            safe_write(ws, r, 4, item['model'])
            safe_write(ws, r, 5, item['qty'])
            safe_write(ws, r, 6, item['usd'])
            safe_write(ws, r, 7, item['usd'] * item['qty'])
            safe_write(ws, r, 9, item['remark'])

        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Click to Download", out.getvalue(), f"{quote_no}.xlsx")
