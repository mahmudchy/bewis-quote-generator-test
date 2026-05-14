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

# --- 1. DATA LOADING ---
@st.cache_data
def load_all_models():
    all_data = []
    # Check for CSV files with 'Model list' in the name
    csv_files = [f for f in os.listdir('.') if 'Model list' in f and f.endswith('.csv')]
    
    for file in csv_files:
        try:
            category = file.replace('Model list', '').replace('.csv', '').replace('-', '').strip()
            df = pd.read_csv(file, header=1).fillna('')
            df.columns = [str(c).strip() for c in df.columns]
            
            # Flexible column detection
            model_col = next((c for c in df.columns if "Model" in c), None)
            
            if model_col:
                for _, row in df.iterrows():
                    m_name = str(row[model_col]).strip()
                    if not m_name or m_name.lower() == "model": continue
                    
                    specs = []
                    # Find Axis and technical data
                    axis_col = next((c for c in df.columns if "axis" in c.lower()), None)
                    if axis_col and row[axis_col]: specs.append(f"Axis: {row[axis_col]}")
                    
                    important = ['accuracy', 'resolution', 'range', 'output']
                    for col in df.columns:
                        if any(k in col.lower() for k in important) and row[col]:
                            specs.append(f"{col}: {row[col]}")
                    
                    all_data.append({
                        "Model": m_name,
                        "Category": category,
                        "Specs": "\n".join(specs)
                    })
        except: continue
    return pd.DataFrame(all_data)

# --- 2. MERGED CELL SAFETY WRITER ---
def safe_write(ws, row, col, value):
    """Writes to a cell even if it is part of a merged range."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        # Find the master cell of the merged range
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

# --- 3. UI SETUP ---
st.set_page_config(layout="wide")
model_db = load_all_models()

st.sidebar.title("Quotation Details")
c_name = st.sidebar.text_input("Customer Name")
c_addr = st.sidebar.text_area("Address")
country_code = st.sidebar.text_input("Country Short Code", value="SA").upper()
rmb_rate = st.sidebar.number_input("RMB to USD", value=7.2)

today = datetime.date.today()
quote_no = f"{today.strftime('%Y%m%d')}-MC-{country_code}"

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

selected_items = []
for i, row in enumerate(st.session_state.rows):
    cols = st.columns([3, 1, 1])
    m_list = [""] + sorted(list(model_db['Model'].unique())) if not model_db.empty else [""]
    choice = cols[0].selectbox(f"Select Product {i+1}", m_list, key=f"sel_{i}")
    qty = cols[1].number_input("Qty", min_value=1, key=f"qty_{i}")
    rmb = cols[2].number_input("RMB Price", key=f"rmb_{i}")
    
    if choice and not model_db.empty:
        db_row = model_db[model_db['Model'] == choice].iloc[0]
        selected_items.append({
            "model": choice, "desc": db_row['Category'], "qty": qty,
            "remark": db_row['Specs'], "usd": round(rmb/rmb_rate, 2)
        })

if st.button("➕ Add Item"):
    st.session_state.rows.append({"model": ""})
    st.rerun()

# --- 4. EXPORT ---
if st.button("🚀 Generate Quote"):
    t_path = 'template.xlsx' if os.path.exists('template.xlsx') else 'template.xls'
    if not os.path.exists(t_path):
        st.error("Missing template.xlsx on GitHub!")
    else:
        wb = load_workbook(t_path)
        ws = wb.active
        
        # Header Info
        safe_write(ws, 4, 9, today.strftime("%B, %dth. %Y")) # Cell I4
        safe_write(ws, 6, 9, quote_no)                      # Cell I6
        safe_write(ws, 10, 2, c_name)                        # Cell B10
        safe_write(ws, 12, 2, c_addr)                        # Cell B12

        # Table Items
        for idx, item in enumerate(selected_items):
            r = 17 + idx
            safe_write(ws, r, 1, item['desc'])   # Col A
            safe_write(ws, r, 4, item['model'])  # Col D
            safe_write(ws, r, 5, item['qty'])    # Col E
            safe_write(ws, r, 6, item['usd'])    # Col F
            safe_write(ws, r, 7, item['usd'] * item['qty']) # Col G
            safe_write(ws, r, 9, item['remark']) # Col I

        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Download Excel", out.getvalue(), f"{quote_no}.xlsx")
