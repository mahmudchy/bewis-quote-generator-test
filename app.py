import streamlit as st
import pandas as pd
import os
import datetime
from io import BytesIO
import openpyxl
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# --- 1. ROBUST DATA LOADING ---
@st.cache_data
def load_all_models():
    all_data = []
    # Find all data files
    files = [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx', '.xls'))]
    
    for file in files:
        try:
            # Skip non-database files
            if any(x in file.lower() for x in ["template", "requirements", "packages", "env"]):
                continue
            
            # Load the data based on extension
            if file.endswith('.csv'):
                df_raw = pd.read_csv(file, header=None).fillna('')
            else:
                df_raw = pd.read_excel(file, header=None).fillna('')
            
            category = file.replace('.csv', '').replace('.xlsx', '').replace('.xls', '').split('-')[-1].strip()
            model_col_idx = -1
            header_row = 0
            
            # Scan for the header row
            for r_idx in range(min(len(df_raw), 25)):
                row_vals = [str(v).strip().lower() for v in df_raw.iloc[r_idx]]
                if 'model' in row_vals or 'bewis no' in row_vals:
                    model_col_idx = row_vals.index('model') if 'model' in row_vals else row_vals.index('bewis no')
                    header_row = r_idx
                    break
            
            if model_col_idx != -1:
                # Reload with proper headers
                if file.endswith('.csv'):
                    df = pd.read_csv(file, header=header_row).fillna('')
                else:
                    df = pd.read_excel(file, header=header_row).fillna('')
                
                df.columns = [str(c).strip() for c in df.columns]
                m_col_name = df.columns[model_col_idx]
                
                for _, row in df.iterrows():
                    m_name = str(row[m_col_name]).strip()
                    if not m_name or m_name.lower() in ['model', 'nan', 'bewis no', 'model list']:
                        continue
                    
                    specs = []
                    # Dynamic spec gathering
                    for col in df.columns:
                        if any(k in col.lower() for k in ['accuracy', 'range', 'axis', 'output', 'resolution']):
                            val = str(row[col]).strip()
                            if val and val.lower() != 'nan':
                                specs.append(f"{col}: {val}")
                    
                    all_data.append({
                        "Model": m_name,
                        "Category": category,
                        "Specs": "\n".join(specs)
                    })
        except Exception:
            continue # If one file fails, keep going to the next one
            
    return pd.DataFrame(all_data)

# --- 2. EXCEL MERGED CELL WRITER ---
def safe_write(ws, row, col, value):
    try:
        cell = ws.cell(row=row, column=col)
        if isinstance(cell, MergedCell):
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                    return
        else:
            cell.value = value
    except Exception:
        pass

# --- 3. UI SETUP ---
st.set_page_config(layout="wide", page_title="BWSensing Quote Tool")

# Initialize State
if 'quote_rows' not in st.session_state:
    st.session_state.quote_rows = [{"model": ""}]

# Load Data
model_db = load_all_models()

# Sidebar
st.sidebar.title("Quotation Settings")
cust_name = st.sidebar.text_input("Customer Name", "Tony Sprent")
cust_addr = st.sidebar.text_area("Address", "University of Tasmania")
short_code = st.sidebar.text_input("Country Short Code", "SA").upper()
rate = st.sidebar.number_input("RMB to USD Rate", value=7.2)

today = datetime.date.today()
quote_id = f"{today.strftime('%Y%m%d')}-MC-{short_code}"

st.title(f"Quote Generator: {quote_id}")

if model_db.empty:
    st.warning("Searching for model files... If this persists, please check your GitHub folder.")
    # Show files for debugging
    st.write("Files found:", [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx'))])
else:
    st.success(f"Database Ready: {len(model_db)} models loaded.")

# Row Builder
final_items = []
for i, _ in enumerate(st.session_state.quote_rows):
    c1, c2, c3 = st.columns([3, 1, 1])
    
    options = [""] + sorted(model_db['Model'].unique().tolist()) if not model_db.empty else [""]
    
    selected = c1.selectbox(f"Select Model {i+1}", options, key=f"mod_{i}")
    qty = c2.number_input("Qty", min_value=1, key=f"q_{i}")
    rmb = c3.number_input("Price (RMB)", key=f"p_{i}")
    
    if selected and not model_db.empty:
        match = model_db[model_db['Model'] == selected].iloc[0]
        final_items.append({
            "model": selected,
            "desc": match['Category'],
            "qty": qty,
            "usd": round(rmb / rate, 2),
            "remark": match['Specs']
        })

if st.button("➕ Add Product"):
    st.session_state.quote_rows.append({"model": ""})
    st.rerun()

# --- 4. EXPORT ---
if st.button("🚀 Generate Excel"):
    template_file = 'template.xlsx'
    if not os.path.exists(template_file):
        st.error("template.xlsx not found on GitHub!")
    elif not final_items:
        st.error("Add a product first.")
    else:
        wb = load_workbook(template_file)
        ws = wb.active
        
        safe_write(ws, 4, 9, today.strftime("%B, %dth. %Y"))
        safe_write(ws, 6, 9, quote_id)
        safe_write(ws, 10, 2, cust_name)
        safe_write(ws, 12, 2, cust_addr)
        
        for idx, item in enumerate(final_items):
            r = 17 + idx
            safe_write(ws, r, 1, item['desc'])
            safe_write(ws, r, 4, item['model'])
            safe_write(ws, r, 5, item['qty'])
            safe_write(ws, r, 6, item['usd'])
            safe_write(ws, r, 7, item['usd'] * item['qty'])
            safe_write(ws, r, 9, item['remark'])
            
        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Download", out.getvalue(), f"{quote_id}.xlsx")
