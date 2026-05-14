import streamlit as st
import pandas as pd
import os
import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# --- 1. SUPER FLEXIBLE DATA LOADING ---
@st.cache_data
def load_all_models():
    all_data = []
    # Look for ANY csv file in the directory
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]
    
    for file in csv_files:
        try:
            # Skip templates or already generated quotes
            if "template" in file.lower() or "Quote" in file:
                continue
            
            # Category name comes from filename
            category = file.split('-')[-1].replace('.csv', '').strip()
            
            # Read file - BWS lists often have headers on row 1 or 2
            df = pd.read_csv(file, header=None).fillna('')
            
            # SEARCH FOR THE HEADER ROW
            model_col_idx = -1
            header_row_idx = 0
            
            for r_idx in range(min(len(df), 15)): # Check first 15 rows for headers
                row_values = [str(val).strip().lower() for val in df.iloc[r_idx]]
                # Look for "Model", "Bewis No", or "Model list"
                if any(x in row_values for x in ['model', 'bewis no', 'model list']):
                    if 'model' in row_values:
                        model_col_idx = row_values.index('model')
                    elif 'bewis no' in row_values:
                        model_col_idx = row_values.index('bewis no')
                    header_row_idx = r_idx
                    break
            
            if model_col_idx != -1:
                # Re-read the file starting from the header we found
                clean_df = pd.read_csv(file, header=header_row_idx).fillna('')
                clean_df.columns = [str(c).strip() for c in clean_df.columns]
                m_col_name = clean_df.columns[model_col_idx]
                
                for _, row in clean_df.iterrows():
                    m_name = str(row[m_col_name]).strip()
                    # Skip empty rows or the header repeating
                    if not m_name or m_name.lower() in ['model', 'nan', 'bewis no', 'model list']:
                        continue
                    
                    specs = []
                    # 1. ADD AXIS OPTION (from your request)
                    axis_col = next((c for c in clean_df.columns if "axis" in c.lower()), None)
                    if axis_col and row[axis_col]:
                        specs.append(f"Axis: {row[axis_col]}")
                    
                    # 2. Key parameters
                    keys = ['accuracy', 'resolution', 'range', 'output', 'power']
                    for col in clean_df.columns:
                        if any(k in col.lower() for k in keys) and row[col]:
                            specs.append(f"{col}: {row[col]}")
                    
                    all_data.append({
                        "Model": m_name,
                        "Category": category,
                        "Specs": "\n".join(specs)
                    })
        except Exception as e:
            continue
            
    return pd.DataFrame(all_data)

# --- 2. MERGED CELL WRITER ---
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
st.set_page_config(layout="wide")
model_db = load_all_models()

st.sidebar.title("Quotation Settings")
c_name = st.sidebar.text_input("Customer Name")
c_addr = st.sidebar.text_area("Address")
country_code = st.sidebar.text_input("Country Short Code (e.g. SA, AU)", value="SA").upper()
rmb_rate = st.sidebar.number_input("RMB to USD Exchange Rate", value=7.2)

today = datetime.date.today()
quote_no = f"{today.strftime('%Y%m%d')}-MC-{country_code}"

st.header(f"Quotation Generator: {quote_no}")

if model_db.empty:
    st.error("❌ No models found. Please check that your CSV files are in the main folder on GitHub.")
    # Debug info for you
    all_files = [f for f in os.listdir('.') if f.endswith('.csv')]
    st.info(f"Files currently seen by app: {all_files}")
else:
    st.success(f"✅ Successfully loaded {len(model_db)} models from your files.")

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

selected_items = []
for i, row in enumerate(st.session_state.rows):
    cols = st.columns([3, 1, 1])
    # The selectbox allows you to type to search
    m_options = [""] + sorted(model_db['Model'].unique().tolist()) if not model_db.empty else [""]
    choice = cols[0].selectbox(f"Product {i+1}", m_options, key=f"sel_{i}")
    qty = cols[1].number_input("Qty", min_value=1, key=f"qty_{i}")
    rmb = cols[2].number_input("RMB Price", key=f"rmb_{i}")
    
    if choice:
        data = model_db[model_db['Model'] == choice].iloc[0]
        selected_items.append({
            "model": choice, "desc": data['Category'], "qty": qty,
            "remark": data['Specs'], "usd": round(rmb/rmb_rate, 2)
        })

if st.button("➕ Add Item"):
    st.session_state.rows.append({"model": ""})
    st.rerun()

# --- 4. EXPORT ---
if st.button("🚀 Generate Official Quote"):
    t_path = 'template.xlsx' if os.path.exists('template.xlsx') else 'template.xls'
    if not os.path.exists(t_path):
        st.error("Missing template.xlsx! Please upload your original Excel file to GitHub.")
    else:
        wb = load_workbook(t_path)
        ws = wb.active
        
        # Fill Header
        safe_write(ws, 4, 9, today.strftime("%B, %dth. %Y")) # I4
        safe_write(ws, 6, 9, quote_no)                      # I6
        safe_write(ws, 10, 2, c_name)                        # B10
        safe_write(ws, 12, 2, c_addr)                        # B12

        # Fill Table
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
