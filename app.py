import streamlit as st
import pandas as pd
import os
import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# --- 1. ROBUST DATA LOADING ---
@st.cache_data
def load_all_models():
    all_data = []
    # Get all CSV files from the folder
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]
    
    for file in csv_files:
        try:
            # Skip non-data files
            if any(x in file.lower() for x in ["template", "quote", "requirements"]):
                continue
            
            # Use the filename (minus the extension) as the category
            category = file.replace('.csv', '').split('-')[-1].strip()
            
            # Read file and look for the header row
            df_raw = pd.read_csv(file, header=None).fillna('')
            
            model_col_idx = -1
            header_row = 0
            
            # Scan first 20 rows to find "Model" or "Bewis No"
            for r_idx in range(min(len(df_raw), 20)):
                row_vals = [str(v).strip().lower() for v in df_raw.iloc[r_idx]]
                if 'model' in row_vals or 'bewis no' in row_vals:
                    model_col_idx = row_vals.index('model') if 'model' in row_vals else row_vals.index('bewis no')
                    header_row = r_idx
                    break
            
            if model_col_idx != -1:
                df = pd.read_csv(file, header=header_row).fillna('')
                df.columns = [str(c).strip() for c in df.columns]
                m_col_name = df.columns[model_col_idx]
                
                for _, row in df.iterrows():
                    m_name = str(row[m_col_name]).strip()
                    if not m_name or m_name.lower() in ['model', 'nan', 'bewis no']:
                        continue
                    
                    # Gather technical specs for the Remark column
                    specs = []
                    # Check for Axis
                    ax_col = next((c for c in df.columns if "axis" in c.lower()), None)
                    if ax_col and row[ax_col]: specs.append(f"Axis: {row[ax_col]}")
                    
                    # Check for key performance metrics
                    for col in df.columns:
                        if any(k in col.lower() for k in ['accuracy', 'range', 'output', 'resolution']):
                            if row[col]: specs.append(f"{col}: {row[col]}")
                    
                    all_data.append({
                        "Model": m_name,
                        "Category": category,
                        "Specs": "\n".join(specs)
                    })
        except:
            continue
    return pd.DataFrame(all_data)

# --- 2. EXCEL MERGED CELL SAFETY ---
def safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

# --- 3. INTERFACE ---
st.set_page_config(layout="wide", page_title="BWSensing Quote Tool")
model_db = load_all_models()

# Initialize session state with a unique name to avoid 'method' error
if 'quote_rows' not in st.session_state:
    st.session_state.quote_rows = [{"model": ""}]

st.sidebar.title("Quotation Settings")
cust_name = st.sidebar.text_input("Customer Name", "Tony Sprent")
cust_addr = st.sidebar.text_area("Address", "University of Tasmania")
short_code = st.sidebar.text_input("Country Short Code", "SA").upper()
exchange_rate = st.sidebar.number_input("RMB to USD Rate", value=7.2)

today = datetime.date.today()
quote_id = f"{today.strftime('%Y%m%d')}-MC-{short_code}"

st.title(f"Official Quote: {quote_id}")

if model_db.empty:
    st.warning("⚠️ Reading files... If this stays, ensure your CSVs are uploaded to GitHub.")
else:
    st.info(f"Loaded {len(model_db)} models from your database.")

# Selection Table
final_items = []
for i, row in enumerate(st.session_state.quote_rows):
    c1, c2, c3 = st.columns([3, 1, 1])
    
    options = [""] + sorted(model_db['Model'].unique().tolist()) if not model_db.empty else [""]
    
    selected_model = c1.selectbox(f"Product {i+1}", options, key=f"model_{i}")
    qty = c2.number_input("Quantity", min_value=1, key=f"qty_{i}")
    price_rmb = c3.number_input("Price (RMB)", key=f"price_{i}")
    
    if selected_model and not model_db.empty:
        match = model_db[model_db['Model'] == selected_model].iloc[0]
        final_items.append({
            "model": selected_model,
            "desc": match['Category'],
            "qty": qty,
            "usd": round(price_rmb / exchange_rate, 2),
            "remark": match['Specs']
        })

if st.button("➕ Add Item Row"):
    st.session_state.quote_rows.append({"model": ""})
    st.rerun()

# --- 4. EXCEL EXPORT ---
if st.button("🚀 Generate Quote Excel"):
    template = 'template.xlsx'
    if not os.path.exists(template):
        st.error("Please upload 'template.xlsx' to your GitHub repository.")
    elif not final_items:
        st.error("Please select at least one product.")
    else:
        wb = load_workbook(template)
        ws = wb.active
        
        # Static Header Info
        safe_write(ws, 4, 9, today.strftime("%B, %dth. %Y")) # I4: Date
        safe_write(ws, 6, 9, quote_id)                      # I6: Quote No
        safe_write(ws, 10, 2, cust_name)                     # B10: Name
        safe_write(ws, 12, 2, cust_addr)                     # B12: Address
        
        # Product Table (Starts row 17)
        for idx, item in enumerate(final_items):
            row_idx = 17 + idx
            safe_write(ws, row_idx, 1, item['desc'])
            safe_write(ws, row_idx, 4, item['model'])
            safe_write(ws, row_idx, 5, item['qty'])
            safe_write(ws, row_idx, 6, item['usd'])
            safe_write(ws, row_idx, 7, item['usd'] * item['qty'])
            safe_write(ws, row_idx, 9, item['remark'])
            
        output = BytesIO()
        wb.save(output)
        st.download_button("📥 Download Final Excel", output.getvalue(), f"{quote_id}.xlsx")
