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
    files = [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx'))]
    for file in files:
        try:
            if any(x in file.lower() for x in ["template", "requirements"]): continue
            df_raw = pd.read_csv(file, header=None).fillna('') if file.endswith('.csv') else pd.read_excel(file, header=None).fillna('')
            category = file.replace('.csv', '').replace('.xlsx', '').split('-')[-1].strip()
            
            model_col_idx = -1
            header_row = 0
            for r_idx in range(min(len(df_raw), 25)):
                row_vals = [str(v).strip().lower() for v in df_raw.iloc[r_idx]]
                if 'model' in row_vals or 'bewis no' in row_vals:
                    model_col_idx = row_vals.index('model') if 'model' in row_vals else row_vals.index('bewis no')
                    header_row = r_idx
                    break
            
            if model_col_idx != -1:
                df = pd.read_csv(file, header=header_row).fillna('') if file.endswith('.csv') else pd.read_excel(file, header=header_row).fillna('')
                df.columns = [str(c).strip() for c in df.columns]
                m_col = df.columns[model_col_idx]
                for _, row in df.iterrows():
                    m_name = str(row[m_col]).strip()
                    if not m_name or m_name.lower() in ['model', 'nan']: continue
                    specs = []
                    for col in df.columns:
                        if any(k in col.lower() for k in ['accuracy', 'range', 'axis', 'output']):
                            val = str(row[col]).strip()
                            if val and val.lower() != 'nan': specs.append(f"{col}: {val}")
                    all_data.append({"Model": m_name, "Category": category, "Specs": "\n".join(specs)})
        except: continue
    return pd.DataFrame(all_data)

# --- 2. IMPROVED IMAGE FETCHING ---
def get_bwsensing_image(model_name):
    """Scrapes bwsensing.com search and extracts product thumbnail."""
    search_url = f"https://www.bwsensing.com/search.html?q={model_name}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        response = requests.get(search_url, timeout=10, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        # Check specifically for the product list images
        img_tag = soup.select_one('.product-list img') or soup.find('img', {'class': 'lazy'})
        if img_tag:
            # Try to get data-original (lazy load) or standard src
            src = img_tag.get('data-original') or img_tag.get('src')
            if src:
                return src if src.startswith('http') else f"https://www.bwsensing.com{src}"
    except: return None
    return None

# --- 3. WRITING & MERGING LOGIC ---
def safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

# --- 4. UI ---
st.set_page_config(layout="wide")
model_db = load_all_models()

if 'quote_models' not in st.session_state:
    st.session_state.quote_models = [{"model": ""}]

st.sidebar.title("Customer Details")
c_name = st.sidebar.text_input("Name")
c_contact = st.sidebar.text_input("Customer Contact")
c_addr = st.sidebar.text_area("Address")
c_phone = st.sidebar.text_input("Phone Number")
c_email = st.sidebar.text_input("Email")
short_code = st.sidebar.text_input("Country Code", "SA").upper()

today = datetime.date.today()
quote_id = f"{today.strftime('%Y%m%d')}-MC-{short_code}"

final_blocks = []
for i, _ in enumerate(st.session_state.quote_models):
    with st.expander(f"Product Block {i+1}", expanded=True):
        col1, col2 = st.columns([3, 1])
        options = [""] + sorted(model_db['Model'].unique().tolist()) if not model_db.empty else [""]
        selected = col1.selectbox(f"Select Model", options, key=f"m_{i}")
        
        if selected:
            match = model_db[model_db['Model'] == selected].iloc[0]
            p_cols = st.columns(3)
            p1 = p_cols[0].number_input("Price (1 pc)", key=f"p1_{i}", format="%.2f")
            p10 = p_cols[1].number_input("Price (10 pcs)", key=f"p10_{i}", format="%.2f")
            p100 = p_cols[2].number_input("Price (100 pcs)", key=f"p100_{i}", format="%.2f")
            
            final_blocks.append({
                "model": selected, "category": match['Category'], "specs": match['Specs'],
                "tiers": [(1, p1), (10, p10), (100, p100)]
            })

if st.button("➕ Add Another Model"):
    st.session_state.quote_models.append({"model": ""})
    st.rerun()

# --- 5. EXPORT ---
if st.button("🚀 Generate Quote (3-Tier View)"):
    template = 'template.xlsx'
    if os.path.exists(template) and final_blocks:
        wb = load_workbook(template)
        ws = wb.active
        
        # Fill Header Info
        safe_write(ws, 4, 9, today.strftime("%B, %dth. %Y"))
        safe_write(ws, 6, 9, quote_id)
        safe_write(ws, 10, 2, c_name)
        safe_write(ws, 11, 2, c_contact)
        safe_write(ws, 12, 2, c_addr)
        safe_write(ws, 13, 2, c_phone)
        safe_write(ws, 14, 2, c_email)
        
        current_row = 17
        for block in final_blocks:
            # 1. Handle Merges (Description, Model, Picture, Remark)
            # Merge 3 rows for Column A (Description/Category)
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row+2, end_column=1)
            safe_write(ws, current_row, 1, block['category'])
            
            # Merge 3 rows for Column D (Bewis No / Model)
            ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row+2, end_column=4)
            safe_write(ws, current_row, 4, block['model'])
            
            # Merge 3 rows for Column I (Remark / Specs)
            ws.merge_cells(start_row=current_row, start_column=9, end_row=current_row+2, end_column=9)
            safe_write(ws, current_row, 9, block['specs'])

            # Merge 3 rows for Column H (Picture)
            ws.merge_cells(start_row=current_row, start_column=8, end_row=current_row+2, end_column=8)
            
            # 2. Fetch and Add Image
            img_url = get_bwsensing_image(block['model'])
            if img_url:
                try:
                    img_data = requests.get(img_url, timeout=5).content
                    img_obj = XLImage(BytesIO(img_data))
                    # Scale image to fit the 3-row merged cell
                    img_obj.width, img_obj.height = (90, 90)
                    ws.add_image(img_obj, f'H{current_row}')
                except: pass

            # 3. Write Price Rows (Columns E, F, G are NOT merged)
            for j, (qty, price) in enumerate(block['tiers']):
                row_num = current_row + j
                safe_write(ws, row_num, 5, qty)    # Qty
                safe_write(ws, row_num, 6, price)  # Unit Price
                safe_write(ws, row_num, 7, qty * price) # Total
            
            current_row += 3 # Shift to next block
            
        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Download Official Quote", out.getvalue(), f"{quote_id}.xlsx")
