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

# --- 1. DATA LOADING & CLEANING ---
@st.cache_data
def load_all_models():
    all_data = []
    files = [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx'))]
    for file in files:
        try:
            if any(x in file.lower() for x in ["template", "requirements"]): continue
            df_raw = pd.read_csv(file, header=None).fillna('') if file.endswith('.csv') else pd.read_excel(file, header=None).fillna('')
            
            # Clean category: Logic to avoid "ALL" or "BWSENSING NEW"
            raw_cat = file.replace('.csv', '').replace('.xlsx', '').upper()
            # Split by common delimiters and remove noise
            parts = [p for p in raw_cat.split('-') if p not in ['ALL', 'NEW', 'BWSENSING', 'LIST', 'MODEL']]
            category = " ".join(parts).strip() if parts else "INCLINOMETER"
            
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
                    if not m_name or m_name.lower() in ['model', 'nan', 'all']: continue
                    specs = []
                    for col in df.columns:
                        if any(k in col.lower() for k in ['accuracy', 'range', 'axis', 'output']):
                            val = str(row[col]).strip()
                            if val and val.lower() != 'nan': specs.append(f"{col}: {val}")
                    all_data.append({"Model": m_name, "Category": category, "Specs": "\n".join(specs)})
        except: continue
    return pd.DataFrame(all_data)

# --- 2. THE IMAGE SCRAPER (Fixed for Lazy Loading) ---
def get_bws_product_image(model_name):
    base_url = "https://www.bwsensing.com"
    search_url = f"{base_url}/search.html?q={model_name}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        # 1. Search
        res = requests.get(search_url, timeout=10, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        link_tag = soup.select_one('.product-list a')
        if not link_tag: return None
        
        # 2. Detail Page
        detail_url = base_url + link_tag['href'] if link_tag['href'].startswith('/') else link_tag['href']
        detail_res = requests.get(detail_url, timeout=10, headers=headers)
        dsoup = BeautifulSoup(detail_res.text, 'html.parser')
        
        # 3. Find Image (checking data-original for lazy load)
        img_tag = dsoup.select_one('.product-info .left-img img')
        if img_tag:
            src = img_tag.get('data-original') or img_tag.get('src')
            if src:
                return src if src.startswith('http') else base_url + src
    except: return None
    return None

# --- 3. WRITING LOGIC ---
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

with st.sidebar:
    st.title("Settings")
    exch_rate = st.number_input("RMB to USD Exchange Rate", value=7.2, step=0.01)
    st.divider()
    c_name = st.text_input("Name")
    c_contact = st.text_input("Customer Contact")
    c_addr = st.text_area("Address")
    c_phone = st.text_input("Phone Number")
    c_email = st.text_input("Email")
    short_code = st.text_input("Country Short Code", "SA").upper()

today = datetime.date.today()
valid_until = today + datetime.timedelta(days=30)
quote_id = f"BW-{today.strftime('%Y%m%d')}-MC-{short_code}"

st.title(f"Quote: {quote_id}")

final_blocks = []
for i, _ in enumerate(st.session_state.quote_models):
    with st.expander(f"Product Block {i+1}", expanded=True):
        opts = [""] + sorted(model_db['Model'].unique().tolist())
        selected = st.selectbox(f"Select Model", opts, key=f"m_{i}")
        if selected:
            m = model_db[model_db['Model'] == selected].iloc[0]
            p_cols = st.columns(3)
            # Input is in RMB
            rmb1 = p_cols[0].number_input("Price 1pc (RMB)", key=f"r1_{i}")
            rmb10 = p_cols[1].number_input("Price 10pcs (RMB)", key=f"r10_{i}")
            rmb100 = p_cols[2].number_input("Price 100pcs (RMB)", key=f"r100_{i}")
            
            final_blocks.append({
                "model": selected, "cat": m['Category'], "specs": m['Specs'],
                "tiers": [
                    (1, round(rmb1 / exch_rate, 2)),
                    (10, round(rmb10 / exch_rate, 2)),
                    (100, round(rmb100 / exch_rate, 2))
                ]
            })

if st.button("➕ Add Another Model"):
    st.session_state.quote_models.append({"model": ""})
    st.rerun()

# --- 5. EXPORT ---
if st.button("🚀 Generate Final Excel"):
    template = 'template.xlsx'
    if os.path.exists(template) and final_blocks:
        wb = load_workbook(template)
        ws = wb.active
        
        # Header Info
        safe_write(ws, 4, 9, today.strftime("%B, %dth. %Y"))
        safe_write(ws, 5, 9, valid_until.strftime("%B, %dth. %Y"))
        safe_write(ws, 6, 9, quote_id)
        
        # Contact Rows
        for idx, val in enumerate([c_name, c_contact, c_addr, c_phone, c_email], 10):
            safe_write(ws, idx, 2, val)
        
        current_row = 17
        for block in final_blocks:
            # 1. Merge and Write Fixed Data
            for col in [1, 4, 8, 9]:
                ws.merge_cells(start_row=current_row, start_column=col, end_row=current_row+2, end_column=col)
            
            safe_write(ws, current_row, 1, block['cat'])
            safe_write(ws, current_row, 4, block['model'])
            safe_write(ws, current_row, 9, block['specs'])
            
            # 2. Pricing (Numeric Only, No $)
            for j, (qty, usd_price) in enumerate(block['tiers']):
                r_num = current_row + j
                safe_write(ws, r_num, 5, qty)
                safe_write(ws, r_num, 6, usd_price)
                safe_write(ws, r_num, 7, round(qty * usd_price, 2))
            
            # 3. Image Fetch
            img_url = get_bws_product_image(block['model'])
            if img_url:
                try:
                    img_data = requests.get(img_url, timeout=10).content
                    img_obj = XLImage(BytesIO(img_data))
                    img_obj.width, img_obj.height = (90, 90)
                    ws.add_image(img_obj, f'H{current_row}')
                except: pass
                
            current_row += 3
            
        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Download Official Quote", out.getvalue(), f"{quote_id}.xlsx")
