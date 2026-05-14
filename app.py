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
            
            # Clean category: Remove 'bwsensing', 'new', etc.
            raw_name = file.replace('.csv', '').replace('.xlsx', '').lower()
            category = raw_name.replace('bwsensing', '').replace('new', '').replace('model list', '').replace('-', '').strip().upper()
            
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

# --- 2. MULTI-STEP PRODUCT SCRAPER ---
def get_bws_official_image(model_name):
    """Searches and enters product page to find the main gallery image."""
    base = "https://www.bwsensing.com"
    search_url = f"{base}/search.html?q={model_name}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    try:
        # Step 1: Get search results
        search_res = requests.get(search_url, timeout=10, headers=headers)
        search_soup = BeautifulSoup(search_res.text, 'html.parser')
        
        # Step 2: Find the first product detail link
        link_tag = search_soup.select_one('.product-list a') or search_soup.select_one('.pro_list a')
        if link_tag:
            detail_path = link_tag['href']
            detail_url = base + detail_path if detail_path.startswith('/') else detail_path
            
            # Step 3: Scrape the detail page for the main image
            detail_res = requests.get(detail_url, timeout=10, headers=headers)
            detail_soup = BeautifulSoup(detail_res.text, 'html.parser')
            
            # Target the highlighted image area from your screenshot
            img_tag = detail_soup.select_one('.product-info .left-img img') or detail_soup.select_one('.detail-pic img')
            if img_tag:
                src = img_tag.get('src')
                return src if src.startswith('http') else base + src
    except: return None
    return None

# --- 3. EXCEL HELPER ---
def safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

# --- 4. APP INTERFACE ---
st.set_page_config(layout="wide")
model_db = load_all_models()

if 'quote_blocks' not in st.session_state:
    st.session_state.quote_blocks = [{"model": ""}]

with st.sidebar:
    st.header("Quote Details")
    c_name = st.text_input("Name")
    c_contact = st.text_input("Customer Contact")
    c_addr = st.text_area("Address")
    c_phone = st.text_input("Phone Number")
    c_email = st.text_input("Email")
    short_code = st.text_input("Country Code (e.g. SA)", "SA").upper()

today = datetime.date.today()
valid_until = today + datetime.timedelta(days=30)
quote_id = f"BW-{today.strftime('%Y%m%d')}-MC-{short_code}"

st.title(f"Official Quote: {quote_id}")

final_blocks = []
for i, _ in enumerate(st.session_state.quote_blocks):
    with st.expander(f"Product {i+1}", expanded=True):
        opts = [""] + sorted(model_db['Model'].unique().tolist()) if not model_db.empty else [""]
        sel = st.selectbox(f"Select Model", opts, key=f"sel_{i}")
        if sel:
            m = model_db[model_db['Model'] == sel].iloc[0]
            p_cols = st.columns(3)
            p1 = p_cols[0].number_input("1 pc Price", key=f"p1_{i}", value=0.0)
            p10 = p_cols[1].number_input("10 pc Price", key=f"p10_{i}", value=0.0)
            p100 = p_cols[2].number_input("100 pc Price", key=f"p100_{i}", value=0.0)
            final_blocks.append({"model": sel, "cat": m['Category'], "specs": m['Specs'], "tiers": [(1, p1), (10, p10), (100, p100)]})

if st.button("➕ Add Another Product"):
    st.session_state.quote_blocks.append({"model": ""})
    st.rerun()

# --- 5. EXCEL GENERATION ---
if st.button("🚀 Create Quote"):
    if os.path.exists('template.xlsx') and final_blocks:
        wb = load_workbook('template.xlsx')
        ws = wb.active
        
        # Header & Dates
        safe_write(ws, 4, 9, today.strftime("%B %d, %Y"))
        safe_write(ws, 5, 9, valid_until.strftime("%B %d, %Y")) # Auto-calc 1 month
        safe_write(ws, 6, 9, quote_id)
        
        # Customer Info
        for idx, val in enumerate([c_name, c_contact, c_addr, c_phone, c_email], 10):
            safe_write(ws, idx, 2, val)
            
        curr_row = 17
        for block in final_blocks:
            # Merging 3 rows for fixed columns
            for col in [1, 4, 8, 9]: # Description, Model, Picture, Remark
                ws.merge_cells(start_row=curr_row, start_column=col, end_row=curr_row+2, end_column=col)
            
            safe_write(ws, curr_row, 1, block['cat'])
            safe_write(ws, curr_row, 4, block['model'])
            safe_write(ws, curr_row, 9, block['specs'])
            
            # Image Scraping
            img_url = get_bws_official_image(block['model'])
            if img_url:
                try:
                    img_data = requests.get(img_url, timeout=10).content
                    img_file = XLImage(BytesIO(img_data))
                    img_file.width, img_file.height = (80, 80)
                    ws.add_image(img_file, f'H{curr_row}')
                except: pass

            # Pricing Tiers (Rows 1, 2, 3 of the block)
            for j, (qty, prc) in enumerate(block['tiers']):
                r = curr_row + j
                safe_write(ws, r, 5, qty)
                safe_write(ws, r, 6, prc) # No dollar sign
                safe_write(ws, r, 7, qty * prc)
            
            curr_row += 3
            
        buf = BytesIO()
        wb.save(buf)
        st.download_button("📥 Download Excel", buf.getvalue(), f"{quote_id}.xlsx")
