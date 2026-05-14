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
                    all_data.append({"Model": m_name, "Specs": "\n".join(specs)})
        except: continue
    return pd.DataFrame(all_data)

# --- 2. THE IMAGE SCRAPER (Strict Path Version) ---
def get_bws_direct_image(model_name):
    base = "https://www.bwsensing.com"
    search_url = f"{base}/search.html?q={model_name}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        # Step 1: Find product link
        res = requests.get(search_url, timeout=10, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        link = soup.select_one('.product-list a')
        if not link: return None
        
        # Step 2: Extract image from detail page
        detail_url = base + link['href'] if link['href'].startswith('/') else link['href']
        detail_res = requests.get(detail_url, timeout=10, headers=headers)
        dsoup = BeautifulSoup(detail_res.text, 'html.parser')
        
        # Look for the specific 'left-img' class from your screenshot
        img_tag = dsoup.select_one('.left-img img') or dsoup.select_one('.product-info img')
        if img_tag:
            # Bypass lazy loading by checking data-original first
            src = img_tag.get('data-original') or img_tag.get('src')
            if src:
                return src if src.startswith('http') else base + src
    except: return None
    return None

# --- 3. EXCEL HELPERS ---
def safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

# --- 4. STREAMLIT UI ---
st.set_page_config(layout="wide")
model_db = load_all_models()

if 'quote_list' not in st.session_state:
    st.session_state.quote_list = [{"model": ""}]

with st.sidebar:
    st.title("Quotation Config")
    rate = st.number_input("RMB to USD Rate", value=7.2)
    st.divider()
    c_name = st.text_input("Customer Name")
    c_contact = st.text_input("Customer Contact")
    c_addr = st.text_area("Address")
    c_phone = st.text_input("Phone Number")
    c_email = st.text_input("Email")
    country_code = st.text_input("Country Code", "SA").upper()

today = datetime.date.today()
expiry = today + datetime.timedelta(days=30)
quote_id = f"BW-{today.strftime('%Y%m%d')}-MC-{country_code}"

st.title(f"Generate Quote: {quote_id}")

final_selections = []
for i, _ in enumerate(st.session_state.quote_list):
    with st.expander(f"Product Block {i+1}", expanded=True):
        opts = [""] + sorted(model_db['Model'].unique().tolist())
        sel_model = st.selectbox(f"Search Model", opts, key=f"m_sel_{i}")
        if sel_model:
            match = model_db[model_db['Model'] == sel_model].iloc[0]
            col_p = st.columns(3)
            r1 = col_p[0].number_input("RMB Price (1pc)", key=f"r1_{i}")
            r10 = col_p[1].number_input("RMB Price (10pcs)", key=f"r10_{i}")
            r100 = col_p[2].number_input("RMB Price (100pcs)", key=f"r100_{i}")
            
            final_selections.append({
                "model": sel_model,
                "specs": match['Specs'],
                "tiers": [round(r1/rate, 2), round(r10/rate, 2), round(r100/rate, 2)]
            })

if st.button("➕ Add More Products"):
    st.session_state.quote_list.append({"model": ""})
    st.rerun()

# --- 5. GENERATE EXCEL ---
if st.button("🚀 Export to Excel"):
    if os.path.exists('template.xlsx') and final_selections:
        wb = load_workbook('template.xlsx')
        ws = wb.active
        
        # Header/Date Info
        safe_write(ws, 4, 9, today.strftime("%B %d, %Y"))
        safe_write(ws, 5, 9, expiry.strftime("%B %d, %Y"))
        safe_write(ws, 6, 9, quote_id)
        
        # Customer Info
        for idx, info in enumerate([c_name, c_contact, c_addr, c_phone, c_email], 10):
            safe_write(ws, idx, 2, info)
            
        row_cursor = 17
        for block in final_selections:
            # 1. Merge and Clear Columns
            # Column 1 (Description) -> Keeping it EMPTY as requested
            ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor+2, end_column=1)
            safe_write(ws, row_cursor, 1, "") 
            
            # Column 4 (Bewis No)
            ws.merge_cells(start_row=row_cursor, start_column=4, end_row=row_cursor+2, end_column=4)
            safe_write(ws, row_cursor, 4, block['model'])
            
            # Column 9 (Remark)
            ws.merge_cells(start_row=row_cursor, start_column=9, end_row=row_cursor+2, end_column=9)
            safe_write(ws, row_cursor, 9, block['specs'])

            # 2. Pricing Tiers (Rows 1-3)
            qty_list = [1, 10, 100]
            for j in range(3):
                curr = row_cursor + j
                safe_write(ws, curr, 5, qty_list[j])
                safe_write(ws, curr, 6, block['tiers'][j])
                # REMOVED safe_write for Column 7 (Line Total) to let template formula work
            
            # 3. Image Handling
            img_url = get_bws_direct_image(block['model'])
            if img_url:
                try:
                    img_data = requests.get(img_url, timeout=10).content
                    xl_img = XLImage(BytesIO(img_data))
                    xl_img.width, xl_img.height = (90, 90)
                    ws.add_image(xl_img, f'H{row_cursor}')
                except: pass
            
            row_cursor += 3
            
        output = BytesIO()
        wb.save(output)
        st.download_button("📥 Click to Download Quote", output.getvalue(), f"{quote_id}.xlsx")
