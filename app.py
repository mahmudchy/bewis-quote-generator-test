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
from fpdf import FPDF

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

# --- 2. NEW SCRAPER (bw-sensing.com) ---
def get_bw_sensing_image(model_name):
    base = "https://www.bw-sensing.com"
    # Search URL for the new site structure
    search_url = f"{base}/search.html?q={model_name}"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(search_url, timeout=10, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        # Target the first product item link
        link = soup.select_one('.product-list a') or soup.select_one('.pro_list a')
        if not link: return None
        
        detail_url = base + link['href'] if link['href'].startswith('/') else link['href']
        d_res = requests.get(detail_url, timeout=10, headers=headers)
        dsoup = BeautifulSoup(d_res.text, 'html.parser')
        
        # Target the main product gallery image
        img_tag = dsoup.select_one('.product-info img') or dsoup.select_one('.left-img img')
        if img_tag:
            src = img_tag.get('data-original') or img_tag.get('src')
            if src: return src if src.startswith('http') else base + src
    except: return None
    return None

# --- 3. UI SETUP ---
st.set_page_config(layout="wide", page_title="BWS Quote Generator")
model_db = load_all_models()

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

with st.sidebar:
    st.title("Settings")
    exch_rate = st.number_input("RMB to USD Rate", value=6.82, step=0.01) # Default 6.82
    st.divider()
    c_name = st.text_input("Name")
    c_contact = st.text_input("Contact")
    c_addr = st.text_area("Address")
    c_phone = st.text_input("Phone")
    c_email = st.text_input("Email")
    code = st.text_input("Country Code", "SA").upper()

today = datetime.date.today()
expiry = today + datetime.timedelta(days=30)
quote_id = f"BW-{today.strftime('%Y%m%d')}-MC-{code}"

st.title(f"Quote Generator: {quote_id}")

final_data = []
for i, _ in enumerate(st.session_state.rows):
    with st.expander(f"Product {i+1}", expanded=True):
        opts = [""] + sorted(model_db['Model'].unique().tolist())
        sel = st.selectbox("Model", opts, key=f"sel_{i}")
        if sel:
            m = model_db[model_db['Model'] == sel].iloc[0]
            p_cols = st.columns(3)
            # Rounded RMB inputs
            r1 = p_cols[0].number_input("RMB (1pc)", key=f"r1_{i}", step=1, format="%d")
            r10 = p_cols[1].number_input("RMB (10pcs)", key=f"r10_{i}", step=1, format="%d")
            r100 = p_cols[2].number_input("RMB (100pcs)", key=f"r100_{i}", step=1, format="%d")
            
            final_data.append({
                "model": sel, "specs": m['Specs'],
                "tiers": [
                    {"qty": 1, "usd": r1/exch_rate},
                    {"qty": 10, "usd": r10/exch_rate},
                    {"qty": 100, "usd": r100/exch_rate}
                ]
            })

if st.button("➕ Add Product"):
    st.session_state.rows.append({"model": ""})
    st.rerun()

# --- 4. PREVIEW TABLE ---
if final_data:
    st.subheader("Live Quote Preview")
    preview_list = []
    for f in final_data:
        for t in f['tiers']:
            preview_list.append({
                "Model": f['model'],
                "Qty": t['qty'],
                "Unit Price (USD)": round(t['usd'], 2),
                "Total (USD)": round(t['qty'] * t['usd'], 2)
            })
    st.table(pd.DataFrame(preview_list))

# --- 5. EXPORT BUTTONS ---
col_ex1, col_ex2 = st.columns(2)

if col_ex1.button("🚀 Generate Excel"):
    if os.path.exists('template.xlsx'):
        wb = load_workbook('template.xlsx')
        ws = wb.active
        # Header/Dates
        ws.cell(4, 9).value = today.strftime("%B %d, %Y")
        ws.cell(5, 9).value = expiry.strftime("%B %d, %Y")
        ws.cell(6, 9).value = quote_id
        # Contact
        for idx, val in enumerate([c_name, c_contact, c_addr, c_phone, c_email], 10):
            ws.cell(idx, 2).value = val
            
        cur_row = 17
        for block in final_data:
            # Merge fixed columns
            for col in [1, 4, 8, 9]:
                ws.merge_cells(start_row=cur_row, start_column=col, end_row=cur_row+2, end_column=col)
            
            ws.cell(cur_row, 4).value = block['model']
            ws.cell(cur_row, 9).value = block['specs']
            
            # Pricing (No formula override for Column 7)
            for j, t in enumerate(block['tiers']):
                ws.cell(cur_row+j, 5).value = t['qty']
                ws.cell(cur_row+j, 6).value = t['usd']
            
            # Image
            img_url = get_bw_sensing_image(block['model'])
            if img_url:
                try:
                    res = requests.get(img_url, timeout=5)
                    img = XLImage(BytesIO(res.content))
                    img.width, img.height = (90, 90)
                    ws.add_image(img, f'H{cur_row}')
                except: pass
            cur_row += 3
            
        out_b = BytesIO()
        wb.save(out_b)
        st.download_button("📥 Download
