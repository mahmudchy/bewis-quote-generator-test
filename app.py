import streamlit as st
import pandas as pd
import os
import datetime
import requests
from bs4 import BeautifulSoup
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Side

# --- 1. DATA LOADING (Standard) ---
@st.cache_data
def load_all_models():
    all_data = []
    files = [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx'))]
    for file in files:
        try:
            if any(x in file.lower() for x in ["template", "requirements"]): continue
            df_raw = pd.read_csv(file, header=None).fillna('')
            model_col_idx = -1
            header_row = 0
            for r_idx in range(min(len(df_raw), 25)):
                row_vals = [str(v).strip().lower() for v in df_raw.iloc[r_idx]]
                if 'model' in row_vals or 'bewis no' in row_vals:
                    model_col_idx = row_vals.index('model') if 'model' in row_vals else row_vals.index('bewis no')
                    header_row = r_idx
                    break
            if model_col_idx != -1:
                df = pd.read_excel(file, header=header_row).fillna('') if file.endswith('.xlsx') else pd.read_csv(file, header=header_row).fillna('')
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

# --- 2. IMAGE SCRAPER ---
def get_bw_sensing_image(model_name):
    base = "https://www.bw-sensing.com"
    search_url = f"{base}/search.html?q={model_name}"
    try:
        res = requests.get(search_url, timeout=5)
        soup = BeautifulSoup(res.text, 'html.parser')
        link = soup.select_one('.product-list a') or soup.select_one('.pro_list a')
        if not link: return None
        detail_url = base + link['href'] if link['href'].startswith('/') else link['href']
        d_res = requests.get(detail_url, timeout=5)
        dsoup = BeautifulSoup(d_res.text, 'html.parser')
        img_tag = dsoup.select_one('.product-info img') or dsoup.select_one('.left-img img')
        if img_tag:
            src = img_tag.get('data-original') or img_tag.get('src')
            return src if src.startswith('http') else base + src
    except: return None

# --- 3. UI SETUP ---
st.set_page_config(layout="wide", page_title="BWS Quote Gen")
model_db = load_all_models()

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

with st.sidebar:
    st.title("Quote Settings")
    exch_rate = st.number_input("RMB to USD Rate", value=7.24, step=0.01)
    c_name = st.text_input("Customer Name")
    c_contact = st.text_input("Customer Contact")
    c_addr = st.text_area("Address")
    c_phone = st.text_input("Phone Number")
    c_email = st.text_input("Email")
    country_code = st.text_input("Country Code", "SA").upper()

today = datetime.date.today()
expiry = today + datetime.timedelta(days=30)
quote_id = f"BW-{today.strftime('%Y%m%d')}-MC-{country_code}"

st.title(f"Quote Generator: {quote_id}")

final_data = []
for i, _ in enumerate(st.session_state.rows):
    with st.expander(f"Product {i+1}", expanded=True):
        opts = [""] + sorted(model_db['Model'].unique().tolist())
        sel = st.selectbox("Select Model", opts, key=f"sel_{i}")
        if sel:
            m = model_db[model_db['Model'] == sel].iloc[0]
            p_cols = st.columns(3)
            r1 = p_cols[0].text_input("RMB (1pc)", "0", key=f"r1_{i}")
            r10 = p_cols[1].text_input("RMB (10pcs)", "0", key=f"r10_{i}")
            r100 = p_cols[2].text_input("RMB (100pcs)", "0", key=f"r100_{i}")
            final_data.append({
                "model": sel, "specs": m['Specs'],
                "tiers": [{"qty": 1, "rmb": float(r1 or 0)}, {"qty": 10, "rmb": float(r10 or 0)}, {"qty": 100, "rmb": float(r100 or 0)}]
            })

if st.button("➕ Add Another Model"):
    st.session_state.rows.append({"model": ""})
    st.rerun()

# --- 4. EXPORT LOGIC ---
if st.button("🚀 Export to Excel"):
    if os.path.exists('template.xlsx') and final_data:
        wb = load_workbook('template.xlsx')
        ws = wb.active
        
        # Header Info
        ws['I4'], ws['I5'], ws['I6'] = today.strftime("%B %d, %Y"), expiry.strftime("%B %d, %Y"), quote_id
        ws['B10'], ws['B11'], ws['B12'] = c_name, c_contact, c_addr
        ws['B13'], ws['B14'] = c_phone, c_email

        # Find the "Remarks" row to push it down
        footer_anchor = 17
        for r in range(1, 100):
            if str(ws.cell(row=r, column=1).value).strip() == "Remarks":
                footer_anchor = r
                break
        
        # Insert exactly enough rows (3 per product + 1 spacer row)
        ws.insert_rows(footer_anchor, len(final_data) * 4)

        thin = Side(style='thin')
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

        current_row = footer_anchor
        for product in final_data:
            # A. MERGE CELLS MANUALLY FOR NEW ROWS
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row+2, end_column=1) # Description
            ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row+2, end_column=4) # Model
            ws.merge_cells(start_row=current_row, start_column=8, end_row=current_row+2, end_column=8) # Picture
            ws.merge_cells(start_row=current_row, start_column=9, end_row=current_row+2, end_column=9) # Specs

            # B. WRITE CONTENT
            ws.cell(row=current_row, column=1).value = "Inclinometer"
            ws.cell(row=current_row, column=4).value = product['model']
            ws.cell(row=current_row, column=9).value = product['specs']

            # C. WRITE PRICING & BORDERS
            for i, tier in enumerate(product['tiers']):
                r_idx = current_row + i
                ws.cell(row=r_idx, column=5).value = tier['qty']
                if tier['rmb'] > 0:
                    usd_val = round(tier['rmb'] / exch_rate, 2)
                    ws.cell(row=r_idx, column=6).value = usd_val
                    ws.cell(row=r_idx, column=7).value = usd_val * tier['qty']
                
                # Apply borders to the full block (Cols A to I)
                for c in range(1, 10):
                    ws.cell(row=r_idx, column=c).border = border
                    ws.cell(row=r_idx, column=c).alignment = center_align

            # D. IMAGE
            img_url = get_bw_sensing_image(product['model'])
            if img_url:
                try:
                    res = requests.get(img_url, timeout=5)
                    img = XLImage(BytesIO(res.content))
                    img.width, img.height = (80, 80)
                    ws.add_image(img, f'H{current_row}')
                except: pass

            current_row += 4 # Skip 3 rows + 1 spacer

        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Download Quote", out.getvalue(), f"{quote_id}.xlsx")
