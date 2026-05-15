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

# --- 1. DATA LOADING ---
@st.cache_data
def load_all_models():
    all_data = []
    files = [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx'))]
    for file in files:
        try:
            if any(x in file.lower() for x in ["template", "requirements"]): continue
            df = pd.read_excel(file).fillna('') if file.endswith('.xlsx') else pd.read_csv(file).fillna('')
            df.columns = [str(c).strip().lower() for c in df.columns]
            model_col = next((c for c in df.columns if any(k in c for k in ['model', 'bewis', 'part'])), None)
            if model_col:
                for _, row in df.iterrows():
                    m_name = str(row[model_col]).strip()
                    if not m_name or m_name.lower() in ['model', 'nan', '']: continue
                    specs = [f"{c.title()}: {row[c]}" for c in df.columns if any(k in c for k in ['accuracy', 'range', 'axis', 'output']) if str(row[c]).strip()]
                    all_data.append({"Model": m_name, "Specs": "\n".join(specs)})
        except: continue
    return pd.DataFrame(all_data)

# --- 2. IMAGE SCRAPER ---
def get_bw_sensing_image(model_name):
    base = "https://www.bw-sensing.com"
    try:
        res = requests.get(f"{base}/search.html?q={model_name}", timeout=5)
        soup = BeautifulSoup(res.text, 'html.parser')
        link = soup.select_one('.product-list a') or soup.select_one('.pro_list a')
        if link:
            d_res = requests.get(base + link['href'] if link['href'].startswith('/') else link['href'], timeout=5)
            dsoup = BeautifulSoup(d_res.text, 'html.parser')
            img = dsoup.select_one('.product-info img') or dsoup.select_one('.left-img img')
            if img:
                src = img.get('data-original') or img.get('src')
                return src if src.startswith('http') else base + src
    except: return None

# --- 3. UI SETUP ---
st.set_page_config(layout="wide", page_title="BWS Quote Gen")
model_db = load_all_models()

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

with st.sidebar:
    st.title("Quote Settings")
    exch_rate = st.number_input("Exchange Rate (RMB/USD)", value=7.24)
    c_name = st.text_input("Customer Name")
    c_contact = st.text_input("Contact Person")
    c_addr = st.text_area("Address")
    c_phone = st.text_input("Phone")
    c_email = st.text_input("Email")
    country = st.text_input("Country Code", "SA")

today = datetime.date.today()
quote_id = f"BW-{today.strftime('%Y%m%d')}-MC-{country.upper()}"

st.title(f"Quote: {quote_id}")

final_data = []
for i, _ in enumerate(st.session_state.rows):
    with st.expander(f"Model {i+1}", expanded=True):
        opts = [""] + sorted(model_db['Model'].unique().tolist()) if not model_db.empty else [""]
        sel = st.selectbox("Select Model", opts, key=f"sel_{i}")
        if sel:
            m = model_db[model_db['Model'] == sel].iloc[0]
            cols = st.columns(3)
            r1 = cols[0].text_input("Price (1pc)", "0", key=f"p1_{i}")
            r10 = cols[1].text_input("Price (10pcs)", "0", key=f"p10_{i}")
            r100 = cols[2].text_input("Price (100pcs)", "0", key=f"p100_{i}")
            final_data.append({
                "model": sel, "specs": m['Specs'],
                "tiers": [{"qty": 1, "rmb": float(r1 or 0)}, {"qty": 10, "rmb": float(r10 or 0)}, {"qty": 100, "rmb": float(r100 or 0)}]
            })

if st.button("➕ Add Another Model"):
    st.session_state.rows.append({"model": ""})
    st.rerun()

# --- 4. EXPORT ENGINE ---
if st.button("🚀 Export to Excel"):
    if os.path.exists('template.xlsx') and final_data:
        wb = load_workbook('template.xlsx')
        ws = wb.active
        
        # 1. Update Header Info
        ws['I4'], ws['I6'] = today.strftime("%B %d, %Y"), quote_id
        ws['B10'], ws['B11'], ws['B12'] = c_name, c_contact, c_addr
        ws['B13'], ws['B14'] = c_phone, c_email

        # Formatting Constants
        thin = Side(style='thin')
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        center = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # START POSITION
        # Model 1 uses 17, 18, 19. 
        # Row 20 is the empty spacer before Remarks.
        
        for idx, product in enumerate(final_data):
            if idx == 0:
                # First model goes into the existing template rows
                write_pos = 17
            else:
                # For every additional model, insert 3 rows ABOVE the footer (currently at row 20)
                # This pushes everything from row 20 downwards
                write_pos = 17 + (idx * 3)
                ws.insert_rows(write_pos, 3)

            # A. Apply Merges
            ws.merge_cells(start_row=write_pos, start_column=1, end_row=write_pos+2, end_column=1)
            ws.merge_cells(start_row=write_pos, start_column=4, end_row=write_pos+2, end_column=4)
            ws.merge_cells(start_row=write_pos, start_column=8, end_row=write_pos+2, end_column=8)
            ws.merge_cells(start_row=write_pos, start_column=9, end_row=write_pos+2, end_column=9)

            # B. Add Values
            ws.cell(row=write_pos, column=1).value = "Inclinometer"
            ws.cell(row=write_pos, column=4).value = product['model']
            ws.cell(row=write_pos, column=9).value = product['specs']

            # C. Pricing & Borders
            for sub_r in range(3):
                row_idx = write_pos + sub_r
                tier = product['tiers'][sub_r]
                
                ws.cell(row=row_idx, column=5).value = tier['qty']
                if tier['rmb'] > 0:
                    u_usd = round(tier['rmb'] / exch_rate, 2)
                    ws.cell(row=row_idx, column=6).value = u_usd
                    ws.cell(row=row_idx, column=7).value = round(u_usd * tier['qty'], 2)
                
                # Apply Borders and Alignment to the whole row
                for col_idx in range(1, 10):
                    ws.cell(row=row_idx, column=col_idx).border = border
                    ws.cell(row=row_idx, column=col_idx).alignment = center

            # D. Image
            img_url = get_bw_sensing_image(product['model'])
            if img_url:
                try:
                    res = requests.get(img_url, timeout=5)
                    img = XLImage(BytesIO(res.content))
                    img.width, img.height = (80, 80)
                    ws.add_image(img, f'H{write_pos}')
                except: pass

        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Download Quote", out.getvalue(), f"{quote_id}.xlsx")
