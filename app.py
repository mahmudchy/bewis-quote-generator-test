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

# --- 2. IMAGE SCRAPER ---
def get_bw_sensing_image(model_name):
    base = "https://www.bw-sensing.com"
    search_url = f"{base}/search.html?q={model_name}"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(search_url, timeout=7, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        link = soup.select_one('.product-list a') or soup.select_one('.pro_list a')
        if not link: return None
        detail_url = base + link['href'] if link['href'].startswith('/') else link['href']
        d_res = requests.get(detail_url, timeout=7, headers=headers)
        dsoup = BeautifulSoup(d_res.text, 'html.parser')
        img_tag = dsoup.select_one('.product-info img') or dsoup.select_one('.left-img img')
        if img_tag:
            src = img_tag.get('data-original') or img_tag.get('src')
            if src: return src if src.startswith('http') else base + src
    except: return None
    return None

# --- 3. THE STYLE ENGINE ---
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'), 
    top=Side(style='thin'), bottom=Side(style='thin')
)

def apply_block_styles(ws, start_row):
    """Applies borders and sets basic row heights for the 3-row product block."""
    for r in range(start_row, start_row + 3):
        ws.row_dimensions[r].height = 25 
        for c in range(1, 10): # Columns A to I
            ws.cell(row=r, column=c).border = thin_border

# --- 4. UI SETUP ---
st.set_page_config(layout="wide", page_title="BWS Quote Generator")
model_db = load_all_models()

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

with st.sidebar:
    st.title("Settings")
    exch_rate = st.number_input("RMB to USD Rate", value=6.82, step=0.01)
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

st.title(f"Quote: {quote_id}")

final_data = []
for i, _ in enumerate(st.session_state.rows):
    with st.expander(f"Product {i+1}", expanded=True):
        opts = [""] + sorted(model_db['Model'].unique().tolist())
        sel = st.selectbox("Search & Select Model", opts, key=f"sel_{i}")
        if sel:
            m = model_db[model_db['Model'] == sel].iloc[0]
            p_cols = st.columns(3)
            r1 = st.text_input("RMB (1pc)", "0", key=f"r1_{i}")
            r10 = st.text_input("RMB (10pcs)", "0", key=f"r10_{i}")
            r100 = st.text_input("RMB (100pcs)", "0", key=f"r100_{i}")
            
            final_data.append({
                "model": sel, "specs": m['Specs'],
                "tiers": [
                    {"qty": 1, "rmb": float(r1.replace(',', '') or 0)},
                    {"qty": 10, "rmb": float(r10.replace(',', '') or 0)},
                    {"qty": 100, "rmb": float(r100.replace(',', '') or 0)}
                ]
            })

if st.button("➕ Add Another Product Line"):
    st.session_state.rows.append({"model": ""})
    st.rerun()

# --- 5. EXPORT ---
if st.button("🚀 Export to Excel"):
    if os.path.exists('template.xlsx') and final_data:
        wb = load_workbook('template.xlsx')
        ws = wb.active
        
        # 1. Update Headers & Customer Info
        ws['I4'] = today.strftime("%B %d, %Y")
        ws['I5'] = expiry.strftime("%B %d, %Y")
        ws['I6'] = quote_id
        ws['B10'], ws['B11'], ws['B12'] = c_name, c_contact, c_addr
        ws['B13'], ws['B14'] = c_phone, c_email
            
        start_row = 17
        
        # 2. CLEAR TEMPLATE DUMMY DATA & INSERT SPACE
        # We find where the footer starts to avoid overwriting it
        footer_row = 20
        for r in range(17, 50):
            if "Subtotal" in str(ws.cell(row=r, column=6).value) or "Total" in str(ws.cell(row=r, column=6).value):
                footer_row = r
                break
        
        # Insert exactly (4 rows per product)
        # 3 rows for data + 1 row for gap
        ws.insert_rows(start_row, len(final_data) * 4)

        for idx, block in enumerate(final_data):
            cur_top = start_row + (idx * 4) 
            
            # 3. APPLY MERGING (Description, Model, Picture, Remark)
            ws.merge_cells(start_row=cur_top, start_column=1, end_row=cur_top+2, end_column=1)
            ws.merge_cells(start_row=cur_top, start_column=4, end_row=cur_top+2, end_column=4)
            ws.merge_cells(start_row=cur_top, start_column=8, end_row=cur_top+2, end_column=8)
            ws.merge_cells(start_row=cur_top, start_column=9, end_row=cur_top+2, end_column=9)
            
            apply_block_styles(ws, cur_top)
            
            # 4. WRITE DATA
            ws.cell(row=cur_top, column=1).value = "Inclinometer" # or block['model']
            ws.cell(row=cur_top, column=4).value = block['model']
            ws.cell(row=cur_top, column=9).value = block['specs']
            
            for c_idx in [1, 4, 9]:
                ws.cell(row=cur_top, column=c_idx).alignment = Alignment(vertical='center', horizontal='center', wrapText=True)
            
            # 5. WRITE PRICE TIERS
            for j, t in enumerate(block['tiers']):
                r_idx = cur_top + j
                ws.cell(row=r_idx, column=5).value = t['qty']
                if t['rmb'] > 0:
                    u_usd = round(t['rmb'] / exch_rate, 2)
                    ws.cell(row=r_idx, column=6).value = u_usd
                    ws.cell(row=r_idx, column=7).value = round(u_usd * t['qty'], 2)

            # 6. PRODUCT IMAGE
            img_url = get_bw_sensing_image(block['model'])
            if img_url:
                try:
                    res = requests.get(img_url, timeout=5)
                    img = XLImage(BytesIO(res.content))
                    img.width, img.height = (85, 85)
                    ws.add_image(img, f'H{cur_top}')
                except: pass
            
        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Download Quote", out.getvalue(), f"{quote_id}.xlsx")
