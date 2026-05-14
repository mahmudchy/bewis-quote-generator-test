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
    # Identify local data files
    files = [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx'))]
    for file in files:
        try:
            if any(x in file.lower() for x in ["template", "requirements"]): continue
            df_raw = pd.read_csv(file, header=None).fillna('') if file.endswith('.csv') else pd.read_excel(file, header=None).fillna('')
            
            model_col_idx = -1
            header_row = 0
            # Find the header row by looking for 'model' or 'bewis no'
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

# --- 3. EXCEL MERGE HANDLING ---
def safe_write(ws, row, col, value):
    """Writes to a cell, specifically targeting the top-left of a merged range if applicable."""
    cell = ws.cell(row=row, column=col)
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
            return
    cell.value = value

# --- 4. UI SETUP ---
st.set_page_config(layout="wide", page_title="BWS Quote Gen")

# Custom CSS for clean numeric entry
st.markdown("""
    <style>
    input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { 
        -webkit-appearance: none; margin: 0; 
    }
    input[type=number] { -moz-appearance: textfield; }
    input::-webkit-clear-button, input::-webkit-search-cancel-button { display: none; -webkit-appearance: none; }
    </style>
""", unsafe_allow_html=True)

model_db = load_all_models()

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

with st.sidebar:
    st.title("Control Panel")
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
            r1_raw = p_cols[0].text_input("RMB (1pc)", key=f"r1_{i}", value="")
            r10_raw = p_cols[1].text_input("RMB (10pcs)", key=f"r10_{i}", value="")
            r100_raw = p_cols[2].text_input("RMB (100pcs)", key=f"r100_{i}", value="")
            
            def to_num(val):
                try: return float(val.replace(',', '')) if val.strip() else 0.0
                except: return 0.0

            r1, r10, r100 = to_num(r1_raw), to_num(r10_raw), to_num(r100_raw)
            
            # Add data only if at least one price is entered
            if r1 > 0 or r10 > 0 or r100 > 0:
                final_data.append({
                    "model": sel, "specs": m['Specs'],
                    "tiers": [
                        {"qty": 1, "rmb": r1},
                        {"qty": 10, "rmb": r10},
                        {"qty": 100, "rmb": r100}
                    ]
                })

if st.button("➕ Add Another Product Line"):
    st.session_state.rows.append({"model": ""})
    st.rerun()

# --- 5. PREVIEW ---
if final_data:
    st.markdown("---")
    st.subheader("👁️ Live Quote Preview")
    preview_rows = []
    for f in final_data:
        for t in f['tiers']:
            if t['rmb'] > 0:
                usd_price = t['rmb'] / exch_rate
                preview_rows.append({
                    "Model": f['model'], "Qty": t['qty'],
                    "Unit Price (USD)": f"{usd_price:.2f}",
                    "Subtotal": f"{(t['qty'] * usd_price):.2f}"
                })
    if preview_rows:
        st.table(pd.DataFrame(preview_rows))

# --- 6. EXPORT ---
c1, c2 = st.columns(2)

if c1.button("🚀 Export to Excel"):
    if os.path.exists('template.xlsx'):
        wb = load_workbook('template.xlsx')
        ws = wb.active
        
        # Header Metadata
        safe_write(ws, 4, 9, today.strftime("%B %d, %Y"))
        safe_write(ws, 5, 9, expiry.strftime("%B %d, %Y"))
        safe_write(ws, 6, 9, quote_id)
        
        # Customer Details
        safe_write(ws, 10, 2, c_name)
        safe_write(ws, 11, 2, c_contact)
        safe_write(ws, 12, 2, c_addr)
        safe_write(ws, 13, 2, c_phone)
        safe_write(ws, 14, 2, c_email)
            
        start_row = 17
        for block_idx, block in enumerate(final_data):
            # Calculate the top row of the current 3-row block
            current_top = start_row + (block_idx * 3)
            
            # 1. Model Name (Column D / 4)
            safe_write(ws, current_top, 4, block['model'])
            
            # 2. Specs/Remark (Column I / 9)
            safe_write(ws, current_top, 9, block['specs'])
            
            # 3. Tiered Quantities and Prices
            # Col 5 = Qty, Col 6 = Unit Price USD
            for tier_idx, tier in enumerate(block['tiers']):
                row_idx = current_top + tier_idx
                safe_write(ws, row_idx, 5, tier['qty'])
                if tier['rmb'] > 0:
                    safe_write(ws, row_idx, 6, round(tier['rmb'] / exch_rate, 2))
            
            # 4. Fetch and Insert Image (Column H / 8)
            img_url = get_bw_sensing_image(block['model'])
            if img_url:
                try:
                    res = requests.get(img_url, timeout=5)
                    img = XLImage(BytesIO(res.content))
                    img.width, img.height = (90, 90)
                    ws.add_image(img, f'H{current_top}')
                except: pass
            
        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Download Excel", out.getvalue(), f"{quote_id}.xlsx", key="dl_xl_final")

if c2.button("📄 Export to PDF"):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, f"Quotation: {quote_id}", ln=True, align='C')
    pdf_out = pdf.output(dest='S').encode('latin-1')
    st.download_button("📥 Download PDF", pdf_out, f"{quote_id}.pdf", key="dl_pdf_final")
