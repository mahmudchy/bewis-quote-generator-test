import streamlit as st
import pandas as pd
import os
import datetime
import requests
from bs4 import BeautifulSoup
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

# --- 1. ENHANCED COUNTRY LOGIC ---
COUNTRY_LOOKUP = {
    "saudi": "SA", "ksa": "SA", "australia": "AU", "united states": "US", 
    "usa": "US", "uk": "UK", "germany": "DE", "china": "CN", "india": "IN"
}

def detect_country_code(address, selected_country):
    addr_lower = address.lower()
    for name, code in COUNTRY_LOOKUP.items():
        if name in addr_lower:
            return code
    return selected_country

# --- 2. ADVANCED DATA SEARCH (Including Axis) ---
@st.cache_data
def load_all_models():
    all_data = []
    # Identify all your uploaded CSV files
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv') and 'Model list' in f]
    
    for file in csv_files:
        try:
            # Detect category from filename
            category = file.split('-')[-1].replace('.csv', '').strip()
            df = pd.read_csv(file, header=1).fillna('')
            df.columns = df.columns.str.strip()
            
            if 'Model' in df.columns:
                for _, row in df.iterrows():
                    # Build the automated Remark string
                    specs = []
                    
                    # 1. Pull Axis Information specifically
                    axis_val = ""
                    for col in df.columns:
                        if "axis" in col.lower():
                            axis_val = row[col]
                            break
                    if axis_val:
                        specs.append(f"Axis: {axis_val}")
                    
                    # 2. Pull other key parameters
                    important_cols = ['Accuracy', 'Resolution', 'Range', 'Measurement range', 'Output mode']
                    for col in df.columns:
                        if any(key in col for key in important_cols) and row[col]:
                            specs.append(f"{col}: {row[col]}")
                    
                    all_data.append({
                        "Model": str(row['Model']),
                        "Category": category,
                        "Specs": "\n".join(specs)
                    })
        except: continue
    return pd.DataFrame(all_data)

# --- 3. BROWSER-MIMIC IMAGE SCRAPER ---
def fetch_image_pro(model_name):
    url = f"https://www.bwsensing.com/search.html?q={model_name}"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    try:
        res = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(res.text, 'html.parser')
        # BWSensing specific image target
        img = soup.find('img', {'class': 'lazy'}) or soup.find('img', src=True)
        if img:
            src = img['src']
            return src if src.startswith('http') else "https://www.bwsensing.com" + src
    except: return None

# --- UI SETUP ---
st.set_page_config(layout="wide", page_title="BWSensing Official Quote")
model_db = load_all_models()

st.sidebar.title("Quotation Settings")
c_name = st.sidebar.text_input("Customer Name")
c_addr = st.sidebar.text_area("Address (Include Country)")
manual_code = st.sidebar.selectbox("Manual Country Override", ["AU", "SA", "US", "UK", "DE", "CN", "IN"])
rmb_rate = st.sidebar.number_input("RMB to USD Exchange Rate", value=7.2)

# Date & Quote ID Logic
today = datetime.date.today()
final_code = detect_country_code(c_addr, manual_code)
quote_no = f"{today.strftime('%Y%m%d')}-MC-{final_code}"

st.header(f"Official Quotation: {quote_no}")

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

def add_item(): st.session_state.rows.append({"model": ""})

# Product Selection
selected_items = []
for i, row in enumerate(st.session_state.rows):
    cols = st.columns([3, 1, 1])
    # The selectbox handles the auto-suggestion requirement
    choice = cols[0].selectbox(f"Select Model {i+1}", [""] + sorted(list(model_db['Model'].unique())), key=f"sel_{i}")
    qty = cols[1].number_input("Qty", min_value=1, key=f"qty_{i}")
    rmb = cols[2].number_input("RMB Price", key=f"rmb_{i}")
    
    if choice:
        data = model_db[model_db['Model'] == choice].iloc[0]
        selected_items.append({
            "model": choice, "desc": data['Category'], "qty": qty, 
            "rmb": rmb, "remark": data['Specs'], "usd": round(rmb/rmb_rate, 2)
        })

st.button("➕ Add Another Product", on_click=add_item)

# --- 4. THE TEMPLATE EXPORT (Preserves Design/Logos) ---
if st.button("🚀 Generate Official Quote (XLSX)"):
    if not os.path.exists('template.xls') and not os.path.exists('template.xlsx'):
        st.error("STOP: Please upload your original Excel file to GitHub and rename it to 'template.xls'")
    else:
        # Load the designed template (works with .xls or .xlsx)
        template_name = 'template.xls' if os.path.exists('template.xls') else 'template.xlsx'
        wb = load_workbook(template_name) 
        ws = wb.active
        
        # Mapping values to the exact cells from your shared file
        ws['I4'] = today.strftime("%B, %dth. %Y")
        ws['I6'] = quote_no
        ws['B10'] = c_name
        ws['B12'] = c_addr

        curr_row = 17
        for item in selected_items:
            ws.cell(row=curr_row, column=1).value = item['desc']
            ws.cell(row=curr_row, column=4).value = item['model']
            ws.cell(row=curr_row, column=5).value = item['qty']
            ws.cell(row=curr_row, column=6).value = item['usd']
            ws.cell(row=curr_row, column=7).value = item['usd'] * item['qty']
            ws.cell(row=curr_row, column=9).value = item['remark']
            
            # Image Injection
            img_url = fetch_image_pro(item['model'])
            if img_url:
                try:
                    img_data = requests.get(img_url, timeout=5).content
                    img_obj = XLImage(BytesIO(img_data))
                    img_obj.width, img_obj.height = (75, 75)
                    ws.add_image(img_obj, f'H{curr_row}')
                except: pass
            curr_row += 1
        
        # Final Output
        out_buffer = BytesIO()
        wb.save(out_buffer)
        st.download_button("📥 Click to Download Official Quote", out_buffer.getvalue(), f"{quote_no}.xlsx")
