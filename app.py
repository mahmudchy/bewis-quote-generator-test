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

# --- 2. ADVANCED DATA SEARCH (Fixes KeyError) ---
@st.cache_data
def load_all_models():
    all_data = []
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv') and 'Model list' in f]
    
    for file in csv_files:
        try:
            # Detect category from filename
            category = file.split('-')[-1].replace('.csv', '').strip()
            df = pd.read_csv(file, header=1).fillna('')
            
            # CRITICAL FIX: Clean column names and find the Model column
            df.columns = [str(c).strip() for c in df.columns]
            model_col = next((c for c in df.columns if "Model" in c), None)
            
            if model_col:
                for _, row in df.iterrows():
                    model_val = str(row[model_col]).strip()
                    if not model_val or "Model" in model_val: continue
                    
                    specs = []
                    # 1. Axis Logic
                    axis_val = ""
                    for col in df.columns:
                        if "axis" in col.lower():
                            axis_val = row[col]
                            break
                    if axis_val: specs.append(f"Axis: {axis_val}")
                    
                    # 2. Key Params Logic
                    important_keys = ['Accuracy', 'Resolution', 'Range', 'Measurement range', 'Output mode']
                    for col in df.columns:
                        if any(key.lower() in col.lower() for key in important_keys) and row[col]:
                            specs.append(f"{col}: {row[col]}")
                    
                    all_data.append({
                        "Model": model_val,
                        "Category": category,
                        "Specs": "\n".join(specs)
                    })
        except: continue
    return pd.DataFrame(all_data) if all_data else pd.DataFrame(columns=["Model", "Category", "Specs"])

# --- 3. SCRAPER ---
def fetch_image_pro(model_name):
    url = f"https://www.bwsensing.com/search.html?q={model_name}"
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        res = requests.get(url, headers=headers, timeout=5)
        soup = BeautifulSoup(res.text, 'html.parser')
        img = soup.find('img', {'class': 'lazy'}) or soup.find('img', src=True)
        if img:
            src = img['src']
            return src if src.startswith('http') else "https://www.bwsensing.com" + src
    except: return None

# --- UI SETUP ---
st.set_page_config(layout="wide")
model_db = load_all_models()

st.sidebar.title("Quotation Settings")
c_name = st.sidebar.text_input("Customer Name")
c_addr = st.sidebar.text_area("Address")
manual_code = st.sidebar.selectbox("Manual Country Override", ["AU", "SA", "US", "UK", "DE", "CN", "IN"])
rmb_rate = st.sidebar.number_input("RMB to USD", value=7.2)

today = datetime.date.today()
final_code = detect_country_code(c_addr, manual_code)
quote_no = f"{today.strftime('%Y%m%d')}-MC-{final_code}"

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

# Product Selection
selected_items = []
for i, row in enumerate(st.session_state.rows):
    cols = st.columns([3, 1, 1])
    # The selectbox handles the auto-suggestion requirement
    model_list = sorted(list(model_db['Model'].unique())) if not model_db.empty else []
    choice = cols[0].selectbox(f"Select Model {i+1}", [""] + model_list, key=f"sel_{i}")
    qty = cols[1].number_input("Qty", min_value=1, key=f"qty_{i}")
    rmb = cols[2].number_input("RMB Price", key=f"rmb_{i}")
    
    if choice and not model_db.empty:
        data = model_db[model_db['Model'] == choice].iloc[0]
        selected_items.append({
            "model": choice, "desc": data['Category'], "qty": qty, 
            "rmb": rmb, "remark": data['Specs'], "usd": round(rmb/rmb_rate, 2)
        })

if st.button("➕ Add Item"):
    st.session_state.rows.append({"model": ""})
    st.rerun()

# --- 4. EXPORT ---
if st.button("🚀 Generate Official Quote"):
    template_path = 'template.xls' if os.path.exists('template.xls') else 'template.xlsx'
    if not os.path.exists(template_path):
        st.error("Please upload 'template.xls' to GitHub!")
    else:
        wb = load_workbook(template_path)
        ws = wb.active
        ws['I4'], ws['I6'] = today.strftime("%B, %dth. %Y"), quote_no
        ws['B10'], ws['B12'] = c_name, c_addr

        for idx, item in enumerate(selected_items):
            r = 17 + idx
            ws.cell(row=r, column=1, value=item['desc'])
            ws.cell(row=r, column=4, value=item['model'])
            ws.cell(row=r, column=5, value=item['qty'])
            ws.cell(row=r, column=6, value=item['usd'])
            ws.cell(row=r, column=7, value=item['usd'] * item['qty'])
            ws.cell(row=r, column=9, value=item['remark'])
            
            img_url = fetch_image_pro(item['model'])
            if img_url:
                try:
                    img_data = requests.get(img_url).content
                    img_obj = XLImage(BytesIO(img_data))
                    img_obj.width, img_obj.height = (70, 70)
                    ws.add_image(img_obj, f'H{r}')
                except: pass
        
        out = BytesIO()
        wb.save(out)
        st.download_button("📥 Download XLSX", out.getvalue(), f"{quote_no}.xlsx")
