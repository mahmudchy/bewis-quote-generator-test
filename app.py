import streamlit as st
import pandas as pd
import os
import datetime
import requests
from bs4 import BeautifulSoup
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

# --- 1. DATA LOADING ENGINE (Updated for Multi-Sheet Excel) ---
@st.cache_data
def load_all_models():
    all_data = []
    # The exact name of your uploaded master file
    master_file = 'Model list  all-BWSENSING new.xlsx'
    
    if not os.path.exists(master_file):
        return pd.DataFrame()

    try:
        # Load all sheets at once into a dictionary
        sheets_dict = pd.read_excel(master_file, sheet_name=None)
        
        for sheet_name, df in sheets_dict.items():
            # Clean up column names (remove spaces/newlines)
            df.columns = [str(c).strip() for c in df.columns]
            
            # Find the "Model" column
            model_col = next((c for c in df.columns if "Model" in c), None)
            
            if model_col:
                for _, row in df.iterrows():
                    m_name = str(row[model_col]).strip()
                    # Skip empty rows or header-looking rows
                    if not m_name or m_name.lower() in ["model", "nan"]: 
                        continue
                    
                    specs = []
                    
                    # 1. Axis Logic (Requirement: Single/Dual Axis in Remark)
                    axis_col = next((c for c in df.columns if "axis" in c.lower()), None)
                    if axis_col and str(row[axis_col]).strip() and str(row[axis_col]).lower() != "nan":
                        specs.append(f"Axis: {row[axis_col]}")
                    
                    # 2. Key Parameter Logic
                    important_keys = ['accuracy', 'resolution', 'range', 'output', 'power supply']
                    for col in df.columns:
                        if any(k in col.lower() for k in important_keys) and str(row[col]).strip() and str(row[col]).lower() != "nan":
                            # Avoid duplicating Axis if already added
                            if "axis" not in col.lower():
                                specs.append(f"{col}: {row[col]}")
                    
                    all_data.append({
                        "Model": m_name,
                        "Category": sheet_name, # Sheet name acts as the category
                        "Specs": "\n".join(specs)
                    })
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        
    return pd.DataFrame(all_data)

# --- 2. IMAGE SCRAPER ---
def fetch_product_image(model_name):
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

# --- 3. UI ---
st.set_page_config(layout="wide", page_title="BWSensing Quote Gen")
model_db = load_all_models()

st.sidebar.title("1. Quote Details")
c_name = st.sidebar.text_input("Customer Name")
c_addr = st.sidebar.text_area("Address")
country_code = st.sidebar.text_input("Country Short Code (e.g. SA, AU, US)", value="SA").upper()
rmb_rate = st.sidebar.number_input("RMB to USD Rate", value=7.2)

today = datetime.date.today()
quote_no = f"{today.strftime('%Y%m%d')}-MC-{country_code}"

st.title(f"Quotation: {quote_no}")

if model_db.empty:
    st.error(f"⚠️ Master file 'Model list  all-BWSENSING new.xlsx' not found or contains no 'Model' columns.")

if 'rows' not in st.session_state:
    st.session_state.rows = [{"model": ""}]

selected_items = []
for i, row in enumerate(st.session_state.rows):
    cols = st.columns([3, 1, 1])
    m_options = [""] + sorted(list(model_db['Model'].unique())) if not model_db.empty else [""]
    
    choice = cols[0].selectbox(f"Select Product {i+1}", m_options, key=f"sel_{i}")
    qty = cols[1].number_input("Qty", min_value=1, key=f"qty_{i}")
    rmb = cols[2].number_input("RMB Price", key=f"rmb_{i}")
    
    if choice and not model_db.empty:
        item_data = model_db[model_db['Model'] == choice].iloc[0]
        selected_items.append({
            "model": choice, "desc": item_data['Category'], "qty": qty,
            "remark": item_data['Specs'], "usd": round(rmb/rmb_rate, 2)
        })

if st.button("➕ Add Another Item"):
    st.session_state.rows.append({"model": ""})
    st.rerun()

# --- 4. EXPORT ---
if st.button("🚀 Generate Official Excel"):
    template_path = 'template.xlsx'
    if not os.path.exists(template_path):
        st.error("Missing 'template.xlsx' on GitHub!")
    else:
        wb = load_workbook(template_path)
        ws = wb.active
        ws['I4'], ws['I6'] = today.strftime("%B, %dth. %Y"), quote_no
        ws['B10'], ws['B12'] = c_name, c_addr

        for idx, item in enumerate(selected_items):
            row_num = 17 + idx
            ws.cell(row=row_num, column=1, value=item['desc'])
            ws.cell(row=row_num, column=4, value=item['model'])
            ws.cell(row=row_num, column=5, value=item['qty'])
            ws.cell(row=row_num, column=6, value=item['usd'])
            ws.cell(row=row_num, column=7, value=item['usd'] * item['qty'])
            ws.cell(row=row_num, column=9, value=item['remark'])
            
            img_url = fetch_product_image(item['model'])
            if img_url:
                try:
                    img_data = requests.get(img_url, timeout=5).content
                    img_obj = XLImage(BytesIO(img_data))
                    img_obj.width, img_obj.height = (70, 70)
                    ws.add_image(img_obj, f'H{row_num}')
                except: pass

        output = BytesIO()
        wb.save(output)
        st.download_button("📥 Download Final Quotation", output.getvalue(), f"{quote_no}.xlsx")
