import streamlit as st
import pandas as pd
import datetime
import requests
from bs4 import BeautifulSoup
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as XLImage
import os

# --- CONFIGURATION ---
COUNTRY_MAP = {
    "Australia": "AU", "United States": "US", "United Kingdom": "UK",
    "Germany": "DE", "China": "CN", "India": "IN", "Canada": "CA"
}

# --- HELPER FUNCTIONS ---
def get_quote_dates():
    today = datetime.date.today()
    valid_until = today + datetime.timedelta(days=30)
    if valid_until.weekday() == 5: valid_until += datetime.timedelta(days=2)
    elif valid_until.weekday() == 6: valid_until += datetime.timedelta(days=1)
    return today.strftime("%B, %dth. %Y"), valid_until.strftime("%B, %dth. %Y"), today.strftime("%Y%m%d")

def search_product_data(model_query):
    if not model_query:
        return "", ""
    # Look for CSV files in the current directory
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv') and 'Model list' in f]
    for file in csv_files:
        try:
            df = pd.read_csv(file, header=1).fillna('')
            df.columns = df.columns.str.strip()
            if 'Model' in df.columns:
                match = df[df['Model'].astype(str).str.contains(model_query, case=False, na=False)]
                if not match.empty:
                    row = match.iloc[0]
                    # Category is derived from the filename
                    category = file.split('-')[-1].replace('.csv', '').strip()
                    specs = []
                    for col in df.columns:
                        if col not in ['Model', 'Remark'] and row[col]:
                            specs.append(f"{col}: {row[col]}")
                    return category, "\n".join(specs)
        except:
            continue
    return "Product", ""

def fetch_bwsensing_image(model_name):
    search_url = f"https://www.bwsensing.com/search.html?q={model_name}"
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(search_url, headers=headers, timeout=5)
        soup = BeautifulSoup(response.text, 'html.parser')
        img_tag = soup.find('img', class_='lazy') or soup.find('img')
        if img_tag and 'src' in img_tag.attrs:
            img_url = img_tag['src']
            return img_url if img_url.startswith('http') else "https://www.bwsensing.com" + img_url
    except:
        return None

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWSensing Quote Tool", layout="wide")
st.title("🛡️ BWSensing Automated Quotation System")

with st.sidebar:
    st.header("1. Customer Details")
    c_name = st.text_input("Customer Name")
    c_addr = st.text_area("Address")
    c_country = st.selectbox("Country (ID Code)", list(COUNTRY_MAP.keys()))
    rmb_rate = st.number_input("RMB to USD Rate", value=7.2)

date_disp, valid_disp, date_id = get_quote_dates()
quote_no = f"{date_id}-MC-{COUNTRY_MAP[c_country]}"
st.info(f"**Quote:** {quote_no} | **Valid Until:** {valid_disp}")

# Initialize state with a different name to avoid method conflict
if 'quote_items' not in st.session_state:
    st.session_state.quote_items = [{"model": "", "desc": "", "qty": 1, "rmb": 0.0, "remark": ""}]

def add_row():
    st.session_state.quote_items.append({"model": "", "desc": "", "qty": 1, "rmb": 0.0, "remark": ""})

# Display Rows
for i, item in enumerate(st.session_state.quote_items):
    cols = st.columns([2, 2, 1, 1, 3])
    
    # Model Input
    new_model = cols[0].text_input(f"Model Number", value=item['model'], key=f"m_{i}")
    
    if new_model != item['model']:
        cat, spec = search_product_data(new_model)
        st.session_state.quote_items[i]['model'] = new_model
        st.session_state.quote_items[i]['desc'] = cat
        st.session_state.quote_items[i]['remark'] = spec
        st.rerun()

    item['desc'] = cols[1].text_input(f"Description", value=item['desc'], key=f"d_{i}")
    item['qty'] = cols[2].number_input(f"Qty", value=item['qty'], min_value=1, key=f"q_{i}")
    item['rmb'] = cols[3].number_input(f"RMB Price", value=item['rmb'], key=f"p_{i}")
    item['remark'] = cols[4].text_area(f"Technical Remark", value=item['remark'], key=f"r_{i}")

st.button("➕ Add Item", on_click=add_row)

# --- EXCEL EXPORT ---
if st.button("📦 Generate Final XLSX"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Quotation"
    
    ws.merge_cells('A2:K2')
    ws['A2'] = "Wuxi Bewis Sensing Technology LLC"
    ws['A2'].font = Font(bold=True, size=14)
    ws['G4'], ws['I4'] = "Date:", date_disp
    ws['G6'], ws['I6'] = "Quote Number:", quote_no
    ws['A10'], ws['B10'] = "Customer:", c_name
    
    header_labels = ["Description", "", "", "Bewis NO", "Qty(Set)", "Unit Price/USD", "Line Total", "Picture", "Remark"]
    for idx, text in enumerate(header_labels, 1):
        ws.cell(row=16, column=idx, value=text).font = Font(bold=True)

    row_idx = 17
    for item in st.session_state.quote_items:
        usd_unit = round(item['rmb'] / rmb_rate, 2)
        ws.cell(row=row_idx, column=1, value=item['desc'])
        ws.cell(row=row_idx, column=4, value=item['model'])
        ws.cell(row=row_idx, column=5, value=item['qty'])
        ws.cell(row=row_idx, column=6, value=usd_unit)
        ws.cell(row=row_idx, column=7, value=usd_unit * item['qty'])
        ws.cell(row=row_idx, column=9, value=item['remark'])
        ws.cell(row=row_idx, column=12, value=item['rmb'])
        
        img_url = fetch_bwsensing_image(item['model'])
        if img_url:
            try:
                res = requests.get(img_url, timeout=5)
                img = XLImage(BytesIO(res.content))
                img.width, img.height = (70, 70)
                ws.add_image(img, f'H{row_idx}')
            except: pass
        row_idx += 1

    ws.column_dimensions['L'].visible = False
    buffer = BytesIO()
    wb.save(buffer)
    st.download_button("📥 Download Official Quote", buffer.getvalue(), f"{quote_no}.xlsx")
