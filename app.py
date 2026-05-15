import streamlit as st
import pandas as pd
import os
import datetime
import requests
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Side

# --- 1. DATA & IMAGE LOGIC ---
@st.cache_data
def load_models():
    # ... (Same loading logic as before to find BWSENSING models)
    pass

def get_product_image(model_name):
    # Scrapes the official site for the product photo seen in your browser tabs
    pass

# --- 2. THE DYNAMIC EXCEL ENGINE ---
def generate_quote(final_data, template_path, exch_rate):
    wb = load_workbook(template_path)
    ws = wb.active
    
    # FIND THE FOOTER ANCHOR
    # We search for "Remarks" in Column A to know where the footer starts
    footer_start_row = 20 # Default fallback
    for r in range(1, 100):
        if str(ws.cell(row=r, column=1).value).strip() == "Remarks":
            footer_start_row = r
            break

    # INSERT SPACE
    # Each model block is 3 rows. We insert (3 * number of models) rows
    # plus a 1-row gap between models.
    rows_to_insert = len(final_data) * 4
    ws.insert_rows(footer_start_row, rows_to_insert)

    current_pos = footer_start_row - rows_to_insert # Start writing here
    
    thin = Side(style='thin')
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    for product in final_data:
        # MERGE BLOCKS (A, D, H, I) for 3 rows
        ws.merge_cells(start_row=current_pos, start_column=1, end_row=current_pos+2, end_column=1) # Description
        ws.merge_cells(start_row=current_pos, start_column=4, end_row=current_pos+2, end_column=4) # Bewis No
        ws.merge_cells(start_row=current_pos, start_column=8, end_row=current_pos+2, end_column=8) # Picture
        ws.merge_cells(start_row=current_pos, start_column=9, end_row=current_pos+2, end_column=9) # Remark

        # SET DATA
        ws.cell(row=current_pos, column=1).value = "Inclinometer"
        ws.cell(row=current_pos, column=4).value = product['model']
        ws.cell(row=current_pos, column=9).value = product['specs']
        
        # TIERS (1, 10, 100)
        for i, tier in enumerate(product['tiers']):
            row_idx = current_pos + i
            ws.cell(row=row_idx, column=5).value = tier['qty']
            usd_price = round(tier['rmb'] / exch_rate, 2)
            ws.cell(row=row_idx, column=6).value = usd_price
            ws.cell(row=row_idx, column=7).value = usd_price * tier['qty']
            
            # Apply borders to the tier grid
            for c in range(1, 10):
                ws.cell(row=row_idx, column=c).border = border
                ws.cell(row=row_idx, column=c).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        current_pos += 4 # Move down to next block (3 rows + 1 spacer)

    return wb

# --- 3. STREAMLIT UI ---
st.title("BWSENSING Quote Generator")

# (Standard input fields for RMB prices and Model selection)
#
