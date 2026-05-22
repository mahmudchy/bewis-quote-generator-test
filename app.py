# --- REPLACE YOUR EXISTING BORDER LOGIC WITH THIS ---
        # Get the source border object
        ref_cell = ws.cell(row=17, column=1)
        b = ref_cell.border
        
        # Create a new Border object that is NOT bound to the old workbook's proxy
        new_border = Border(
            left=Side(style=b.left.style, color=b.left.color),
            right=Side(style=b.right.style, color=b.right.color),
            top=Side(style=b.top.style, color=b.top.color),
            bottom=Side(style=b.bottom.style, color=b.bottom.color)
        )
        
        for idx, block in enumerate(final_data):
            cur_top = 17 + (idx * 3)
            # ... (your existing ultra_safe_write calls) ...

            # Apply the new_border instead of the old one
            for r in range(cur_top, cur_top + 3):
                for c in range(1, 10):
                    cell = ws.cell(row=r, column=c)
                    cell.border = new_border
                    cell.alignment = Alignment(vertical='center', horizontal='center', wrapText=True)
            # ... (rest of your logic)
