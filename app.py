import os

# --- ADVANCED SEARCH LOGIC ---
def search_product_in_csvs(model_query):
    """Searches across all uploaded CSV files for the model number."""
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv') and 'Model list' in f]
    
    for file in csv_files:
        try:
            # We use header=1 or 2 based on your file structure
            df = pd.read_csv(file, header=1) 
            # Clean column names
            df.columns = df.columns.str.strip()
            
            # Look for a match in the 'Model' column
            if 'Model' in df.columns:
                match = df[df['Model'].str.contains(model_query, case=False, na=False)]
                if not match.empty:
                    # Get the Sheet Name from the filename (e.g., 'Tilt sensor')
                    category = file.split('-')[-1].replace('.csv', '').strip()
                    # Combine parameters into a remark
                    row = match.iloc[0]
                    specs = "\n".join([f"{k}: {v}" for k, v in row.items() if k not in ['Model', 'Remark']])
                    return category, specs
        except:
            continue
    return "Product", ""

# --- UPDATED UI ROW ---
# In your row loop, add a 'Search' button or trigger:
if model:
    category_found, specs_found = search_product_in_csvs(model)
    # This automatically updates the fields
    item['desc'] = category_found
    item['remark'] = specs_found