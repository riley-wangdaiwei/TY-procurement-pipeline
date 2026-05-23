import os
import pandas as pd
import numpy as np
from datetime import datetime
import difflib

# 1. Force jump to the exact absolute folder of this repo to ensure reliable execution
project_root = os.path.dirname(os.path.abspath(os.getcwd()))
if os.path.exists(os.path.join(project_root, 'source')):
    os.chdir(project_root)
else:
    # If already at root, stay here
    pass

# 2. Define paths matching the TY_REPORT_NEW directory structure
source_dir = "./source"
output_dir = "./output"

sales_report_file = os.path.join(source_dir, "销售报表202605.xls")
order_detail_file = os.path.join(source_dir, "订单明细2026年.xls")
warehouse_output_file = os.path.join(output_dir, "仓库报表_2026-05-22_更新版.xlsx")

# 3. Read the source files
df_sales = pd.read_excel(sales_report_file)
# Reading the '5月' sheet from the order details file
df_orders = pd.read_excel(order_detail_file, sheet_name='5月')

# Strip column spaces for structural safety
df_sales.columns = [str(c).strip() for c in df_sales.columns]
df_orders.columns = [str(c).strip() for c in df_orders.columns]

# copy baseline strcuture
# 1. Standardize dates in the existing sales report to find the latest processed date
df_sales['订单日期'] = pd.to_datetime(df_sales['订单日期'], errors='coerce')
latest_sales_date = df_sales['订单日期'].max()

# 2. Filter out empty/summary rows from order details and standardize its date column
df_orders_clean = df_orders.dropna(subset=['订单单号', '品号']).copy()
df_orders_clean['订单日期'] = pd.to_datetime(df_orders_clean['订单日期'], errors='coerce')

# 3. Incremental check: Filter only for records with dates newer than your latest sales report entry
if not pd.isna(latest_sales_date):
    df_incremental_orders = df_orders_clean[df_orders_clean['订单日期'] > latest_sales_date].copy()
    print(f"Incremental filtering active. Latest processed date found: {latest_sales_date.strftime('%Y-%m-%d')}")
else:
    # If the sales report template is empty, grab all clean records
    df_incremental_orders = df_orders_clean.copy()
    print("Existing sales report appears empty. Processing all records from order details.")

# 4. Initialize the new Sales DataFrame framework using only the incremental subset index
df_new_sales = pd.DataFrame(index=df_incremental_orders.index)

# 5. Directly mapping columns A-E, J from the incremental slice using Chinese column names
df_new_sales['订单日期'] = df_incremental_orders['订单日期']    # Col A
df_new_sales['订单单号'] = df_incremental_orders['订单单号']    # Col B
df_new_sales['品号']     = df_incremental_orders['品号']        # Col C
df_new_sales['品名']     = df_incremental_orders['品名']        # Col D

# Ensure numeric values are safely cast to float to avoid calculation crashes
df_new_sales['数量']     = pd.to_numeric(df_incremental_orders['交易数量'], errors='coerce').fillna(0.0) # Col E
df_new_sales['销售单价'] = pd.to_numeric(df_incremental_orders['含税单价'], errors='coerce').fillna(0.0) # Col J

if '业务员' in df_incremental_orders.columns:
    df_new_sales['业务员'] = df_incremental_orders['业务员'] # Col G

# Fetch 采购单价 and Inject into H
from difflib import SequenceMatcher

# 1. Extract all sheet names from the processed warehouse file as the match library
warehouse_sheets = pd.read_excel(warehouse_output_file, sheet_name=None)
warehouse_names = [str(k).strip() for k in warehouse_sheets.keys()]

# 2. Pre-extract the latest moving average price from each warehouse tab
warehouse_prices = {}
for tab in warehouse_names:
    df_tab = warehouse_sheets[tab].copy()
    df_tab.columns = [str(c).strip() for c in df_tab.columns]
    if '均价' in df_tab.columns:
        df_tab_valid = df_tab.dropna(subset=['均价'])
        warehouse_prices[tab] = float(df_tab_valid['均价'].iloc[-1]) if not df_tab_valid.empty else 0.0
    else:
        warehouse_prices[tab] = 0.0

# 3. Apply your exact row-by-row matching formula chain
matched_prices = []
match_scores = []

for idx, row in df_new_sales.iterrows():
    code_name = str(row['品号']).strip()
    
    # Mirror process.extractOne using built-in sequence matcher
    matches = difflib.get_close_matches(code_name, warehouse_names, n=1, cutoff=0.1)
    if matches:
        match_tab = matches[0]
        score = int(difflib.SequenceMatcher(None, code_name, match_tab).ratio() * 100)
        
        # Your core containment rule: force 100 if one string contains the other
        if code_name.upper() in match_tab.upper() or match_tab.upper() in code_name.upper():
            score = 100
    else:
        match_tab = "None"
        score = 0
        
    # Extract the average price if it passes your 80-point threshold; otherwise default to 0.0
    price_val = warehouse_prices.get(match_tab, 0.0) if score >= 80 else 0.0
    
    matched_prices.append(price_val)
    match_scores.append(score)

# 4. Write back the columns to the Sales DataFrame framework
df_new_sales['采购单价'] = matched_prices
df_new_sales['匹配度'] = match_scores

# KPI Calculations
# Aliasing columns into exact mathematical variables specified in your formulas
E = df_new_sales['数量'].astype(float)
J = df_new_sales['销售单价'].astype(float)
H = df_new_sales['采购单价'].astype(float)

# 1. E * J = K
df_new_sales['销售总价'] = E * J                          # Col K

# 2. (J - H) * E = N
df_new_sales['差价'] = (J - H) * E                       # Col N

# 3. (J + H) * E * 0.0003 = O
df_new_sales['印花税'] = (J + H) * E * 0.0003            # Col O

# 4. P = (J - H) / 1.13 * 0.13 * E
df_new_sales['增值税'] = ((J - H) / 1.13) * 0.13 * E        # Col P

# 5. Q = P * 0.12
df_new_sales['附加税'] = df_new_sales['增值税'] * 0.12     # Col Q

# 6. R = E * 10
df_new_sales['运费'] = E * 10                             # Col R

# 7. T = N - O - P - R - Q
df_new_sales['利润'] = (
    df_new_sales['差价'] 
    - df_new_sales['印花税'] 
    - df_new_sales['增值税'] 
    - df_new_sales['运费'] 
    - df_new_sales['附加税']
)                                                            # Col T

# 1. Clean the original sales report to remove any ghost rows filled with zeros
if '数量' in df_sales.columns:
    df_sales = df_sales[df_sales['数量'] > 0].copy()

# 2. Add template baseline placeholder to the new data
df_new_sales['供应商'] = 'A'

# 3. Align exactly with your standard corporate template column sequence
export_columns = [
    '订单日期', '订单单号', '品号', '品名', '数量', '业务员', '采购单价', 
    '供应商', '销售单价', '销售总价', '差价', '印花税', '增值税', '附加税', '运费', '利润'
]
df_incremental_export = df_new_sales[[col for col in export_columns if col in df_new_sales.columns]].copy()

# 4. Clean the new batch data to keep only rows with physical quantities
df_incremental_export = df_incremental_export[df_incremental_export['数量'] > 0]

# 5. Format the date column to string if it was converted to datetime
if not df_incremental_export.empty and pd.api.types.is_datetime64_any_dtype(df_incremental_export['订单日期']):
    df_incremental_export['订单日期'] = df_incremental_export['订单日期'].dt.strftime('%Y-%m-%d')

# 6. Append directly onto the freshly cleaned original sales report framework
df_final_combined = pd.concat([df_sales, df_incremental_export], ignore_index=True)

# 7. Dynamic naming matching corporate style: 销售报表_YYYY-MM-DD_更新版.xlsx
current_date_str = datetime.now().strftime('%Y-%m-%d')
output_sales_path = os.path.join(output_dir, f"销售报表_{current_date_str}_更新版.xlsx")

# 8. Save clean data to Excel
df_final_combined.to_excel(output_sales_path, index=False)