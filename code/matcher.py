import os
import sys
import pandas as pd
import difflib
from datetime import datetime

# 1. Force jump to the exact absolute folder of this repo
project_root = os.path.dirname(os.path.abspath(os.getcwd()))
if os.path.exists(os.path.join(project_root, 'source')):
    os.chdir(project_root)
else:
    # If already at root, stay here
    pass

# 2. Import the updated procurement sheet from the 'output' directory
df_m = pd.concat(pd.read_excel([os.path.join('output', f) for f in os.listdir('output') if "采购明细表" in f][0], sheet_name=None, engine='openpyxl').values(), ignore_index=True)

# 3. Load Warehouse tab names directly
warehouse_file = [os.path.join('source', f) for f in os.listdir('source') if "仓库报表" in f][0]
warehouse_tabs = list(pd.ExcelFile(warehouse_file, engine='xlrd').sheet_names)

# clean procurement sheet
warehouse_sheets_dict = pd.read_excel(warehouse_file, sheet_name=None, engine='xlrd')

def parse_date_safely(series):
    s_str = series.dropna().astype(str).str.split('.').str[0].str.strip()
    valid_mask = s_str.str.contains('-') | (s_str.str.startswith('20') & (s_str.str.len() == 8))
    s_str = s_str[valid_mask]
    return pd.to_datetime(s_str, errors='coerce', format='mixed')

max_w_date = pd.concat([parse_date_safely(df.iloc[:, 0]) for df in warehouse_sheets_dict.values() if not df.empty]).max()

# only retain dates not in the warehouse
date_col = next((c for c in df_m.columns if '日期' in c or '时间' in c), '采购日期')

df_m = df_m[df_m.apply(lambda r: (
    pd.to_datetime(str(r[date_col]).split('.')[0][:8], format='%Y%m%d', errors='coerce') 
    if '-' not in str(r[date_col]) else pd.to_datetime(r[date_col], errors='coerce')
) > max_w_date, axis=1)].copy()

df_m = df_m.iloc[::-1].reset_index(drop=True)

# Fuzzy matching 品名
# 1. Prepare your variables to mirror your exact production setup
warehouse_names = list(warehouse_tabs)  # Tab names extracted from your warehouse file
my_purchase = df_m.copy()               # Flat procurement data loaded in Step 1

match_results = []

# 2. Run your original row-by-row matching logic
for idx, row in my_purchase.iterrows():
    if pd.isna(row['品号']):
        continue
    
    code_name = str(row['品号']).strip()
    
    # Mirror process.extractOne using built-in sequence matcher
    matches = difflib.get_close_matches(code_name, warehouse_names, n=1, cutoff=0.1)
    if matches:
        match = matches[0]
        score = int(difflib.SequenceMatcher(None, code_name, match).ratio() * 100)
        # Apply substring rule to guarantee 100% score for solid containment matches
        if code_name.upper() in match.upper() or match.upper() in code_name.upper():
            score = 100
    else:
        match = "None"
        score = 0
    
    match_results.append({
        '采购品号': code_name,
        '仓库名': match,
        '匹配分数': score
    })

# 3. Create the final tracking dataframe
match_df = pd.DataFrame(match_results)

# 4. Filter for valid fuzzy matches (score >= 80) 
passed_df = match_df[match_df['匹配分数'] >= 80].copy()

# add 进仓日期 and 进仓数量 and 进仓价格
df_m.columns = [str(c).strip() for c in df_m.columns]

date_col = None
qty_col = None
price_col = None

for col in df_m.columns:
    if '日期' in col or '时间' in col:
        date_col = col
    if '数量' in col or '进仓' in col or '采购数量' in col:
        qty_col = col
    if '单价' in col or '含税单价' in col or '价格' in col:
        price_col = col

if not date_col: date_col = '采购日期'
if not qty_col: qty_col = '数量'
if not price_col: price_col = '含税单价'

# 2. Specialized inline helper function to fix the scientific float dates into clean string format
def clean_financial_date_to_str(val):
    if pd.isna(val):
        return ""
    str_val = f"{int(val)}"
    if len(str_val) >= 8:
        date_str = str_val[:8]
        try:
            return pd.to_datetime(date_str, format='%Y%m%d').strftime('%Y-%m-%d')
        except Exception:
            return ""
    return ""


# 3. Re-initialize raw warehouse dict container
warehouse_sheets_dict = pd.read_excel(warehouse_file, sheet_name=None, engine='xlrd')

# 4. Step-by-step row processing for successfully matched entities without printing logs
for idx, row in passed_df.iterrows():
    code_name = row['采购品号']
    tab_name = row['仓库名']
    
    # Extract exact purchase line items from the flattened df_m
    orig_row = df_m.loc[idx]
    purchase_date = orig_row[date_col]
    purchase_qty = pd.to_numeric(orig_row[qty_col], errors='coerce')
    if pd.isna(purchase_qty): purchase_qty = 0.0
    
    # Retrieve the target warehouse tracking sheet
    df_tab = warehouse_sheets_dict[tab_name].copy()
    df_tab.columns = [str(c).strip() for c in df_tab.columns]
    
    # Ensure baseline calculation columns are present
    if '进仓数量' not in df_tab.columns or '存货数量' not in df_tab.columns:
        continue
        
    # Clean previous inventory rows to capture the last real stock balance value
    df_tab_clean = df_tab.dropna(subset=['存货数量']).copy()
    previous_stock = 0.0
    if not df_tab_clean.empty:
        previous_stock = float(pd.to_numeric(df_tab_clean['存货数量'].iloc[-1], errors='coerce'))
        if pd.isna(previous_stock): previous_stock = 0.0

    # 5. Compute inventory balance update using financial formula
    new_stock_balance = previous_stock + purchase_qty
    
    # 6. Build the new transaction row data structure
    new_row_data = {}
    for col in df_tab.columns:
        if '日期' in col:
            new_row_data[col] = purchase_date
        elif '进仓数量' in col:
            new_row_data[col] = purchase_qty
        # ---------------------------------------------------------
        # 【新增修改点】：在这里把采购表的单价写入老表的“进货单价”列
        # ---------------------------------------------------------
        elif '进货单价' in col:
            new_row_data[col] = pd.to_numeric(orig_row[price_col], errors='coerce') if not pd.isna(orig_row[price_col]) else 0.0
        elif '存货数量' in col:
            new_row_data[col] = new_stock_balance
        else:
            new_row_data[col] = pd.NA
            
    # Append the calculated row right onto the bottom of the dataframe silently
    df_tab = pd.concat([df_tab, pd.DataFrame([new_row_data])], ignore_index=True)
    warehouse_sheets_dict[tab_name] = df_tab

# add new tabs for 新品
# Identify records that failed to match the automated threshold
failed_df = match_df[match_df['匹配分数'] < 80].copy()

if not failed_df.empty:
    # Template schema headers matching your exact sheet structure
    row_tuples_template = [
        ('日期', pd.NA), ('进仓数量', 0.0), ('进仓单号', pd.NA), ('车牌号', pd.NA),
        ('进货单价', 0.0), ('进货总价', 0.0), ('运费（元/吨）', pd.NA), ('运费', pd.NA),
        ('卸货单价（元/吨）', pd.NA), ('卸货费', pd.NA), ('出仓数量', 0.0), ('出仓单号', pd.NA),
        ('车牌号', pd.NA), ('装货单价', pd.NA), ('装货费', pd.NA), ('存货数量', 0.0),
        ('均价', 0.0), ('仓储单价（元/吨/日）', pd.NA), ('存储天数', pd.NA),
        ('仓储费（/日）', pd.NA), ('仓库名', pd.NA), ('占位符', pd.NA), ('库存金额', 0.0)
    ]
    _keys = [t[0] for t in row_tuples_template]
    
    new_pipeline_records = []

    for idx, row in failed_df.iterrows():
        new_tab_name = str(row['采购品号']).strip()
        
        # Pull actual procurement details directly from df_m
        orig_row = df_m.loc[idx]
        purchase_date = orig_row[date_col]
        purchase_qty = pd.to_numeric(orig_row[qty_col], errors='coerce')
        if pd.isna(purchase_qty): purchase_qty = 0.0
        
        # ---------------------------------------------------------
        # 【新增修改 1】：从采购表里提取含税单价（如果没有 price_col 请直接改为 '含税单价'）
        # ---------------------------------------------------------
        purchase_price = pd.to_numeric(orig_row[price_col], errors='coerce') if 'price_col' in locals() else pd.to_numeric(orig_row['含税单价'], errors='coerce')
        if pd.isna(purchase_price): purchase_price = 0.0
        
        # Row 0: Historical Empty Baseline (so max_idx - 1 exists downstream)
        vals_row0 = [pd.NA] * len(_keys)
        df_r0 = pd.DataFrame([vals_row0], columns=_keys)
        df_r0['存货数量'] = 0.0
        df_r0['均价'] = 0.0
        
        # Row 1: Active Inbound Row (with actual data injected immediately)
        vals_row1 = [pd.NA] * len(_keys)
        df_r1 = pd.DataFrame([vals_row1], columns=_keys)
        df_r1['日期'] = purchase_date
        df_r1['进仓数量'] = purchase_qty
        
        # ---------------------------------------------------------
        # 【新增修改 2】：将单价和总价代入 Row 1
        # ---------------------------------------------------------
        df_r1['进货单价'] = purchase_price
        df_r1['进货总价'] = purchase_qty * purchase_price
        
        df_r1['存货数量'] = purchase_qty  # First transaction: inventory = inbound qty
        df_r1['均价'] = 0.0               # Let the downstream manage-loop compute this moving average
        
        # Combine into a perfect 2-row structure
        df_new_tab = pd.concat([df_r0, df_r1], ignore_index=True)
        
        # Safely inject into the freshly re-initialized dictionary
        warehouse_sheets_dict[new_tab_name] = df_new_tab
        
        # Track mapping info to pass onward
        new_pipeline_records.append({
            '采购品号': new_tab_name,
            '仓库名': new_tab_name,
            '匹配分数': row['匹配分数']
        })
        
    # Append these new entries to passed_df so they flow into the downstream calculation loop smoothly
    df_new_mapping = pd.DataFrame(new_pipeline_records)
    passed_df = pd.concat([passed_df, df_new_mapping], ignore_index=True)
    
    print(f"Successfully appended {len(failed_df)} new initialized tabs to warehouse_sheets_dict.")

# manage warehouse data
# 1. Loop through each tab inside the storage container
for tab_name, df_tab in warehouse_sheets_dict.items():
    df_tab.columns = [str(c).strip() for c in df_tab.columns]
    
    if df_tab.empty:
        continue
        
    # Locate column names dynamically using substring matching
    stock_qty_col = next((c for c in df_tab.columns if '存货数量' in c), None)
    inbound_qty_col = next((c for c in df_tab.columns if '进仓数量' in c), None)
    inbound_price_col = next((c for c in df_tab.columns if '单价' in c or '进货单价' in c), None)
    avg_price_col = next((c for c in df_tab.columns if '均价' in c or '平均单价' in c), None)
    
    if not all([stock_qty_col, inbound_qty_col, inbound_price_col, avg_price_col]):
        continue

    # 2. Get the target row index (the newly appended transaction)
    max_idx = df_tab.index[-1]
    
    # 3. Create a temporary filled look-up table to catch the last valid values instantly
    # ffill() automatically cascades the last valid non-null numerical value forward
    df_filled = df_tab.ffill()
    
    # 4. Extract historical balances safely from the second-to-last row of the filled lookup table
    if max_idx > 0:
        prev_idx = max_idx - 1
        prev_stock_qty = pd.to_numeric(df_filled.loc[prev_idx, stock_qty_col], errors='coerce')
        prev_avg_price = pd.to_numeric(df_filled.loc[prev_idx, avg_price_col], errors='coerce')
    else:
        prev_stock_qty = 0.0
        prev_avg_price = 0.0

    # Fallback to zero if the history table has no entries at all
    if pd.isna(prev_stock_qty): prev_stock_qty = 0.0
    if pd.isna(prev_avg_price): prev_avg_price = 0.0

    # 5. Fetch the new transaction numbers from the real, unaltered dataframe at max_idx
    current_inbound_qty = pd.to_numeric(df_tab.loc[max_idx, inbound_qty_col], errors='coerce')
    current_inbound_price = pd.to_numeric(df_tab.loc[max_idx, inbound_price_col], errors='coerce')
    
    if pd.isna(current_inbound_qty): current_inbound_qty = 0.0
    if pd.isna(current_inbound_price): current_inbound_price = 0.0

    # 6. Apply moving average and total volume equations
    total_volume_pool = prev_stock_qty + current_inbound_qty
    
    if total_volume_pool > 0:
        calculated_moving_avg = (
            (prev_stock_qty * prev_avg_price) + (current_inbound_qty * current_inbound_price)
        ) / total_volume_pool
    else:
        calculated_moving_avg = 0.0

    # 7. Write the finalized tokens directly back onto the real dataframe
    df_tab.loc[max_idx, stock_qty_col] = total_volume_pool
    df_tab.loc[max_idx, avg_price_col] = round(calculated_moving_avg, 4)
    
    # Commit changes back to memory storage
    warehouse_sheets_dict[tab_name] = df_tab

# output
# 1. Ensure the output directory exists
output_dir = 'output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 2. Get the current system date dynamically (Format: YYYY-MM-DD)
current_date_str = datetime.now().strftime('%Y-%m-%d')

# 3. Define the dynamic output file path
output_file_name = f'仓库报表_{current_date_str}_更新版.xlsx'
output_file_path = os.path.join(output_dir, output_file_name)

# 4. Stream all modified dataframes seamlessly into a single Excel file
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    for tab_name, df_tab in warehouse_sheets_dict.items():
        # index=False drops the unneeded row tracking indices from the final spreadsheet
        df_tab.to_excel(writer, sheet_name=tab_name, index=False)