import os
import sys
import pandas as pd
from datetime import datetime
import difflib

# 1. Force jump to the exact absolute folder of this repo
if os.path.basename(os.getcwd()) == 'code':
    os.chdir('..')
elif os.path.exists(os.path.join(os.path.dirname(os.path.abspath(os.getcwd())), 'source')):
    os.chdir(os.path.dirname(os.path.abspath(os.getcwd())))

source_dir = "./source"
output_dir = "./output"

# Retrieve the current system date string matching the file pattern (YYYY-MM-DD)
current_date_str = datetime.now().strftime('%Y-%m-%d')

# 2. Target the latest generated operational files located inside the output directory
updated_sales_file = os.path.join(output_dir, f"销售报表_{current_date_str}_更新版.xlsx")
updated_procure_file = os.path.join(output_dir, f"202605采购明细表_{current_date_str}更新.xlsx")

# 3. Target both separate journal files remaining inside the source directory
current_month_str = current_date_str.replace('-', '')[:6]
cash_journal_m = os.path.join(source_dir, f"日记账20230506起{current_month_str}.xls")
cash_journal_h = os.path.join(source_dir, "2026年日记账.xlsx")

# 4. Read the operational datasets into dataframes
df_journal_m = pd.read_excel(cash_journal_m)
df_journal_h = pd.read_excel(cash_journal_h)
df_procure   = pd.read_excel(updated_procure_file)
df_sales     = pd.read_excel(updated_sales_file)

# 5. Standardize column names by stripping trailing spaces
df_journal_m.columns = [str(c).strip() for c in df_journal_m.columns]
df_journal_h.columns = [str(c).strip() for c in df_journal_h.columns]
df_procure.columns   = [str(c).strip() for c in df_procure.columns]
df_sales.columns     = [str(c).strip() for c in df_sales.columns]


# ==========================================
# CHRONOLOGICAL TRUNCATION (FIXED LOGIC)
# ==========================================

# 1. Create independent working copies to shield raw imported matrices
df_mine = df_journal_m.copy()
df_hers = df_journal_h.copy()

def safe_parse_yyyymmdd(df_target):
    """Safely handles float/string dates, isolating valid 8-digit integers."""
    df_target = df_target.dropna(subset=['日期']).copy()
    # Force string conversion, split off decimals, clean up padding spaces
    cleaned_strings = df_target['日期'].astype(str).str.split('.').str[0].str.strip()
    df_target['日期'] = pd.to_datetime(cleaned_strings, format='%Y%m%d', errors='coerce')
    # Immediately drop rows that failed parsing to prevent numeric fallback corruption
    df_target = df_target.dropna(subset=['日期']).copy()
    return df_target

# 2. Extract and format my tracking data dates securely
df_mine = safe_parse_yyyymmdd(df_mine)
max_my_date = df_mine['日期'].max()

# 3. Clean her date column entries identically
df_hers = safe_parse_yyyymmdd(df_hers)

# 4. Filter out and retain only her rows that contain strictly higher dates
df_journal = df_hers[df_hers['日期'] > max_my_date].copy()

# 5. Format sales tracking date column safely
df_sales['订单日期'] = pd.to_datetime(df_sales['订单日期'], errors='coerce')


# ==========================================
# RECONCILIATION PROCESSING
# ==========================================

# 1. Standardize and force numeric values across all transaction channels
bank_cols = [c for c in ['农行', '建行', '工行', '兴业'] if c in df_journal.columns]
for col in bank_cols:
    df_journal[col] = pd.to_numeric(df_journal[col], errors='coerce').fillna(0.0)

# Target verified sales ledger column headers
sales_amount_col = '销售总价'
df_sales[sales_amount_col] = pd.to_numeric(df_sales[sales_amount_col], errors='coerce').fillna(0.0)
sales_company_col = '客户名称'
sales_companies = df_sales[sales_company_col].dropna().astype(str).str.strip().unique().tolist()

# Target verified procurement ledger column headers
procure_amount_col = '价税合计'
df_procure[procure_amount_col] = pd.to_numeric(df_procure[procure_amount_col], errors='coerce').fillna(0.0)
procure_company_col = '供应商名称'
procure_companies = df_procure[procure_company_col].dropna().astype(str).str.strip().unique().tolist()

# 2. Iterate row by row to process transaction metrics independently
variance_column = []

for idx, row in df_journal.iterrows():
    company_name = str(row['摘要']).strip()
    current_flow = sum(float(row[col]) for col in bank_cols)
    
    # Filter: Skip completely empty lines or placeholder notes
    if current_flow == 0.0 or pd.isna(row['摘要']) or company_name in ['nan', '', 'None']:
        variance_column.append(0.0)
        continue
    
    # ROUTE A: Positive bank flow -> Match with Sales Ledger
    if current_flow > 0.0:
        matches = difflib.get_close_matches(company_name, sales_companies, n=1, cutoff=0.1)
        if matches:
            matched_company = matches[0]
            target_sales_amount = df_sales[df_sales[sales_company_col].astype(str).str.strip() == matched_company][sales_amount_col].sum()
            variance_column.append(current_flow - target_sales_amount)
        else:
            variance_column.append(current_flow)
            
    # ROUTE B: Negative bank flow -> Match with Procurement Ledger
    else:
        absolute_flow = abs(current_flow)
        matches = difflib.get_close_matches(company_name, procure_companies, n=1, cutoff=0.1)
        if matches:
            matched_company = matches[0]
            target_procure_amount = df_procure[df_procure[procure_company_col].astype(str).str.strip() == matched_company][procure_amount_col].sum()
            # Maintain standard structural negative format for expenditures
            variance_column.append(-(absolute_flow - target_procure_amount))
        else:
            variance_column.append(current_flow)

# Inject calculations directly into the active tracking dataframe
df_journal['对账差额'] = variance_column


# ==========================================
# RECOMBINATION AND EXPORT
# ==========================================

# 1. Standardize dates on my original master set to clean datetime format
df_master_output = df_journal_m.copy()
cleaned_master_strings = df_master_output['日期'].astype(str).str.split('.').str[0].str.strip()
df_master_output['日期'] = pd.to_datetime(cleaned_master_strings, format='%Y%m%d', errors='coerce')
df_master_output = df_master_output.dropna(subset=['日期']).copy()

# 2. Recombine the historical master dataset with the newly processed records
df_final_ledger = pd.concat([df_master_output, df_journal], axis=0, ignore_index=True)

# 3. CRITICAL: Reformat the datetime column to a clean string format before exporting
# This completely prevents Excel from showing '1970-01-01' or '00:00:00' hours timestamps
df_final_ledger['日期'] = df_final_ledger['日期'].dt.strftime('%Y-%m-%d')

# 4. Generate the export file path dynamically using the current runtime date
export_file_name = f"日记账_更新版_{current_date_str}.xlsx"
export_file_path = os.path.join(output_dir, export_file_name)

# 5. Export clean database records to destination file path
df_final_ledger.to_excel(export_file_path, index=False)

print(f"Success! Output exported nicely with clean formatting: {export_file_path}")