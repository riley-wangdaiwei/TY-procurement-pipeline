import os
import pandas as pd
from datetime import datetime

# input files
if os.path.basename(os.getcwd()) == 'code':
    os.chdir('..')

df_m = pd.concat(pd.read_excel([os.path.join('source', f) for f in os.listdir('source') if "采购明细表" in f][0], sheet_name=None, engine='xlrd').values(), ignore_index=True)
df_h = pd.concat(pd.read_excel([os.path.join('source', f) for f in os.listdir('source') if "订货明细" in f][0], sheet_name=None, engine='xlrd').values(), ignore_index=True)

# 1) Rename the first two columns of df_m directly using a dictionary mapping
df_m = df_m.rename(columns={df_m.columns[0]: '采购日期', df_m.columns[1]: '采购单号'})

# 2) Convert dates to datetime format
df_m['采购日期'] = pd.to_datetime(df_m['采购日期'], errors='coerce')
df_h.iloc[:, 0] = pd.to_datetime(df_h.iloc[:, 0], errors='coerce')

# 3) Find rows with dates in df_h that do not exist in df_m
new_data = df_h[~df_h['采购日期'].isin(df_m['采购日期'])]

# 4) Append the new rows into df_m
df_m = pd.concat([df_m, new_data], ignore_index=True)

# 5) Combine 供应商 as duplicated columns
if '供应商名称' in df_m.columns and '供应商全称' in df_m.columns:
    df_m['供应商名称'] = df_m['供应商名称'].fillna(df_m['供应商全称'])
    df_m = df_m.drop(columns=['供应商全称'])
elif '供应商全称' in df_m.columns:
    df_m = df_m.rename(columns={'供应商全称': '供应商名称'})

# Calculate total tax-inclusive amount: 数量 (Quantity) * 含税单价 (Unit Price) = 价税合计 (Total Amount)
df_m['数量'] = pd.to_numeric(df_m['数量'], errors='coerce').fillna(0)
df_m['含税单价'] = pd.to_numeric(df_m['含税单价'], errors='coerce').fillna(0)
df_m['价税合计'] = df_m['数量'] * df_m['含税单价']

# Output 
# Sort the dataframe by '采购日期' in descending order (newest dates on top)
df_m = df_m.sort_values(by='采购日期', ascending=False, na_position='last')

# Re-format the datetime back to a clean string format ('YYYY-MM-DD') for the final Excel report
df_m['采购日期'] = pd.to_datetime(df_m['采购日期'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')

# Construct the original filename pattern with the specific update date suffix
export_now = datetime.now()
yr_str = export_now.strftime("%Y年")
mo_str = export_now.strftime("%Y%m")
date_suffix = export_now.strftime("%Y-%m-%d") + "更新"
export_filename = f"{mo_str}采购明细表_{date_suffix}.xlsx"

# Export the sorted dataframe to the output directory
os.makedirs('output', exist_ok=True)
df_m.to_excel(os.path.join('output', export_filename), index=False)