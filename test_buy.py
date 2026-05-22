import os
import pandas as pd
from datetime import datetime

def test():
    """ Runs integration tests to verify data correctness """
    now = datetime.now()
    mo = now.strftime("%Y%m")     
    today = now.strftime("%Y-%m-%d") 
    
    # 0. File finder helper
    def find(hint, dyn): 
        for f in os.listdir('source'):
            if hint in f and dyn in f: return os.path.join('source', f)
        return None
    
    # Load baselines
    my_old_p = find("采购明细表", mo)
    her_p = find("订货明细", now.strftime("%Y年"))
    
    if not my_old_p or not her_p:
        print(f"[FAIL] Baseline files missing in 'source' folder.")
        return

    df_old = pd.read_excel(my_old_p, engine='xlrd')
    df_her = pd.read_excel(her_p, engine='xlrd')

    # --- Dynamic Sheet Selection for df_her
    xl = pd.ExcelFile(her_p, engine='xlrd')
    # Filter sheet names that follow the 'X月' pattern and find the maximum one
    sheet_months = [s for s in xl.sheet_names if '月' in s]
    if not sheet_months:
        print("[FAIL] Test 1: No monthly sheets found in her file!")
        return
    
    # Sort by extracting the number before '月' to get the latest sheet (e.g., '5月')
    latest_sheet = max(sheet_months, key=lambda x: int(''.join(filter(str.isdigit, x))))
    print(f"[INFO] Test 1: Loading her latest sheet -> '{latest_sheet}'")
    df_her = pd.read_excel(her_p, sheet_name=latest_sheet, engine='xlrd')

    # 2. Target updated output file name
    out_p = os.path.join("output", f"{mo}采购明细表_{today}更新.xlsx")
    if not os.path.exists(out_p): 
        print(f"[FAIL] Output missing: {out_p}")
        return
    df_out = pd.read_excel(out_p)

    print("\n--- Step 2: INTEGRATION TESTS ---")
    
    # Test 1: Latest date alignment check
    date_col = '采购日期' # Replace with your actual date column name if different
    try:
        her_latest = pd.to_datetime(df_her[date_col]).max()
        out_latest = pd.to_datetime(df_out[date_col]).max()
        
        if her_latest == out_latest:
            print(f"[PASS] Test 1: Timeline alignment success ({out_latest.strftime('%Y-%m-%d')}).")
        else:
            print(f"[FAIL] Test 1: Timeline Mismatch! Source: {her_latest}, Output: {out_latest}")
    except KeyError:
        print(f"[FAIL] Test 1: Column '{date_col}' not found in the files!")

    # Test 2: Historical Integrity (ID matching)
    s1 = df_old.iloc[:, 1].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    s_out = df_out['采购单号'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    if all(s1.isin(s_out)):
        print(f"[PASS] Test 2: Historical records (N={len(s1)}) are intact.")
    else:
        print("[FAIL] Test 2: Historical data records missing from output!")


def run():
    """ Executes the Notebook Pipeline and triggers tests """
    print("Executing Notebook Pipeline...")
    cmd = "jupyter nbconvert --to notebook --execute --inplace code/buy_process.ipynb"
    exit_code = os.system(cmd)
    
    if exit_code == 0:
        print("Pipeline finished successfully. Starting Automated Tests...")
        test()  # Call the test function directly
    else:
        print(f"Pipeline crashed with exit code: {exit_code}. Skipping tests.")
        
    return exit_code


if __name__ == "__main__":
    run()