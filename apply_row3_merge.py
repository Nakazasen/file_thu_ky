import sys
import os
import time
import shutil
import win32com.client
import pywintypes

sys.stdout.reconfigure(encoding='utf-8')

file_path = os.path.abspath(r"d:\Sandbox\file_thu_ky\スマートファクトリー分科会_管理ツール案(修正版) .xlsx")
backup_path = os.path.abspath(r"d:\Sandbox\file_thu_ky\スマートファクトリー分科会_管理ツール案(修正版)_backup_before_merge_r3.xlsx")
target_sync_path = os.path.abspath(r"d:\Sandbox\file_thu_ky\smart_factory_3ngu_v3.xlsx")

print("1. Creating backup...")
shutil.copy2(file_path, backup_path)
print("Backup created at:", backup_path)

SHEET_MERGE_RANGES = {
    "03_WBS": "A3:Q3",
    "04_課題ToDo": "A3:W3",
    "05_決定事項": "A3:Z3",
    "06_KPI": "A3:W3",
    "07_更新履歴": "A3:Q3"
}

def safe_com_action(func, retries=10, delay=0.5):
    for i in range(retries):
        try:
            return func()
        except pywintypes.com_error as e:
            if e.hresult == -2147418111 or 'rejected' in str(e).lower():
                time.sleep(delay)
            else:
                raise
    return func()

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

try:
    print("2. Opening workbook with Excel COM...")
    wb = safe_com_action(lambda: xl.Workbooks.Open(file_path))
    time.sleep(1)
    
    for sheet_name, merge_range in SHEET_MERGE_RANGES.items():
        print(f"\n--- Processing {sheet_name} (Merge {merge_range}) ---")
        ws = safe_com_action(lambda: wb.Worksheets(sheet_name))
        
        # Unprotect if protected
        safe_com_action(lambda: ws.Unprotect() if ws.ProtectContents else None)
        time.sleep(0.2)
        
        cell_a3 = ws.Range("A3")
        cell_a3.WrapText = True
        cell_a3.HorizontalAlignment = -4131 # xlLeft
        cell_a3.VerticalAlignment = -4108   # xlCenter
        
        rng = ws.Range(merge_range)
        if not rng.MergeCells:
            safe_com_action(lambda: rng.Merge())
            print(f"  Merged {merge_range}")
        else:
            print(f"  Already merged {merge_range}")
            
        # Set row height
        ws.Rows(3).RowHeight = 38
        
        # Ensure formula is intact
        formula_cur = str(cell_a3.Formula2)
        if "PARA_HOWTO" not in formula_cur:
            cell_a3.Formula2 = '=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("PARA_HOWTO",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))'
            
        rng.Locked = True
        
        # Protect worksheet
        safe_com_action(lambda: ws.Protect("", True, True, True, False, True, True, True, False, False, False, False, True, True, True, False))
        ws.EnableSelection = 0 # xlNoRestrictions
        
        print(f"  Row 3 Height: {ws.Rows(3).RowHeight}")
        print(f"  A3 Formula: {cell_a3.Formula2}")
        print(f"  A3 Text preview:\n{cell_a3.Text}")
        time.sleep(0.2)
        
    print("\n3. Saving workbook...")
    safe_com_action(lambda: wb.Save())
    print("Workbook saved!")
    safe_com_action(lambda: wb.Close(SaveChanges=False))
    print("Workbook closed cleanly!")

except Exception as e:
    print("Error during execution:", e)
    if 'wb' in locals():
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
    raise
finally:
    try:
        xl.Quit()
    except Exception:
        pass

print("4. Syncing to smart_factory_3ngu_v3.xlsx...")
shutil.copy2(file_path, target_sync_path)
print("Synchronization complete!")
