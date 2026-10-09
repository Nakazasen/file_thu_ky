import sys
import os
import shutil
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

file_path = os.path.abspath(r"d:\Sandbox\file_thu_ky\スマートファクトリー分科会_管理ツール案(修正版_20261002) .xlsx")
target_sync_path = os.path.abspath(r"d:\Sandbox\file_thu_ky\smart_factory_3ngu_v3.xlsx")

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

try:
    wb = xl.Workbooks.Open(file_path)
    ws = wb.Worksheets("01_全体サマリー")
    if ws.ProtectContents:
        ws.Unprotect()
        
    # Fix D5
    ws.Range("D5").Formula2 = '=COUNTIF(\'03_WBS\'!$T$5:$T$205,INDEX(_翻訳マスタ!$B$92:$D$92,1,MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0)))'
    
    # Fix C10:C19
    for r in range(10, 20):
        ws.Range(f"C{r}").Formula2 = f'=COUNTIFS(\'03_WBS\'!$B$5:$B$205,$A{r},\'03_WBS\'!$T$5:$T$205,INDEX(_翻訳マスタ!$B$92:$D$92,1,MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0)))'
        
    ws.Protect("", True, True, True, False, True, True, True, False, False, False, False, True, True, True, False)
    ws.EnableSelection = 0
    
    xl.Calculate()
    wb.Save()
    wb.Close(SaveChanges=False)
    print("Fixed REF errors in 01_全体サマリー successfully!")

finally:
    xl.Quit()

shutil.copy2(file_path, target_sync_path)
print("Synced to smart_factory_3ngu_v3.xlsx!")
