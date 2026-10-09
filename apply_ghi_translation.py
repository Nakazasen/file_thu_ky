import sys
import os
import time
import shutil
import win32com.client
import pywintypes

sys.stdout.reconfigure(encoding='utf-8')

file_path = os.path.abspath(r"d:\Sandbox\file_thu_ky\スマートファクトリー分科会_管理ツール案(修正版_20261002) .xlsx")
backup_path = os.path.abspath(r"d:\Sandbox\file_thu_ky\スマートファクトリー分科会_管理ツール案(修正版_20261002)_backup_before_ghi_trans.xlsx")
target_sync_path = os.path.abspath(r"d:\Sandbox\file_thu_ky\smart_factory_3ngu_v3.xlsx")

print("1. Ensuring safety backup...")
if not os.path.exists(backup_path):
    shutil.copy2(file_path, backup_path)
    print("Backup created at:", backup_path)
else:
    print("Backup already exists at:", backup_path)

def com(func, retries=25, delay=0.5):
    for i in range(retries):
        try:
            return func()
        except pywintypes.com_error as e:
            if e.hresult == -2147418111 or 'rejected' in str(e).lower() or 'busy' in str(e).lower():
                time.sleep(delay)
            else:
                raise
    return func()

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False

try:
    print("2. Opening workbook with Excel COM...")
    wb = com(lambda: xl.Workbooks.Open(file_path))
    time.sleep(1)
    
    ws = com(lambda: wb.Worksheets("03_WBS"))
    if com(lambda: ws.ProtectContents):
        com(lambda: ws.Unprotect())
    time.sleep(0.5)
    
    print("3. Updating Row 4 Headers for Columns F, G, H, I...")
    # Copy format from K4 (Navy blue, white text) to G4:I4
    com(lambda: ws.Range("K4").Copy())
    com(lambda: ws.Range("G4:I4").PasteSpecial(-4122)) # xlPasteFormats
    xl.CutCopyMode = False
    time.sleep(0.3)
    
    # Formulas
    f4_formula = '=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("WBS_CATEGORY",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_INPUT_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))'
    g4_formula = '=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("WBS_CATEGORY",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_VI_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))'
    h4_formula = '=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("WBS_CATEGORY",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_JP_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))'
    i4_formula = '=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("WBS_CATEGORY",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_CN_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))'
    
    com(lambda: setattr(ws.Range("F4"), 'Formula2', f4_formula))
    time.sleep(0.2)
    com(lambda: setattr(ws.Range("G4"), 'Formula2', g4_formula))
    time.sleep(0.2)
    com(lambda: setattr(ws.Range("H4"), 'Formula2', h4_formula))
    time.sleep(0.2)
    com(lambda: setattr(ws.Range("I4"), 'Formula2', i4_formula))
    time.sleep(0.3)
    
    print("4. Copying cell formats from K5:K205 to G5:I205...")
    com(lambda: ws.Range("K5:K205").Copy())
    com(lambda: ws.Range("G5:I205").PasteSpecial(-4122)) # xlPasteFormats
    xl.CutCopyMode = False
    time.sleep(0.3)
    
    print("5. Batch assigning translation formulas to Rows 5-205...")
    com(lambda: setattr(ws.Range("G5:G205"), 'Formula2', '=IF(F5="","",TRANSLATE(F5,"","vi"))'))
    time.sleep(0.2)
    com(lambda: setattr(ws.Range("H5:H205"), 'Formula2', '=IF(F5="","",TRANSLATE(F5,"","ja"))'))
    time.sleep(0.2)
    com(lambda: setattr(ws.Range("I5:I205"), 'Formula2', '=IF(F5="","",TRANSLATE(F5,"","zh-Hans"))'))
    time.sleep(0.2)
    
    # Complete K and W down to row 205
    com(lambda: setattr(ws.Range("K5:K205"), 'Formula2', '=IF(J5="","",TRANSLATE(J5,"","vi"))'))
    time.sleep(0.2)
    com(lambda: setattr(ws.Range("W5:W205"), 'Formula2', '=IF(V5="","",TRANSLATE(V5,"","vi"))'))
    time.sleep(0.3)
    
    # Format font and locks for G5:I205
    rng_ghi = com(lambda: ws.Range("G5:I205"))
    com(lambda: setattr(rng_ghi.Font, 'Name', 'Meiryo UI'))
    com(lambda: setattr(rng_ghi.Font, 'Size', 11))
    com(lambda: setattr(rng_ghi.Font, 'Color', 0xC07000)) # Blue text #0070C0
    com(lambda: setattr(rng_ghi, 'HorizontalAlignment', -4131)) # xlLeft
    com(lambda: setattr(rng_ghi, 'VerticalAlignment', -4108))   # xlCenter
    com(lambda: setattr(rng_ghi, 'Locked', True))
    
    # Ensure locks
    com(lambda: setattr(ws.Range("K5:M205"), 'Locked', True))
    com(lambda: setattr(ws.Range("W5:Y205"), 'Locked', True))
    com(lambda: setattr(ws.Range("F5:F205"), 'Locked', False))
    com(lambda: setattr(ws.Range("J5:J205"), 'Locked', False))
    com(lambda: setattr(ws.Range("V5:V205"), 'Locked', False))

    print("6. Setting column widths for G, H, I...")
    com(lambda: setattr(ws.Columns("G"), 'ColumnWidth', 18.0))
    com(lambda: setattr(ws.Columns("H"), 'ColumnWidth', 18.0))
    com(lambda: setattr(ws.Columns("I"), 'ColumnWidth', 18.0))
    com(lambda: setattr(ws.Columns("K"), 'ColumnWidth', 18.0))
    com(lambda: setattr(ws.Columns("L"), 'ColumnWidth', 18.0))
    com(lambda: setattr(ws.Columns("M"), 'ColumnWidth', 18.0))

    print("7. Protecting sheet 03_WBS...")
    com(lambda: ws.Protect("", True, True, True, False, True, True, True, False, False, False, False, True, True, True, False))
    com(lambda: setattr(ws, 'EnableSelection', 0)) # xlNoRestrictions
    
    print("8. Calculating and saving...")
    com(lambda: xl.Calculate())
    time.sleep(1)
    
    com(lambda: wb.Save())
    print("Workbook saved successfully!")
    com(lambda: wb.Close(SaveChanges=False))
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

print("9. Syncing to smart_factory_3ngu_v3.xlsx...")
shutil.copy2(file_path, target_sync_path)
print("Synchronization complete!")
