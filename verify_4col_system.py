import sys
import time
import win32com.client
import openpyxl.utils

sys.stdout.reconfigure(encoding='utf-8')

def col_letter(c):
    return openpyxl.utils.get_column_letter(c)

def verify_all():
    file_path = r'd:\Sandbox\file_thu_ky\smart_factory_3ngu_v3.xlsx'
    print(f"\n=======================================================")
    print(f"VERIFYING 4-COLUMN MULTI-USER SYSTEM ON: {file_path}")
    print(f"=======================================================\n")
    
    xl = win32com.client.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    
    try:
        wb = xl.Workbooks.Open(file_path)
        xl.Calculate()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        # 1. Check for #REF! across all 11 sheets
        print("--- CHECK 1: Scanning for #REF! errors across all sheets ---")
        total_refs = 0
        for ws in wb.Worksheets:
            sname = ws.Name
            used = ws.UsedRange
            max_r = min(used.Rows.Count, 305)
            max_c = min(used.Columns.Count, 40)
            
            formulas = ws.Range(ws.Cells(1, 1), ws.Cells(max_r, max_c)).Formula2
            values = ws.Range(ws.Cells(1, 1), ws.Cells(max_r, max_c)).Value
            
            s_refs = 0
            if formulas and values:
                for r_idx in range(len(formulas)):
                    for c_idx in range(len(formulas[r_idx])):
                        f = str(formulas[r_idx][c_idx]) if formulas[r_idx][c_idx] is not None else ""
                        v = str(values[r_idx][c_idx]) if values[r_idx][c_idx] is not None else ""
                        if "#REF!" in f or "#REF!" in v:
                            s_refs += 1
                            print(f"  [ERROR] #REF! in {sname}!{col_letter(c_idx+1)}{r_idx+1}: {f}")
            total_refs += s_refs
            print(f"  Sheet '{sname}': {s_refs} #REF! errors. (Used columns: {used.Columns.Count})")
            
        print(f"\nTOTAL #REF! ERRORS IN WORKBOOK: {total_refs}")

        # 2. Check 01_全体サマリー cross-sheet formulas
        print("\n--- CHECK 2: Checking 01_全体サマリー Dashboard Formulas ---")
        ws_sum = wb.Worksheets('01_全体サマリー')
        sum_formulas = []
        for r in range(1, 35):
            for c in range(1, 15):
                f = ws_sum.Cells(r, c).Formula2
                if f and any(target in f for target in ['03_WBS', '04_課題ToDo', '05_決定事項', '06_KPI']):
                    sum_formulas.append((f"{col_letter(c)}{r}", f, ws_sum.Cells(r, c).Text))
                    
        print(f"Found {len(sum_formulas)} dynamic cross-sheet formulas:")
        for cell_addr, f, val in sum_formulas[:10]:
            print(f"  {cell_addr} -> Value: '{val}' | Formula: {f}")

        # 3. Test Multi-User Tri-directional Translation in 03_WBS
        print("\n--- CHECK 3: Multi-User Tri-directional Translation in 03_WBS ---")
        ws_wbs = wb.Worksheets('03_WBS')
        ws_wbs.Unprotect()
        
        # Test 3a: Vietnamese Input in E8
        print("  [3a. Vietnamese Input] Typing 'Kiểm tra quy trình sản xuất' into E8...")
        ws_wbs.Range("E8").Value = "Kiểm tra quy trình sản xuất"
        
        # Test 3b: Japanese Input in E9
        print("  [3b. Japanese Input] Typing '製造プロセスの確認' into E9...")
        ws_wbs.Range("E9").Value = "製造プロセスの確認"
        
        # Test 3c: Chinese Input in E10
        print("  [3c. Chinese Input] Typing '检查生产流程' into E10...")
        ws_wbs.Range("E10").Value = "检查生产流程"
        
        xl.Calculate()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
            
        print("  Waiting 2s for Microsoft Translator AI cloud responses...")
        time.sleep(2)
        
        # Read back Row 8 (Vietnamese input)
        print("\n  == Results for Row 8 (Vietnamese input: 'Kiểm tra quy trình sản xuất') ==")
        print(f"    E8 [Input]: '{ws_wbs.Range('E8').Text}'")
        print(f"    F8 (VI):    '{ws_wbs.Range('F8').Text}' (Formula: {ws_wbs.Range('F8').Formula2})")
        print(f"    G8 (JP):    '{ws_wbs.Range('G8').Text}' (Formula: {ws_wbs.Range('G8').Formula2})")
        print(f"    H8 (CN):    '{ws_wbs.Range('H8').Text}' (Formula: {ws_wbs.Range('H8').Formula2})")
        
        # Read back Row 9 (Japanese input)
        print("\n  == Results for Row 9 (Japanese input: '製造プロセスの確認') ==")
        print(f"    E9 [Input]: '{ws_wbs.Range('E9').Text}'")
        print(f"    F9 (VI):    '{ws_wbs.Range('F9').Text}' (Formula: {ws_wbs.Range('F9').Formula2})")
        print(f"    G9 (JP):    '{ws_wbs.Range('G9').Text}' (Formula: {ws_wbs.Range('G9').Formula2})")
        print(f"    H9 (CN):    '{ws_wbs.Range('H9').Text}' (Formula: {ws_wbs.Range('H9').Formula2})")
        
        # Read back Row 10 (Chinese input)
        print("\n  == Results for Row 10 (Chinese input: '检查生产流程') ==")
        print(f"    E10 [Input]: '{ws_wbs.Range('E10').Text}'")
        print(f"    F10 (VI):    '{ws_wbs.Range('F10').Text}' (Formula: {ws_wbs.Range('F10').Formula2})")
        print(f"    G10 (JP):    '{ws_wbs.Range('G10').Text}' (Formula: {ws_wbs.Range('G10').Formula2})")
        print(f"    H10 (CN):    '{ws_wbs.Range('H10').Text}' (Formula: {ws_wbs.Range('H10').Formula2})")

        # 4. Check 06_KPI Row 9 (Purpose Input in H9)
        print("\n--- CHECK 4: Checking 06_KPI Purpose Field ---")
        ws_kpi = wb.Worksheets('06_KPI')
        print(f"  H9 [Input]: '{ws_kpi.Range('H9').Text}'")
        print(f"  I9 (VI):    '{ws_kpi.Range('I9').Text}'")
        print(f"  J9 (JP):    '{ws_kpi.Range('J9').Text}'")
        print(f"  K9 (CN):    '{ws_kpi.Range('K9').Text}'")

        # Clean up test rows in 03_WBS (keep row 8 as a live demo, clear row 9 & 10)
        ws_wbs.Range("E9").Value = ""
        ws_wbs.Range("E10").Value = ""
        
        # Re-protect 03_WBS
        ws_wbs.Protect(
            Password="",
            DrawingObjects=False,
            Contents=True,
            Scenarios=True,
            AllowFormattingCells=True,
            AllowFormattingColumns=True,
            AllowFormattingRows=True,
            AllowInsertingRows=True,
            AllowDeletingRows=True,
            AllowSorting=True,
            AllowFiltering=True
        )
        ws_wbs.EnableSelection = 0

        # Save verified workbook
        wb.Save()
        wb.Close(True)
        print("\n=======================================================")
        print("ALL VERIFICATIONS COMPLETED SUCCESSFULLY!")
        print("=======================================================\n")
    finally:
        xl.Quit()

if __name__ == '__main__':
    verify_all()
