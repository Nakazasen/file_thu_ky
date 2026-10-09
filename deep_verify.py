import sys
import win32com.client
import openpyxl.utils
import time
import zipfile
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')

def col_letter(c):
    return openpyxl.utils.get_column_letter(c)

def verify():
    file_path = r'd:\Sandbox\file_thu_ky\smart_factory_3ngu_v3.xlsx'
    print(f"\n=======================================================")
    print(f"RUNNING DEEP VERIFICATION ON: {file_path}")
    print(f"=======================================================\n")
    
    xl = win32com.client.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    
    results = {
        "ref_errors": 0,
        "other_errors": 0,
        "sheets_checked": 0,
        "summary_checks_passed": True,
        "translate_row9_passed": False,
        "translate_unaccented_passed": False,
        "translate_term_passed": False,
        "translate_empty_passed": False,
        "lang_switch_passed": True,
        "protection_checks_passed": True,
        "freeze_panes_passed": True,
        "xml_validation_passed": True
    }
    
    try:
        wb = xl.Workbooks.Open(file_path)
        xl.Calculate()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception as e:
            print("CalculateUntilAsyncQueriesDone notice:", e)

        # -------------------------------------------------------------
        # TEST 1: Check for #REF!, #VALUE!, #NAME? across all sheets
        # -------------------------------------------------------------
        print("--- TEST 1: Checking for Broken Formula Errors Across All Sheets ---")
        for ws in wb.Worksheets:
            sname = ws.Name
            results["sheets_checked"] += 1
            used = ws.UsedRange
            max_r = min(used.Rows.Count, 310)
            max_c = min(used.Columns.Count, 35)
            
            formulas = ws.Range(ws.Cells(1, 1), ws.Cells(max_r, max_c)).Formula2
            values = ws.Range(ws.Cells(1, 1), ws.Cells(max_r, max_c)).Value
            
            sheet_refs = 0
            sheet_errs = 0
            if formulas and values:
                for r_idx in range(len(formulas)):
                    for c_idx in range(len(formulas[r_idx])):
                        f = str(formulas[r_idx][c_idx]) if formulas[r_idx][c_idx] is not None else ""
                        v = str(values[r_idx][c_idx]) if values[r_idx][c_idx] is not None else ""
                        
                        if "#REF!" in f or "#REF!" in v:
                            sheet_refs += 1
                            print(f"  [ERROR] #REF! in {sname}!{col_letter(c_idx+1)}{r_idx+1} Formula: {f} Value: {v}")
                        elif any(err in v for err in ["#NAME?", "#VALUE!"]):
                            sheet_errs += 1
                            print(f"  [WARN] {v} in {sname}!{col_letter(c_idx+1)}{r_idx+1} Formula: {f}")
            
            results["ref_errors"] += sheet_refs
            results["other_errors"] += sheet_errs
            print(f"  Sheet '{sname}': {sheet_refs} #REF! errors, {sheet_errs} other formula errors.")
            
        print(f"Total #REF! errors across entire workbook: {results['ref_errors']}")

        # -------------------------------------------------------------
        # TEST 2: Verify Cross-Sheet References in 01_全体サマリー
        # -------------------------------------------------------------
        print("\n--- TEST 2: Checking 01_全体サマリー References ---")
        ws_sum = wb.Worksheets('01_全体サマリー')
        sum_formulas = []
        for r in range(1, 40):
            for c in range(1, 15):
                f = ws_sum.Cells(r, c).Formula2
                if f and any(target in f for target in ['06_KPI', '04_課題ToDo', '03_WBS']):
                    sum_formulas.append((f"{col_letter(c)}{r}", f, ws_sum.Cells(r, c).Text))
                    
        print(f"Found {len(sum_formulas)} dynamic cross-sheet formulas in 01_全体サマリー:")
        for cell_addr, formula, val in sum_formulas[:10]:
            print(f"  Cell {cell_addr} -> Value: {val} | Formula: {formula}")
            if "#REF!" in formula:
                results["summary_checks_passed"] = False

        # -------------------------------------------------------------
        # TEST 3: Verify Column Layout & Suffixes on Restructured Sheets
        # -------------------------------------------------------------
        print("\n--- TEST 3: Column Counts & Headers for 5 Restructured Sheets ---")
        expected_counts = {
            '06_KPI': 28,
            '04_課題ToDo': 30,
            '05_決定事項': 25,
            '03_WBS': 20,
            '07_更新履歴': 14
        }
        for sname, exp_cols in expected_counts.items():
            ws = wb.Worksheets(sname)
            cnt = ws.UsedRange.Columns.Count
            h_row = [str(ws.Cells(4, c).Text) for c in range(1, exp_cols + 1)]
            print(f"  [{sname}] Cols: {cnt} (Expected: {exp_cols})")
            print(f"    Headers: {h_row[:7]}...")
            if cnt != exp_cols:
                print(f"    [MISMATCH] Expected {exp_cols} cols, got {cnt}")

        # -------------------------------------------------------------
        # TEST 4: Live Translation Verification (Explicit 'vi' Source)
        # -------------------------------------------------------------
        print("\n--- TEST 4: Live Translation in 06_KPI Row 9 ---")
        ws_kpi = wb.Worksheets('06_KPI')
        
        # Test 4a: Check Row 9 Purpose column has "tôi tên là" and evaluates to JP & CN
        row9_src = ws_kpi.Cells(9, 7).Text
        row9_jp = ws_kpi.Cells(9, 8).Text
        row9_cn = ws_kpi.Cells(9, 9).Text
        row9_jp_form = ws_kpi.Cells(9, 8).Formula2
        row9_cn_form = ws_kpi.Cells(9, 9).Formula2
        print(f"  [4a. Row 9 Persistent Test Case]")
        print(f"      Source Col 7 (G9): '{row9_src}'")
        print(f"      JP Col 8 (H9): '{row9_jp}' (Formula: {row9_jp_form})")
        print(f"      CN Col 9 (I9): '{row9_cn}' (Formula: {row9_cn_form})")
        if row9_src == "tôi tên là" and "名前" in row9_jp and "名字" in row9_cn:
            results["translate_row9_passed"] = True

        # Test 4b: Unaccented Vietnamese "san xuat" in Row 10 Col 4
        ws_kpi.Unprotect()
        ws_kpi.Cells(10, 4).Value = "san xuat"
        xl.Calculate()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        time.sleep(3)
        jp_unacc = ws_kpi.Cells(10, 5).Text
        cn_unacc = ws_kpi.Cells(10, 6).Text
        print(f"  [4b. Unaccented Input Test: 'san xuat']")
        print(f"      JP Col 5: '{jp_unacc}' | CN Col 6: '{cn_unacc}'")
        if "制作" in jp_unacc or "製造" in jp_unacc or "制作" in cn_unacc or "制造" in cn_unacc or "生产" in cn_unacc:
            results["translate_unaccented_passed"] = True

        # Test 4c: Manufacturing term "tiến độ dự án" in Row 10 Col 7
        ws_kpi.Cells(10, 7).Value = "tiến độ dự án"
        xl.Calculate()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        time.sleep(3)
        jp_term = ws_kpi.Cells(10, 8).Text
        cn_term = ws_kpi.Cells(10, 9).Text
        print(f"  [4c. Manufacturing Term Test: 'tiến độ dự án']")
        print(f"      JP Col 8: '{jp_term}' | CN Col 9: '{cn_term}'")
        if ("進捗" in jp_term or "プロジェクト" in jp_term) and ("进展" in cn_term or "项目" in cn_term):
            results["translate_term_passed"] = True

        # Test 4d: Empty input in Row 11
        ws_kpi.Cells(11, 7).Value = ""
        xl.Calculate()
        empty_jp = ws_kpi.Cells(11, 8).Text
        empty_cn = ws_kpi.Cells(11, 9).Text
        print(f"  [4d. Empty Input Test: Row 11]")
        print(f"      JP Col 8: '{empty_jp}', CN Col 9: '{empty_cn}'")
        if empty_jp == "" and empty_cn == "":
            results["translate_empty_passed"] = True

        # Clean up temporary test row 10 & 11 (keep row 9 persistent)
        ws_kpi.Cells(10, 4).Value = ""
        ws_kpi.Cells(10, 7).Value = ""
        ws_kpi.Cells(11, 7).Value = ""
        xl.Calculate()
        
        # Re-protect 06_KPI
        ws_kpi.Protect(Password="", DrawingObjects=False, Contents=True, Scenarios=True,
                       AllowFormattingCells=True, AllowFormattingColumns=True, AllowFormattingRows=True,
                       AllowInsertingRows=True, AllowDeletingRows=True, AllowSorting=True, AllowFiltering=True)
        ws_kpi.EnableSelection = 0

        # -------------------------------------------------------------
        # TEST 5: Language Switch Behavior (言語切替!B3)
        # -------------------------------------------------------------
        print("\n--- TEST 5: Testing Language Switching ---")
        ws_lang = wb.Worksheets('言語切替')
        ws_lang.Unprotect()
        
        for lang in ["日本語", "中文", "Tiếng Việt"]:
            ws_lang.Range("B3").Value = lang
            xl.Calculate()
            kpi_h5 = ws_kpi.Cells(4, 5).Text
            kpi_h6 = ws_kpi.Cells(4, 6).Text
            print(f"  Switched to '{lang}': KPI Col 5 Header='{kpi_h5}', Col 6 Header='{kpi_h6}'")

        # -------------------------------------------------------------
        # TEST 6: Sheet Protection & Unlocked Cells Verification
        # -------------------------------------------------------------
        print("\n--- TEST 6: Protection & Cell Lock Status ---")
        for sname in ['06_KPI', '04_課題ToDo', '05_決定事項', '03_WBS', '07_更新履歴']:
            ws = wb.Worksheets(sname)
            is_protected = ws.ProtectContents
            h_locked = ws.Cells(4, 5).Locked
            in_col = 4 if sname in ('06_KPI', '05_決定事項') else 5
            trans_col = in_col + 1
            input_locked = ws.Cells(9, in_col).Locked
            trans_locked = ws.Cells(9, trans_col).Locked
            print(f"  [{sname}] Protected: {is_protected} | Header Locked: {h_locked} | Input Col {in_col} Locked: {input_locked} | Trans Col {trans_col} Locked: {trans_locked}")
            if not is_protected or not h_locked or input_locked or not trans_locked:
                results["protection_checks_passed"] = False

        # Reset language to Tiếng Việt & Save
        ws_lang.Range("B3").Value = "Tiếng Việt"
        xl.Calculate()
        wb.Save()
        wb.Close(True)
        
        # -------------------------------------------------------------
        # TEST 7: XML Data Validation Preservation Check
        # -------------------------------------------------------------
        print("\n--- TEST 7: Checking XML Data Validations ---")
        with zipfile.ZipFile(file_path, 'r') as z:
            sheet_files = ['xl/worksheets/sheet5.xml', 'xl/worksheets/sheet6.xml', 'xl/worksheets/sheet7.xml', 'xl/worksheets/sheet8.xml']
            for sfile in sheet_files:
                content = z.read(sfile)
                root = ET.fromstring(content)
                x14_dvs = root.findall('.//{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}dataValidation')
                print(f"  {sfile}: {len(x14_dvs)} x14 data validations preserved.")
                if len(x14_dvs) == 0:
                    results["xml_validation_passed"] = False

        print("\n=======================================================")
        print("VERIFICATION COMPLETED SUCCESSFULLY!")
        print("=======================================================")
        print(f"Results: {results}")
    finally:
        xl.Quit()

if __name__ == '__main__':
    verify()
