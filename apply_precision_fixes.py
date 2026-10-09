import sys
import win32com.client
import openpyxl.utils
import time
import shutil

sys.stdout.reconfigure(encoding='utf-8')

def col_letter(c):
    return openpyxl.utils.get_column_letter(c)

def com_retry(func, retries=30, delay=0.5):
    for attempt in range(retries):
        try:
            return func()
        except Exception as e:
            err_str = str(e).lower()
            if "-2147418111" in err_str or "rejected" in err_str or "busy" in err_str or "-2147417848" in err_str:
                time.sleep(delay)
            else:
                raise
    return func()

def set_formula2(target, val):
    com_retry(lambda: setattr(target, 'Formula2', val))

def set_value(target, val):
    com_retry(lambda: setattr(target, 'Value', val))

def fix_translations(file_path):
    print(f"\n=======================================================")
    print(f"Applying Precision Fixes to: {file_path}")
    print(f"=======================================================")
    
    xl = win32com.client.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    xl.EnableEvents = False
    
    try:
        wb = xl.Workbooks.Open(file_path)
        try:
            xl.Calculation = -4135 # xlCalculationManual
        except Exception:
            pass
            
        target_sheets = ['03_WBS', '04_課題ToDo', '05_決定事項', '06_KPI', '07_更新履歴']
        for sname in target_sheets:
            try:
                com_retry(lambda s=sname: wb.Worksheets(s).Unprotect())
            except Exception as e:
                print(f"Unprotect {sname}: {e}")

        # ---------------------------------------------------------
        # 1. Sheet 06_KPI
        # ---------------------------------------------------------
        print("\n--- 1. Updating 06_KPI formulas to use explicit 'vi' source ---")
        ws_kpi = wb.Worksheets('06_KPI')
        kpi_trans = [
            (4, 5, 6),   # Name (VI, JP, CN)
            (7, 8, 9),   # Purpose (VI, JP, CN)
            (10, 11, 12),# Formula (VI, JP, CN)
            (23, 24, 25),# Action (VI, JP, CN)
            (26, 27, 28) # Note (VI, JP, CN)
        ]
        max_kpi_row = 204
        for vi_c, jp_c, cn_c in kpi_trans:
            jp_formulas = []
            cn_formulas = []
            for r in range(9, max_kpi_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
            set_formula2(ws_kpi.Range(f"{col_letter(jp_c)}9:{col_letter(jp_c)}{max_kpi_row}"), jp_formulas)
            set_formula2(ws_kpi.Range(f"{col_letter(cn_c)}9:{col_letter(cn_c)}{max_kpi_row}"), cn_formulas)

        # Populate user test case in Row 9 Purpose column
        print("Populating test case in 06_KPI Row 9 (Col 7 / G9): 'tôi tên là'")
        set_value(ws_kpi.Cells(9, 7), "tôi tên là")

        # ---------------------------------------------------------
        # 2. Sheet 04_課題ToDo
        # ---------------------------------------------------------
        print("\n--- 2. Updating 04_課題ToDo formulas to use explicit 'vi' source ---")
        ws_todo = wb.Worksheets('04_課題ToDo')
        todo_trans = [
            (5, 6, 7),   # Title
            (8, 9, 10),  # BG
            (11, 12, 13),# Action
            (21, 22, 23),# Cond
        ]
        max_todo_row = 304
        for vi_c, jp_c, cn_c in todo_trans:
            jp_formulas = []
            cn_formulas = []
            for r in range(9, max_todo_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
            set_formula2(ws_todo.Range(f"{col_letter(jp_c)}9:{col_letter(jp_c)}{max_todo_row}"), jp_formulas)
            set_formula2(ws_todo.Range(f"{col_letter(cn_c)}9:{col_letter(cn_c)}{max_todo_row}"), cn_formulas)

        # For Result (Col 24, 25, 26) - rows 5 to 304
        jp_res_formulas = []
        cn_res_formulas = []
        for r in range(5, max_todo_row + 1):
            vi_addr = f"X{r}"
            jp_res_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
            cn_res_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
        set_formula2(ws_todo.Range(f"Y5:Y{max_todo_row}"), jp_res_formulas)
        set_formula2(ws_todo.Range(f"Z5:Z{max_todo_row}"), cn_res_formulas)

        # ---------------------------------------------------------
        # 3. Sheet 05_決定事項
        # ---------------------------------------------------------
        print("\n--- 3. Updating 05_決定事項 formulas to use explicit 'vi' source ---")
        ws_dec = wb.Worksheets('05_決定事項')
        dec_trans = [
            (4, 5, 6),   # Title
            (7, 8, 9),   # Content
            (10, 11, 12),# Reason
            (16, 17, 18),# Scope
            (23, 24, 25) # Note
        ]
        max_dec_row = 204
        for vi_c, jp_c, cn_c in dec_trans:
            jp_formulas = []
            cn_formulas = []
            for r in range(6, max_dec_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
            set_formula2(ws_dec.Range(f"{col_letter(jp_c)}6:{col_letter(jp_c)}{max_dec_row}"), jp_formulas)
            set_formula2(ws_dec.Range(f"{col_letter(cn_c)}6:{col_letter(cn_c)}{max_dec_row}"), cn_formulas)

        # ---------------------------------------------------------
        # 4. Sheet 03_WBS
        # ---------------------------------------------------------
        print("\n--- 4. Updating 03_WBS formulas to use explicit 'vi' source ---")
        ws_wbs = wb.Worksheets('03_WBS')
        wbs_trans = [
            (5, 6, 7),   # Task
            (18, 19, 20) # Note
        ]
        max_wbs_row = 204
        for vi_c, jp_c, cn_c in wbs_trans:
            jp_formulas = []
            cn_formulas = []
            for r in range(8, max_wbs_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
            set_formula2(ws_wbs.Range(f"{col_letter(jp_c)}8:{col_letter(jp_c)}{max_wbs_row}"), jp_formulas)
            set_formula2(ws_wbs.Range(f"{col_letter(cn_c)}8:{col_letter(cn_c)}{max_wbs_row}"), cn_formulas)

        # ---------------------------------------------------------
        # 5. Sheet 07_更新履歴
        # ---------------------------------------------------------
        print("\n--- 5. Updating 07_更新履歴 formulas to use explicit 'vi' source ---")
        ws_rev = wb.Worksheets('07_更新履歴')
        rev_trans = [
            (5, 6, 7),   # Content
            (8, 9, 10),  # Reason
            (12, 13, 14) # Note
        ]
        max_rev_row = 104
        for vi_c, jp_c, cn_c in rev_trans:
            jp_formulas = []
            cn_formulas = []
            for r in range(6, max_rev_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
            set_formula2(ws_rev.Range(f"{col_letter(jp_c)}6:{col_letter(jp_c)}{max_rev_row}"), jp_formulas)
            set_formula2(ws_rev.Range(f"{col_letter(cn_c)}6:{col_letter(cn_c)}{max_rev_row}"), cn_formulas)

        # ---------------------------------------------------------
        # 6. Update _翻訳マスタ guidance text (Row 249)
        # ---------------------------------------------------------
        print("\n--- 6. Updating _翻訳マスタ ---")
        ws_m = wb.Worksheets('_翻訳マスタ')
        set_value(ws_m.Range("B249"), "自由記入欄（ベトナム語）に入力すると、隣接するJP/CN列がTRANSLATE関数で自動翻訳される（M365・ネット要）。")
        set_value(ws_m.Range("C249"), "越南语列正常输入。相邻的JP/CN列用TRANSLATE自动翻译（需M365和网络）。")
        set_value(ws_m.Range("D249"), "Cột tiếng Việt nhập liệu bình thường. Các cột JP/CN kề cạnh tự động dịch bằng hàm TRANSLATE (cần M365 + mạng).")

        # ---------------------------------------------------------
        # 7. Configure Sheet Protection
        # ---------------------------------------------------------
        print("\n--- 7. Re-applying Protection & Unlocked Cells ---")
        sheet_configs = [
            (ws_kpi, 204, 28, [5, 6, 8, 9, 11, 12, 20, 21, 24, 25, 27, 28]),
            (ws_todo, 304, 30, [6, 7, 9, 10, 12, 13, 22, 23, 25, 26, 28, 29]),
            (ws_dec, 204, 25, [5, 6, 8, 9, 11, 12, 17, 18, 24, 25]),
            (ws_wbs, 204, 20, [6, 7, 13, 15, 16, 19, 20]),
            (ws_rev, 104, 14, [6, 7, 9, 10, 13, 14]),
        ]
        
        for w, max_r, max_c, locked_cols in sheet_configs:
            com_retry(lambda w=w, c=max_c: setattr(w.Range(w.Cells(1, 1), w.Cells(4, c)), 'Locked', True))
            com_retry(lambda w=w, r=max_r, c=max_c: setattr(w.Range(w.Cells(5, 1), w.Cells(r, c)), 'Locked', False))
            for lc in locked_cols:
                com_retry(lambda w=w, r=max_r, col=lc: setattr(w.Range(w.Cells(5, col), w.Cells(r, col)), 'Locked', True))
                
            com_retry(lambda w=w: w.Protect(
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
            ))
            com_retry(lambda w=w: setattr(w, 'EnableSelection', 0)) # xlNoRestrictions

        # ---------------------------------------------------------
        # 8. Set Freeze Panes
        # ---------------------------------------------------------
        print("\n--- 8. Setting Freeze Panes ---")
        freeze_map = [
            (ws_kpi, "D5"),   # KPI: Freeze Cols A-C (ID, WG, Level)
            (ws_todo, "E5"),  # ToDo: Freeze Cols A-D (ID, Type, WG, Site)
            (ws_dec, "D5"),   # Dec: Freeze Cols A-C (ID, Date, Forum)
            (ws_wbs, "E5"),   # WBS: Freeze Cols A-D (ID, WG, Site, Category)
            (ws_rev, "E5"),   # Rev: Freeze Cols A-D (No, Date, Author, Target)
        ]
        for w, freeze_cell in freeze_map:
            com_retry(lambda w=w: w.Activate())
            com_retry(lambda: setattr(xl.ActiveWindow, 'FreezePanes', False))
            com_retry(lambda w=w, f=freeze_cell: w.Range(f).Select())
            com_retry(lambda: setattr(xl.ActiveWindow, 'FreezePanes', True))

        com_retry(lambda: wb.Worksheets('言語切替').Activate())

        xl.EnableEvents = True
        xl.ScreenUpdating = True
        try:
            xl.Calculation = -4105 # xlCalculationAutomatic
        except Exception:
            pass

        # ---------------------------------------------------------
        # 9. Save and Calculate
        # ---------------------------------------------------------
        print("\n--- 9. Saving & Calculating ---")
        com_retry(lambda: wb.Save())
        
        xl.Calculate()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception as e:
            print("CalculateUntilAsyncQueriesDone notice:", e)
            
        time.sleep(3)
        com_retry(lambda: wb.Save())
        com_retry(lambda: wb.Close(True))
        print("Precision fixes successfully applied and saved!")
    finally:
        xl.Application.Quit()

if __name__ == '__main__':
    prod_file = r'd:\Sandbox\file_thu_ky\smart_factory_3ngu_v3.xlsx'
    fix_translations(prod_file)
