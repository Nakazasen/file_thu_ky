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

def insert_col(ws, col_idx):
    com_retry(lambda: ws.Columns(col_idx).Insert())

def delete_range(ws, range_str):
    com_retry(lambda: ws.Range(range_str).Delete())

def transform_workbook(file_path):
    print(f"\n=======================================================")
    print(f"Applying Option A (Side-by-Side) to: {file_path}")
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
                
        # =========================================================================
        # 1. SHEET 06_KPI
        # =========================================================================
        print("\n--- 1. Restructuring 06_KPI ---")
        ws_kpi = wb.Worksheets('06_KPI')
        com_retry(lambda: ws_kpi.Unprotect())
        
        kpi_specs = [
            (18, "KPI_COL_NOTE", ["COMMON_EXAMPLE"]*4),
            (17, "KPI_COL_ACTION", ["KPI_EX1_ACTION", "KPI_EX2_ACTION", "KPI_EX3_ACTION", "KPI_EX4_ACTION"]),
            (6, "KPI_COL_FORMULA", ["KPI_EX1_FORMULA", "KPI_EX2_FORMULA", "KPI_EX3_FORMULA", "KPI_EX4_FORMULA"]),
            (5, "KPI_COL_PURPOSE", ["KPI_EX1_PURPOSE", "KPI_EX2_PURPOSE", "KPI_EX3_PURPOSE", "KPI_EX4_PURPOSE"]),
            (4, "KPI_COL_NAME", ["KPI_EX1_NAME", "KPI_EX2_NAME", "KPI_EX3_NAME", "KPI_EX4_NAME"]),
        ]
        
        for orig_c, key, sample_keys in kpi_specs:
            insert_col(ws_kpi, orig_c + 1)
            insert_col(ws_kpi, orig_c + 1)
            
        time.sleep(0.5)
        delete_range(ws_kpi, f"{col_letter(29)}:{col_letter(38)}")
        time.sleep(0.5)
        print("06_KPI columns count:", ws_kpi.UsedRange.Columns.Count)
        
        new_kpi_trans = [
            (4, 5, 6, "KPI_COL_NAME", ["KPI_EX1_NAME", "KPI_EX2_NAME", "KPI_EX3_NAME", "KPI_EX4_NAME"]),
            (7, 8, 9, "KPI_COL_PURPOSE", ["KPI_EX1_PURPOSE", "KPI_EX2_PURPOSE", "KPI_EX3_PURPOSE", "KPI_EX4_PURPOSE"]),
            (10, 11, 12, "KPI_COL_FORMULA", ["KPI_EX1_FORMULA", "KPI_EX2_FORMULA", "KPI_EX3_FORMULA", "KPI_EX4_FORMULA"]),
            (23, 24, 25, "KPI_COL_ACTION", ["KPI_EX1_ACTION", "KPI_EX2_ACTION", "KPI_EX3_ACTION", "KPI_EX4_ACTION"]),
            (26, 27, 28, "KPI_COL_NOTE", ["COMMON_EXAMPLE"]*4),
        ]
        
        max_kpi_row = 204
        for vi_c, jp_c, cn_c, key, sample_keys in new_kpi_trans:
            set_formula2(ws_kpi.Cells(4, vi_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_VI_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_kpi.Cells(4, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_JP_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_kpi.Cells(4, cn_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_CN_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            
            for i, ex_k in enumerate(sample_keys):
                r = 5 + i
                set_formula2(ws_kpi.Cells(r, vi_c), f'=INDEX(_翻訳マスタ!$D$2:$D$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
                set_formula2(ws_kpi.Cells(r, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$B$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
                set_formula2(ws_kpi.Cells(r, cn_c), f'=INDEX(_翻訳マスタ!$C$2:$C$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
                
            jp_formulas = []
            cn_formulas = []
            for r in range(9, max_kpi_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
                
            set_formula2(ws_kpi.Range(f"{col_letter(jp_c)}9:{col_letter(jp_c)}{max_kpi_row}"), jp_formulas)
            set_formula2(ws_kpi.Range(f"{col_letter(cn_c)}9:{col_letter(cn_c)}{max_kpi_row}"), cn_formulas)

        set_value(ws_kpi.Cells(9, 7), "tôi tên là")

        time.sleep(1.0)

        # =========================================================================
        # 2. SHEET 04_課題ToDo
        # =========================================================================
        print("\n--- 2. Restructuring 04_課題ToDo ---")
        ws_todo = wb.Worksheets('04_課題ToDo')
        com_retry(lambda: ws_todo.Unprotect())
        
        todo_specs = [
            (16, "ISSUE_COL_RESULT", [None]*4),
            (15, "ISSUE_COL_COND", ["ISSUE_EX1_COND", "ISSUE_EX2_COND", "ISSUE_EX3_COND", "ISSUE_EX4_COND"]),
            (7, "ISSUE_COL_ACTION", ["ISSUE_EX1_ACT", "ISSUE_EX2_ACT", "ISSUE_EX3_ACT", "ISSUE_EX4_ACT"]),
            (6, "ISSUE_COL_BG", ["ISSUE_EX1_BG", "ISSUE_EX2_BG", "ISSUE_EX3_BG", "ISSUE_EX4_BG"]),
            (5, "ISSUE_COL_TITLE", ["ISSUE_EX1_TITLE", "ISSUE_EX2_TITLE", "ISSUE_EX3_TITLE", "ISSUE_EX4_TITLE"]),
        ]
        
        for orig_c, key, sample_keys in todo_specs:
            insert_col(ws_todo, orig_c + 1)
            insert_col(ws_todo, orig_c + 1)
            
        time.sleep(0.5)
        delete_range(ws_todo, f"{col_letter(31)}:{col_letter(40)}")
        time.sleep(0.5)
        print("04_課題ToDo columns count:", ws_todo.UsedRange.Columns.Count)
        
        new_todo_trans = [
            (5, 6, 7, "ISSUE_COL_TITLE", ["ISSUE_EX1_TITLE", "ISSUE_EX2_TITLE", "ISSUE_EX3_TITLE", "ISSUE_EX4_TITLE"]),
            (8, 9, 10, "ISSUE_COL_BG", ["ISSUE_EX1_BG", "ISSUE_EX2_BG", "ISSUE_EX3_BG", "ISSUE_EX4_BG"]),
            (11, 12, 13, "ISSUE_COL_ACTION", ["ISSUE_EX1_ACT", "ISSUE_EX2_ACT", "ISSUE_EX3_ACT", "ISSUE_EX4_ACT"]),
            (21, 22, 23, "ISSUE_COL_COND", ["ISSUE_EX1_COND", "ISSUE_EX2_COND", "ISSUE_EX3_COND", "ISSUE_EX4_COND"]),
            (24, 25, 26, "ISSUE_COL_RESULT", [None]*4),
        ]
        
        max_todo_row = 304
        for vi_c, jp_c, cn_c, key, sample_keys in new_todo_trans:
            set_formula2(ws_todo.Cells(4, vi_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_VI_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_todo.Cells(4, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_JP_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_todo.Cells(4, cn_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_CN_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            
            for i, ex_k in enumerate(sample_keys):
                r = 5 + i
                if ex_k:
                    set_formula2(ws_todo.Cells(r, vi_c), f'=INDEX(_翻訳マスタ!$D$2:$D$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
                    set_formula2(ws_todo.Cells(r, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$B$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
                    set_formula2(ws_todo.Cells(r, cn_c), f'=INDEX(_翻訳マスタ!$C$2:$C$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
                else:
                    vi_addr = f"{col_letter(vi_c)}{r}"
                    set_formula2(ws_todo.Cells(r, jp_c), f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))')
                    set_formula2(ws_todo.Cells(r, cn_c), f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))')
                    
            jp_formulas = []
            cn_formulas = []
            for r in range(9, max_todo_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
                
            set_formula2(ws_todo.Range(f"{col_letter(jp_c)}9:{col_letter(jp_c)}{max_todo_row}"), jp_formulas)
            set_formula2(ws_todo.Range(f"{col_letter(cn_c)}9:{col_letter(cn_c)}{max_todo_row}"), cn_formulas)

        time.sleep(1.0)

        # =========================================================================
        # 3. SHEET 05_決定事項
        # =========================================================================
        print("\n--- 3. Restructuring 05_決定事項 ---")
        ws_dec = wb.Worksheets('05_決定事項')
        com_retry(lambda: ws_dec.Unprotect())
        
        dec_specs = [
            (15, "DEC_COL_NOTE", ["COMMON_EXAMPLE"]),
            (10, "DEC_COL_SCOPE", ["COMMON_ALL_WG"]),
            (6, "DEC_COL_REASON", ["DEC_EX1_REASON"]),
            (5, "DEC_COL_CONTENT", ["DEC_EX1_CONTENT"]),
            (4, "DEC_COL_TITLE", ["DEC_EX1_TITLE"]),
        ]
        
        for orig_c, key, sample_keys in dec_specs:
            insert_col(ws_dec, orig_c + 1)
            insert_col(ws_dec, orig_c + 1)
            
        time.sleep(0.5)
        delete_range(ws_dec, f"{col_letter(26)}:{col_letter(35)}")
        time.sleep(0.5)
        print("05_決定事項 columns count:", ws_dec.UsedRange.Columns.Count)
        
        new_dec_trans = [
            (4, 5, 6, "DEC_COL_TITLE", ["DEC_EX1_TITLE"]),
            (7, 8, 9, "DEC_COL_CONTENT", ["DEC_EX1_CONTENT"]),
            (10, 11, 12, "DEC_COL_REASON", ["DEC_EX1_REASON"]),
            (16, 17, 18, "DEC_COL_SCOPE", ["COMMON_ALL_WG"]),
            (23, 24, 25, "DEC_COL_NOTE", ["COMMON_EXAMPLE"]),
        ]
        
        max_dec_row = 204
        for vi_c, jp_c, cn_c, key, sample_keys in new_dec_trans:
            set_formula2(ws_dec.Cells(4, vi_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_VI_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_dec.Cells(4, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_JP_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_dec.Cells(4, cn_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_CN_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            
            ex_k = sample_keys[0]
            set_formula2(ws_dec.Cells(5, vi_c), f'=INDEX(_翻訳マスタ!$D$2:$D$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
            set_formula2(ws_dec.Cells(5, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$B$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
            set_formula2(ws_dec.Cells(5, cn_c), f'=INDEX(_翻訳マスタ!$C$2:$C$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
            
            jp_formulas = []
            cn_formulas = []
            for r in range(6, max_dec_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
                
            set_formula2(ws_dec.Range(f"{col_letter(jp_c)}6:{col_letter(jp_c)}{max_dec_row}"), jp_formulas)
            set_formula2(ws_dec.Range(f"{col_letter(cn_c)}6:{col_letter(cn_c)}{max_dec_row}"), cn_formulas)

        time.sleep(1.0)

        # =========================================================================
        # 4. SHEET 03_WBS
        # =========================================================================
        print("\n--- 4. Restructuring 03_WBS ---")
        ws_wbs = wb.Worksheets('03_WBS')
        com_retry(lambda: ws_wbs.Unprotect())
        
        wbs_specs = [
            (16, "WBS_NOTE", ["COMMON_EXAMPLE"]*3),
            (5, "WBS_TASK", ["WBS_TASK_1", "WBS_TASK_2", "WBS_TASK_3"]),
        ]
        
        for orig_c, key, sample_keys in wbs_specs:
            insert_col(ws_wbs, orig_c + 1)
            insert_col(ws_wbs, orig_c + 1)
            
        time.sleep(0.5)
        delete_range(ws_wbs, f"{col_letter(21)}:{col_letter(24)}")
        time.sleep(0.5)
        print("03_WBS columns count:", ws_wbs.UsedRange.Columns.Count)
        
        new_wbs_trans = [
            (5, 6, 7, "WBS_TASK", ["WBS_TASK_1", "WBS_TASK_2", "WBS_TASK_3"]),
            (18, 19, 20, "WBS_NOTE", ["COMMON_EXAMPLE"]*3),
        ]
        
        max_wbs_row = 204
        for vi_c, jp_c, cn_c, key, sample_keys in new_wbs_trans:
            set_formula2(ws_wbs.Cells(4, vi_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_VI_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_wbs.Cells(4, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_JP_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_wbs.Cells(4, cn_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_CN_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            
            for i, ex_k in enumerate(sample_keys):
                r = 5 + i
                set_formula2(ws_wbs.Cells(r, vi_c), f'=INDEX(_翻訳マスタ!$D$2:$D$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
                set_formula2(ws_wbs.Cells(r, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$B$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
                set_formula2(ws_wbs.Cells(r, cn_c), f'=INDEX(_翻訳マスタ!$C$2:$C$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
                
            jp_formulas = []
            cn_formulas = []
            for r in range(8, max_wbs_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
                
            set_formula2(ws_wbs.Range(f"{col_letter(jp_c)}8:{col_letter(jp_c)}{max_wbs_row}"), jp_formulas)
            set_formula2(ws_wbs.Range(f"{col_letter(cn_c)}8:{col_letter(cn_c)}{max_wbs_row}"), cn_formulas)

        time.sleep(1.0)

        # =========================================================================
        # 5. SHEET 07_更新履歴
        # =========================================================================
        print("\n--- 5. Restructuring 07_更新履歴 ---")
        ws_rev = wb.Worksheets('07_更新履歴')
        com_retry(lambda: ws_rev.Unprotect())
        
        rev_specs = [
            (8, "REV_COL_NOTE", ["COMMON_EXAMPLE"]),
            (6, "REV_COL_REASON", ["REV_EX1_REASON"]),
            (5, "REV_COL_CONTENT", ["REV_EX1_CONTENT"]),
        ]
        
        for orig_c, key, sample_keys in rev_specs:
            insert_col(ws_rev, orig_c + 1)
            insert_col(ws_rev, orig_c + 1)
            
        time.sleep(0.5)
        delete_range(ws_rev, f"{col_letter(15)}:{col_letter(20)}")
        time.sleep(0.5)
        print("07_更新履歴 columns count:", ws_rev.UsedRange.Columns.Count)
        
        new_rev_trans = [
            (5, 6, 7, "REV_COL_CONTENT", ["REV_EX1_CONTENT"]),
            (8, 9, 10, "REV_COL_REASON", ["REV_EX1_REASON"]),
            (12, 13, 14, "REV_COL_NOTE", ["COMMON_EXAMPLE"]),
        ]
        
        max_rev_row = 104
        for vi_c, jp_c, cn_c, key, sample_keys in new_rev_trans:
            set_formula2(ws_rev.Cells(4, vi_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_VI_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_rev.Cells(4, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_JP_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            set_formula2(ws_rev.Cells(4, cn_c), f'=INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("{key}",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))&INDEX(_翻訳マスタ!$B$2:$D$300,MATCH("NEW_CN_SUFFIX",_翻訳マスタ!$A$2:$A$300,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))')
            
            ex_k = sample_keys[0]
            set_formula2(ws_rev.Cells(5, vi_c), f'=INDEX(_翻訳マスタ!$D$2:$D$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
            set_formula2(ws_rev.Cells(5, jp_c), f'=INDEX(_翻訳マスタ!$B$2:$B$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
            set_formula2(ws_rev.Cells(5, cn_c), f'=INDEX(_翻訳マスタ!$C$2:$C$300,MATCH("{ex_k}",_翻訳マスタ!$A$2:$A$300,0))')
            
            jp_formulas = []
            cn_formulas = []
            for r in range(6, max_rev_row + 1):
                vi_addr = f"{col_letter(vi_c)}{r}"
                jp_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","ja"))'])
                cn_formulas.append([f'=IF({vi_addr}="","",TRANSLATE({vi_addr},"vi","zh-Hans"))'])
                
            set_formula2(ws_rev.Range(f"{col_letter(jp_c)}6:{col_letter(jp_c)}{max_rev_row}"), jp_formulas)
            set_formula2(ws_rev.Range(f"{col_letter(cn_c)}6:{col_letter(cn_c)}{max_rev_row}"), cn_formulas)

        time.sleep(1.0)

        # =========================================================================
        # 6. UPDATE _翻訳マスタ (Row 249 PARA_DESC)
        # =========================================================================
        print("\n--- 6. Updating _翻訳マスタ ---")
        ws_m = wb.Worksheets('_翻訳マスタ')
        set_value(ws_m.Range("B249"), "自由記入欄（ベトナム語）に入力すると、隣接するJP/CN列がTRANSLATE関数で自動翻訳される（M365・ネット要）。")
        set_value(ws_m.Range("C249"), "越南语列正常输入。相邻的JP/CN列用TRANSLATE自动翻译（需M365和网络）。")
        set_value(ws_m.Range("D249"), "Cột tiếng Việt nhập liệu bình thường. Các cột JP/CN kề cạnh tự động dịch bằng hàm TRANSLATE (cần M365 + mạng).")

        time.sleep(1.0)

        # =========================================================================
        # 7. CONFIGURE PROTECTION & UNLOCKED INPUT CELLS
        # =========================================================================
        print("\n--- 7. Configuring Protection & Unlocked Cells ---")
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

        time.sleep(1.0)

        # =========================================================================
        # 8. CONFIGURE FREEZE PANES
        # =========================================================================
        print("\n--- 8. Setting Freeze Panes ---")
        freeze_map = [
            (ws_kpi, "D5"),   # KPI: Freeze ID, WG, Level (Cols A-C)
            (ws_todo, "E5"),  # ToDo: Freeze ID, Type, WG, Site (Cols A-D)
            (ws_dec, "D5"),   # Dec: Freeze ID, Date, Forum (Cols A-C)
            (ws_wbs, "E5"),   # WBS: Freeze ID, WG, Site, Category (Cols A-D)
            (ws_rev, "E5"),   # Rev: Freeze No, Date, Author, Target (Cols A-D)
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

        # =========================================================================
        # 9. SAVE WORKBOOK & TEST CALCULATION
        # =========================================================================
        print("\n--- 9. Saving & Calculating ---")
        com_retry(lambda: wb.Save())
        print("Workbook saved successfully!")
        
        xl.Calculate()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception as e:
            print("CalculateUntilAsyncQueriesDone:", e)
            
        com_retry(lambda: wb.Close(True))
        print("All sheets successfully transformed to Option A Side-by-Side!")
    finally:
        xl.Application.Quit()

if __name__ == '__main__':
    prod_file = r'd:\Sandbox\file_thu_ky\smart_factory_3ngu_v3.xlsx'
    pre_backup = r'd:\Sandbox\file_thu_ky\smart_factory_3ngu_v3_backup_pre_option_a.xlsx'
    print(f"Creating safety backup to: {pre_backup}")
    shutil.copyfile(prod_file, pre_backup)
    transform_workbook(prod_file)
