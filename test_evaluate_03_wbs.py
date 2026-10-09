import sys
import os
import time
import win32com.client
from PIL import ImageGrab

sys.stdout.reconfigure(encoding='utf-8')

file_path = os.path.abspath(r"d:\Sandbox\file_thu_ky\スマートファクトリー分科会_管理ツール案(修正版_20261002) .xlsx")

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True
wb = xl.Workbooks.Open(file_path)

ws_lang = wb.Worksheets("言語切替")
ws_wbs = wb.Worksheets("03_WBS")

for lang in ["日本語", "Tiếng Việt", "中文"]:
    ws_lang.Range("B3").Value = lang
    xl.CalculateUntilAsyncQueriesDone()
    xl.Calculate()
    time.sleep(1)
    
    print(f"\n=================== Language: {lang} ===================")
    print("Headers (Row 4):")
    for col in ["F", "G", "H", "I", "J", "K", "L", "M"]:
        print(f"  {col}4: {ws_wbs.Range(f'{col}4').Text}")
        
    print("\nData Row 8 (Tự động hóa cấp phát):")
    print(f"  F8 [Input]: {ws_wbs.Range('F8').Text}")
    print(f"  G8 (VI)   : {ws_wbs.Range('G8').Text}")
    print(f"  H8 (JP)   : {ws_wbs.Range('H8').Text}")
    print(f"  I8 (CN)   : {ws_wbs.Range('I8').Text}")
    
    print("\nData Row 10 (Cơ chế cấp phát):")
    print(f"  F10 [Input]: {ws_wbs.Range('F10').Text}")
    print(f"  G10 (VI)   : {ws_wbs.Range('G10').Text}")
    print(f"  H10 (JP)   : {ws_wbs.Range('H10').Text}")
    print(f"  I10 (CN)   : {ws_wbs.Range('I10').Text}")
    
    if lang == "日本語":
        rng = ws_wbs.Range("D1:N12")
        rng.CopyPicture(1, 2)
        time.sleep(0.5)
        img = ImageGrab.grabclipboard()
        if img:
            art_path = os.path.abspath(r"C:\Users\tvn183660\.gemini\antigravity\brain\7e67b3ac-1a49-4e1c-b681-79d667e82134\wbs_ghi_japanese_preview.png")
            img.save(art_path)
            print("Saved JP screenshot:", art_path)
            
    elif lang == "Tiếng Việt":
        rng = ws_wbs.Range("D1:N12")
        rng.CopyPicture(1, 2)
        time.sleep(0.5)
        img = ImageGrab.grabclipboard()
        if img:
            art_path = os.path.abspath(r"C:\Users\tvn183660\.gemini\antigravity\brain\7e67b3ac-1a49-4e1c-b681-79d667e82134\wbs_ghi_vietnamese_preview.png")
            img.save(art_path)
            print("Saved VI screenshot:", art_path)

# Set language back to Vietnamese (or as preferred)
ws_lang.Range("B3").Value = "Tiếng Việt"
xl.Calculate()
wb.Save()
wb.Close(False)
xl.Quit()
print("\nEvaluation test complete!")
