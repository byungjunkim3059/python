import tkinter as tk
from tkinter import filedialog
import openpyxl

import pandas as pd

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls, *.xlsx")])

def is_cell_merged(workbook_path, sheet_name, target_cell_address):
    # 엑셀 파일 열기
    workbook = openpyxl.load_workbook(workbook_path)

    # 특정 시트 선택
    sheet = workbook[sheet_name]

    # 병합된 셀의 범위 확인
    merged_cells = sheet.merged_cells.ranges

    # 주어진 셀이 병합된 셀에 속하는지 여부 확인
    for merged_cell in merged_cells:
        if target_cell_address in merged_cell:
            # 주어진 셀이 병합된 셀에 속함
            workbook.close()
            return True

    # 주어진 셀이 병합된 셀에 속하지 않음
    workbook.close()
    return False

# 사용 예제
workbook_path = file_path  # 실제 파일 경로로 대체
sheet_name = "Sheet1"  # 시트 이름으로 대체
target_cell_address = "A7"  # 확인하고자 하는 셀의 주소로 대체

if is_cell_merged(workbook_path, sheet_name, target_cell_address):
    print(f"셀 {target_cell_address}은 병합된 셀입니다.")
else:
    print(f"셀 {target_cell_address}은 병합된 셀이 아닙니다.")