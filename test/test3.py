import tkinter as tk
from tkinter import filedialog
import openpyxl

import pandas as pd

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls, *.xlsx")])


def get_merged_cells(workbook_path, sheet_name):
    # 엑셀 파일 열기
    workbook = openpyxl.load_workbook(workbook_path)

    # 특정 시트 선택
    sheet = workbook.active

    # 병합된 셀의 범위 확인
    merged_cells = sheet.merged_cells.ranges

    # 엑셀 파일 닫기
    workbook.close()

    return merged_cells

# 사용 예제
workbook_path = file_path  # 실제 파일 경로로 대체
sheet_name = "Sheet1"  # 시트 이름으로 대체

merged_cells = get_merged_cells(workbook_path, sheet_name)

if merged_cells:
    print("병합된 셀이 있습니다:")
    for merged_cell in merged_cells:
        print(f"병합된 범위: {merged_cell}")
else:
    print("병합된 셀이 없습니다.")