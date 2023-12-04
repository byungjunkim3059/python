import xlrd
import tkinter as tk
from tkinter import filedialog
import openpyxl

import pandas as pd

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls")])

def get_merged_cells(file_path, sheet_name):
    # .xls 파일 열기
    workbook = xlrd.open_workbook(file_path)

    try:
        # 특정 시트 선택
        sheet = workbook[sheet_name]

        # 병합된 셀의 범위 확인
        merged_cells = []
        for merged_range in sheet.merged_cells:
            merged_cells.append({
                'start_row': merged_range[0],
                'end_row': merged_range[1],
                'start_col': merged_range[2],
                'end_col': merged_range[3]
            })

        return merged_cells
    except xlrd.XLRDError:
        # 시트가 존재하지 않는 경우
        return None
    finally:
        # 엑셀 파일 닫기
        workbook.release_resources()

# 사용 예제
xls_file_path = file_path  # 실제 .xls 파일 경로로 대체
target_sheet_name = "PL"  # 실제 시트 이름으로 대체

merged_cells = get_merged_cells(xls_file_path, target_sheet_name)

if merged_cells:
    print("병합된 셀이 있습니다:")
    for merged_cell in merged_cells:
        print(f"행: {merged_cell['start_row']} - {merged_cell['end_row']}, 열: {merged_cell['start_col']} - {merged_cell['end_col']}")
else:
    print("병합된 셀이 없습니다.")