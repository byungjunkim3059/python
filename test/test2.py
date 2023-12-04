# 병합된 셀 바로 옆 셀 찾기

import tkinter as tk
from tkinter import filedialog

import pandas as pd

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls, *.xlsx")])

import openpyxl

def find_cell_with_value(workbook_path, sheet_name, target_value):
    # 엑셀 파일 열기
    workbook = openpyxl.load_workbook(workbook_path)

    # 특정 시트 선택
    sheet = workbook[sheet_name]

    # 엑셀 시트를 순회하며 특정 값이 있는 셀 찾기
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == target_value:
                cell_address = cell.coordinate
                print(f"'{target_value}'가 있는 셀의 주소: {cell_address}")

    # 엑셀 파일 닫기
    workbook.close()

# 사용 예제
workbook_path = file_path  # 실제 파일 경로로 대체
sheet_name = "Sheet1"  # 시트 이름으로 대체
target_value = "ITEM"  # 실제로 찾고자 하는 문자열로 대체

find_cell_with_value(workbook_path, sheet_name, target_value)