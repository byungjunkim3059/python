import tkinter as tk
from tkinter import filedialog
import xlrd
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook

import os

import pandas as pd
from datetime import datetime

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls")])



workbook = xlrd.open_workbook(file_path)

sheet = workbook.sheet_by_name("PL")

cell_value = sheet.cell_value(0, 0)

if "PACKING LIST" in cell_value:
    print("ok")

df = pd.read_excel(file_path, engine="xlrd")


def extract_row_index(target_df, target_str):
    row_indices = target_df[target_df.apply(lambda row: row.astype(str).str.contains(target_str).any(), axis=1)].index
    row_index = row_indices[0]
    return row_index

def extract_col_index(target_df, target_str):
    col_indices = [i for i, col in enumerate(target_df.columns) if target_df[col].astype(str).str.contains(target_str).any()]
    if len(col_indices) == 0:
        col_index = None
    else:
        col_index = col_indices[0]
    return col_index


def extract_customer_name_and_outbound_date(target_df):
    target_str_lst = ["수         신 :", "입   고   일:"]
    customer_name_and_outbound_date_lst = []

    for target_str in target_str_lst:
        row_index = extract_row_index(target_df, target_str)
        col_index = extract_col_index(target_df, target_str)
        result = target_df.iloc[row_index, col_index + 1]
        customer_name_and_outbound_date_lst.append(result)
    return customer_name_and_outbound_date_lst


customer_name_and_outbound_date_lst = extract_customer_name_and_outbound_date(df)


# 고객명
customer_info = customer_name_and_outbound_date_lst[0]

# 출고일
outbound_date_info = customer_name_and_outbound_date_lst[1]

# 쉽핑마크 문자열
sm_str_row_index = extract_row_index(df, "S/M")
sm_str_col_index = extract_col_index(df, "S/M")
sm_data = df.iloc[sm_str_row_index + 1:, sm_str_col_index]
sm_data = sm_data.dropna()

sm_lst = []
sm_str = ""
print(len(sm_data))
for data in sm_data:
    data = data.replace("\n", " ")
    if data not in sm_lst:
        sm_lst.append(data)

i = 1
for data in sm_lst:
    if i == len(sm_lst):
        sm_str += str(data)
    else:
        sm_str += str(data) + " / "
    i += 1

print(sm_str)


def remove_duplicates(data_series):
    lst = []
    strg = ""
    for data in data_series:
        data = data.replace("\n", " ")
        if data not in lst:
            lst.append(data)

    i = 1
    for data in lst:
        if i == len(lst):
            strg += str(data)
        else:
            strg += str(data) + " / "
        i += 1
    return strg

sm_str = remove_duplicates(sm_data)


# ==================================================================================

def get_day_of_week(day_of_week):
    
            days = ["월", "화", "수", "목", "금", "토", "일"]
            day_name = days[day_of_week]

            return day_name
        
year = outbound_date_info.year
month = outbound_date_info.month
day = outbound_date_info.day


date_object = datetime(year, month, (day - 1))
day_of_week = date_object.weekday()

day_name = get_day_of_week(day_of_week)

year_str = str(year)
if month < 10:
    month_str = "0" + str(month)
else:
    month_str = str(month)
if (day - 1) < 10:
    day_str = "0" + str(day - 1)
else:
    day_str = str(day - 1)

date_str = year_str + "-" + month_str + "-" + day_str
date_str2 = year_str + month_str + day_str

folder_name = customer_info + "_" + date_str2
# ==================================================================================


path = "C:\\Users\\bnj30\Desktop\\출고서류\\" + year_str + "년\\" + month_str + "월\\" + date_str + "(" + day_name + ")\\" + folder_name

try:
    # 폴더 생성
    if os.path.exists(path):
        user_response = input(f"폴더가 이미 존재합니다. 덮어쓰시겠습니까? (Y/N): ")
        if user_response.lower() == 'y':
            i = 1
    else:
        os.makedirs(path)
        print(f"폴더가 성공적으로 생성되었습니다. 경로: {path}")
except FileExistsError:
    print(f"폴더가 이미 존재합니다. 경로: {path}")
except Exception as e:
    print(f"폴더 생성 중 오류가 발생했습니다. 오류 내용: {e}")
