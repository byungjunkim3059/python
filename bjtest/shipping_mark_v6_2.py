import tkinter as tk
from tkinter import filedialog

import pandas as pd
import numpy as np
import os

def extract_row_index(target_df, target_str):
    row_indices = target_df[target_df.apply(lambda row: row.astype(str).str.contains(target_str).any(), axis=1)].index
    if len(row_indices) == 0:
        row_index = None
    else:
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


def extract_data_lst(df, col_str):
    item_row_index = extract_row_index(df, "ITEM")

    col_index = extract_col_index(df, col_str)
    data = df.iloc[item_row_index + 1 : item_row_index + 1 + len(item_data), col_index]
    
    lst = []
    for d in data:
        lst.append(d)

    if col_str == "PACKING":
        for i, d in enumerate(lst):
            if pd.notna(d):
                d = d.replace(" ", "")
                d = d.replace("\n", "")
                print(d)

                str1 = ""
                gh_str = ""
                if "(" in d:
                    start_index = d.find('(')
                    str1= d[:start_index]
                    gh_str = d[start_index:]
                else:
                    str1 = d
                # print(str1)
                # print(gh_str)

                ct_lst = str1.split('+')
                # print(ct_lst)

                plst = []
                for var in ct_lst:
                    x_index = var.find('x')

                    b_index = 0
                    if 'b' in var:
                        b_index = var.find('b')
                    elif 'C/T' in var:
                        b_index = var.find('C')
                    y_str = var[:x_index] + gh_str
                    print(y_str)
                    bale_num = int(var[x_index + 1:b_index])

                    for i in range(bale_num):
                        plst.append(y_str)

                lst = plst


    return lst

def remove_na_in_lst(lst):
    for i, data in enumerate(lst):
        if pd.isna(data):
            lst[i] = lst[i - 1]
    return lst

def remove_duplicates_in_lst(lst):
    lst2 = []
    for data in lst:
        if data not in lst2:
            lst2.append(data)
    lst = lst2
    return lst

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls *xlsx")])

if file_path:
    try:
        if file_path.endswith("xls"):
            df = pd.read_excel(file_path, engine="xlrd")
        elif file_path.endswith("xlsx"):
            df = pd.read_excel(file_path)

        item_row_index = extract_row_index(df, "ITEM")
        item_col_index = extract_col_index(df, "ITEM")
        column_data = df.iloc[item_row_index, item_col_index:]
        column_data = column_data.dropna()

        column_lst = []
        for data in column_data:
            column_lst.append(data)
        print(column_lst)

        item_data = df.iloc[item_row_index + 1:, item_col_index]
        item_data = item_data.dropna()

        print(item_data)
        print(len(item_data))


# 병합된 셀과 아닌 셀 각각 구하기
        packing_lst = extract_data_lst(df, "PACKING")
        
        merged_lst = []
        lst = []
        for i in range(len(packing_lst)):
            if pd.isna(packing_lst[i]):
                lst.append(i)
                if (i - 1) not in lst:
                    lst.append(i - 1)
                if (i == (len(packing_lst) - 1)):
                    lst.sort()
                    merged_lst.append(lst)
            else:
                if i == (len(packing_lst) - 1):
                    merged_lst.append([i])
                if (i < (len(packing_lst) - 1)) and pd.notna(packing_lst[i + 1]):
                    merged_lst.append([i])
                if len(lst) > 0:
                    lst.sort()
                    merged_lst.append(lst)
                    lst = []

        merged_lst.sort()
        # print(merged_lst)

        # print(packing_lst)
        # packing_lst = remove_na_in_lst(packing_lst)
        # packing_lst = remove_duplicates_in_lst(packing_lst)
        # print(packing_lst)

        def merged_info_adapt_data(df, col_str):
            new_lst = []
            for l in merged_lst:
                s_lst = []
                for j in l:
                    data_lst = extract_data_lst(df, col_str)
                    s_lst.append(data_lst[j])
                s_lst = remove_na_in_lst(s_lst)
                s_lst = remove_duplicates_in_lst(s_lst)
                
                new_lst.append(s_lst)

            return new_lst
        
        def merge_str(lst):
            new_str = ""
            for i, d in enumerate(lst):
                if i < (len(lst) - 1):
                    new_str += str(d) + " / "
                else:
                    new_str += str(d)

            return new_str
        
        for col in column_data:
            new_lst = merged_info_adapt_data(df, col)
            for l in new_lst:
                new_str = merge_str(l)
                # print(new_str)




    except Exception as e:
        print(f"파일 열기 오류: {str(e)}")
else:
    print("파일 선택이 취소되었습니다.")