# -*- coding: utf-8 -*-


import tkinter as tk
from tkinter import filedialog

import pandas as pd


# def extract_row_index(target_df, target_str):
#     row_indices = target_df[target_df.apply(lambda row: row.astype(str).str.contains(target_str).any(), axis=1)].index
#     row_index = row_indices[0]
#     return row_index

def extract_col_index(target_df, target_str):
    col_indices = [i for i, col in enumerate(target_df.columns) if target_df[col].astype(str).str.contains(target_str or "ITEM").any()]
    if len(col_indices) == 0:
        col_index = None
    else:
        col_index = col_indices[0]
    return col_index



root = tk.Tk()
root.withdraw()


file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xls *.xlsx")])


if file_paths:
    try:
        lst = []
        for file_path in file_paths:
            df = pd.read_excel(file_path)
            col_index = extract_col_index(df, "ITEM")
            if col_index != None:
                data = df.iloc[0:, col_index]
                data = data.dropna()
            # print(data)
            # print(file_path)
            for i in data:
                i = str(i)
                if (i not in lst) and ("MADE" not in i) :
                    lst.append(i)

        print(lst)
        print(len(lst))

    except Exception as e:
        print(f"파일 열기 오류: {str(e)}")
else:
    print("파일 선택이 취소되었습니다.")