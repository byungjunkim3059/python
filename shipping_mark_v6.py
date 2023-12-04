import tkinter as tk
from tkinter import filedialog

import pandas as pd


# ================================================================================================================================================ 
# 1. 고객명, 출고일(입고일 - 1), 주문 정보 데이터프레임 도출
# 2. 병합된 행(row) 주소값 구하고 병합된 셀들에 값 입력해주기, 병합되지 않은 행(row) 주소값도 구하기
# 3. 




# ================================================================================================================================================
# 함수 zone
# 1.

def extract_row_index(target_df, target_str):
    row_indices = target_df[target_df.apply(lambda row: row.astype(str).str.contains(target_str).any(), axis=1)].index
    if len(row_indices) == 0:
        col_index = None
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


def extract_order_info_and_create_new_df(target_df):

    new_df_columns = ["BUYER", "ITEM", "COLOR", "STYLE NO", "PACKING", "Bale No.", "S/M"]
    new_df_data = []

    # 최초 행(row) 수 구하기
    item_row_index = extract_row_index(target_df, "ITEM")
    item_col_index = extract_col_index(target_df, "ITEM")

    item_data = df.iloc[item_row_index + 1:, item_col_index]
    item_data = item_data.dropna()

    row_length = len(item_data)
    row_start_index = item_row_index + 1

    for row_num in range(row_length):
        lst = []
        for column in new_df_columns:
            col_index = extract_col_index(target_df, column)
            if col_index == None:
                new_df_columns.remove(column)
                continue
            data = target_df.iloc[row_start_index + row_num, col_index]
            lst.append(data)
        new_df_data.append(lst)
    
    new_df = pd.DataFrame(columns=new_df_columns, data=new_df_data)
    return new_df


# 2.

def extract_merged_row_address_and_fill_nan_cells(target_df):
    merged_row_address_lst = []
    
    for i in range(len(target_df["PACKING"])):
        if pd.isna(target_df["PACKING"][i]):
            target_df["PACKING"][i] = target_df["PACKING"][i - 1]
            target_df["Bale No."][i] = target_df["Bale No."][i - 1]
            merged_row_address_lst.append(i)

    return merged_row_address_lst


def extract_not_merged_row_address(target_df, target_merged_row_address_lst):
    not_merged_row_address_lst = []
    row_length = len(target_df["ITEM"])

    for i in range(row_length):
        for j in target_merged_row_address_lst:
            if i != j:
                not_merged_row_address_lst.append(i)

    return not_merged_row_address_lst






# ================================================================================================================================================
# 파일 열기 (여러 개 가능)

root = tk.Tk()
root.withdraw()

packing_list_initialdir = "C:\\Users\\bnj30\\Desktop\\10월 출고서류"

file_paths = filedialog.askopenfilenames(title="패킹리스트 파일 선택", initialdir=packing_list_initialdir, filetypes=[("Excel Files", "*.xls")])
# print(file_path)

# ================================================================================================================================================
if file_paths:
    try:
        for file_path in file_paths:
            df = pd.read_excel(file_path, engine="xlrd")
            print("---")
            df = extract_order_info_and_create_new_df(df)

            print(df)

            merged_row_address_lst = extract_merged_row_address_and_fill_nan_cells(df)

            print(df)
            print(merged_row_address_lst)

            not_merged_row_address_lst = extract_not_merged_row_address(df, merged_row_address_lst)

            print(not_merged_row_address_lst)


    except Exception as e:
        print(f"파일 열기 오류: {str(e)}")
else:
    print("파일 선택이 취소되었습니다.")