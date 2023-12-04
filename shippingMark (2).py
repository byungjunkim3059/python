import pandas as pd
import xlrd
from tkinter import filedialog
import tkinter as tk


root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls")])

if file_path:
    try:
        df = pd.read_excel(file_path, engine="xlrd")


        item = []
        width = []
        color = []
        styleNo = []
        packing = []

        item_series = df.iloc[7:,1]
        item_series = item_series.dropna()
        for data in item_series:
            item.append(data)

        width_series = df.iloc[7:,3]
        width_series = width_series.dropna()
        for data in width_series:
            width.append(data)

        color_series = df.iloc[7:,4]
        color_series = color_series.dropna()
        for data in color_series:
            color.append(data)

        styleNo_series = df.iloc[7:,7]
        styleNo_series = styleNo_series.dropna()
        for data in styleNo_series:
            styleNo.append(data)


        merged_cells = []

        packing_series = df.iloc[7:,10]
        i = 0
        for data in packing_series:
            if pd.isna(data):
                packing.append(packing[i - 1])
                merged_cells.append(i)
            else:
                packing.append(data)
            i = i + 1
            if i >= len(styleNo):
                break
        
        data = []
        for i in range(len(item)):
            lst = [item[i], width[i], color[i], styleNo[i], packing[i]]
            data.append(lst)

        columns = ['item', 'width', 'color', 'styleNo', 'packing']




        # 병합된 셀 열 주소값 구하기
        merged_cells_arr = []
        arr = []
        k = 0
        for index, value in enumerate(merged_cells):
            if k >= len(merged_cells) - 1:
                arr.append(value)
                merged_cells_arr.append(arr)
                break
            if merged_cells[index + 1] - value > 1:
                arr.append(value)
                merged_cells_arr.append(arr)
                arr = []
            else:
                arr.append(value)
            k = k + 1


        merged_cells_arr2 = []

        for lst in merged_cells_arr:
            lst.append(lst[0] - 1)
            sorted_lst = sorted(lst)
            merged_cells_arr2.append(sorted_lst)

        # 병합된 셀 리스트
        # print(merged_cells_arr2)

        # 병합되지 않은 셀 주소값 구하기
        lst1 = []
        for i in range(len(item)):
            lst1.append(i)


        lst2 = []
        for i in range(len(merged_cells_arr2)):
            for value in merged_cells_arr2[i]:
               lst2.append(value) 

        unmerged_arr = []

        for value in lst1:
            if value in lst2:
                continue
            else:
                unmerged_arr.append(value)

        # print(unmerged_arr)

        # 최종 구성
        final_arr = []
        if unmerged_arr:
            final_arr.append(unmerged_arr)
        for lst in merged_cells_arr2:
            final_arr.append(lst)

        new_df = pd.DataFrame(data=data, columns=columns)
        print(new_df)
        print(final_arr)
        



    except Exception as e:
        print(f"파일 열기 오류: {str(e)}")
else:
    print("파일 선택이 취소되었습니다.")
