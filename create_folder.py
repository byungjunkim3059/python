import os
import calendar
from datetime import datetime
import openpyxl
from openpyxl.styles import Font

year = 2023

def create_folder(path):
            try:
                # 폴더 생성
                os.makedirs(path)
                print(f"폴더가 성공적으로 생성되었습니다. 경로: {path}")
            except FileExistsError:
                print(f"폴더가 이미 존재합니다. 경로: {path}")
            except Exception as e:
                print(f"폴더 생성 중 오류가 발생했습니다. 오류 내용: {e}")
folder_path = "C:\\Users\\bnj30\Desktop\\출고서류\\" + str(year) +"년\\"
create_folder(folder_path)




def create_excel_table_with_columns(file_path, columns):
    # 엑셀 워크북 및 워크시트 생성
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    # 헤더(컬럼) 추가
    for col_num, column in enumerate(columns, 1):
        cell = worksheet.cell(row=1, column=col_num, value=column)
        cell.font = Font(bold=True)

    # 엑셀 파일 저장
    workbook.save(file_path)

columns = ['NO.', 'CUSTOMER_INFO', 'BUYER', 'ITEM', 'COLOR', 'STYLE NO', 'PACKING', 'Bale No.', 'S/M']
seperate_amount_columns = ['NO.', 'CUSTOMER', 'ITEM', 'COLOR', "Q'TY"]


def days_in_month(year, month):
    return calendar.monthrange(year, month)[1]

def get_day_of_week(day_of_week):
    
    days = ["월", "화", "수", "목", "금", "토", "일"]
    day_name = days[day_of_week]

    return day_name

for month in range(12):
    month = month + 1
    if month < 10:
        path = folder_path + "0" +str(month) + "월\\"
    else:
        path = folder_path + str(month) + "월\\"
    create_folder(path)
    for day in range(days_in_month(year, month)):
        day = day + 1
        date_object = datetime(year, month, day)
        day_of_week = date_object.weekday()

        if day_of_week < 5:
            day_name = get_day_of_week(day_of_week)

            if day < 10:
                 day = "0" + str(day)
            else:
                 day = str(day)

            if month < 10:
                 monthstr = "0" + str(month)
            else:
                 monthstr = str(month)

            inner_path = path + str(year) + '-' + monthstr + '-' + day + "(" + day_name + ")\\"

            create_folder(inner_path)

            excel_file_path1 = inner_path + str(year) + '-' + monthstr + '-' + day + "(" + day_name + ")_출고.xlsx" 
            create_excel_table_with_columns(excel_file_path1, columns)
            
            excel_file_path2 = inner_path + str(year) + '-' + monthstr + '-' + day + "(" + day_name + ")_똥갈이.xlsx"
            create_excel_table_with_columns(excel_file_path2, seperate_amount_columns)
