import pandas as pd
import xlrd
import tkinter as tk
from tkinter import filedialog
from openpyxl.styles import Font
from openpyxl.styles import Alignment


root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls")])

def convert_xls_to_xlsx(input_file, output_file):
    # xls 파일을 DataFrame으로 읽기
    df = pd.read_excel(input_file, sheet_name=None)

    # ExcelWriter 객체 생성
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # DataFrame을 각각의 시트로 쓰기
        for sheet, data in df.items():
            data.to_excel(writer, sheet_name=sheet, index=False)

if __name__ == "__main__":
    # 파일 선택
    input_file_path = file_path

    if not input_file_path:
        print("No file selected. Exiting.")
    else:
        # 변환할 파일의 경로 출력
        print(f"Selected file: {input_file_path}")

        # 파일 이름과 경로에서 확장자를 바꿔서 출력 파일 경로 생성
        output_file_path = ".".join(input_file_path.split(".")[:-1]) + "_쉽핑마크_데이터.xlsx"

        # 함수 호출하여 변환 수행
        convert_xls_to_xlsx(input_file_path, output_file_path)

        print(f"Conversion complete. Output file saved at: {output_file_path}")

if file_path:
    try:
        df = pd.read_excel(file_path, engine="xlrd")

        item_row_indices = df[df.apply(lambda row: row.astype(str).str.contains('ITEM').any(), axis=1)].index
        item_columns_indices = [i for i, col in enumerate(df.columns) if df[col].astype(str).str.contains('ITEM').any()]

        # ITEM이라는 값이 들어 있는 셀의 행과 열 주소값
        item_row_index = item_row_indices[0]
        item_column_index = item_columns_indices[0]

        item_data = df.iloc[item_row_index + 1:, item_column_index]
        item_data = item_data.dropna()

        row_start_index = item_row_index + 1
        row_end_index = row_start_index + len(item_data)

        # buyer, color, styleNo, packing, baleNo 열 주소값 구하기
        buyer_columns_indices = [i for i, col in enumerate(df.columns) if df[col].astype(str).str.contains('BUYER').any()]
        color_columns_indices = [i for i, col in enumerate(df.columns) if df[col].astype(str).str.contains('COLOR').any()]
        styleNo_columns_indices = [i for i, col in enumerate(df.columns) if df[col].astype(str).str.contains('STYLE NO').any()]
        packing_columns_indices = [i for i, col in enumerate(df.columns) if df[col].astype(str).str.contains('PACKING').any()]
        baleNo_columns_indices = [i for i, col in enumerate(df.columns) if df[col].astype(str).str.contains('Bale No.').any()]

        buyer_column_index = buyer_columns_indices[0]
        color_column_index = color_columns_indices[0]
        styleNo_column_index = styleNo_columns_indices[0]
        packing_column_index = packing_columns_indices[0]
        baleNo_column_index = baleNo_columns_indices[0]

        # 5개 속성에 대해 데이터 구하기
        buyer_data = df.iloc[row_start_index:row_end_index, buyer_column_index]
        color_data = df.iloc[row_start_index:row_end_index, color_column_index]
        styleNo_data = df.iloc[row_start_index:row_end_index,styleNo_column_index]
        packing_data = df.iloc[row_start_index:row_end_index, packing_column_index]
        baleNo_data = df.iloc[row_start_index:row_end_index, baleNo_column_index]
        

        for k in range(len(baleNo_data)):
            i = k + row_start_index
            if pd.isna(baleNo_data[i]):
                baleNo_data[i] = (baleNo_data[i - 1])
            baleNo_data[i] = str(baleNo_data[i])
        


        data = []
        for k in range(len(item_data)):
            i = k + row_start_index
            lst = [buyer_data[i], item_data[i], color_data[i], styleNo_data[i], packing_data[i], baleNo_data[i]]
            data.append(lst)

        columns = ['BUYER', 'ITEM', 'COLOR', 'STYLE NO', 'PACKING', 'Bale No.']
        new_df = pd.DataFrame(data=data, columns=columns)

        
        
        # 병합된 셀 리스트
        merged_row_index_lst = []

        # packing 열에서 빈칸인 경우 바로 앞의 데이터를 가져오는 코드
        for i in range(len(new_df['PACKING'])):
            if pd.isna(new_df['PACKING'][i]):
                new_df['PACKING'][i] = new_df['PACKING'][i - 1]
                merged_row_index_lst.append(i)
        
        merged_lst = []
        lst = []
        for i in range(len(merged_row_index_lst)):
            if i >= len(merged_row_index_lst) - 1:
                lst.append(merged_row_index_lst[i])
                merged_lst.append(lst)
                break
            lst.append(merged_row_index_lst[i])
            if merged_row_index_lst[i + 1] - merged_row_index_lst[i] > 1:
                merged_lst.append(lst)
                lst = []

        for i in range(len(merged_lst)):
            first_value = merged_lst[i][0] - 1
            merged_lst[i].append(first_value)
            merged_lst[i] = sorted(merged_lst[i])





        # 병합되지 않은 셀들의 인덱스 구하기
        merged_values = []
        for lst in merged_lst:
            for value in lst:
                merged_values.append(value)
        

        for i in range(len(item_data)):
            if i not in merged_values:
                lst = [i]
                merged_lst.append(lst)
        # 최종 셀들 병합 여부 저장된 리스트
        merged_lst = sorted(merged_lst)

        

        # 'COLOR'에 W, B를 WHITE / BLACK으로 넣어주기
        for index in range(len(new_df['COLOR'])):
            new_df['COLOR'][index] = new_df['COLOR'][index].replace("W", "WHITE")
            new_df['COLOR'][index] = new_df['COLOR'][index].replace("B", "BLACK")
            new_df['COLOR'][index] = new_df['COLOR'][index].replace("CH", "CHACOAL")

        # 'PACKING' 분해하기
        for index in range(len(new_df['PACKING'])):
            if not(pd.isna(new_df['PACKING'][index])):
                new_df['PACKING'][index] = new_df['PACKING'][index].replace(" ", "")
                new_df['PACKING'][index] = new_df['PACKING'][index].replace("\n", "")
                
                str1 = ""
                gh_str = ""
                if "(" in new_df['PACKING'][index]:
                    start_index = new_df['PACKING'][index].find('(')
                    str1= new_df['PACKING'][index][:start_index]
                    gh_str = " " + new_df['PACKING'][index][start_index:]
                else:
                    str1 = new_df['PACKING'][index]

                ct_lst = str1.split('+')

                lst = []
                for var in ct_lst:
                    x_index = var.find('x')

                    b_index = 0
                    if 'b' in var:
                        b_index = var.find('b')
                    elif 'C/T' in var:
                        b_index = var.find('C')
                    y_str = var[:x_index] + gh_str
                    bale_num = int(var[x_index + 1:b_index])

                    for i in range(bale_num):
                        lst.append(y_str)

                new_df['PACKING'][index] = lst
                

        def remove_duplicates_function(lst):
            str2 = ""
            for i in range(len(lst)):
                str2 += lst[i]
                if i >= len(lst) - 1:
                    break
                str2 += " / "
            return str2


        final_lst = []
        for lst in merged_lst:
            # 여기가 하나의 packing -> 속성들 중복 제거하는 단위!!
            buyer_lst = []
            item_lst = []
            color_lst = []
            styleNo_lst = []
            packing_lst = []
            baleNo_lst = []

            
            for index in lst:
                if new_df['BUYER'][index] not in buyer_lst:
                    buyer_lst.append(new_df['BUYER'][index])

                if new_df['ITEM'][index] not in item_lst:
                    item_lst.append(new_df['ITEM'][index])

                if new_df['COLOR'][index] not in color_lst:
                    color_lst.append(new_df['COLOR'][index])

                if new_df['STYLE NO'][index] not in styleNo_lst:
                    styleNo_lst.append(new_df['STYLE NO'][index])
                
                if new_df['PACKING'][index] != packing_lst:
                    packing_lst = new_df['PACKING'][index]

                if new_df['Bale No.'][index] not in baleNo_lst:
                    baleNo_lst.append(new_df['Bale No.'][index])


            buyer_str = remove_duplicates_function(buyer_lst)
            item_str = remove_duplicates_function(item_lst)
            color_str = remove_duplicates_function(color_lst)
            styleNo_str = remove_duplicates_function(styleNo_lst)
            baleNo_str = remove_duplicates_function(baleNo_lst)

            
            final_lst.append([buyer_str, item_str,color_str, styleNo_str, packing_lst, baleNo_str])



        new_data = []

        for i in final_lst:
            for bale in i[4]:
                lst = [i[0], i[1], i[2], i[3], bale, i[5]]
                new_data.append(lst)


        result_df = pd.DataFrame(data=new_data, columns=columns)
        result_df = result_df.transpose()

        print(result_df)
        # 첫 번째 행을 컬럼으로 추가
        result_df.insert(0, ' ', result_df.index)
        



        lst = ['  ']
        for i in range(len(result_df.columns)):
            if i > 0:
                lst.append(i)

        result_df.columns = lst

        
        def add_dataframe_to_excel(existing_file, new_sheet_name, new_dataframe):
        # Excel 파일 열기
            with pd.ExcelWriter(existing_file, engine='openpyxl', mode='a') as writer:
                # DataFrame을 새로운 시트로 추가
                new_dataframe.to_excel(writer, sheet_name=new_sheet_name, index=False)

                # 엑셀 시트 객체 가져오기
                sheet = writer.sheets[new_sheet_name]

                # 인덱스 값에 볼드 스타일 적용
                for cell in sheet['A']:
                    cell.font = Font(bold=True)

                sheet['A1'].alignment = Alignment(horizontal='left')

        existing_excel_file = output_file_path

        # 기존 Excel 파일 경로 및 시트 이름 설정
        existing_excel_file = output_file_path
        new_sheet_name = '쉽핑마크 데이터'

        # 함수 호출하여 새로운 시트에 데이터프레임 추가
        add_dataframe_to_excel(existing_excel_file, new_sheet_name, result_df)

        print(f"Dataframe added to {existing_excel_file} under sheet {new_sheet_name}")
   
    except Exception as e:
        print(f"파일 열기 오류: {str(e)}")
else:
    print("파일 선택이 취소되었습니다.")