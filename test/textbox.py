import openpyxl

def find_shapes(workbook_path, sheet_name):
    # 엑셀 파일 열기
    workbook = openpyxl.load_workbook(workbook_path)

    # 특정 시트 선택
    sheet = workbook[sheet_name]

    # 시트에서 도형 찾기
    for shape in sheet.shapes:
        print(f"도형 타입: {shape.type}")
        print(f"좌표: ({shape.left}, {shape.top})")
        print(f"너비: {shape.width}, 높이: {shape.height}")

    # 엑셀 파일 닫기
    workbook.close()


# 사용 예제
workbook_path = "C:\\Users\\bnj30\\Desktop\\쉬핑마크\\���θ�ũ\\MS D&M.xlsx"  # 실제 파일 경로로 대체
sheet_name = "Sheet1"  # 시트 이름으로 대체

find_shapes(workbook_path, sheet_name)