import openpyxl

def extract_shape_text(file_path):
    wb = openpyxl.load_workbook(file_path)
    
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        
        for drawing in sheet.drawing_shapes:
            if isinstance(drawing, openpyxl.drawing.textbox.TextBox):
                print(f"도형({drawing.name})의 문자열: {drawing.text}")

# 엑셀 파일 경로
excel_file_path = r"C:\Users\bnj30\Desktop\10월 출고서류\10월 출고서류\[국내출고] [나디아] 24'SS 내셔널지오그래픽 신규 스타일 출고요청_10_18(수) _DNW 로지스틱스_ 입고\SHIPPING MARK_VNC(VNG).xlsx"  # 본인의 파일 경로로 수정

# 텍스트 상자에서 문자열 추출
extract_shape_text(excel_file_path)
