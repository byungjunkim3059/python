import tkinter as tk
from tkinter import filedialog
import os

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls")])

def extract_filename(file_path):
    try:
        # 파일 이름 추출
        filename = os.path.basename(file_path)
        # 유니코드 디코딩 시도
        decoded_filename = filename.encode('utf-8').decode('utf-8')
        return decoded_filename
    except UnicodeDecodeError:
        # 디코딩 실패 시 원본 파일 이름 반환
        return filename

# 사용 예제
filename = extract_filename(file_path)

print("전체 파일 경로:", file_path)
print("파일 이름:", filename)
