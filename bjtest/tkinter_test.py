import tkinter as tk
import tkinter.font
from tkinter import filedialog
import tkinter.ttk
from tkinter import messagebox

import os

folder_path = r"C:\Users\bnj30\Desktop\쉬핑마크"  # 실제 폴더 경로로 바꿔주세요

# 해당 폴더의 모든 파일 목록을 가져옴
file_names = os.listdir(folder_path)

# 파일 목록을 리스트에 저장
file_list = []
for file_name in file_names:
    file_list.append(os.path.basename(file_name))

file_paths = []
def add_item():
    global file_paths
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xls")])

    for file_path in file_paths:
        name = os.path.basename(file_path)
        if (file_path not in listbox1.get(0, tk.END)) and (name in file_list):
            listbox1.insert(listbox1.size(), file_path)
            listbox2.insert(listbox2.size(), name)


def delete_selected_item():
    selected_items = listbox1.curselection()
    if not selected_items:
        messagebox.showinfo("경고", "삭제할 항목을 선택하세요.")
        return
    
    for index in reversed(selected_items):
        listbox1.delete(index)
        listbox2.delete(index)




window = tk.Tk()
window.geometry("1600x1200+100+50")
window.resizable(False, False)

label_font=tk.font.Font(family="맑은 고딕", size=20, weight="bold")
lst_font=tk.font.Font(family="맑은 고딕", size=16)

frame1=tk.Frame(window, relief="solid", bd=1)
frame1.pack(side="left", fill="both", expand=True)

frame2=tk.Frame(window, relief="solid")
frame2.pack(side="left", fill="x")

frame3=tk.Frame(window, relief="solid", bd=1)
frame3.pack(side="right", fill="both", expand=True)

# progressbar=tkinter.ttk.Progressbar(frame2, maximum=100, mode="indeterminate")
# progressbar.pack()

# progressbar.start(50)

add_button = tk.Button(frame2, text="패킹리스트 추가 [+]", command=add_item, font=14)
add_button.pack()
delete_button = tk.Button(frame2, text="패킹리스트 삭제 [-]", command=delete_selected_item, font=14)
delete_button.pack()


label1=tk.Label(frame1, text="패킹리스트", font=label_font)
label1.pack(side="top", pady=10)

label2=tk.Label(frame3, text="쉽핑마크", font=label_font)
label2.pack(side="top", pady=10)



listbox1 = tk.Listbox(frame1, selectmode='extended', height=0, font=lst_font, bd=1)
listbox1.pack(padx=10, pady=10, fill="both", expand=True)

listbox2 = tk.Listbox(frame3, height=0, font=lst_font, bd=1)
listbox2.pack(padx=10, pady=10, fill="both", expand=True)


for i, file_path in enumerate(file_paths):
    name = os.path.basename(file_path)
    if name in file_list:
        listbox1.insert(i, file_path)
        listbox2.insert(i, name)
window.mainloop()

# import tkinter as tk

# # Tkinter 윈도우 생성
# window = tk.Tk()
# window.title("2x2 모양의 프레임 레이아웃")

# # 2x2 모양의 프레임 생성
# frame1 = tk.Frame(window, relief="solid", bd=1)
# frame1.grid(row=0, column=0, padx=5, pady=5)

# frame2 = tk.Frame(window, relief="solid", bd=1)
# frame2.grid(row=0, column=1, padx=5, pady=5)

# frame3 = tk.Frame(window, relief="solid", bd=1)
# frame3.grid(row=1, column=0, padx=5, pady=5)

# frame4 = tk.Frame(window, relief="solid", bd=1)
# frame4.grid(row=1, column=1, padx=5, pady=5)

# # 프레임에 내용 추가 (예: 라벨)
# label1 = tk.Label(frame1, text="Frame 1")
# label1.pack(side="left", padx=10, pady=10)

# label2 = tk.Label(frame2, text="Frame 2")
# label2.pack(padx=10, pady=10)

# label3 = tk.Label(frame3, text="Frame 3")
# label3.pack(padx=10, pady=10)

# label4 = tk.Label(frame4, text="Frame 4")
# label4.pack(padx=10, pady=10)

# # Tkinter 윈도우 실행
# window.mainloop()