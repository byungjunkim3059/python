import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

# file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls *.xlsx")])
# file_path2 = filedialog.askdirectory()
file_path3 = filedialog.askopenfile()
print(file_path3.name)
