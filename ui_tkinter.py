import tkinter as tk

def delete_selected():
    selected_indices = listbox.curselection()
    for index in reversed(selected_indices):
        listbox.delete(index)

def modify_selected():
    selected_indices = listbox.curselection()
    if len(selected_indices) == 1:
        selected_index = selected_indices[0]
        new_value = entry.get()
        listbox.delete(selected_index)
        listbox.insert(selected_index, new_value)
        entry.delete(0, tk.END)  # 수정 후 Entry 초기화

root = tk.Tk()
root.title("Listbox Modification Example")

frame1 = tk.Frame(root)
frame1.pack(fill=tk.BOTH, expand=True)

listbox = tk.Listbox(frame1, selectmode='extended', height=10)
listbox.pack(padx=10, fill="both", expand=True)

# 예시로 몇 가지 항목을 추가
for i in range(10):
    listbox.insert(tk.END, f"Item {i}")

delete_button = tk.Button(root, text="Delete Selected", command=delete_selected)
delete_button.pack(pady=5)

modify_button = tk.Button(root, text="Modify Selected", command=modify_selected)
modify_button.pack(pady=5)

entry = tk.Entry(root)
entry.pack(pady=5)

root.mainloop()
