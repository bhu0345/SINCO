import tkinter as tk
from tkinter import ttk

# 简单测试GUI是否能显示
root = tk.Tk()
root.title("测试窗口")
root.geometry("400x300")

label = ttk.Label(root, text="如果你能看到这个文字，说明tkinter工作正常", font=("", 14))
label.pack(pady=50)

button = ttk.Button(root, text="关闭", command=root.destroy)
button.pack(pady=20)

print("GUI窗口应该已经显示，如果看不到请检查是否被其他窗口遮挡")
root.mainloop()
