import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("布局测试")
root.geometry("800x600")

# 测试混合使用 grid 和 pack
top = ttk.Frame(root, relief="solid", borderwidth=2)
top.pack(fill="x", padx=10, pady=10)

ttk.Label(top, text="标签1", background="lightblue").grid(row=0, column=0, sticky="w", padx=5, pady=5)
ttk.Entry(top, width=20).grid(row=0, column=1, padx=5, pady=5)

ttk.Label(top, text="标签2", background="lightgreen").grid(row=1, column=0, sticky="w", padx=5, pady=5)
ttk.Entry(top, width=20).grid(row=1, column=1, padx=5, pady=5)

btn_frame = ttk.Frame(top, relief="solid", borderwidth=1)
btn_frame.grid(row=0, column=2, rowspan=2, padx=10)
ttk.Button(btn_frame, text="按钮1").pack(pady=2)
ttk.Button(btn_frame, text="按钮2").pack(pady=2)

# 分隔线
ttk.Separator(root).pack(fill="x", padx=10, pady=10)

# 主内容区
main = ttk.Frame(root, relief="solid", borderwidth=2)
main.pack(fill="both", expand=True, padx=10, pady=10)

left = ttk.Frame(main, relief="solid", borderwidth=1)
left.pack(side="left", fill="both", expand=True, padx=5, pady=5)

ttk.Label(left, text="左侧内容区域", background="lightyellow").pack(pady=10)
ttk.Button(left, text="测试按钮").pack(pady=5)

right = ttk.Frame(main, relief="solid", borderwidth=1)
right.pack(side="right", fill="y", padx=5, pady=5)

ttk.Label(right, text="右侧内容区域", background="lightpink").pack(pady=10)
text = tk.Text(right, height=10, width=30)
text.pack(fill="both", expand=True, padx=5, pady=5)
text.insert("1.0", "这是一些测试文本\n" * 10)

print("如果你能看到窗口中的彩色区域和文字，说明布局正常")
root.mainloop()
