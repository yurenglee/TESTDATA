# TKINTER GUI EXAMPLE
import os
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

root = tk.Tk()
root.title("我的第一个Tkinter程序")
root.geometry("520x320")

# 添加标签组件
label = ttk.Label(root, text="欢迎使用Tkinter!")
label.pack(pady=80)
# 添加按钮组件
button = ttk.Button(root, text="点击我", 
                   command=lambda: print("按钮被点击了"))
button.pack()

text = tk.Text(root, wrap="word")
text.pack(fill="both", expand=True)
text.insert("3.0", "在这里输入一些文字，然后右键试试。")

menu = tk.Menu(root, tearoff=0)
menu.add_command(label="剪切", command=lambda: text.event_generate("<<Cut>>"))
menu.add_command(label="复制", command=lambda: text.event_generate("<<Copy>>"))
menu.add_command(label="粘贴", command=lambda: text.event_generate("<<Paste>>"))
menu.add_separator()
menu.add_command(label="全选", command=lambda: text.tag_add("sel", "1.0", "end"))

def popup(event):
    menu.post(event.x_root, event.y_root)

text.bind("<Button-3>", popup)
text.bind("<Button-2>", popup)

root.mainloop()
