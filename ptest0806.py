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
# 曲線繪圖工具 PUSH GITHUB 20260813 #

class ExcelPlotApp:
    MAX_SERIES = 5

    def __init__(self, root):
        self.root = root
        self.root.title("Excel 曲線繪圖工具")
        self.root.geometry("1700x1040")

        self.df = None
        self.current_path = ""
        self.plot_panels = []
        self.selection_vars = {}
        self.search_var = tk.StringVar(value="")
        self.search_var.trace_add("write", lambda *_: self.refresh_selection_list())
        self.linewidth_top_var = tk.DoubleVar(value=1.0)
        self.linewidth_bottom_var = tk.DoubleVar(value=1.0)
        self.cursor_display_top_var = tk.BooleanVar(value=True)
        self.cursor_display_bottom_var = tk.BooleanVar(value=True)
        self.dual_axis_top_var = tk.BooleanVar(value=False)
        self.dual_axis_bottom_var = tk.BooleanVar(value=False)

        self.build_ui()

    def build_ui(self):
        # ── 最外層容器：整個視窗的根 Frame，四周留 5px 邊距 ──
        main = ttk.Frame(self.root, padding=5)
        main.pack(fill="both", expand=True)
        # ── 頂部工具列：開檔按鈕 + 檔名標籤 + 狀態文字 ──
        top = ttk.Frame(main)
        top.pack(fill="x", pady=(0, 6))
        # 按鈕：點擊後呼叫 self.open_excel_file 開啟 Excel 檔
        ttk.Button(top, text="A. 開啟 Excel 檔", command=self.open_excel_file).pack(side="left")
        self.file_label = ttk.Label(top, text="尚未開啟 Excel 檔")
        self.file_label.pack(side="left", padx=(10, 0))
        # 顯示目前已開啟的檔名（尚未開檔時顯示提示文字）
        self.status_var = tk.StringVar(value="請先開啟 Excel 檔")
        ttk.Label(top, textvariable=self.status_var).pack(side="left", padx=(20, 0))
        # 狀態列文字，用 StringVar 綁定，之後可用 self.status_var.set(...) 更新
        content = ttk.Frame(main)
        content.pack(fill="both", expand=True)
        # ── 主要內容區：左右兩欄，用grid管理比例（左 3 : 右 5）──
        content.columnconfigure(0, weight=2)   # 左欄（圖表區）較寬
        content.columnconfigure(1, weight=1)   # 右欄（控制區）較窄
        content.rowconfigure(0, weight=1)
        # ── 左側：圖表區 ──
        left_frame = ttk.Frame(content)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 0))
        left_frame.columnconfigure(0, weight=1)   
        left_frame.rowconfigure(1, weight=1)      
        # 垂直方向的可拖曳分割視窗，讓使用者能自行調整上下兩個圖表的高度比例
        plots_frame = ttk.Panedwindow(left_frame, orient="vertical")
        plots_frame.pack(fill="both", expand=True)

        # ── 右側：控制面板區（上：分頁籤，下：座標點列表）──
        right_frame = ttk.Frame(content)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 0))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)

        # -- 右上：分頁籤容器 --
        right_top = ttk.Frame(right_frame, padding=(0, 0, 0, 4))
        right_top.grid(row=0, column=0, sticky="nsew")
        right_top.columnconfigure(0, weight=1)
        right_top.rowconfigure(0, weight=1)
        # Notebook：多分頁控制面板，固定高度 300px
        self.control_notebook = ttk.Notebook(right_top)
        self.control_notebook.grid(row=0, column=0, sticky="nsew")
        self.control_notebook.configure(height=300)
        # 分頁一：曲線資料選取（勾選要顯示哪些欄位/變數)
        self.selection_tab = ttk.Frame(self.control_notebook)
        self.selection_tab.pack(fill="both", expand=True)
        self.control_notebook.add(self.selection_tab, text="資料選取")
        # 分頁二：控制與設定（縮放、線寬、雙 Y 軸等）
        self.control_settings_tab = ttk.Frame(self.control_notebook)
        self.control_settings_tab.pack(fill="both", expand=True)
        self.control_notebook.add(self.control_settings_tab, text="控制與設定")
        # 分頁三：預留給未來擴充功能的空分頁
        self.extra_tab = ttk.Frame(self.control_notebook)
        self.extra_tab.pack(fill="both", expand=True)
        self.control_notebook.add(self.extra_tab, text="擴充功能")
        ttk.Label(self.extra_tab, text="後續新功能將放在這裡").pack(anchor="w", padx=12, pady=12)

        # ── 建立兩個圖表面板（上視窗 index=1、下視窗 index=2），加入 Panedwindow ── 
        for index in range(2):
            panel = self._create_panel(plots_frame, index + 1)
            self.plot_panels.append(panel)
            plots_frame.add(panel["frame"], weight=1)
            self.attach_plot_interaction(panel)
        ## 分頁一內容：共用曲線選取區 ──
        selection_frame = ttk.Labelframe(self.selection_tab, text="共用曲線選取", padding=6)
        selection_frame.pack(fill="both", expand=True, padx=4, pady=4)
        selection_frame.columnconfigure(0, weight=1)
        selection_frame.rowconfigure(1, weight=1)
        # 搜尋列：輸入關鍵字篩選變數名稱，並可一鍵清除
        search_frame = ttk.Frame(selection_frame)
        search_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Label(search_frame, text="搜尋變數：").pack(side="left")
        ttk.Entry(search_frame, textvariable=self.search_var, width=10).pack(side="left", fill="x", expand=True, padx=(6, 0))
        ttk.Button(search_frame, text="清除", command=self._clear_search).pack(side="left", padx=(4, 0))

       # 可捲動的勾選清單：因為 ttk 元件本身不支援捲動，
       # 所以用 Canvas 當「視窗」，裡面放一個 Frame（checkbox_container）承載所有勾選框，
       # 再搭配 Scrollbar 控制 Canvas 的可視範圍。
        self.selection_canvas = tk.Canvas(selection_frame, height=160)
        self.selection_canvas.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(selection_frame, orient="vertical", command=self.selection_canvas.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
      
        self.checkbox_container = ttk.Frame(self.selection_canvas)
       # 當 checkbox_container 內容大小改變時（例如動態加入勾選框），
       # 重新計算 Canvas 的捲動範圍，否則捲軸範圍不會更新
        
        #self.checkbox_container.bind(
        #    "<Configure>",
        #    lambda event: self.selection_canvas.configure(scrollregion=self.selection_canvas.bbox("all")),
        #)

       # 1. 綁定 Canvas 尺寸變更事件：讓裡面的 Frame 寬度永遠跟 Canvas 一樣寬
        self.selection_canvas.bind(
            "<Configure>",
            lambda event: self.selection_canvas.itemconfig(canvas_window, width=event.width)
        )

        # 2. 綁定 Frame 內容變更事件：更新 Scrollbar 捲動範圍
        self.checkbox_container.bind(
            "<Configure>",
            lambda event: self.selection_canvas.configure(scrollregion=self.selection_canvas.bbox("all")),
        )

        # 3. 建立視窗並保存物件引用 (canvas_window)
        canvas_window = self.selection_canvas.create_window((0, 0), window=self.checkbox_container, anchor="nw")
        self.selection_canvas.configure(yscrollcommand=scrollbar.set)


        # 把 checkbox_container 這個 Frame 嵌入 Canvas 內，錨點在左上角
        self.selection_canvas.create_window((0, 0), window=self.checkbox_container, anchor="nw")
        self.selection_canvas.configure(yscrollcommand=scrollbar.set)
        # ────── 縮放範圍輸入框對應的變數（上/下視窗各自獨立）──────
        self.x_min_top_var = tk.StringVar()
        self.x_max_top_var = tk.StringVar()
        self.x_min_bottom_var = tk.StringVar()
        self.x_max_bottom_var = tk.StringVar()

        # ── 分頁二內容：控制與設定 ──
        control_settings_frame = ttk.Frame(self.control_settings_tab, padding=6)
        control_settings_frame.pack(fill="both", expand=True)
        control_settings_frame.columnconfigure(0, weight=1)

        control_settings = ttk.Labelframe(control_settings_frame, text="控制與設定", padding=6)
        control_settings.grid(row=0, column=0, sticky="nsew")
        # -- 上視窗 X 軸縮放區 --
        zoom_top = ttk.Labelframe(control_settings, text="X 縮放 - 上視窗", padding=6)
        zoom_top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(zoom_top, text="X 開始：").grid(row=0, column=0, sticky="w")
        ttk.Entry(zoom_top, textvariable=self.x_min_top_var, width=8).grid(row=0, column=1, padx=(4, 0))
        ttk.Label(zoom_top, text="~").grid(row=0, column=2, padx=4)
        ttk.Entry(zoom_top, textvariable=self.x_max_top_var, width=8).grid(row=0, column=3)

        # panel_index=1 代表這兩個按鈕作用在「上視窗」（第一個圖表面板）
        ttk.Button(zoom_top, text="放大(上)", command=lambda: self.apply_zoom(panel_index=1)).grid(row=0, column=4, padx=(8, 0))
        ttk.Button(zoom_top, text="重設(上)", command=lambda: self.reset_zoom(panel_index=1)).grid(row=0, column=5, padx=(8, 0))

        zoom_bottom = ttk.Labelframe(control_settings, text="X 縮放 - 下視窗", padding=6)
        zoom_bottom.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(zoom_bottom, text="X 開始：").grid(row=0, column=0, sticky="w")
        ttk.Entry(zoom_bottom, textvariable=self.x_min_bottom_var, width=8).grid(row=0, column=1, padx=(4, 0))
        ttk.Label(zoom_bottom, text="~").grid(row=0, column=2, padx=4)
        ttk.Entry(zoom_bottom, textvariable=self.x_max_bottom_var, width=8).grid(row=0, column=3)
        ttk.Button(zoom_bottom, text="放大(下)", command=lambda: self.apply_zoom(panel_index=2)).grid(row=0, column=4, padx=(8, 0))
        ttk.Button(zoom_bottom, text="重設(下)", command=lambda: self.reset_zoom(panel_index=2)).grid(row=0, column=5, padx=(8, 0))

        panel_settings = ttk.Labelframe(control_settings, text="圖窗顯示設定", padding=6)
        panel_settings.grid(row=2, column=0, sticky="ew")
        panel_settings.columnconfigure(1, weight=1)
        ttk.Label(panel_settings, text="上視窗線條粗細：").grid(row=0, column=0, sticky="w")
        ttk.Entry(panel_settings, textvariable=self.linewidth_top_var, width=8).grid(row=0, column=1, sticky="w", padx=(4, 0))
        ttk.Label(panel_settings, text="(預設 1.0)").grid(row=0, column=2, sticky="w", padx=(4, 0))
        ttk.Checkbutton(panel_settings, text="使用雙 Y 軸", variable=self.dual_axis_top_var, command=self.refresh_all_plots).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Checkbutton(panel_settings, text="顯示游標座標", variable=self.cursor_display_top_var).grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))

        bottom_panel_settings = ttk.Labelframe(control_settings, text="下視窗設定", padding=6)
        bottom_panel_settings.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        bottom_panel_settings.columnconfigure(1, weight=1)
        ttk.Label(bottom_panel_settings, text="下視窗線條粗細：").grid(row=0, column=0, sticky="w")
        ttk.Entry(bottom_panel_settings, textvariable=self.linewidth_bottom_var, width=8).grid(row=0, column=1, sticky="w", padx=(4, 0))
        ttk.Label(bottom_panel_settings, text="(預設 1.0)").grid(row=0, column=2, sticky="w", padx=(4, 0))
        ttk.Checkbutton(bottom_panel_settings, text="使用雙 Y 軸", variable=self.dual_axis_bottom_var, command=self.refresh_all_plots).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Checkbutton(bottom_panel_settings, text="顯示游標座標", variable=self.cursor_display_bottom_var).grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))

        right_bottom = ttk.Frame(right_frame, padding=(0, 4, 0, 0))
        right_bottom.grid(row=1, column=0, sticky="nsew")
        right_bottom.columnconfigure(0, weight=1)
        right_bottom.columnconfigure(1, weight=1)
        right_bottom.rowconfigure(0, weight=1)

        points_frame = ttk.Labelframe(right_bottom, text="Cursor 座標點", padding=6)
        points_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
        points_frame.columnconfigure(0, weight=1)
        points_frame.rowconfigure(1, weight=1)
        points_frame.rowconfigure(3, weight=1)

        ttk.Label(points_frame, text="點擊圖形可新增座標點，座標文字較小").grid(row=0, column=0, sticky="w", pady=(0, 8))

        ttk.Label(points_frame, text="上視窗:").grid(row=1, column=0, sticky="w")
        self.points_list_top = tk.Listbox(points_frame, height=6)
        self.points_list_top.grid(row=2, column=0, sticky="nsew", pady=(4, 8))

        ttk.Label(points_frame, text="下視窗:").grid(row=3, column=0, sticky="w")
        self.points_list_bottom = tk.Listbox(points_frame, height=6)
        self.points_list_bottom.grid(row=4, column=0, sticky="nsew", pady=(4, 0))

    def _create_panel(self, container, panel_index):
        frame = ttk.LabelFrame(container, text=f"圖窗 {panel_index}", padding=2)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        header = ttk.Frame(frame)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 1))
        header.columnconfigure(0, weight=1)
        header.columnconfigure(2, weight=0)
        header.columnconfigure(3, weight=0)

        title_var = tk.StringVar(value=f"圖窗 {panel_index}")
        ttk.Label(header, text="標題：").grid(row=0, column=0, sticky="w")
        ttk.Entry(header, textvariable=title_var, width=24).grid(row=0, column=1, sticky="w", padx=(6, 0))

        ## cursor_label = ttk.Label(header, text="Cursor: --", font=(None, 9))
        cursor_label = ttk.Label(header, text="Cursor: --", font=(None, 9), width=32, anchor="w")
        cursor_label.grid(row=0, column=2, sticky="e", padx=(10, 0))

        toolbar_frame = ttk.Frame(header)
        toolbar_frame.grid(row=0, column=3, sticky="e")

        plot_area = ttk.Frame(frame)
        plot_area.grid(row=1, column=0, sticky="nsew")
        plot_area.columnconfigure(0, weight=1)
        plot_area.rowconfigure(0, weight=1)

        fig = plt.Figure(figsize=(14, 6.0), dpi=100)
        ax = fig.add_subplot(111)
        ax.set_title(title_var.get())
        ax.grid(True, alpha=0.3)
        fig.subplots_adjust(top=0.92, bottom=0.10)

        canvas = FigureCanvasTkAgg(fig, master=plot_area)
        canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")
        canvas.draw()

        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame, pack_toolbar=False)
        toolbar.update()
        toolbar.pack(side="right", fill="x")

        return {
            "frame": frame,
            "fig": fig,
            "ax": ax,
            "canvas": canvas,
            "toolbar": toolbar,
            "title_var": title_var,
            "panel_index": panel_index,
            "cursor_label": cursor_label,
            "cursor_vline": None,
            "cursor_hline": None,
            "saved_points": [],
        }

    def open_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔",
            filetypes=[("Excel 檔", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        if not file_path:
            return

        try:
            if file_path.lower().endswith(".csv"):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)

            if df is None or df.empty:
                raise ValueError("Excel 內容為空")

            self.df = df
            self.current_path = file_path
            self.file_label.configure(text=os.path.basename(file_path))
            self.status_var.set(f"已讀取: {file_path}")
            self.setup_column_selection()
            self.refresh_all_plots()
        except Exception as exc:
            messagebox.showerror("讀取失敗", str(exc))

    def setup_column_selection(self):
        if self.df is None:
            return

        columns = [str(col) for col in self.df.columns if pd.notna(col)]
        for widget in self.checkbox_container.winfo_children():
            widget.destroy()

        self.selection_vars = {}
        self.dual_axis_top_var.set(False)
        self.dual_axis_bottom_var.set(False)
        self.linewidth_top_var.set(1.0)
        self.linewidth_bottom_var.set(1.0)
        self.cursor_display_top_var.set(True)
        self.cursor_display_bottom_var.set(True)
        self.x_min_top_var.set("")
        self.x_max_top_var.set("")
        self.x_min_bottom_var.set("")
        self.x_max_bottom_var.set("")

        if not columns:
            ttk.Label(self.checkbox_container, text="沒有可選欄位").pack(anchor="w")
            return

        for col in columns:
            up_var = tk.BooleanVar(value=False)
            down_var = tk.BooleanVar(value=False)
            x_var = tk.BooleanVar(value=False)
            self.selection_vars[col] = {"up": up_var, "down": down_var, "x": x_var}

        self._render_selection_rows(columns)

    def _clear_search(self):
        self.search_var.set("")

    def refresh_selection_list(self):
        if self.df is None or not hasattr(self, "checkbox_container"):
            return
        if not self.selection_vars:
            return
        self._render_selection_rows(list(self.selection_vars.keys()))

    def _render_selection_rows(self, columns):
        for widget in self.checkbox_container.winfo_children():
            widget.destroy()

        visible_columns = self._get_filtered_columns(columns)
        if not visible_columns:
            ttk.Label(self.checkbox_container, text="沒有符合的欄位").pack(anchor="w")
            return

        for col in visible_columns:
            up_var = self.selection_vars[col]["up"]
            down_var = self.selection_vars[col]["down"]
            x_var = self.selection_vars[col]["x"]

            row = ttk.Frame(self.checkbox_container)
            row.pack(fill="x", pady=1)
            # width=10 是設定這個 Label 元件的固定寬度為 9個字元
            # ttk.Label(row, text=col, width=9, anchor="w").pack(side="left")
            # 將 width=9 移除，並加上 fill="x", expand=True，讓欄位名稱自動擠滿剩餘寬度
            ttk.Label(row, text=col, anchor="w").pack(side="left", fill="x", expand=True, padx=(0, 4))
            ttk.Checkbutton(row, text="上", variable=up_var, command=lambda c=col: self.on_selection_changed(c, "up")).pack(side="right", padx=1)
            ttk.Checkbutton(row, text="下", variable=down_var, command=lambda c=col: self.on_selection_changed(c, "down")).pack(side="right", padx=1)
            ttk.Checkbutton(row, text="X", variable=x_var, command=lambda c=col: self.on_selection_changed(c, "x")).pack(side="right", padx=1)

            #ttk.Checkbutton(row, text="上", variable=up_var, command=lambda c=col: self.on_selection_changed(c, "up")).pack(side="left", padx=(1, 0))
            #ttk.Checkbutton(row, text="下", variable=down_var, command=lambda c=col: self.on_selection_changed(c, "down")).pack(side="left", padx=(1, 0))
            #ttk.Checkbutton(row, text="X", variable=x_var, command=lambda c=col: self.on_selection_changed(c, "x")).pack(side="left", padx=(1, 0))

        self.checkbox_container.update_idletasks()
        self._scroll_to_first_match()

    def _get_filtered_columns(self, columns):
        keyword = self.search_var.get().strip().lower()
        if not keyword:
            return list(columns)
        return [col for col in columns if keyword in str(col).lower()]

    def _scroll_to_first_match(self):
        if not hasattr(self, "selection_canvas") or self.selection_canvas is None:
            return
        if not self.selection_canvas.winfo_exists():
            return

        self.selection_canvas.configure(scrollregion=self.selection_canvas.bbox("all"))
        keyword = self.search_var.get().strip().lower()
        if not keyword:
            self.selection_canvas.yview_moveto(0)
            return

        children = self.checkbox_container.winfo_children()
        if not children:
            return

        first_row = children[0]
        if not first_row.winfo_exists():
            return

        total_height = max(1, self.checkbox_container.winfo_reqheight())
        target_y = first_row.winfo_y()
        if target_y <= 0:
            self.selection_canvas.yview_moveto(0)
            return

        ratio = target_y / total_height
        self.selection_canvas.yview_moveto(max(0.0, min(1.0, ratio)))

    def on_selection_changed(self, col, side):
        if side not in {"up", "down", "x"}:
            return

        if side == "x":
            x_count = sum(1 for c in self.selection_vars if self.selection_vars[c]["x"].get())
            if x_count > 1:
                messagebox.showwarning("X 軸限制", "只能選一個 X 軸變數")
                self.selection_vars[col]["x"].set(False)
                return

        top_count = sum(1 for c in self.selection_vars if self.selection_vars[c]["up"].get())
        bottom_count = sum(1 for c in self.selection_vars if self.selection_vars[c]["down"].get())

        if top_count > self.MAX_SERIES:
            messagebox.showwarning("超過上限", f"上視窗最多只能選 {self.MAX_SERIES} 條曲線")
            self.selection_vars[col]["up"].set(False)
            return
        if bottom_count > self.MAX_SERIES:
            messagebox.showwarning("超過上限", f"下視窗最多只能選 {self.MAX_SERIES} 條曲線")
            self.selection_vars[col]["down"].set(False)
            return

        self.refresh_all_plots()

    def refresh_all_plots(self):
        if self.df is None:
            return

        for panel in self.plot_panels:
            self._plot_panel(panel)

    def _plot_panel(self, panel):
        fig = panel["fig"]
        fig.clf()
        ax = fig.add_subplot(111)
        panel["ax"] = ax

        if panel["panel_index"] == 1:
            selected_columns = [col for col in self.selection_vars if self.selection_vars[col]["up"].get()]
        else:
            selected_columns = [col for col in self.selection_vars if self.selection_vars[col]["down"].get()]

        if not selected_columns:
            ax.text(0.5, 0.5, "請勾選至少 1 條曲線", ha="center", va="center", transform=ax.transAxes)
            ax.set_axis_off()
            fig.tight_layout()
            panel["canvas"].draw()
            return

        selected_columns = selected_columns[: self.MAX_SERIES]
        x_values = self._get_x_values(selected_columns)

        colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"]
        linewidth = max(0.5, float(self.linewidth_top_var.get() if panel["panel_index"] == 1 else self.linewidth_bottom_var.get()))
        left_handles = []
        right_handles = []
        if (self.dual_axis_top_var.get() if panel["panel_index"] == 1 else self.dual_axis_bottom_var.get()) and len(selected_columns) >= 2:
            ax2 = ax.twinx()
            for index, col in enumerate(selected_columns):
                y_values = self._prepare_series(self.df[col])
                color = colors[index % len(colors)]
                if index % 2 == 0:
                    line, = ax.plot(x_values, y_values, label=col, color=color, linewidth=linewidth)
                    left_handles.append(line)
                else:
                    line, = ax2.plot(x_values, y_values, label=col, color=color, linewidth=linewidth)
                    right_handles.append(line)
            ax.set_ylabel("Master Y axis")
            ax2.set_ylabel("Slave Y axis")
        else:
            for index, col in enumerate(selected_columns):
                y_values = self._prepare_series(self.df[col])
                color = colors[index % len(colors)]
                line, = ax.plot(x_values, y_values, label=col, color=color, linewidth=linewidth)
                left_handles.append(line)
            ax.set_ylabel("Y 軸")

        ax.set_title(panel["title_var"].get())
        ax.grid(True, alpha=0.3)
        use_dual_for_legend = self.dual_axis_top_var.get() if panel["panel_index"] == 1 else self.dual_axis_bottom_var.get()
        if use_dual_for_legend and len(selected_columns) >= 2:
            if left_handles:
                ax.legend(handles=left_handles, loc="upper left")
            if right_handles:
                ax2.legend(handles=right_handles, loc="upper right")
        else:
            ax.legend(loc="upper left")

        panel["cursor_vline"] = ax.axvline(color="gray", lw=0.8, alpha=0.7, visible=False)
        panel["cursor_hline"] = ax.axhline(color="gray", lw=0.8, alpha=0.7, visible=False)
        panel["cursor_label"].configure(text="Cursor: --")
        panel["ax2"] = None
        if (self.dual_axis_top_var.get() if panel["panel_index"] == 1 else self.dual_axis_bottom_var.get()) and len(selected_columns) >= 2:
            panel["ax2"] = ax.twinx() if len(selected_columns) >= 2 else None

        self._redraw_saved_points(panel, ax)

        try:
            self.update_point_list_display(panel)
        except Exception:
            pass

        fig.subplots_adjust(top=0.92, bottom=0.10)
        panel["canvas"].draw()

    def attach_plot_interaction(self, panel):
        canvas = panel.get("canvas")
        if canvas is None:
            return

        def motion(event, panel=panel):
            self.on_plot_motion(event, panel)

        def click(event, panel=panel):
            self.on_plot_click(event, panel)

        canvas.mpl_connect("motion_notify_event", motion)
        canvas.mpl_connect("button_press_event", click)
        canvas.mpl_connect("scroll_event", lambda event, panel=panel: self.on_plot_scroll(event, panel))

    def on_plot_motion(self, event, panel):
        panel_cursor_var = self.cursor_display_top_var if panel["panel_index"] == 1 else self.cursor_display_bottom_var
        if not panel_cursor_var.get() or event.inaxes is None or event.xdata is None or event.ydata is None:
            if panel.get("cursor_vline") is not None:
                panel["cursor_vline"].set_visible(False)
                panel["cursor_hline"].set_visible(False)
                panel["cursor_label"].configure(text="Cursor: off")
                panel["canvas"].draw_idle()
            return

        ax = panel["ax"]
        snapped = self._find_nearest_data_point(panel, event)
        if snapped is None:
            panel["cursor_vline"].set_visible(False)
            panel["cursor_hline"].set_visible(False)
            panel["cursor_label"].configure(text="Cursor: off")
            panel["canvas"].draw_idle()
            return

        x_value, y_value = snapped
        panel["cursor_vline"].set_visible(True)
        panel["cursor_hline"].set_visible(True)
        panel["cursor_vline"].set_xdata([x_value, x_value])
        panel["cursor_hline"].set_ydata([y_value, y_value])
        panel["cursor_label"].configure(text=f"Cursor: {self._format_plot_value(x_value)}, {self._format_plot_value(y_value)}")
        panel["canvas"].draw_idle()

    def on_plot_click(self, event, panel):
        if event.inaxes is None or event.xdata is None or event.ydata is None:
            return

        # If the toolbar zoom mode is active, ignore left-click additions
        toolbar = panel.get("toolbar")
        toolbar_mode = ""
        if toolbar is not None:
            toolbar_mode = getattr(toolbar, "mode", "") or ""

        if event.button == 1:
            if "zoom" in toolbar_mode.lower():
                # zoom active: do not add point with left click
                panel["cursor_label"].configure(text="Zoom active - left-click disabled")
                return

            snapped = self._find_nearest_data_point(panel, event)
            if snapped is None:
                x_value, y_value = event.xdata, event.ydata
            else:
                x_value, y_value = snapped

            pid = self._get_next_point_id(panel)
            point = {"id": pid, "x": x_value, "y": y_value}
            panel["saved_points"].append(point)
            self._draw_saved_point(panel, point["id"], point["x"], point["y"])
            panel["cursor_label"].configure(text=f"Added: {point['id']} {self._format_plot_value(point['x'])}, {self._format_plot_value(point['y'])}")
            panel["canvas"].draw_idle()
            self.update_point_list_display(panel)
        elif event.button == 3:
            index = self._find_nearest_saved_point_index(panel, event)
            if index is not None:
                removed = panel["saved_points"].pop(index)
                self._plot_panel(panel)
                panel["cursor_label"].configure(text=f"Removed: {removed['id']} {removed['x']:.3f}, {removed['y']:.3f}")
                self.update_point_list_display(panel)
            else:
                panel["cursor_label"].configure(text="No saved point nearby")

    def on_plot_scroll(self, event, panel):
        if event.inaxes is None:
            return

        ax = panel.get("ax")
        if ax is None:
            return

        # Determine scroll direction/step
        try:
            step = event.step
        except Exception:
            step = 1 if getattr(event, "button", "") in ("up", "wheel_up") else -1

        base_scale = 1.2
        if step > 0:
            scale = 1.0 / base_scale
        else:
            scale = base_scale

        xdata = event.xdata
        ydata = event.ydata
        cur_xlim = ax.get_xlim()
        cur_ylim = ax.get_ylim()

        # Compute new limits keeping mouse position as center point
        left = xdata - (xdata - cur_xlim[0]) * scale
        right = xdata + (cur_xlim[1] - xdata) * scale
        bottom = ydata - (ydata - cur_ylim[0]) * scale
        top = ydata + (cur_ylim[1] - ydata) * scale

        ax.set_xlim(left, right)
        ax.set_ylim(bottom, top)
        panel["canvas"].draw_idle()

    def _find_nearest_saved_point_index(self, panel, event):
        saved_points = panel.get("saved_points", [])
        if not saved_points:
            return None

        ax = panel["ax"]
        try:
            event_xy = ax.transData.transform((event.xdata, event.ydata))
        except Exception:
            return None

        best_index = None
        best_dist = 12.0
        for idx, pt in enumerate(saved_points):
            px, py = pt.get("x"), pt.get("y")
            saved_xy = ax.transData.transform((px, py))
            dist = math.hypot(event_xy[0] - saved_xy[0], event_xy[1] - saved_xy[1])
            if dist < best_dist:
                best_dist = dist
                best_index = idx

        return best_index

    def _find_nearest_data_point(self, panel, event):
        if event is None or event.inaxes is None or event.xdata is None or event.ydata is None:
            return None

        axis = event.inaxes
        lines = list(getattr(axis, "lines", []))
        if not lines:
            return None

        try:
            event_xy = axis.transData.transform((event.xdata, event.ydata))
        except Exception:
            return None

        best_point = None
        best_distance = None
        for line in lines:
            xdata = np.asarray(line.get_xdata())
            ydata = np.asarray(line.get_ydata())
            if len(xdata) == 0 or len(ydata) == 0:
                continue

            x_numeric = self._coerce_plot_value(xdata)
            y_numeric = self._coerce_plot_value(ydata)
            if x_numeric is None or y_numeric is None:
                continue

            if np.asarray(x_numeric).shape != np.asarray(y_numeric).shape:
                continue

            dx = np.asarray(x_numeric, dtype=float) - float(event.xdata)
            dy = np.asarray(y_numeric, dtype=float) - float(event.ydata)
            distances = np.hypot(dx, dy)
            idx = int(np.argmin(distances))
            point_xy = axis.transData.transform((float(x_numeric[idx]), float(y_numeric[idx])))
            distance_px = math.hypot(event_xy[0] - point_xy[0], event_xy[1] - point_xy[1])
            if best_distance is None or distance_px < best_distance:
                best_distance = distance_px
                best_point = (float(x_numeric[idx]), float(y_numeric[idx]))

        return best_point

    def _coerce_plot_value(self, values):
        if values is None:
            return None
        try:
            array = np.asarray(values)
        except Exception:
            return None

        if array.size == 0:
            return None

        if np.issubdtype(array.dtype, np.datetime64):
            return mdates.date2num(array)

        if array.dtype == object:
            converted = []
            for item in array:
                if item is None:
                    converted.append(np.nan)
                elif hasattr(item, "to_pydatetime"):
                    converted.append(mdates.date2num(item.to_pydatetime()))
                else:
                    try:
                        converted.append(float(item))
                    except Exception:
                        converted.append(np.nan)
            return np.asarray(converted, dtype=float)

        try:
            return array.astype(float)
        except Exception:
            return None

    def _format_plot_value(self, value):
        if value is None:
            return ""
        if isinstance(value, (np.datetime64,)):
            try:
                return pd.Timestamp(value).strftime("%Y-%m-%d")
            except Exception:
                return str(value)
        if hasattr(value, "to_pydatetime"):
            try:
                return value.to_pydatetime().strftime("%Y-%m-%d")
            except Exception:
                return str(value)
        if isinstance(value, (float, np.floating)):
            return f"{float(value):.3f}"
        return str(value)

    def _draw_saved_point(self, panel, pid, x, y):
        ax = panel["ax"]
        ax.plot(x, y, marker="o", color="black", markersize=5, markerfacecolor="yellow")
        ax.annotate(
            f"{pid}:({x:.3f}, {y:.3f})",
            xy=(x, y),
            xytext=(6, 6),
            textcoords="offset points",
            fontsize=8,
            color="black",
            bbox=dict(boxstyle="round,pad=0.2", fc="white", alpha=0.8),
        )

    def _redraw_saved_points(self, panel, ax):
        for pt in panel.get("saved_points", []):
            x, y = pt.get("x"), pt.get("y")
            pid = pt.get("id")
            ax.plot(x, y, marker="o", color="black", markersize=5, markerfacecolor="yellow")
            ax.annotate(
                f"{pid}:({x:.3f}, {y:.3f})",
                xy=(x, y),
                xytext=(6, 6),
                textcoords="offset points",
                fontsize=8,
                color="black",
                bbox=dict(boxstyle="round,pad=0.2", fc="white", alpha=0.8),
            )
        

    def _get_next_point_id(self, panel):
        # find smallest missing integer index starting from 1
        existing = set()
        for pt in panel.get("saved_points", []):
            pid = pt.get("id", "")
            if pid.startswith("NO"):
                try:
                    n = int(pid[2:])
                    existing.add(n)
                except Exception:
                    continue
        i = 1
        while True:
            if i not in existing:
                return f"NO{i}"
            i += 1

    def update_point_list_display(self, panel):
        # update the right-side listbox for top/bottom
        idx = panel.get("panel_index", 1)
        lst = self.points_list_top if idx == 1 else self.points_list_bottom
        lst.delete(0, tk.END)
        for pt in panel.get("saved_points", []):
            pid = pt.get("id")
            x = pt.get("x")
            y = pt.get("y")
            lst.insert(tk.END, f"{pid}: {x:.3f}, {y:.3f}")
    def _get_x_values(self, selected_columns):
        x_cols = [col for col in self.selection_vars if self.selection_vars[col]["x"].get()]
        if x_cols:
            x_col = x_cols[0]
            values = self.df[x_col]
            numeric = pd.to_numeric(values, errors="coerce")
            if numeric.notna().any():
                return numeric.to_numpy(dtype=float)
            try:
                return pd.to_datetime(values).to_numpy()
            except Exception:
                pass

        for col in self.df.columns:
            if str(col) in selected_columns:
                continue
            values = self.df[col]
            numeric = pd.to_numeric(values, errors="coerce")
            if numeric.notna().any():
                return numeric.to_numpy(dtype=float)
            try:
                return pd.to_datetime(values).to_numpy()
            except Exception:
                continue
        return np.arange(len(self.df), dtype=float)

    def _prepare_series(self, series):
        values = pd.to_numeric(series, errors="coerce")
        if values.notna().any():
            return values.to_numpy(dtype=float)
        try:
            return pd.to_datetime(series).to_numpy()
        except Exception:
            return np.asarray(series)

    def apply_zoom(self, panel_index=None):
        # panel_index: None -> apply to all; 1 -> top only; 2 -> bottom only
        if panel_index == 1:
            xmin = self.x_min_top_var.get().strip()
            xmax = self.x_max_top_var.get().strip()
        elif panel_index == 2:
            xmin = self.x_min_bottom_var.get().strip()
            xmax = self.x_max_bottom_var.get().strip()
        else:
            # apply per-panel inputs for each panel
            xmin = None
            xmax = None

        if not xmin or not xmax:
            messagebox.showwarning("輸入錯誤", "請輸入 X 範圍的開始與結束值")
            return

        try:
            if panel_index == 1:
                selected_columns = [col for col in self.selection_vars if self.selection_vars[col]["up"].get()]
                panels = [self.plot_panels[0]]
            elif panel_index == 2:
                selected_columns = [col for col in self.selection_vars if self.selection_vars[col]["down"].get()]
                panels = [self.plot_panels[1]]
            else:
                selected_columns = [col for col in self.selection_vars if self.selection_vars[col]["up"].get() or self.selection_vars[col]["down"].get()]
                panels = self.plot_panels

            if panel_index in (1, 2):
                x_data = self._get_x_values(selected_columns)
                xmin_val = self._parse_zoom_value(xmin, x_data)
                xmax_val = self._parse_zoom_value(xmax, x_data)
                for panel in panels:
                    if panel.get("ax") is not None:
                        panel["ax"].set_xlim(xmin_val, xmax_val)
                        panel["canvas"].draw()
            else:
                # apply per-panel limits from their respective inputs
                for i, panel in enumerate(panels, start=1):
                    if i == 1:
                        xmin_s = self.x_min_top_var.get().strip()
                        xmax_s = self.x_max_top_var.get().strip()
                    else:
                        xmin_s = self.x_min_bottom_var.get().strip()
                        xmax_s = self.x_max_bottom_var.get().strip()
                    if not xmin_s or not xmax_s:
                        continue
                    x_cols = [col for col in self.selection_vars if self.selection_vars[col]["up"].get() or self.selection_vars[col]["down"].get()]
                    x_data = self._get_x_values(x_cols)
                    xmin_val = self._parse_zoom_value(xmin_s, x_data)
                    xmax_val = self._parse_zoom_value(xmax_s, x_data)
                    if panel.get("ax") is not None:
                        panel["ax"].set_xlim(xmin_val, xmax_val)
                        panel["canvas"].draw()
        except Exception as exc:
            messagebox.showerror("放大失敗", f"無法設定 X 範圍：{exc}")

    def reset_zoom(self, panel_index=None):
        # reset zoom for specific panel or all
        if panel_index == 1:
            panels = [self.plot_panels[0]]
        elif panel_index == 2:
            panels = [self.plot_panels[1]]
        else:
            panels = self.plot_panels

        if panel_index == 1:
            self.x_min_top_var.set("")
            self.x_max_top_var.set("")
        elif panel_index == 2:
            self.x_min_bottom_var.set("")
            self.x_max_bottom_var.set("")
        else:
            self.x_min_top_var.set("")
            self.x_max_top_var.set("")
            self.x_min_bottom_var.set("")
            self.x_max_bottom_var.set("")

        for panel in panels:
            if panel.get("ax") is not None:
                panel["ax"].relim()
                panel["ax"].autoscale_view()
                panel["canvas"].draw()

    def _parse_zoom_value(self, text, x_data):
        try:
            return float(text)
        except ValueError:
            return pd.to_datetime(text).to_numpy()

    def reset_zoom_all(self):
        self.x_min_top_var.set("")
        self.x_max_top_var.set("")
        self.x_min_bottom_var.set("")
        self.x_max_bottom_var.set("")
        for panel in self.plot_panels:
            if panel.get("ax") is not None:
                panel["ax"].relim()
                panel["ax"].autoscale_view()
                panel["canvas"].draw()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelPlotApp(root)
    root.mainloop()
