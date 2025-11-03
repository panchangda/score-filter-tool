# app_gui.py
# -*- coding: utf-8 -*-
import sys
import os
from pathlib import Path
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from score_filter_core import process_file, DEFAULT_PUBCLASS_QUALIFIED_NUM

class App:
    def __init__(self, root):
        self.root = root
        root.title("学业预警筛选工具（GUI）")
        root.geometry("820x520")

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        # 文件选择
        file_row = ttk.Frame(frm)
        file_row.pack(fill=tk.X, pady=4)
        ttk.Label(file_row, text="输入 Excel 文件:").pack(side=tk.LEFT)
        self.path_var = tk.StringVar()
        self.entry_file = ttk.Entry(file_row, textvariable=self.path_var, width=70)
        self.entry_file.pack(side=tk.LEFT, padx=6)
        ttk.Button(file_row, text="浏览", command=self.browse_file).pack(side=tk.LEFT)

        # 参数
        param_row = ttk.Frame(frm)
        param_row.pack(fill=tk.X, pady=4)
        ttk.Label(param_row, text="公选课学分阈值：").pack(side=tk.LEFT)
        self.pubspin = tk.Spinbox(param_row, from_=1, to=100, width=6)
        self.pubspin.delete(0, tk.END)
        self.pubspin.insert(0, str(DEFAULT_PUBCLASS_QUALIFIED_NUM))
        self.pubspin.pack(side=tk.LEFT, padx=6)
        self.divide_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(param_row, text="分开输出各规则 CSV（divide_output）", variable=self.divide_var)\
            .pack(side=tk.LEFT, padx=10)

        # 按钮
        btn_row = ttk.Frame(frm)
        btn_row.pack(fill=tk.X, pady=6)
        self.run_btn = ttk.Button(btn_row, text="运行（Run）", command=self.on_run)
        self.run_btn.pack(side=tk.LEFT)
        ttk.Button(btn_row, text="打开输出目录", command=self.open_out_dir).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="退出", command=root.quit).pack(side=tk.RIGHT)

        # 日志
        ttk.Label(frm, text="日志：").pack(anchor=tk.W, pady=(8,0))
        self.txt = tk.Text(frm, height=18)
        self.txt.pack(fill=tk.BOTH, expand=True)
        self.txt.configure(state=tk.DISABLED)

        # 状态栏
        self.status_var = tk.StringVar(value="Ready")
        status = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status.pack(fill=tk.X, side=tk.BOTTOM)

        self._last_out_dir = None

    # ---------- UI handlers ----------
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="请选择成绩表（Excel）",
            filetypes=[("Excel files", ".xlsx .xls"), ("All files", "*.*")]
        )
        if filename:
            self.path_var.set(filename)

    def log(self, s=""):
        self.txt.configure(state=tk.NORMAL)
        self.txt.insert(tk.END, s + "\n")
        self.txt.see(tk.END)
        self.txt.configure(state=tk.DISABLED)

    def on_run(self):
        infile = self.path_var.get().strip()
        if not infile:
            messagebox.showwarning("未选择文件", "请先选择输入 Excel 文件。")
            return
        if not Path(infile).exists():
            messagebox.showerror("文件不存在", "指定的输入文件不存在，请检查路径。")
            return

        try:
            pub_val = int(self.pubspin.get())
        except Exception:
            pub_val = DEFAULT_PUBCLASS_QUALIFIED_NUM

        divide_output = bool(self.divide_var.get())
        self.run_btn.config(state=tk.DISABLED)
        self.status_var.set("运行中...")

        def worker():
            success, msg = process_file(infile, pub_val, divide_output, log_fn=self.log)
            self.run_btn.config(state=tk.NORMAL)
            self.status_var.set("就绪" if success else "出错（请查看日志）")
            if success:
                messagebox.showinfo("完成", "处理完成！输出已写入输入文件同一目录。")
                self._last_out_dir = str(Path(infile).parent)
            else:
                messagebox.showerror("运行出错", "处理时发生错误，请查看日志。")
        threading.Thread(target=worker, daemon=True).start()

    def open_out_dir(self):
        path = self._last_out_dir or (Path(self.path_var.get()).parent if self.path_var.get() else None)
        if path and Path(path).exists():
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f"open '{path}'")
            else:
                os.system(f"xdg-open '{path}'")
        else:
            messagebox.showinfo("无可打开目录", "还没有输出目录，请先运行一次处理。")
