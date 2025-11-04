# app_gui.py
# -*- coding: utf-8 -*-
import sys
import os
import webbrowser
from pathlib import Path
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from score_filter_core import (
    process_files, process_one_file,
    DEFAULT_PUBCLASS_QUALIFIED_NUM, SUPPORTED_EXTS
)

# 在这里放你的 GitHub 仓库链接（可点击打开）
GITHUB_URL = "https://github.com/panchangda/score-filter-tool"  # TODO: 替换为你的实际地址

class App:
    def __init__(self, root):
        self.root = root
        root.title("学业预警筛选工具（GUI）")
        root.geometry("900x800")

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        # ========== 文件区 ==========
        file_frame = ttk.LabelFrame(frm, text="输入文件（可多选）", padding=8)
        file_frame.pack(fill=tk.BOTH, expand=False)

        btns_row = ttk.Frame(file_frame)
        btns_row.pack(fill=tk.X, pady=4)
        ttk.Button(btns_row, text="添加文件…", command=self.add_files).pack(side=tk.LEFT)
        ttk.Button(btns_row, text="添加文件夹…", command=self.add_folder).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns_row, text="移除选中", command=self.remove_selected).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns_row, text="清空列表", command=self.clear_list).pack(side=tk.LEFT, padx=6)
        # 新增：介绍按钮
        ttk.Button(btns_row, text="介绍", command=self.show_about).pack(side=tk.RIGHT)

        self.listbox = tk.Listbox(file_frame, selectmode=tk.EXTENDED, height=8)
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        # ========== 输出设置 ==========
        out_frame = ttk.LabelFrame(frm, text="输出设置", padding=8)
        out_frame.pack(fill=tk.X, expand=False, pady=(10, 0))

        self.use_input_dir = tk.BooleanVar(value=True)
        cb = ttk.Checkbutton(out_frame, text="输出到各输入文件所在目录（默认）", variable=self.use_input_dir, command=self._toggle_outdir_state)
        cb.pack(anchor=tk.W)

        row_out = ttk.Frame(out_frame)
        row_out.pack(fill=tk.X, pady=4)
        ttk.Label(row_out, text="统一输出目录（可选）：").pack(side=tk.LEFT)

        self.outdir_var = tk.StringVar()
        self.outdir_entry = ttk.Entry(row_out, textvariable=self.outdir_var, width=70, state=tk.DISABLED)
        self.outdir_entry.pack(side=tk.LEFT, padx=6)
        ttk.Button(row_out, text="浏览", command=self.browse_outdir).pack(side=tk.LEFT)

        # ========== 参数区 ==========
        param_row = ttk.LabelFrame(frm, text="处理参数", padding=8)
        param_row.pack(fill=tk.X, expand=False, pady=(10, 0))

        row1 = ttk.Frame(param_row)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="公选课学分阈值：").pack(side=tk.LEFT)
        self.pubspin = tk.Spinbox(row1, from_=1, to=100, width=6)
        self.pubspin.delete(0, tk.END)
        self.pubspin.insert(0, str(DEFAULT_PUBCLASS_QUALIFIED_NUM))
        self.pubspin.pack(side=tk.LEFT, padx=6)

        self.divide_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row1, text="分开输出各规则 CSV（divide_output）", variable=self.divide_var)\
            .pack(side=tk.LEFT, padx=10)

        # ========== 操作区 ==========
        ctl_row = ttk.Frame(frm)
        ctl_row.pack(fill=tk.X, pady=10)
        self.run_btn = ttk.Button(ctl_row, text="批量运行（Run）", command=self.on_run)
        self.run_btn.pack(side=tk.LEFT)
        ttk.Button(ctl_row, text="退出", command=root.quit).pack(side=tk.RIGHT)

        # 进度条
        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(0, 6))

        # 日志
        ttk.Label(frm, text="日志：").pack(anchor=tk.W, pady=(4, 0))
        self.txt = tk.Text(frm, height=14)
        self.txt.pack(fill=tk.BOTH, expand=True)
        self.txt.configure(state=tk.DISABLED)

        # 记录最近输出位置（用于“打开目录”扩展时）
        self._last_outdir_used = None

    # ---------- 文件区操作 ----------
    def add_files(self):
        filenames = filedialog.askopenfilenames(
            title="选择 Excel 文件（可多选）",
            filetypes=[("Excel files", ".xlsx .xls"), ("All files", "*.*")]
        )
        if not filenames: return
        for fn in filenames:
            p = Path(fn)
            if p.suffix.lower() in SUPPORTED_EXTS and str(p) not in self.listbox.get(0, tk.END):
                self.listbox.insert(tk.END, str(p))

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if not folder: return
        count = 0
        for ext in SUPPORTED_EXTS:
            for p in Path(folder).glob(f"*{ext}"):
                if str(p) not in self.listbox.get(0, tk.END):
                    self.listbox.insert(tk.END, str(p))
                    count += 1
        if count == 0:
            messagebox.showinfo("未发现文件", "该文件夹内未找到 .xlsx/.xls 文件。")

    def remove_selected(self):
        sel = list(self.listbox.curselection())
        if not sel: return
        sel.reverse()
        for i in sel:
            self.listbox.delete(i)

    def clear_list(self):
        self.listbox.delete(0, tk.END)

    # ---------- 输出目录 ----------
    def _toggle_outdir_state(self):
        if self.use_input_dir.get():
            self.outdir_entry.configure(state=tk.DISABLED)
        else:
            self.outdir_entry.configure(state=tk.NORMAL)

    def browse_outdir(self):
        d = filedialog.askdirectory(title="选择统一输出目录")
        if d:
            self.outdir_var.set(d)
            self.use_input_dir.set(False)
            self._toggle_outdir_state()

    # ---------- 日志 ----------
    def log(self, s=""):
        self.txt.configure(state=tk.NORMAL)
        self.txt.insert(tk.END, s + "\n")
        self.txt.see(tk.END)
        self.txt.configure(state=tk.DISABLED)

    # ---------- 介绍 / 关于 ----------
    def show_about(self):
        # 用 Toplevel 做一个简单“介绍”窗口，可点击打开 GitHub
        top = tk.Toplevel(self.root)
        top.title("介绍 / 关于")
        top.geometry("700x500")
        top.resizable(True, True)

        container = ttk.Frame(top, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        text = (
            "【工具说明】\n"
            "本工具用于对班级成绩表进行筛选与学业预警导出，核心规则如下：\n\n"
            "规则一（公选未达标）：\n"
            "  - 条件：课程为“公共选修课”，且“获得学分” < 阈值，且“成绩” < 60 或空白。\n"
            "  - 作用：识别公选课学分不足或成绩不达标的记录。\n\n"
            "规则二（疑似未开设 / 0分需关注）：\n"
            "  - 对非公选课程：若该课程全班成绩均为 0 或空白，则视为“未开设”（不导出记录，仅在日志与汇总中列出课程名）。\n"
            "  - 否则，非公选课程中成绩为 0 或空白的记录作为“需关注”导出。\n\n"
            "规则三（非公选不及格）：\n"
            "  - 对非公选课程：成绩在 (0, 60) 区间的记录视为不及格并导出。\n\n"
            "其它处理：\n"
            "  - 若存在列“一层节点”为“其它”的行，会被过滤。\n"
            "  - 合并导出前会按 [学号, 课程名称, 学期(可选列)] 去重。\n"
            "  - 可选将三类结果分别另存 CSV。\n\n"
            "字段要求（列名）：\n"
            "  - 第一列作为学号；需包含：一层节点、课程名称、获得学分、成绩（大小写一致）。\n\n"
            "GitHub：点击下方链接打开仓库地址。"
        )

        lbl = ttk.Label(container, text=text, justify=tk.LEFT, anchor=tk.NW)
        lbl.pack(fill=tk.BOTH, expand=True)

        link = ttk.Label(container, text=GITHUB_URL, foreground="blue", cursor="hand2")
        link.pack(anchor=tk.W, pady=(8, 4))
        link.bind("<Button-1>", lambda e: webbrowser.open_new(GITHUB_URL))

        btn_row = ttk.Frame(container)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text="关闭", command=top.destroy).pack(side=tk.RIGHT)

    # ---------- 运行 ----------
    def on_run(self):
        files = list(self.listbox.get(0, tk.END))
        if not files:
            messagebox.showwarning("未选择文件", "请先添加至少一个 Excel 文件。")
            return

        try:
            pub_val = int(self.pubspin.get())
        except Exception:
            pub_val = DEFAULT_PUBCLASS_QUALIFIED_NUM

        divide_output = bool(self.divide_var.get())

        # 输出目录策略
        output_dir = None
        if not self.use_input_dir.get():
            out = self.outdir_var.get().strip()
            if not out:
                messagebox.showwarning("未选择输出目录", "请先选择统一输出目录或勾选“输出到输入文件同目录”。")
                return
            outp = Path(out)
            try:
                outp.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("创建输出目录失败", f"无法创建目录：{outp}\n{e}")
                return
            output_dir = str(outp)

        # UI状态
        self.run_btn.config(state=tk.DISABLED)
        self.progress.start(10)

        def worker():
            self.log(f"开始批量处理（{len(files)} 个文件）...")
            ok, combined, results = process_files(
                files, pubclass_qualified_num=pub_val,
                divide_output=divide_output, output_dir=output_dir, log_fn=self.log
            )
            self.progress.stop()
            self.run_btn.config(state=tk.NORMAL)

            # 记录最近输出目录
            if output_dir:
                self._last_outdir_used = output_dir
            else:
                try:
                    self._last_outdir_used = str(Path(files[0]).parent)
                except Exception:
                    self._last_outdir_used = None

            # 弹窗与日志
            self.log("\n=== 批量汇总 ===")
            self.log(combined)
            if ok:
                messagebox.showinfo("完成", "所有文件处理完毕！")
            else:
                messagebox.showwarning("部分失败", "已完成，但部分文件处理失败，请查看日志。")

        threading.Thread(target=worker, daemon=True).start()
