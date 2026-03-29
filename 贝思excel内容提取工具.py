#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
贝思excel内容提取工具
支持多文件、多表格、百万行数据、拖拽、进度显示、智能表头合并
"""

import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import re

# 尝试导入 tkinterdnd2（拖拽支持）
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False


class ExcelExtractorApp:
    def __init__(self):
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title("贝思excel内容提取工具")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        self.root.configure(bg="#f0f4f8")

        self.file_list = []  # 已添加的文件路径列表
        self.is_running = False
        self.cancel_flag = False

        self._build_ui()

    # ------------------------------------------------------------------ UI --
    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", font=("微软雅黑", 10), padding=6)
        style.configure("TLabel", background="#f0f4f8", font=("微软雅黑", 10))
        style.configure("TLabelframe", background="#f0f4f8")
        style.configure("TLabelframe.Label", background="#f0f4f8", font=("微软雅黑", 10, "bold"))
        style.configure("green.Horizontal.TProgressbar", troughcolor="#dce3eb", background="#4caf50")

        main_frame = tk.Frame(self.root, bg="#f0f4f8", padx=12, pady=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ── 文件区域 ──────────────────────────────────────────────────────────
        file_frame = ttk.LabelFrame(main_frame, text=" 📂 文件列表 ", padding=8)
        file_frame.pack(fill=tk.BOTH, expand=True)

        btn_bar = tk.Frame(file_frame, bg="#f0f4f8")
        btn_bar.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(btn_bar, text="➕ 添加文件", command=self._add_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_bar, text="📁 添加文件夹", command=self._add_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_bar, text="🗑 清空列表", command=self._clear_files).pack(side=tk.LEFT)

        list_container = tk.Frame(file_frame, bg="#f0f4f8")
        list_container.pack(fill=tk.BOTH, expand=True)

        self.file_listbox = tk.Listbox(
            list_container,
            selectmode=tk.EXTENDED,
            bg="#ffffff",
            fg="#333",
            font=("微软雅黑", 9),
            relief=tk.FLAT,
            highlightthickness=1,
            highlightcolor="#4caf50",
            highlightbackground="#ccd",
            height=8,
        )
        sb_x = ttk.Scrollbar(list_container, orient=tk.HORIZONTAL, command=self.file_listbox.xview)
        sb_y = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(xscrollcommand=sb_x.set, yscrollcommand=sb_y.set)
        sb_y.pack(side=tk.RIGHT, fill=tk.Y)
        sb_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        dnd_hint = "（支持拖拽文件/文件夹到此区域）" if HAS_DND else "（安装 tkinterdnd2 后可支持拖拽）"
        tk.Label(file_frame, text=dnd_hint, bg="#f0f4f8", fg="#888", font=("微软雅黑", 8)).pack(anchor=tk.W)

        if HAS_DND:
            self.file_listbox.drop_target_register(DND_FILES)
            self.file_listbox.dnd_bind("<<Drop>>", self._on_drop)

        # ── 关键词区域 ────────────────────────────────────────────────────────
        kw_frame = ttk.LabelFrame(main_frame, text=" 🔍 搜索关键词（每行一个） ", padding=8)
        kw_frame.pack(fill=tk.X, pady=(10, 0))

        self.keyword_text = tk.Text(
            kw_frame, height=5, font=("微软雅黑", 10),
            bg="#ffffff", relief=tk.FLAT,
            highlightthickness=1, highlightbackground="#ccd",
        )
        kw_sb = ttk.Scrollbar(kw_frame, orient=tk.VERTICAL, command=self.keyword_text.yview)
        self.keyword_text.configure(yscrollcommand=kw_sb.set)
        kw_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.keyword_text.pack(fill=tk.X)

        # ── 选项区域 ──────────────────────────────────────────────────────────
        opt_frame = tk.Frame(main_frame, bg="#f0f4f8")
        opt_frame.pack(fill=tk.X, pady=(8, 0))

        self.match_mode = tk.StringVar(value="contains")
        tk.Label(opt_frame, text="匹配方式：", bg="#f0f4f8", font=("微软雅黑", 10)).pack(side=tk.LEFT)
        ttk.Radiobutton(opt_frame, text="包含", variable=self.match_mode, value="contains").pack(side=tk.LEFT, padx=4)
        ttk.Radiobutton(opt_frame, text="完全匹配", variable=self.match_mode, value="exact").pack(side=tk.LEFT, padx=4)
        ttk.Radiobutton(opt_frame, text="开头匹配", variable=self.match_mode, value="startswith").pack(side=tk.LEFT, padx=4)

        self.case_sensitive = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text="区分大小写", variable=self.case_sensitive).pack(side=tk.LEFT, padx=12)

        # ── 输出路径 ──────────────────────────────────────────────────────────
        out_frame = tk.Frame(main_frame, bg="#f0f4f8")
        out_frame.pack(fill=tk.X, pady=(8, 0))

        tk.Label(out_frame, text="输出文件：", bg="#f0f4f8", font=("微软雅黑", 10)).pack(side=tk.LEFT)
        self.output_var = tk.StringVar(value=str(Path.home() / "贝思提取结果.xlsx"))
        tk.Entry(out_frame, textvariable=self.output_var, font=("微软雅黑", 9), width=50).pack(side=tk.LEFT, padx=6)
        ttk.Button(out_frame, text="浏览", command=self._choose_output).pack(side=tk.LEFT)

        # ── 进度区域 ──────────────────────────────────────────────────────────
        prog_frame = ttk.LabelFrame(main_frame, text=" ⚡ 进度 ", padding=8)
        prog_frame.pack(fill=tk.X, pady=(10, 0))

        self.progress_bar = ttk.Progressbar(
            prog_frame, style="green.Horizontal.TProgressbar",
            orient=tk.HORIZONTAL, length=100, mode="determinate"
        )
        self.progress_bar.pack(fill=tk.X)

        info_row = tk.Frame(prog_frame, bg="#f0f4f8")
        info_row.pack(fill=tk.X, pady=(4, 0))
        self.progress_label = tk.Label(info_row, text="就绪", bg="#f0f4f8", fg="#555", font=("微软雅黑", 9))
        self.progress_label.pack(side=tk.LEFT)
        self.speed_label = tk.Label(info_row, text="", bg="#f0f4f8", fg="#4caf50", font=("微软雅黑", 9, "bold"))
        self.speed_label.pack(side=tk.RIGHT)

        # ── 操作按钮 ──────────────────────────────────────────────────────────
        action_frame = tk.Frame(main_frame, bg="#f0f4f8")
        action_frame.pack(pady=(12, 0))

        self.start_btn = ttk.Button(action_frame, text="▶  开始提取", command=self._start_extract)
        self.start_btn.pack(side=tk.LEFT, padx=8)
        self.cancel_btn = ttk.Button(action_frame, text="⏹  取消", command=self._cancel, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=8)

    # --------------------------------------------------------- 文件管理 ------
    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")]
        )
        for p in paths:
            self._add_path(p)

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            for root_dir, _, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith((".xlsx", ".xls", ".xlsm")):
                        self._add_path(os.path.join(root_dir, f))

    def _add_path(self, path):
        if path not in self.file_list:
            self.file_list.append(path)
            self.file_listbox.insert(tk.END, path)

    def _clear_files(self):
        self.file_list.clear()
        self.file_listbox.delete(0, tk.END)

    def _on_drop(self, event):
        raw = event.data
        # tkinterdnd2 返回的路径可能被 {} 包裹
        paths = self.root.tk.splitlist(raw)
        for p in paths:
            p = p.strip()
            if os.path.isdir(p):
                for root_dir, _, files in os.walk(p):
                    for f in files:
                        if f.lower().endswith((".xlsx", ".xls", ".xlsm")):
                            self._add_path(os.path.join(root_dir, f))
            elif os.path.isfile(p) and p.lower().endswith((".xlsx", ".xls", ".xlsm")):
                self._add_path(p)

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title="保存结果文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if path:
            self.output_var.set(path)

    # --------------------------------------------------------- 提取逻辑 ------
    def _get_keywords(self):
        raw = self.keyword_text.get("1.0", tk.END)
        kws = [line.strip() for line in raw.splitlines() if line.strip()]
        return kws

    def _row_matches(self, row_values, keywords, mode, case):
        """检查行是否匹配任意关键词"""
        cell_strs = [str(v) if v is not None else "" for v in row_values]
        row_str = " ".join(cell_strs)
        if not case:
            row_str_cmp = row_str.lower()
        else:
            row_str_cmp = row_str

        for kw in keywords:
            kw_cmp = kw if case else kw.lower()
            if mode == "contains" and kw_cmp in row_str_cmp:
                return True
            elif mode == "exact":
                if any((v.strip() if case else v.strip().lower()) == kw_cmp
                       for v in cell_strs):
                    return True
            elif mode == "startswith":
                if any((v if case else v.lower()).startswith(kw_cmp)
                       for v in cell_strs):
                    return True
        return False

    def _start_extract(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加Excel文件！")
            return
        kws = self._get_keywords()
        if not kws:
            messagebox.showwarning("提示", "请输入至少一个搜索关键词！")
            return
        output_path = self.output_var.get().strip()
        if not output_path:
            messagebox.showwarning("提示", "请设置输出文件路径！")
            return

        self.is_running = True
        self.cancel_flag = False
        self.start_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0

        thread = threading.Thread(
            target=self._extract_worker,
            args=(list(self.file_list), kws, output_path,
                  self.match_mode.get(), self.case_sensitive.get()),
            daemon=True
        )
        thread.start()

    def _cancel(self):
        self.cancel_flag = True
        self._update_status("正在取消...")

    def _update_status(self, msg, progress=None, speed=None):
        def _do():
            self.progress_label.config(text=msg)
            if progress is not None:
                self.progress_bar["value"] = progress
            if speed is not None:
                self.speed_label.config(text=speed)
        self.root.after(0, _do)

    def _finish(self, success, output_path=None):
        def _do():
            self.is_running = False
            self.start_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)
            self.speed_label.config(text="")
            if success:
                self.progress_bar["value"] = 100
                if messagebox.askyesno("完成", f"提取完成！\n结果已保存至：\n{output_path}\n\n是否立即打开？"):
                    os.startfile(output_path) if sys.platform == "win32" else os.system(f'open "{output_path}"')
            else:
                self.progress_bar["value"] = 0
        self.root.after(0, _do)

    # --------------------------------------------------------- 工作线程 ------
    def _extract_worker(self, files, keywords, output_path, mode, case):
        try:
            total_files = len(files)
            # 结果字典：header_key -> list of rows
            # header_key 是列标题的 tuple
            results = {}          # {header_tuple: [row_list, ...]}
            header_order = []     # 保持插入顺序

            start_time = time.time()
            total_rows_scanned = 0
            total_rows_matched = 0

            for file_idx, filepath in enumerate(files):
                if self.cancel_flag:
                    self._update_status("已取消")
                    self._finish(False)
                    return

                fname = os.path.basename(filepath)
                self._update_status(
                    f"正在读取 [{file_idx+1}/{total_files}]：{fname}",
                    progress=int(file_idx / total_files * 90)
                )

                try:
                    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
                except Exception as e:
                    self._update_status(f"跳过（无法读取）：{fname} — {e}")
                    continue

                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    rows_iter = ws.iter_rows(values_only=True)

                    # 读取表头（第一行）
                    try:
                        header_row = next(rows_iter)
                    except StopIteration:
                        continue

                    header = tuple(str(h) if h is not None else "" for h in header_row)

                    if header not in results:
                        results[header] = []
                        header_order.append(header)

                    row_num = 0
                    batch_rows = []
                    for row in rows_iter:
                        if self.cancel_flag:
                            break
                        row_num += 1
                        total_rows_scanned += 1
                        if self._row_matches(row, keywords, mode, case):
                            batch_rows.append(list(row))
                            total_rows_matched += 1

                        # 每 5000 行更新一次速度
                        if row_num % 5000 == 0:
                            elapsed = time.time() - start_time
                            speed = total_rows_scanned / elapsed if elapsed > 0 else 0
                            self._update_status(
                                f"[{file_idx+1}/{total_files}] {fname} / {sheet_name}  "
                                f"已扫描 {total_rows_scanned:,} 行，命中 {total_rows_matched:,} 行",
                                progress=int(file_idx / total_files * 90),
                                speed=f"{speed:,.0f} 行/秒"
                            )

                    results[header].extend(batch_rows)
                    wb.close()

                if self.cancel_flag:
                    self._update_status("已取消")
                    self._finish(False)
                    return

            # ── 写出结果 ──────────────────────────────────────────────────────
            self._update_status("正在写入结果文件...", progress=92)
            out_wb = Workbook()
            out_ws = out_wb.active
            out_ws.title = "提取结果"

            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            center = Alignment(horizontal="center", vertical="center")

            current_row = 1
            first_header = True

            for hdr in header_order:
                matched_rows = results[hdr]
                if not matched_rows:
                    continue

                # 不同表头之间插入空行（第一个不插）
                if not first_header:
                    current_row += 1  # 空行
                first_header = False

                # 写表头
                for col_idx, col_name in enumerate(hdr, start=1):
                    cell = out_ws.cell(row=current_row, column=col_idx, value=col_name)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center
                current_row += 1

                # 写数据
                for data_row in matched_rows:
                    for col_idx, val in enumerate(data_row, start=1):
                        out_ws.cell(row=current_row, column=col_idx, value=val)
                    current_row += 1

                    if current_row % 10000 == 0:
                        self._update_status(f"写入中... 已写 {current_row:,} 行", progress=95)

            out_wb.save(output_path)

            elapsed = time.time() - start_time
            self._update_status(
                f"完成！共扫描 {total_rows_scanned:,} 行，提取 {total_rows_matched:,} 行，耗时 {elapsed:.1f} 秒",
                progress=100,
                speed=f"{total_rows_scanned/elapsed:,.0f} 行/秒" if elapsed > 0 else ""
            )
            self._finish(True, output_path)

        except Exception as e:
            import traceback
            self._update_status(f"错误：{e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"提取过程出错：\n{traceback.format_exc()}"))
            self._finish(False)

    # ----------------------------------------------------------------- 运行 --
    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = ExcelExtractorApp()
    app.run()
