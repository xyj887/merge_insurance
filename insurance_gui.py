# -*- coding: utf-8 -*-
"""
投保数据汇总工具 - 图形界面版

给不懂命令行的同事使用：
  1. 点"选择文件"挑一个投保数据Excel
  2. 点"开始处理"
  3. 处理完在原文件旁边生成 "<原文件名>_汇总.xlsx"

数据全程在本机处理，不上传任何服务器。
"""

import os
import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from merge_insurance import build_result_df, write_excel

APP_TITLE = "投保数据汇总工具 v1.0"
FILE_TYPES = [("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]


def reveal_in_explorer(file_path):
    """在系统文件管理器中定位到文件所在目录（跨平台）"""
    path = os.path.abspath(file_path)
    try:
        if sys.platform.startswith("win"):
            # /select, 可以直接选中该文件
            subprocess.Popen(["explorer", "/select,", path])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path)])
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("720x520")
        self.root.resizable(True, True)

        self.src_path = None
        self.out_path = None
        self.msg_queue = queue.Queue()
        self.running = False

        self._build_ui()
        # 主线程定时从队列取日志，保证多线程下UI更新安全
        self._poll_queue()

    # ---------- 界面搭建 ----------
    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # 标题区
        top = ttk.Frame(self.root)
        top.pack(fill="x", **pad)
        ttk.Label(
            top, text="投保数据汇总工具",
            font=("Microsoft YaHei UI", 16, "bold") if sys.platform.startswith("win")
                 else ("PingFang SC", 16, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            top, text="数据全程在本机处理，不会上传到任何服务器",
            foreground="#666666",
        ).pack(anchor="w")

        # 第一步：选择文件
        step1 = ttk.LabelFrame(self.root, text="第一步：选择要处理的 Excel 文件")
        step1.pack(fill="x", **pad)
        row = ttk.Frame(step1)
        row.pack(fill="x", padx=10, pady=8)
        self.btn_browse = ttk.Button(row, text="选择文件...", command=self.browse)
        self.btn_browse.pack(side="left")
        self.var_file = tk.StringVar(value="尚未选择文件")
        ttk.Label(
            row, textvariable=self.var_file, foreground="#333333", wraplength=520,
        ).pack(side="left", padx=10)

        # 第二步：处理
        step2 = ttk.LabelFrame(self.root, text="第二步：开始处理")
        step2.pack(fill="x", **pad)
        btn_row = ttk.Frame(step2)
        btn_row.pack(fill="x", padx=10, pady=8)
        self.btn_run = ttk.Button(btn_row, text="开始处理", command=self.start, state="disabled")
        self.btn_run.pack(side="left")
        self.var_status = tk.StringVar(value="")
        ttk.Label(btn_row, textvariable=self.var_status, foreground="#0066cc").pack(side="left", padx=10)

        # 日志区
        log_frame = ttk.LabelFrame(self.root, text="处理日志")
        log_frame.pack(fill="both", expand=True, **pad)
        self.txt_log = scrolledtext.ScrolledText(
            log_frame, height=12, font=("Consolas", 10) if sys.platform.startswith("win")
            else ("Menlo", 10),
            state="disabled",
        )
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=8)

        # 底部：打开结果文件夹
        bottom = ttk.Frame(self.root)
        bottom.pack(fill="x", **pad)
        self.btn_open = ttk.Button(
            bottom, text="打开结果所在文件夹", command=self.open_result, state="disabled"
        )
        self.btn_open.pack(side="left")

        self._log("就绪。请先选择要处理的 Excel 文件。")

    # ---------- 日志 ----------
    def _log(self, msg):
        """线程安全写日志：子线程入队，主线程统一写入文本框"""
        self.msg_queue.put(str(msg))

    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                self.txt_log.configure(state="normal")
                self.txt_log.insert("end", msg + "\n")
                self.txt_log.see("end")
                self.txt_log.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)

    # ---------- 交互动作 ----------
    def browse(self):
        if self.running:
            return
        path = filedialog.askopenfilename(title="选择投保数据 Excel 文件", filetypes=FILE_TYPES)
        if not path:
            return
        self.src_path = path
        self.var_file.set(path)
        self.btn_run.configure(state="normal")
        self.btn_open.configure(state="disabled")
        self.var_status.set("")
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")
        self._log(f"已选择文件：{os.path.basename(path)}")

    def start(self):
        if self.running or not self.src_path:
            return
        self.running = True
        self.btn_run.configure(state="disabled")
        self.btn_browse.configure(state="disabled")
        self.btn_open.configure(state="disabled")
        self.var_status.set("正在处理中，请稍候...")
        self.out_path = None
        # 耗时任务放到子线程，避免界面卡死
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            src = self.src_path
            out = str(Path(src).parent / f"{Path(src).stem}_汇总.xlsx")
            self._log("开始处理...")
            # 把日志回调传给处理函数，进度实时显示到界面
            result_df = build_result_df(src, log=self._log)
            if len(result_df) == 0:
                raise ValueError("汇总结果为空，请检查文件内容是否正确。")
            self._log("正在生成 Excel 文件...")
            write_excel(result_df, out)
            self.out_path = out
            self._log(f"完成！共 {len(result_df)} 条汇总记录。")
            self._log(f"输出文件：{out}")
            self.root.after(0, self._on_success)
        except Exception as e:
            self._log("处理失败：" + str(e))
            self._log(traceback.format_exc(limit=2))
            self.root.after(0, lambda: self._on_error(str(e)))

    def _on_success(self):
        self.running = False
        self.var_status.set("处理完成！")
        self.btn_browse.configure(state="normal")
        self.btn_run.configure(state="normal")
        self.btn_open.configure(state="normal")
        if self.out_path:
            messagebox.showinfo(
                "处理完成",
                f"汇总完成！\n\n输出文件：\n{Path(self.out_path).name}\n\n"
                f"位置：原文件所在的同一个文件夹",
            )

    def _on_error(self, msg):
        self.running = False
        self.var_status.set("处理失败")
        self.btn_browse.configure(state="normal")
        self.btn_run.configure(state="normal")
        short = msg.split("\n")[0]
        messagebox.showerror("处理失败", short + "\n\n详情请看下方日志。")

    def open_result(self):
        if self.out_path and os.path.exists(self.out_path):
            reveal_in_explorer(self.out_path)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
