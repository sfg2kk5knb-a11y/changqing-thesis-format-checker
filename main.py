from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from analyzer import ThesisAnalyzer, export_csv, export_html, export_json, summary


def resource_path(relative):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # 明确指定中文界面字体，避免部分 Windows 环境的 Tk 默认字体出现方框/乱码。
        self.option_add("*Font", ("Microsoft YaHei UI", 10))
        self.title("论文格式检查助手")
        self.geometry("1180x720")
        self.minsize(920, 600)
        self.template_var = tk.StringVar()
        self.paper_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择学校模板和待检查论文。")
        self.issues = []
        self.logo_path = resource_path("assets/logo.png")
        self.icon_path = resource_path("assets/app_icon.png")
        self.qr_path = resource_path("assets/qr.png")
        self.feedback_qr_path = resource_path("assets/feedback_qr.png")
        self.logo_image = None
        self.qr_image = None
        if self.icon_path.exists():
            try:
                icon = tk.PhotoImage(file=str(self.icon_path))
                self.iconphoto(True, icon)
                self._icon_keep = icon
            except tk.TclError:
                pass
        self._build()

    def _build(self):
        brand = ttk.Frame(self, padding=(14, 10))
        brand.pack(fill="x")
        if self.logo_path.exists():
            try:
                self.logo_image = tk.PhotoImage(file=str(self.logo_path)).subsample(6, 6)
                ttk.Label(brand, image=self.logo_image).pack(side="left")
            except tk.TclError:
                ttk.Label(brand, text="常青文创设计", font=("Microsoft YaHei", 20, "bold"), foreground="#071952").pack(side="left")
        contact = ttk.Frame(brand)
        contact.pack(side="right", padx=(18, 4))
        ttk.Label(contact, text="合作联系", foreground="#071952", font=("Microsoft YaHei", 13, "bold")).pack(side="left", padx=(0, 8))
        if self.qr_path.exists():
            try:
                self.qr_image = tk.PhotoImage(file=str(self.qr_path)).subsample(7, 7)
                ttk.Label(contact, image=self.qr_image).pack(side="left")
            except tk.TclError:
                ttk.Label(contact, text="扫码添加微信", foreground="#657083").pack(side="left")
        feedback = ttk.Frame(brand)
        feedback.pack(side="right", padx=(8, 4))
        ttk.Label(feedback, text="用户使用问题反馈群", foreground="#071952", font=("Microsoft YaHei", 11, "bold")).pack(side="left", padx=(0, 8))
        if self.feedback_qr_path.exists():
            try:
                self.feedback_qr_image = tk.PhotoImage(file=str(self.feedback_qr_path)).subsample(7, 7)
                ttk.Label(feedback, image=self.feedback_qr_image).pack(side="left")
            except tk.TclError:
                ttk.Label(feedback, text="扫码进群", foreground="#657083").pack(side="left")

        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")
        ttk.Label(top, text="学校模板（.doc/.docx）").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.template_var).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(top, text="选择", command=lambda: self._choose(self.template_var)).grid(row=0, column=2)
        ttk.Label(top, text="待检查论文（.doc/.docx）").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.paper_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(8, 0))
        ttk.Button(top, text="选择", command=lambda: self._choose(self.paper_var)).grid(row=1, column=2, pady=(8, 0))
        self.run_btn = ttk.Button(top, text="开始检查", command=self._run)
        self.run_btn.grid(row=0, column=3, rowspan=2, padx=(12, 0), ipadx=12, ipady=12)
        top.columnconfigure(1, weight=1)

        filters = ttk.Frame(self, padding=(12, 0))
        filters.pack(fill="x")
        ttk.Label(filters, text="筛选：").pack(side="left")
        self.filter_var = tk.StringVar(value="全部")
        for name in ["全部", "严重", "警告", "提示"]:
            ttk.Radiobutton(filters, text=name, value=name, variable=self.filter_var, command=self._populate).pack(side="left", padx=5)
        ttk.Button(filters, text="导出 HTML", command=lambda: self._export("html")).pack(side="right", padx=4)
        ttk.Button(filters, text="导出 CSV", command=lambda: self._export("csv")).pack(side="right", padx=4)
        ttk.Button(filters, text="导出 JSON", command=lambda: self._export("json")).pack(side="right", padx=4)

        cols = ("severity", "category", "location", "message", "expected")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        labels = {"severity":"程度", "category":"类别", "location":"位置", "message":"问题", "expected":"模板要求"}
        widths = {"severity":60, "category":110, "location":220, "message":420, "expected":220}
        for c in cols:
            self.tree.heading(c, text=labels[c])
            self.tree.column(c, width=widths[c], minwidth=50)
        self.tree.tag_configure("严重", background="#ffe8e8")
        self.tree.tag_configure("警告", background="#fff4d9")
        self.tree.tag_configure("提示", background="#eef5ff")
        self.tree.pack(fill="both", expand=True, padx=12, pady=10)
        self.tree.bind("<<TreeviewSelect>>", self._detail)

        self.detail = tk.Text(self, height=7, wrap="word", padx=10, pady=8)
        self.detail.pack(fill="x", padx=12)
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padding=5).pack(fill="x", side="bottom")

    def _choose(self, variable):
        path = filedialog.askopenfilename(filetypes=[("Word 文档", "*.doc;*.docx"), ("所有文件", "*.*")])
        if path:
            variable.set(path)

    def _run(self):
        template, paper = self.template_var.get().strip(), self.paper_var.get().strip()
        if not template or not paper:
            messagebox.showwarning("缺少文件", "请先选择学校模板和待检查论文。")
            return
        if Path(template).suffix.lower() not in {".doc", ".docx"} or Path(paper).suffix.lower() not in {".doc", ".docx"}:
            messagebox.showwarning("文件格式", "请选择 .doc 或 .docx 文件。")
            return
        self.run_btn.config(state="disabled")
        self.status_var.set("正在解析文档，请稍候……")
        threading.Thread(target=self._analyze, args=(template, paper), daemon=True).start()

    def _analyze(self, template, paper):
        try:
            issues = ThesisAnalyzer(template, paper).analyze()
            self.after(0, lambda: self._finish(issues))
        except Exception as exc:
            self.after(0, lambda: self._failed(str(exc)))

    def _finish(self, issues):
        self.issues = issues
        self.run_btn.config(state="normal")
        self._populate()
        s = summary(issues)
        self.status_var.set(f"检查完成：严重 {s['严重']}，警告 {s['警告']}，提示 {s['提示']}，共 {s['总计']} 项。")

    def _failed(self, text):
        self.run_btn.config(state="normal")
        self.status_var.set("检查失败。")
        messagebox.showerror("检查失败", text)

    def _populate(self):
        self.tree.delete(*self.tree.get_children())
        selected = self.filter_var.get()
        for idx, issue in enumerate(self.issues):
            if selected != "全部" and issue.severity != selected:
                continue
            self.tree.insert("", "end", iid=str(idx), values=(issue.severity, issue.category, issue.location, issue.message, issue.expected), tags=(issue.severity,))

    def _detail(self, _event=None):
        selection = self.tree.selection()
        if not selection:
            return
        i = self.issues[int(selection[0])]
        text = f"问题：{i.message}\n位置：{i.location}\n当前情况：{i.actual}\n模板要求：{i.expected}\n修改建议：{i.suggestion}"
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", text)

    def _export(self, kind):
        if not self.issues:
            messagebox.showinfo("尚无结果", "请先完成一次检查。")
            return
        ext = "." + kind
        path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[(kind.upper(), "*" + ext)])
        if not path:
            return
        if kind == "html":
            export_html(path, self.issues, Path(self.template_var.get()).name, Path(self.paper_var.get()).name, self.logo_path)
        elif kind == "csv":
            export_csv(path, self.issues)
        else:
            export_json(path, self.issues)
        self.status_var.set(f"报告已导出：{path}")


if __name__ == "__main__":
    App().mainloop()
