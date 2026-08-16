# -*- coding: utf-8 -*-
"""
company_doc_gui.py - 企业信用报告解析 GUI 界面

功能：
1. 选择包含 docx 文件的目录
2. 执行解析（调用 company_doc.py 的核心逻辑）
3. 显示执行结果
4. 打开结果所在目录

依赖：
  pip install pypandoc pypandoc_binary openpyxl
"""

import json
import logging
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

# 确保可以导入 company_doc 模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import company_doc
except ImportError:
    company_doc = None


class CompanyDocGUI:
    """企业信用报告解析 GUI"""

    def __init__(self, root):
        self.root = root
        self.root.title("企业信用报告解析工具")
        self.root.geometry("720x560")
        self.root.resizable(True, True)

        # 变量
        self.input_dir = tk.StringVar(value=r"D:\Download\word")
        self.output_file = tk.StringVar(value="company_report.xlsx")
        self.gen_json = tk.BooleanVar(value=True)
        self.verbose = tk.BooleanVar(value=False)
        self.is_running = False

        self._build_ui()

    def _build_ui(self):
        """构建界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ===== 输入目录 =====
        input_frame = ttk.LabelFrame(main_frame, text="输入设置", padding=10)
        input_frame.pack(fill=tk.X, pady=(0, 10))

        row1 = ttk.Frame(input_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="Word 目录:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.input_dir, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row1, text="浏览...", command=self._browse_dir).pack(side=tk.LEFT, padx=5)

        row2 = ttk.Frame(input_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="输出文件:").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.output_file, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row2, text="浏览...", command=self._browse_output).pack(side=tk.LEFT, padx=5)

        row3 = ttk.Frame(input_frame)
        row3.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(row3, text="同时输出中间 JSON", variable=self.gen_json).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(row3, text="详细日志输出", variable=self.verbose).pack(side=tk.LEFT, padx=5)

        # ===== 操作按钮 =====
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.run_btn = ttk.Button(btn_frame, text="开始解析", command=self._run_parse)
        self.run_btn.pack(side=tk.LEFT, padx=5)

        self.open_dir_btn = ttk.Button(btn_frame, text="打开结果目录", command=self._open_output_dir, state=tk.DISABLED)
        self.open_dir_btn.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(btn_frame, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, padx=20, fill=tk.X, expand=True)

        # ===== 日志输出 =====
        log_frame = ttk.LabelFrame(main_frame, text="执行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(5, 0))

    # ===== 界面操作 =====

    def _browse_dir(self):
        """选择输入目录"""
        path = filedialog.askdirectory(initialdir=self.input_dir.get() or os.getcwd())
        if path:
            self.input_dir.set(path)

    def _browse_output(self):
        """选择输出文件"""
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile=self.output_file.get() or "company_report.xlsx",
        )
        if path:
            self.output_file.set(path)

    def _open_output_dir(self):
        """打开输出文件所在目录"""
        out_path = Path(self.output_file.get())
        if not out_path.is_absolute():
            out_path = Path.cwd() / out_path
        out_dir = out_path.parent
        if out_dir.exists():
            os.startfile(str(out_dir))  # Windows
        else:
            messagebox.showerror("错误", f"目录不存在: {out_dir}")

    def _log(self, msg: str):
        """向日志区域追加文本"""
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def _set_running(self, running: bool):
        """设置运行状态"""
        self.is_running = running
        self.run_btn.config(state=tk.DISABLED if running else tk.NORMAL)
        if running:
            self.progress.start(10)
            self.status_var.set("正在解析...")
        else:
            self.progress.stop()
            self.status_var.set("完成")

    # ===== 核心逻辑 =====

    def _run_parse(self):
        """执行解析（后台线程）"""
        if self.is_running:
            return

        input_path = self.input_dir.get().strip()
        if not input_path:
            messagebox.showerror("错误", "请选择输入目录")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("错误", f"输入路径不存在: {input_path}")
            return

        self._set_running(True)
        self.log_text.delete(1.0, tk.END)
        self.open_dir_btn.config(state=tk.DISABLED)

        thread = threading.Thread(target=self._parse_worker, args=(input_path,), daemon=True)
        thread.start()

    def _parse_worker(self, input_path: str):
        """后台解析线程"""
        try:
            self._log("=" * 50)
            self._log("企业信用报告解析开始")
            self._log(f"输入目录: {input_path}")
            self._log(f"输出文件: {self.output_file.get()}")
            self._log("")

            if company_doc is not None:
                # 直接调用 company_doc 模块
                result = self._parse_with_module(input_path)
            else:
                # 回退：调用命令行
                result = self._parse_with_subprocess(input_path)

            # 显示结果
            self._log("")
            self._log("=" * 50)
            self._log("解析完成！")
            self._log(f"共解析 {result['companies']} 家公司 / {result['persons']} 位人员（失败 {result['failed']}）")
            self._log(f"输出文件: {result['output']}")

            # 显示身份统计
            stats = result.get("stats", {})
            if stats:
                self._log("")
                self._log("--- 身份统计 ---")
                self._log(f"  法定代表人: {stats.get('legal_count', 0)} 人")
                self._log(f"  投资人:     {stats.get('investor_count', 0)} 人")
                self._log(f"  经营者:     {stats.get('operator_count', 0)} 人")
                self._log(f"  股东:       {stats.get('shareholder_count', 0)} 人")
                self._log(f"  主要成员:   {stats.get('member_count', 0)} 人")
                self._log(f"  人员总数:   {stats.get('person_count', 0)} 人")

                # 显示每个公司的详细统计
                company_details = stats.get("company_details", [])
                if company_details:
                    self._log("")
                    self._log("--- 各公司统计 ---")
                    for cd in company_details:
                        self._log(f"  {cd['company']}")
                        self._log(f"    法定代表人: {cd.get('legal', '') or '-'}")
                        self._log(f"    投资人:     {cd.get('investor', '') or '-'}")
                        self._log(f"    经营者:     {cd.get('operator', '') or '-'}")
                        self._log(f"    股东:       {len(cd.get('shareholders', []))} 人")
                        self._log(f"    主要成员:   {len(cd.get('members', []))} 人")
                        self._log(f"    人员总数:   {cd.get('person_count', 0)} 人")

            if result["failed"] > 0:
                self._log(f"警告: {result['failed']} 个文件解析失败，请查看日志")

            self.open_dir_btn.config(state=tk.NORMAL)
            self.status_var.set(f"完成: {result['companies']} 家公司 / {result['persons']} 位人员")

        except Exception as e:
            self._log(f"错误: {e}")
            messagebox.showerror("错误", str(e))
            self.status_var.set("失败")
        finally:
            self._set_running(False)

    def _parse_with_module(self, input_path: str) -> dict:
        """使用 company_doc 模块直接解析"""
        # 配置日志
        log_file = "company_doc.log"
        company_doc.setup_logging(log_file, verbose=self.verbose.get())

        # 收集 docx 文件
        input_path_obj = Path(input_path)
        if input_path_obj.is_dir():
            docx_files = sorted(input_path_obj.glob("*.docx"))
        elif input_path_obj.is_file():
            docx_files = [input_path_obj]
        else:
            raise FileNotFoundError(f"输入路径不存在: {input_path}")

        if not docx_files:
            raise FileNotFoundError("未找到任何 docx 文件")

        self._log(f"找到 {len(docx_files)} 个 docx 文件")

        # 逐个解析
        all_reports = []
        fail_count = 0
        for i, f in enumerate(docx_files, 1):
            self._log(f"[{i}/{len(docx_files)}] 解析: {f.name}")
            try:
                data = company_doc.parse_report(str(f))
                all_reports.append(data)
                company_name = data.get('公司名称', '')
                legal = data.get('法定代表人', '') or '-'
                investor = data.get('投资人', '') or '-'
                operator = data.get('经营者', '') or '-'
                shareholders = data.get('股东', [])
                members = data.get('主要人员', [])

                # 统计该公司信息
                sh_names = set()
                for sh in shareholders:
                    name = (sh.get('发起人名称') or sh.get('股东名称') or '').strip()
                    if name and name != '-':
                        sh_names.add(name)

                m_names = set()
                for m in members:
                    name = (m.get('姓名') or '').strip()
                    if name and name != '-':
                        m_names.add(name)

                all_persons = set()
                if legal and legal != '-':
                    all_persons.add(legal)
                if investor and investor != '-':
                    all_persons.add(investor)
                if operator and operator != '-':
                    all_persons.add(operator)
                all_persons.update(sh_names)
                all_persons.update(m_names)

                self._log(f"  -> 公司: {company_name}")
                self._log(f"     法定代表人: {legal}")
                self._log(f"     投资人:     {investor}")
                self._log(f"     经营者:     {operator}")
                self._log(f"     股东:       {len(sh_names)} 人")
                self._log(f"     主要成员:   {len(m_names)} 人")
                self._log(f"     人员总数:   {len(all_persons)} 人")
            except Exception as e:
                fail_count += 1
                self._log(f"  -> 失败: {e}")

        if not all_reports:
            raise RuntimeError(f"所有 {len(docx_files)} 个文件均解析失败")

        # 合并人员信息
        self._log("")
        self._log("合并人员信息...")
        rows = company_doc.merge_all_companies(all_reports)

        # 获取统计信息
        stats = company_doc.get_last_stats()

        # 保存 xlsx（含第2个 sheet 统计）
        out_path = self.output_file.get()
        if not os.path.isabs(out_path):
            out_path = os.path.join(os.getcwd(), out_path)
        company_doc.save_xlsx(rows, out_path, stats=stats)

        # 保存 JSON
        if self.gen_json.get():
            json_path = os.path.splitext(out_path)[0] + ".json"
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(all_reports, f, ensure_ascii=False, indent=2, default=str)
            self._log(f"中间 JSON: {json_path}")

        return {
            "companies": len(all_reports),
            "persons": len(rows),
            "failed": fail_count,
            "output": out_path,
            "stats": stats,
        }

    def _parse_with_subprocess(self, input_path: str) -> dict:
        """回退：使用命令行解析"""
        cmd = [
            sys.executable,
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "company_doc.py"),
            input_path,
            "-o", self.output_file.get(),
        ]
        if self.gen_json.get():
            cmd.append("--json")
        if self.verbose.get():
            cmd.append("-v")

        self._log("执行命令: " + " ".join(cmd))
        self._log("")

        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        output_lines = []
        for line in proc.stdout:
            line = line.rstrip()
            output_lines.append(line)
            self._log(line)

        proc.wait()

        # 解析输出
        companies = 0
        persons = 0
        failed = 0
        out_path = self.output_file.get()
        if not os.path.isabs(out_path):
            out_path = os.path.join(os.getcwd(), out_path)

        for line in output_lines:
            if "共解析" in line and "家公司" in line:
                import re
                m = re.search(r"共解析 (\d+) 家公司 / (\d+) 位人员（失败 (\d+)）", line)
                if m:
                    companies = int(m.group(1))
                    persons = int(m.group(2))
                    failed = int(m.group(3))

        return {
            "companies": companies,
            "persons": persons,
            "failed": failed,
            "output": out_path,
        }


def main():
    root = tk.Tk()
    app = CompanyDocGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()