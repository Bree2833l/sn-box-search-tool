#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
双向查询工具：SN ↔ 箱码
模式1：SN查箱码；模式2：箱码反查SN
SN输出格式可选：纵向堆叠 / 横向分列
输出文件名自带查询模式+输出类型+时间戳
【安全说明】
✅ 扫描检索阶段：仅【只读】读取目标文件夹内所有文件，不会修改/写入/覆盖任何源文件
✅ 唯一写操作：新建输出结果表格文件，不会改动被检索的原始文件
✅ 路径记忆、弹窗初始目录、输出不带后缀文件名
"""

import os
import sys
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl.utils import get_column_letter
from datetime import datetime

import pandas as pd

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "search_config.json")

class FileSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SN-箱码双向查询工具")
        self.root.geometry("720x460")
        self.root.resizable(False, False)

        # 变量
        self.query_mode_var = tk.StringVar(value="sn_to_box") # sn_to_box / box_to_sn
        self.sn_output_type_var = tk.StringVar(value="vertical") # vertical纵向 / horizontal横向

        self.query_file_var = tk.StringVar()
        self.target_dir_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.column_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)
        self.search_excel_var = tk.BooleanVar(value=True)
        self.search_text_var = tk.BooleanVar(value=True)

        self._build_ui()
        # 加载上次保存的配置
        self._load_config()
        # 窗口关闭事件：保存配置
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- UI ----------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # 查询模式
        frm_mode = ttk.LabelFrame(self.root, text="查询模式")
        frm_mode.pack(fill="x",** pad)
        ttk.Radiobutton(frm_mode, text="SN → 箱码（原功能）", variable=self.query_mode_var, value="sn_to_box").pack(side="left", padx=10)
        ttk.Radiobutton(frm_mode, text="箱码 → SN（新增）", variable=self.query_mode_var, value="box_to_sn").pack(side="left", padx=10)
        ttk.Label(frm_mode, text="| SN输出方式：").pack(side="left", padx=(20,5))
        ttk.Radiobutton(frm_mode, text="纵向堆叠(一个SN一行)", variable=self.sn_output_type_var, value="vertical").pack(side="left")
        ttk.Radiobutton(frm_mode, text="横向分列(一行箱码多SN列)", variable=self.sn_output_type_var, value="horizontal").pack(side="left")

        # 1. 查询文件
        frm1 = ttk.LabelFrame(self.root, text="① 选择查询文件（xlsx / csv）")
        frm1.pack(fill="x", **pad)
        ttk.Entry(frm1, textvariable=self.query_file_var).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(frm1, text="浏览…", command=self._pick_query_file).pack(side="left", padx=4)

        # 2. 目标文件夹
        frm2 = ttk.LabelFrame(self.root, text="② 选择要搜索的目标文件夹（存放箱码文件）")
        frm2.pack(fill="x",** pad)
        ttk.Entry(frm2, textvariable=self.target_dir_var).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(frm2, text="浏览…", command=self._pick_target_dir).pack(side="left", padx=4)

        # 3. 选项
        frm3 = ttk.LabelFrame(self.root, text="③ 选项")
        frm3.pack(fill="x", **pad)

        row1 = ttk.Frame(frm3)
        row1.pack(fill="x", padx=8, pady=4)
        ttk.Label(row1, text="查询列：").pack(side="left")
        self.column_combo = ttk.Combobox(row1, textvariable=self.column_var, state="readonly", width=30)
        self.column_combo.pack(side="left", padx=4)

        row2 = ttk.Frame(frm3)
        row2.pack(fill="x", padx=8, pady=4)
        ttk.Checkbutton(row2, text="递归搜索子文件夹", variable=self.recursive_var).pack(side="left", padx=8)
        ttk.Checkbutton(row2, text="搜索 xlsx/xls", variable=self.search_excel_var).pack(side="left", padx=8)
        ttk.Checkbutton(row2, text="搜索 csv/txt/md", variable=self.search_text_var).pack(side="left", padx=8)

        # 4. 输出位置
        frm4 = ttk.LabelFrame(self.root, text="④ 结果输出文件夹")
        frm4.pack(fill="x",** pad)
        ttk.Entry(frm4, textvariable=self.output_dir_var).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(frm4, text="浏览…", command=self._pick_output_dir).pack(side="left", padx=4)

        # 5. 开始按钮 + 进度
        frm5 = ttk.Frame(self.root)
        frm5.pack(fill="x", **pad)
        self.start_btn = ttk.Button(frm5, text="开始搜索", command=self._start_search)
        self.start_btn.pack(side="left")
        self.progress = ttk.Progressbar(frm5, mode="determinate", length=420)
        self.progress.pack(side="left", padx=10, fill="x", expand=True)

        self.status_label = ttk.Label(self.root, text="就绪", foreground="gray")
        self.status_label.pack(fill="x", padx=10, pady=(0, 8))

    # ---------- 配置读写（路径记忆） ----------
    def _save_config(self):
        """保存当前界面参数到json"""
        cfg = {
            "query_mode": self.query_mode_var.get(),
            "sn_output_type": self.sn_output_type_var.get(),
            "query_file": self.query_file_var.get(),
            "target_dir": self.target_dir_var.get(),
            "output_dir": self.output_dir_var.get(),
            "select_column": self.column_var.get(),
            "recursive": self.recursive_var.get(),
            "search_excel": self.search_excel_var.get(),
            "search_text_var": self.search_text_var.get(),
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_config(self):
        """加载上次配置，回填界面；路径失效不会报错"""
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.query_mode_var.set(cfg.get("query_mode", "sn_to_box"))
            self.sn_output_type_var.set(cfg.get("sn_output_type", "vertical"))
            self.query_file_var.set(cfg.get("query_file", ""))
            self.target_dir_var.set(cfg.get("target_dir", ""))
            self.output_dir_var.set(cfg.get("output_dir", ""))
            self.column_var.set(cfg.get("select_column", ""))
            self.recursive_var.set(cfg.get("recursive", True))
            self.search_excel_var.set(cfg.get("search_excel_var", True))
            self.search_text_var.set(cfg.get("search_text_var", True))

            q_path = self.query_file_var.get()
            if os.path.isfile(q_path):
                self._load_columns(q_path)
        except Exception:
            pass

    def _on_close(self):
        self._save_config()
        self.root.destroy()

    # ---------- 选择文件/目录 ----------
    def _pick_query_file(self):
        last_file_path = self.query_file_var.get().strip()
        initial_dir = os.path.dirname(last_file_path) if os.path.isfile(last_file_path) else ""

        p = filedialog.askopenfilename(
            title="选择查询文件",
            initialdir=initial_dir,
            filetypes=[("表格文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")],
        )
        if p:
            self.query_file_var.set(p)
            self.output_dir_var.set(os.path.dirname(p))
            self._load_columns(p)

    def _pick_target_dir(self):
        d = filedialog.askdirectory(
            title="选择目标文件夹",
            initialdir=self.target_dir_var.get()
        )
        if d:
            self.target_dir_var.set(d)
            if not self.output_dir_var.get():
                self.output_dir_var.set(d)

    def _pick_output_dir(self):
        d = filedialog.askdirectory(
            title="选择结果输出文件夹",
            initialdir=self.output_dir_var.get()
        )
        if d:
            self.output_dir_var.set(d)

    def _load_columns(self, path):
        try:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str, nrows=5)
            else:
                df = pd.read_excel(path, dtype=str, nrows=5)
            cols = list(df.columns)
            self.column_combo["values"] = cols
            if cols:
                self.column_var.set(cols[0])
        except Exception as e:
            messagebox.showwarning("提示", f"读取查询文件列名失败：\n{e}")

    # ---------- 读取箱文件，提取SN（箱码模式专用，取列名'SN'那一列） ----------
    def _read_sn_list_from_box_file(self, fp):
        ext = os.path.splitext(fp)[1].lower()
        try:
            if ext in (".xlsx", ".xls"):
                df = pd.read_excel(fp, dtype=str)
            else:
                for enc in ("utf-8", "gbk", "gb18030"):
                    try:
                        df = pd.read_csv(fp, dtype=str, encoding=enc)
                        break
                    except:
                        continue
            if "SN" in df.columns:
                sn_list = df["SN"].dropna().astype(str).str.strip().tolist()
                sn_list = [s for s in sn_list if s and s.lower() != "nan"]
                return sn_list
            else:
                return []
        except Exception:
            return []

    def _read_file_content(self, fp):
        ext = os.path.splitext(fp)[1].lower()
        if ext in (".xlsx", ".xls"):
            sheets = pd.read_excel(fp, dtype=str, sheet_name=None)
            parts = []
            for name, df in sheets.items():
                parts.append(df.to_string(index=False, header=False))
            return "\n".join(parts)
        else:
            for enc in ("utf-8", "gbk", "gb18030", "latin-1"):
                try:
                    with open(fp, "r", encoding=enc) as f:
                        return f.read()
                except UnicodeDecodeError:
                    continue
            return ""

    # ---------- 搜索主逻辑 ----------
    def _start_search(self):
        qf = self.query_file_var.get().strip()
        td = self.target_dir_var.get().strip()
        od = self.output_dir_var.get().strip()
        col = self.column_var.get().strip()

        if not qf or not os.path.isfile(qf):
            messagebox.showerror("错误", "请先选择有效的查询文件")
            return
        if not td or not os.path.isdir(td):
            messagebox.showerror("错误", "请先选择有效的目标文件夹")
            return
        if not col:
            messagebox.showerror("错误", "请选择查询列")
            return
        if not od or not os.path.isdir(od):
            messagebox.showerror("错误", "请选择有效的输出文件夹")
            return

        self.start_btn.config(state="disabled")
        self.progress["value"] = 0
        threading.Thread(target=self._run_search, args=(qf, td, od, col), daemon=True).start()

    def _run_search(self, qf, td, od, col):
        try:
            mode = self.query_mode_var.get()
            sn_out_type = self.sn_output_type_var.get()
            now_str = datetime.now().strftime("%Y%m%d_%H%M")

            self._set_status("正在读取查询列表…")
            if qf.lower().endswith(".csv"):
                qdf = pd.read_csv(qf, dtype=str)
            else:
                qdf = pd.read_excel(qf, dtype=str)

            if col not in qdf.columns:
                raise ValueError(f"列 '{col}' 不存在")

            query_items = [str(x).strip() for x in qdf[col].tolist() if str(x).strip() and str(x).strip().lower() != "nan"]
            seen = set()
            query_unique = []
            for item in query_items:
                if item not in seen:
                    seen.add(item)
                    query_unique.append(item)

            self._set_status(f"共 {len(query_unique)} 个查询项，开始扫描文件…")

            # 收集目标文件
            exts = []
            if self.search_excel_var.get():
                exts += [".xlsx", ".xls"]
            if self.search_text_var.get():
                exts += [".csv", ".txt", ".md", ".tsv", ".log"]
            exts = tuple(exts)

            files = []
            if self.recursive_var.get():
                for root, _, names in os.walk(td):
                    for n in names:
                        if n.lower().endswith(exts) and not n.startswith("~$"):
                            full_fp = os.path.join(root, n)
                            bname = os.path.basename(full_fp)
                            box_name_no_ext, _ = os.path.splitext(bname)
                            files.append((full_fp, box_name_no_ext))
            else:
                for n in os.listdir(td):
                    fp = os.path.join(td, n)
                    if os.path.isfile(fp) and n.lower().endswith(exts) and not n.startswith("~$"):
                        bname = os.path.basename(fp)
                        box_name_no_ext, _ = os.path.splitext(bname)
                        files.append((fp, box_name_no_ext))

            if not files:
                raise RuntimeError("目标文件夹下没有找到可搜索的文件")

            self.progress["maximum"] = len(files)

            # 模式1 SN -> 箱码
            if mode == "sn_to_box":
                hits = {k: [] for k in query_unique}
                for idx, (fp, box_name) in enumerate(files):
                    self._set_status(f"扫描 {idx+1}/{len(files)}：{box_name}")
                    content = self._read_file_content(fp)
                    content_lower = content.lower()
                    for sn in query_unique:
                        if sn.lower() in content_lower:
                            hits[sn].append(box_name)
                    self.progress["value"] = idx + 1
                    self.root.update_idletasks()
                # 组装结果
                rows = []
                for sn, box_list in hits.items():
                    rows.append({
                        "查询SN": sn,
                        "是否找到": "是" if box_list else "否",
                        "箱码数量": len(box_list),
                        "箱码列表": "；".join(box_list) if box_list else ""
                    })
                result_df = pd.DataFrame(rows)
                filename = f"结果_SN查箱码_{now_str}.xlsx"

            # 模式2 箱码 -> SN
            else:
                box_to_sn = {}
                for idx, (fp, box_name) in enumerate(files):
                    self._set_status(f"扫描箱码文件 {idx+1}/{len(files)}：{box_name}")
                    sn_list = self._read_sn_list_from_box_file(fp)
                    box_to_sn[box_name] = sn_list
                    self.progress["value"] = idx + 1
                    self.root.update_idletasks()

                # 构建结果DataFrame
                if sn_out_type == "vertical":
                    rows = []
                    for box_code in query_unique:
                        sn_list = box_to_sn.get(box_code, [])
                        if len(sn_list) == 0:
                            rows.append({"箱码": box_code, "是否找到": "否", "SN": ""})
                        else:
                            for sn in sn_list:
                                rows.append({"箱码": box_code, "是否找到": "是", "SN": sn})
                    result_df = pd.DataFrame(rows)
                    filename = f"结果_箱码查SN_纵向堆叠_{now_str}.xlsx"
                else:
                    # 横向分列 SN1 SN2 SN3...
                    rows = []
                    max_sn_count = 0
                    data_rows = []
                    for box_code in query_unique:
                        sn_list = box_to_sn.get(box_code, [])
                        row_dict = {"箱码": box_code, "是否找到": "是" if sn_list else "否"}
                        max_sn_count = max(max_sn_count, len(sn_list))
                        for i, sn in enumerate(sn_list):
                            row_dict[f"SN{i+1}"] = sn
                        data_rows.append(row_dict)
                    # 补全所有SN列
                    columns = ["箱码", "是否找到"] + [f"SN{n+1}" for n in range(max_sn_count)]
                    result_df = pd.DataFrame(data_rows, columns=columns)
                    filename = f"结果_箱码查SN_横向铺开_{now_str}.xlsx"

            # 输出Excel【修复关键点：取消循环全部列设置宽度！只固定前两列，其余自动】
            self._set_status("正在写出结果文件…")
            out_path = os.path.join(od, filename)

            with pd.ExcelWriter(out_path, engine="openpyxl") as w:
                result_df.to_excel(w, index=False, sheet_name="查找结果")
                ws = w.sheets["查找结果"]
                # 只固定A、B两列宽度，不再循环所有列，彻底杜绝A[报错
                ws.column_dimensions["A"].width = 22
                ws.column_dimensions["B"].width = 14

            found_cnt = 0
            if mode == "sn_to_box":
                found_cnt = sum(1 for v in hits.values() if v)
            else:
                if sn_out_type == "vertical":
                    found_cnt = len(set([r["箱码"] for r in rows if r["是否找到"]=="是"]))
                else:
                    found_cnt = sum(1 for r in data_rows if r["是否找到"]=="是")

            self._set_status(f"完成！找到 {found_cnt}/{len(query_unique)} 项。\n结果保存：{out_path}")
            messagebox.showinfo("完成", f"搜索完成！\n找到 {found_cnt}/{len(query_unique)} 个查询项\n文件：{filename}")

        except Exception as e:
            messagebox.showerror("出错了", str(e))
            self._set_status(f"出错：{e}")
        finally:
            self.start_btn.config(state="normal")

    def _set_status(self, text):
        self.status_label.after(0, lambda: self.status_label.config(text=text))


def main():
    root = tk.Tk()
    app = FileSearchApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()