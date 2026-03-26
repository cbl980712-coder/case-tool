# -*- coding: utf-8 -*-
"""
批量文书生成工具 v2 - 极简稳定版
"""
import os, re, sys, traceback, threading
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ── 启动前检查依赖 ──
def check_deps():
    missing = []
    for pkg, imp in [("pandas","pandas"),("openpyxl","openpyxl"),("python-docx","docx")]:
        try:
            __import__(imp)
        except ImportError:
            missing.append(pkg)
    if missing:
        root = Tk(); root.withdraw()
        messagebox.showerror("缺少依赖",
            f"请先运行：\npip install {' '.join(missing)}\n\n"
            "安装后重新打开程序。")
        sys.exit(1)

check_deps()

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from docx import Document
import pandas as pd

# ─────────────────────────────────
# 核心函数
# ─────────────────────────────────
def parse_records(text, fields):
    """解析名单文本，支持中英文冒号，字段名前后空格自动去除"""
    blocks = re.split(r"\n{2,}", text.strip())
    records = []
    for i, block in enumerate(blocks, 1):
        if not block.strip():
            continue
        rec = {"序号": i, "案号": ""}
        for line in block.splitlines():
            line = line.strip()
            m = re.match(r"^([^:：]+)[：:]\s*(.*)", line)
            if m:
                k = m.group(1).strip()
                v = m.group(2).strip()
                if k in fields:
                    rec[k] = v
        for f in fields:
            if f not in rec:
                rec[f] = ""
        records.append(rec)
    return records

def validate(records, fields):
    errors = []
    seen = {}
    ID_RE = re.compile(r"^\d{17}[\dXx]$")
    PH_RE = re.compile(r"^1[3-9]\d{9}$")
    AM_RE = re.compile(r"^\d+(\.\d{1,2})?$")
    for rec in records:
        errs = []
        seq = rec["序号"]
        if not rec.get("姓名","").strip(): errs.append("姓名为空")
        idn = rec.get("证件号","").strip()
        if not idn: errs.append("证件号为空")
        elif not ID_RE.match(idn): errs.append("证件号格式异常")
        else:
            if idn in seen: errs.append(f"证件号重复(序号{seen[idn]})")
            else: seen[idn] = seq
        amt = rec.get("金额","").strip()
        if not amt: errs.append("金额为空")
        elif not AM_RE.match(amt): errs.append("金额格式异常")
        ph = rec.get("手机号","").strip()
        if not ph: errs.append("手机号为空")
        elif not PH_RE.match(ph): errs.append("手机号格式异常")
        if not rec.get("地址","").strip(): errs.append("地址为空")
        if errs:
            errors.append({"序号":seq,"姓名":rec.get("姓名",""),"异常说明":"；".join(errs)})
            rec["状态"] = "异常：" + "；".join(errs)
        else:
            rec["状态"] = "正常"
    return errors

def save_excel(records, fields, path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    cols = ["序号"] + fields + ["案号","状态"]
    wb = Workbook(); ws = wb.active; ws.title = "案件总表"
    hf = Font(bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="2F5496")
    for ci,cn in enumerate(cols,1):
        c = ws.cell(row=1,column=ci,value=cn)
        c.font=hf; c.fill=hfill; c.alignment=Alignment(horizontal="center")
    for rec in records:
        ws.append([rec.get(c,"") for c in cols])
    wb.save(str(path))

def save_error_report(errors, path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    cols = ["序号","姓名","异常说明"]
    wb = Workbook(); ws = wb.active; ws.title = "异常报告"
    hf = Font(bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="C00000")
    for ci,cn in enumerate(cols,1):
        c = ws.cell(row=1,column=ci,value=cn)
        c.font=hf; c.fill=hfill
    for e in errors:
        ws.append([e.get(c,"") for c in cols])
    wb.save(str(path))

def replace_in_para(para, rmap):
    full = "".join(r.text for r in para.runs)
    new = full
    for k,v in rmap.items():
        new = new.replace(k, v)
    if new != full:
        for i,run in enumerate(para.runs):
            run.text = new if i==0 else ""

def replace_in_doc(doc, rmap):
    for p in doc.paragraphs:
        replace_in_para(p, rmap)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_para(p, rmap)
    # 检测残留
    txt = " ".join(p.text for p in doc.paragraphs)
    return re.findall(r"（[^）]{1,20}）", txt)

def gen_doc(rec, fields, tmpl_path, out_dir):
    name = rec.get("姓名","未知").strip() or "未知"
    case_no = str(rec.get("案号","")).strip()
    safe = lambda s: re.sub(r'[\\/:*?"<>|]',"_",s)
    base = f"{safe(name)}_（{safe(case_no)}）_起诉状" if case_no else f"{safe(name)}_起诉状"
    out = Path(out_dir) / (base + ".docx")
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    rmap = {f"（{f}）": str(rec.get(f,"")) for f in fields + ["案号"]}
    doc = Document(str(tmpl_path))
    remaining = replace_in_doc(doc, rmap)
    doc.save(str(out))
    return out.name, remaining

# ─────────────────────────────────
# GUI
# ─────────────────────────────────
class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("批量文书生成工具 v2")
        self.geometry("820x720")
        self.resizable(True, True)
        self._fields = ["姓名","证件号","金额","手机号","地址","备注"]
        self._tmpl = StringVar()
        self._outdir = StringVar(value=str(Path.home() / "文书输出"))
        self._excel = StringVar()
        self._build()

    def _build(self):
        f = ttk.Frame(self, padding=10)
        f.pack(fill=BOTH, expand=True)

        # ① 名单
        ttk.LabelFrame(f, text="① 当事人名单", padding=6).pack(fill=X, pady=(0,6))
        box1 = self.nameframes = ttk.LabelFrame(f, text="① 当事人名单", padding=6)
        box1.pack(fill=X, pady=(0,6))
        row1 = ttk.Frame(box1); row1.pack(fill=X)
        ttk.Button(row1, text="📂 导入文件", command=self._load_file).pack(side=LEFT)
        ttk.Button(row1, text="清空", command=lambda:self.txt.delete("1.0",END)).pack(side=LEFT, padx=6)
        self.txt = scrolledtext.ScrolledText(box1, height=9, font=("Consolas",10))
        self.txt.pack(fill=BOTH, expand=True, pady=(4,0))

        # ② 模板
        box2 = ttk.LabelFrame(f, text="② Word 模板文件（.docx）", padding=6)
        box2.pack(fill=X, pady=(0,6))
        row2 = ttk.Frame(box2); row2.pack(fill=X)
        ttk.Button(row2, text="📂 选择模板", command=self._pick_tmpl).pack(side=LEFT)
        ttk.Label(row2, textvariable=self._tmpl, foreground="#333", wraplength=600).pack(side=LEFT, padx=8)

        # ③ 案号
        box3 = ttk.LabelFrame(f, text="③ 案号补录（可选，立案后导入 Excel）", padding=6)
        box3.pack(fill=X, pady=(0,6))
        row3 = ttk.Frame(box3); row3.pack(fill=X)
        ttk.Button(row3, text="📂 导入 case_list.xlsx", command=self._pick_excel).pack(side=LEFT)
        ttk.Label(row3, textvariable=self._excel, foreground="#333", wraplength=500).pack(side=LEFT, padx=8)

        # ④ 输出目录
        box4 = ttk.LabelFrame(f, text="④ 文书保存目录", padding=6)
        box4.pack(fill=X, pady=(0,6))
        row4 = ttk.Frame(box4); row4.pack(fill=X)
        ttk.Button(row4, text="📁 选择目录", command=self._pick_outdir).pack(side=LEFT)
        ttk.Label(row4, textvariable=self._outdir, foreground="#2255AA",
                  font=("",9,"bold"), wraplength=580).pack(side=LEFT, padx=8)

        # 按钮
        brow = ttk.Frame(f); brow.pack(fill=X, pady=(4,6))
        self._btn = ttk.Button(brow, text="▶  开始生成文书", command=self._run)
        self._btn.pack(side=LEFT)
        ttk.Button(brow, text="📂 打开输出目录", command=self._open_out).pack(side=LEFT, padx=8)
        ttk.Button(brow, text="📊 打开异常报告", command=self._open_report).pack(side=LEFT)

        # 进度
        self._prog = ttk.Progressbar(f, mode="determinate")
        self._prog.pack(fill=X, pady=(0,4))

        # 日志
        logf = ttk.LabelFrame(f, text="运行日志", padding=4)
        logf.pack(fill=BOTH, expand=True)
        self._log_box = scrolledtext.ScrolledText(logf, font=("Consolas",9),
            bg="#1e1e1e", fg="#d4d4d4", state=DISABLED)
        self._log_box.pack(fill=BOTH, expand=True)

        self._log("准备就绪。")

    # ── 日志 ──
    def _log(self, msg):
        self._log_box.config(state=NORMAL)
        self._log_box.insert(END, str(msg)+"\n")
        self._log_box.see(END)
        self._log_box.config(state=DISABLED)
        self._log_box.update()
        self.update_idletasks()

    # ── 选择 ──
    def _load_file(self):
        p = filedialog.askopenfilename(filetypes=[("文本","*.txt"),("全部","*.*")])
        if p:
            self.txt.delete("1.0",END)
            self.txt.insert("1.0", Path(p).read_text(encoding="utf-8"))
            self._log(f"导入：{p}")

    def _pick_tmpl(self):
        p = filedialog.askopenfilename(filetypes=[("Word文档","*.docx"),("全部","*.*")])
        if p:
            self._tmpl.set(p)
            self._log(f"模板：{p}")
            self._scan_tmpl(p)

    def _scan_tmpl(self, path):
        try:
            doc = Document(path)
            txt = " ".join(p.text for p in doc.paragraphs)
            for tbl in doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        txt += " " + " ".join(p.text for p in cell.paragraphs)
            found = list(dict.fromkeys(re.findall(r"（[^）]{1,20}）", txt)))
            if found:
                self._log(f"模板占位符：{' '.join(found)}")
            else:
                self._log("⚠️  模板未检测到占位符，请确认括号是中文全角（）")
        except Exception as e:
            self._log(f"模板扫描失败：{e}")

    def _pick_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx"),("全部","*.*")])
        if p:
            self._excel.set(p)
            self._log(f"Excel：{p}")

    def _pick_outdir(self):
        d = filedialog.askdirectory()
        if d:
            self._outdir.set(d)
            self._log(f"输出目录：{d}")

    def _open_out(self):
        d = self._outdir.get()
        if Path(d).exists():
            os.startfile(d)
        else:
            messagebox.showinfo("提示","目录不存在，请先生成文书。")

    def _open_report(self):
        rp = Path(self._outdir.get()) / "_reports" / "error_report.xlsx"
        if rp.exists():
            os.startfile(str(rp))
        else:
            messagebox.showinfo("提示","异常报告尚未生成。")

    # ── 主流程 ──
    def _run(self):
        self._btn.config(state=DISABLED, text="⏳ 生成中...")
        self.update()
        try:
            self._do_run()
        except Exception as e:
            self._log(f"❌ 严重错误：{e}")
            self._log(traceback.format_exc())
        finally:
            self._btn.config(state=NORMAL, text="▶  开始生成文书")

    def _do_run(self):
        self._log("\n" + "="*40)
        out_dir = Path(self._outdir.get())
        out_dir.mkdir(parents=True, exist_ok=True)
        fields = self._fields

        # 数据来源：Excel 或 文本框
        excel_src = self._excel.get().strip()
        if excel_src and Path(excel_src).exists():
            self._log(f"从 Excel 读取：{excel_src}")
            df = pd.read_excel(excel_src, sheet_name="案件总表", dtype=str).fillna("")
            records = df.to_dict(orient="records")
            for i,r in enumerate(records,1):
                r.setdefault("序号", i)
            errors = []
            self._log(f"读取到 {len(records)} 条")
        else:
            raw = self.txt.get("1.0", END).strip()
            if not raw:
                self._log("❌ 名单为空，请输入或导入名单")
                return
            self._log("解析名单...")
            records = parse_records(raw, fields)
            self._log(f"解析完成：{len(records)} 条")
            if not records:
                self._log("❌ 解析结果为空，请检查名单格式")
                return
            errors = validate(records, fields)
            self._log(f"校验完成：{len(errors)} 条异常")
            xl_path = out_dir / "_excel" / "case_list.xlsx"
            save_excel(records, fields, xl_path)
            self._log(f"Excel 总表：{xl_path}")
            if errors:
                rp = out_dir / "_reports" / "error_report.xlsx"
                save_error_report(errors, rp)
                self._log(f"异常报告：{rp}")
                for e in errors:
                    self._log(f"  ⚠️  序号{e['序号']} {e['姓名']}：{e['异常说明']}")

        # 检查模板
        tmpl = self._tmpl.get().strip()
        if not tmpl or not Path(tmpl).exists():
            self._log("❌ 未选择模板文件")
            return

        # 批量生成
        total = len(records)
        ok = skip = 0
        self._log(f"开始生成 {total} 份文书...")
        self._prog.config(maximum=total, value=0)
        for i, rec in enumerate(records, 1):
            name = rec.get("姓名","?")
            if "异常" in str(rec.get("状态","")):
                self._log(f"  跳过 {name}（有异常）")
                skip += 1
            else:
                try:
                    fname, remaining = gen_doc(rec, fields, tmpl, str(out_dir))
                    self._log(f"  ✅ {fname}")
                    if remaining:
                        self._log(f"     ⚠️ 残留占位符：{remaining}")
                    ok += 1
                except Exception as e:
                    self._log(f"  ❌ {name} 失败：{e}")
            self._prog.config(value=i)
            self.update_idletasks()

        self._log(f"\n✅ 完成！成功 {ok} 份 / 跳过 {skip} 份")
        self._log(f"保存目录：{out_dir}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
