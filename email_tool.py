import os
import sys
import json
import zipfile
import smtplib
from smtplib import SMTP, SMTP_SSL
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# PDF & Excel
from pypdf import PdfReader, PdfWriter
from openpyxl import load_workbook


# ══════════════════════════════════════════════════════════
#  配置文件读写
# ══════════════════════════════════════════════════════════

def get_config_path():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "config.json")


def load_config():
    path = get_config_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(cfg):
    path = get_config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════
#  业务逻辑
# ══════════════════════════════════════════════════════════

def split_pdf_by_order(pdf_path, output_dir):
    """按订单号拆分 PDF，返回 {订单号: pdf路径}"""
    reader = PdfReader(pdf_path)
    order_pages = {}
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        m = re.search(r'订单号\s*(\d+)', text)
        if m:
            order_pages.setdefault(m.group(1), []).append(i)

    os.makedirs(output_dir, exist_ok=True)
    result = {}
    for order_no, pages in order_pages.items():
        writer = PdfWriter()
        for p in pages:
            writer.add_page(reader.pages[p])
        out_path = os.path.join(output_dir, f"{order_no}.pdf")
        with open(out_path, "wb") as f:
            writer.write(f)
        result[order_no] = out_path
    return result


def find_coa_files(order_no, coa_dir, order_batch_map):
    """根据订单号找 COA 文件，返回路径列表"""
    found = []
    if not order_no or order_no not in order_batch_map:
        return found
    for device, batch_str in order_batch_map[order_no]:
        if not device or not batch_str:
            continue
        batches = [b.strip() for b in str(batch_str).split("+") if b.strip()]
        for batch in batches:
            target = f"{device}-{batch}"
            for fname in os.listdir(coa_dir):
                fpath = os.path.join(coa_dir, fname)
                if not os.path.isfile(fpath):
                    continue
                if target.lower() in fname.lower():
                    found.append(fpath)
                    break
    # 去重
    seen, unique = set(), []
    for f in found:
        if f not in seen:
            seen.add(f)
            unique.append(f)
    return unique


def zip_files(file_paths, zip_path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fp in file_paths:
            zf.write(fp, os.path.basename(fp))


def refresh_pivot_tables(excel_path):
    """
    刷新 Excel 数据透视表（仅 Windows + 已安装 Excel 环境支持）
    成功返回 True，失败返回 False
    """
    if not sys.platform.startswith("win"):
        return False
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()  # 等待刷新完成
        wb.Save()
        wb.Close()
        excel.Quit()
        return True
    except Exception:
        return False


def read_email_data(excel_path):
    """
    读取 Excel：
    - 「改单邮件」Sheet：按订单号去重
    - 「订单批次」Sheet：装置 + 批次号
    返回 (email_data列表, order_batch字典)
    """
    wb = load_workbook(excel_path, data_only=True)

    ws = wb["改单邮件"]
    order_email_map = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        order_no = row[0]
        if order_no is None:
            continue
        order_str = str(int(order_no)) if isinstance(order_no, (int, float)) else str(order_no).strip()
        if order_str in order_email_map:
            continue
        title = row[1] if row[1] else ""
        body  = row[3] if row[3] and str(row[3]).strip() != "." else ""
        order_email_map[order_str] = {
            "订单号": order_str,
            "邮件标题": str(title),
            "邮件正文": str(body) if body else "您好，请查收附件。"
        }
    email_data = list(order_email_map.values())

    ws2 = wb["订单批次"]
    order_batch = {}
    for row in ws2.iter_rows(min_row=2, values_only=True):
        order_no, device, batch = row[0], row[1], row[2]
        if order_no is None or not device:
            continue
        order_str = str(int(order_no)) if isinstance(order_no, (int, float)) else str(order_no).strip()
        order_batch.setdefault(order_str, []).append((str(device), str(batch) if batch else ""))

    return email_data, order_batch


def send_email(smtp_server, smtp_port, username, password,
               recipients, subject, body, attachments=None):
    """
    发送邮件，根据端口自动选择加密方式：
    - 465：SMTP_SSL（全程加密）
    - 其他：starttls（明文升级加密）
    """
    msg = MIMEMultipart()
    msg["From"] = username
    msg["To"] = recipients
    msg["Subject"] = Header(subject, "utf-8")
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if attachments:
        for fp in attachments:
            if not os.path.exists(fp):
                continue
            part = MIMEBase("application", "octet-stream")
            with open(fp, "rb") as f:
                part.set_payload(f.read())
            encoders.encode_base64(part)
            fname = os.path.basename(fp)
            part.add_header("Content-Disposition", "attachment",
                            filename=("utf-8", "", fname))
            msg.attach(part)

    # 根据端口选择连接方式
    use_ssl = (str(smtp_port) == "465")
    if use_ssl:
        server = SMTP_SSL(smtp_server, smtp_port, timeout=30)
    else:
        server = SMTP(smtp_server, smtp_port, timeout=30)
        server.starttls()

    server.login(username, password)
    rcpt_list = [r.strip() for r in recipients.split(";") if r.strip()]
    server.sendmail(username, rcpt_list, msg.as_string())
    server.quit()


# ══════════════════════════════════════════════════════════
#  颜色主题（B 端专业风格）
# ══════════════════════════════════════════════════════════

class Theme:
    PRIMARY       = "#2B5AED"
    PRIMARY_HOVER= "#1E45C0"
    SUCCESS       = "#52C41A"
    WARNING       = "#FAAD14"
    ERROR         = "#F5222D"
    BG_PAGE       = "#F5F7FA"
    BG_CARD       = "#FFFFFF"
    BG_INPUT      = "#F7F8FA"
    BORDER        = "#E4E7ED"
    TEXT_PRIMARY   = "#1D2129"
    TEXT_SECONDARY = "#86909C"
    TEXT_LABEL     = "#4E5969"
    DIVIDER        = "#EBEBEB"


# ══════════════════════════════════════════════════════════
#  GUI 主程序
# ══════════════════════════════════════════════════════════

class EmailToolApp:

    def __init__(self, root):
        self.root = root
        self.root.title("改单邮件批量发送工具")
        self.root.geometry("780x880")
        self.root.resizable(False, False)
        self.root.configure(bg=Theme.BG_PAGE)

        # 界面变量
        self.excel_path   = tk.StringVar()
        self.coa_dir      = tk.StringVar()
        self.pdf_path     = tk.StringVar()
        self.output_dir    = tk.StringVar()   # 新增：输出目录
        self.smtp_server  = tk.StringVar(value="smtp.exmail.qq.com")
        self.smtp_port    = tk.StringVar(value="465")
        self.email_user   = tk.StringVar()
        self.email_pass   = tk.StringVar()
        self.recipients   = tk.StringVar()

        self._load_saved_config()
        self._setup_styles()
        self._build_ui()

    # ───── 配置持久化 ─────

    def _load_saved_config(self):
        cfg = load_config()
        if cfg.get("excel_path"):   self.excel_path.set(cfg["excel_path"])
        if cfg.get("coa_dir"):      self.coa_dir.set(cfg["coa_dir"])
        if cfg.get("pdf_path"):     self.pdf_path.set(cfg["pdf_path"])
        if cfg.get("output_dir"):   self.output_dir.set(cfg["output_dir"])
        if cfg.get("smtp_server"): self.smtp_server.set(cfg["smtp_server"])
        if cfg.get("smtp_port"):   self.smtp_port.set(cfg["smtp_port"])
        if cfg.get("email_user"):   self.email_user.set(cfg["email_user"])
        if cfg.get("email_pass"):   self.email_pass.set(cfg["email_pass"])
        if cfg.get("recipients"):  self.recipients.set(cfg["recipients"])

    def _save_current_config(self):
        cfg = {
            "excel_path":  self.excel_path.get(),
            "coa_dir":     self.coa_dir.get(),
            "pdf_path":    self.pdf_path.get(),
            "output_dir":  self.output_dir.get(),
            "smtp_server": self.smtp_server.get(),
            "smtp_port":   self.smtp_port.get(),
            "email_user":  self.email_user.get(),
            "email_pass":  self.email_pass.get(),
            "recipients":  self.recipients.get(),
        }
        save_config(cfg)

    # ───── 样式 ─────

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Card.TLabelframe",
                        background=Theme.BG_CARD,
                        bordercolor=Theme.BORDER,
                        relief="solid",
                        borderwidth=1)
        style.configure("Card.TLabelframe.Label",
                        background=Theme.BG_CARD,
                        foreground=Theme.TEXT_PRIMARY,
                        font=("Microsoft YaHei UI", 11, "bold"),
                        padding=(12, 6))

        style.configure("Input.TEntry",
                        fieldbackground=Theme.BG_INPUT,
                        bordercolor=Theme.BORDER,
                        focuscolor=Theme.PRIMARY,
                        padding=(8, 6))

        style.configure("Primary.TButton",
                        background=Theme.PRIMARY,
                        foreground="white",
                        font=("Microsoft YaHei UI", 11, "bold"),
                        padding=(28, 10),
                        borderwidth=0)
        style.map("Primary.TButton",
                  background=[("active", "#1E45C0"), ("pressed", "#1635A1")])

        style.configure("Secondary.TButton",
                        background=Theme.BG_INPUT,
                        foreground=Theme.TEXT_PRIMARY,
                        font=("Microsoft YaHei UI", 9),
                        padding=(12, 5),
                        borderwidth=1,
                        bordercolor=Theme.BORDER)
        style.map("Secondary.TButton",
                  background=[("active", "#E8E8E8")])

        style.configure("Horizontal.TProgressbar",
                        troughcolor=Theme.BG_INPUT,
                        background=Theme.PRIMARY,
                        borderwidth=0,
                        thickness=6)

    # ───── 界面搭建 ─────

    def _build_ui(self):
        page = tk.Frame(self.root, bg=Theme.BG_PAGE)
        page.pack(fill="both", expand=True, padx=20, pady=(16, 0))

        # 标题区
        header = tk.Frame(page, bg=Theme.BG_PAGE)
        header.pack(fill="x", pady=(0, 12))
        tk.Label(header, text="改单邮件批量发送工具",
                  font=("Microsoft YaHei UI", 16, "bold"),
                  fg=Theme.TEXT_PRIMARY, bg=Theme.BG_PAGE).pack(side="left")
        tk.Label(header, text="v1.2",
                  font=("Microsoft YaHei UI", 9),
                  fg=Theme.TEXT_SECONDARY, bg=Theme.BG_PAGE).pack(side="left", padx=(8, 0))

        # 卡片 1：数据文件
        self._build_file_card(page)

        # 卡片 2：邮箱配置
        self._build_smtp_card(page)

        # 卡片 3：收件人
        self._build_recipient_card(page)

        # 操作区
        self._build_action_area(page)

        # 日志区
        self._build_log_area(page)

        # 状态栏
        self._build_statusbar()

    # ─── 数据文件卡片 ───

    def _build_file_card(self, parent):
        card = ttk.LabelFrame(parent, text="  📂  数据文件  ", style="Card.TLabelframe")
        card.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(card, bg=Theme.BG_CARD, padx=16, pady=8)
        inner.pack(fill="x")

        rows = [
            ("Excel 文件",   self.excel_path, self.select_excel,   "xlsx / xls"),
            ("COA 文件夹",  self.coa_dir,    self.select_coa_dir, "选择文件夹"),
            ("大 PDF 文件",  self.pdf_path,  self.select_pdf,     "pdf"),
            ("输出目录",     self.output_dir,  self.select_output,  "拆分文件存放位置"),
        ]
        for i, (label, var, cmd, _hint) in enumerate(rows):
            tk.Label(inner, text=label, font=("Microsoft YaHei UI", 9),
                     fg=Theme.TEXT_LABEL, bg=Theme.BG_CARD, width=10, anchor="w"
                      ).grid(row=i, column=0, sticky="w", pady=6)
            entry = ttk.Entry(inner, textvariable=var, style="Input.TEntry", width=52)
            entry.grid(row=i, column=1, sticky="ew", padx=(8, 8), pady=6)
            ttk.Button(inner, text="浏览…", style="Secondary.TButton",
                        command=cmd).grid(row=i, column=2, pady=6)

        inner.columnconfigure(1, weight=1)

    def _build_smtp_card(self, parent):
        card = ttk.LabelFrame(parent, text="  ⚙️  邮箱配置  ", style="Card.TLabelframe")
        card.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(card, bg=Theme.BG_CARD, padx=16, pady=8)
        inner.pack(fill="x")

        r = 0
        tk.Label(inner, text="SMTP 服务器", font=("Microsoft YaHei UI", 9),
                 fg=Theme.TEXT_LABEL, bg=Theme.BG_CARD, width=10, anchor="w"
                  ).grid(row=r, column=0, sticky="w", pady=6)
        ttk.Entry(inner, textvariable=self.smtp_server, style="Input.TEntry", width=28
                   ).grid(row=r, column=1, sticky="w", padx=(8, 20), pady=6)
        tk.Label(inner, text="端口", font=("Microsoft YaHei UI", 9),
                 fg=Theme.TEXT_LABEL, bg=Theme.BG_CARD
                  ).grid(row=r, column=2, sticky="w", pady=6)
        ttk.Entry(inner, textvariable=self.smtp_port, style="Input.TEntry", width=8
                   ).grid(row=r, column=3, sticky="w", padx=(8, 0), pady=6)

        r = 1
        tk.Label(inner, text="邮箱账号", font=("Microsoft YaHei UI", 9),
                 fg=Theme.TEXT_LABEL, bg=Theme.BG_CARD, width=10, anchor="w"
                  ).grid(row=r, column=0, sticky="w", pady=6)
        ttk.Entry(inner, textvariable=self.email_user, style="Input.TEntry", width=28
                   ).grid(row=r, column=1, sticky="w", padx=(8, 20), pady=6)
        tk.Label(inner, text="授权码", font=("Microsoft YaHei UI", 9),
                 fg=Theme.TEXT_LABEL, bg=Theme.BG_CARD
                  ).grid(row=r, column=2, sticky="w", pady=6)
        ttk.Entry(inner, textvariable=self.email_pass, style="Input.TEntry", width=18, show="•"
                   ).grid(row=r, column=3, sticky="w", padx=(8, 0), pady=6)

    def _build_recipient_card(self, parent):
        card = ttk.LabelFrame(parent, text="  📧  收件人  ", style="Card.TLabelframe")
        card.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(card, bg=Theme.BG_CARD, padx=16, pady=4)
        inner.pack(fill="x")
        ttk.Entry(inner, textvariable=self.recipients, style="Input.TEntry", width=82
                   ).pack(fill="x", pady=6)
        tk.Label(inner, text="多个收件人用英文分号 ; 分隔",
                  font=("Microsoft YaHei UI", 8),
                  fg=Theme.TEXT_SECONDARY, bg=Theme.BG_CARD, anchor="w"
                   ).pack(fill="x")

    def _build_action_area(self, parent):
        frm = tk.Frame(parent, bg=Theme.BG_PAGE)
        frm.pack(fill="x", pady=(4, 8))

        self.btn_run = ttk.Button(frm, text="▶  开始执行", style="Primary.TButton",
                                   command=self.on_start)
        self.btn_run.pack(side="left")

        self.progress = ttk.Progressbar(frm, mode="indeterminate",
                                         style="Horizontal.TProgressbar", length=400)
        self.progress.pack(side="left", padx=(20, 0), fill="x", expand=True)

    def _build_log_area(self, parent):
        card = ttk.LabelFrame(parent, text="  📋  执行日志  ", style="Card.TLabelframe")
        card.pack(fill="both", expand=True, pady=(0, 6))

        self.log_text = scrolledtext.ScrolledText(
            card, height=12, state="disabled",
            font=("Consolas", 9),
            bg="#FAFBFC", fg=Theme.TEXT_PRIMARY,
            borderwidth=0, highlightthickness=0,
            insertbackground=Theme.TEXT_PRIMARY,
            selectbackground=Theme.PRIMARY,
            selectforeground="white",
            padx=12, pady=8
        )
        self.log_text.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        # 日志颜色标签
        self.log_text.tag_configure("red",   foreground=Theme.ERROR)
        self.log_text.tag_configure("green", foreground=Theme.SUCCESS)

    def _build_statusbar(self):
        bar = tk.Frame(self.root, bg=Theme.BG_CARD, height=28,
                        highlightbackground=Theme.BORDER, highlightthickness=1)
        bar.pack(fill="x", side="bottom")
        self.status_label = tk.Label(
            bar, text="  就绪  |  配置自动保存至本地",
            font=("Microsoft YaHei UI", 8),
            fg=Theme.TEXT_SECONDARY, bg=Theme.BG_CARD, anchor="w"
        )
        self.status_label.pack(fill="x", padx=12, pady=4)

    # ───── 文件选择 ─────

    def select_excel(self):
        p = filedialog.askopenfilename(title="选择 Excel 文件",
                                        filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if p:
            self.excel_path.set(p)
            # 默认输出目录设为 Excel 同级目录
            if not self.output_dir.get():
                self.output_dir.set(os.path.join(os.path.dirname(p), "output"))
            self._save_current_config()

    def select_coa_dir(self):
        p = filedialog.askdirectory(title="选择 COA 文件夹")
        if p:
            self.coa_dir.set(p)
            self._save_current_config()

    def select_pdf(self):
        p = filedialog.askopenfilename(title="选择大 PDF 文件",
                                        filetypes=[("PDF 文件", "*.pdf")])
        if p:
            self.pdf_path.set(p)
            self._save_current_config()

    def select_output(self):
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            self.output_dir.set(p)
            self._save_current_config()

    # ───── 日志 ─────

    def log(self, msg, color=None):
        self.log_text.config(state="normal")
        if color:
            self.log_text.insert("end", msg + "\n", color)
        else:
            self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    # ───── 启动 ─────

    def on_start(self):
        missing = []
        if not self.excel_path.get(): missing.append("Excel 文件")
        if not self.coa_dir.get():    missing.append("COA 文件夹")
        if not self.pdf_path.get():   missing.append("大 PDF 文件")
        if not self.output_dir.get():  missing.append("输出目录")
        if not self.email_user.get(): missing.append("邮箱账号")
        if not self.email_pass.get(): missing.append("授权码")
        if not self.recipients.get(): missing.append("收件人")

        if missing:
            messagebox.showerror("缺少必填项",
                                 "请先填写：\n" + "\n".join(f"  • {m}" for m in missing))
            return

        self._save_current_config()
        threading.Thread(target=self.do_task, daemon=True).start()

    # ───── 主任务 ─────

    def do_task(self):
        self.btn_run.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="  ⏳ 执行中…")

        try:
            excel_path   = self.excel_path.get()
            coa_dir      = self.coa_dir.get()
            pdf_path     = self.pdf_path.get()
            output_dir   = self.output_dir.get()
            smtp_server  = self.smtp_server.get()
            smtp_port    = self.smtp_port.get()
            username     = self.email_user.get()
            password     = self.email_pass.get()
            recipients   = self.recipients.get()

            os.makedirs(output_dir, exist_ok=True)

            # ─── Step 0：刷新透视表 ───
            self.log("正在刷新 Excel 数据透视表…")
            refreshed = refresh_pivot_tables(excel_path)
            if not refreshed:
                self.log("❌ 无法刷新数据透视表！", "red")
                self.log("   原因：当前环境不支持（需 Windows + Excel）", "red")
                self.root.after(0, lambda: messagebox.showerror(
                    "透视表刷新失败",
                    "无法自动刷新 Excel 数据透视表。\n\n"
                    "请确保：\n"
                    "  1. 本机已安装 Microsoft Excel\n"
                    "  2. Excel 文件中的透视表已手动刷新并保存\n\n"
                    "请手动刷新透视表后重新执行。"
                ))
                self.status_label.config(text="  ❌ 透视表刷新失败，已停止")
                return  # ← 停止执行

            self.log("  ✅ 透视表刷新成功", "green")

            # ─── Step 1：读取 Excel ───
            email_data, order_batch = read_email_data(excel_path)
            total = len(email_data)

            # ─── Step 2：拆分 PDF ───
            self.log(f"正在拆分 PDF，共 {total} 个订单待处理…")
            pdf_map = split_pdf_by_order(pdf_path, output_dir)

            # ─── Step 3：逐订单处理 ───
            ok_cnt    = 0
            err_cnt   = 0
            coa_missing = []   # COA 缺失的订单号
            pdf_missing  = []   # PDF 中没有的订单号

            for i, item in enumerate(email_data, 1):
                order_no = item["订单号"]
                self.log(f"\n[{i}/{total}] 订单 {order_no}")

                # 检查 PDF 是否有该订单
                pdf_file = pdf_map.get(order_no)
                if not pdf_file or not os.path.exists(pdf_file):
                    self.log(f"  ❌ PDF 中无此订单，已跳过", "red")
                    pdf_missing.append(order_no)
                    err_cnt += 1
                    continue

                # COA 打包
                coa_files = find_coa_files(order_no, coa_dir, order_batch)
                zip_path = os.path.join(output_dir, f"{order_no}.zip")
                coa_ok   = bool(coa_files)

                if coa_ok:
                    zip_files(coa_files, zip_path)
                else:
                    coa_missing.append(order_no)

                # 邮件标题
                subject = item["邮件标题"]
                if not coa_ok:
                    subject = f"【COA待确认】{subject}"

                # 邮件正文
                body = item["邮件正文"]

                # 附件
                attachments = [pdf_file]
                if coa_ok and os.path.exists(zip_path):
                    attachments.append(zip_path)

                # 发送
                try:
                    send_email(smtp_server, smtp_port, username, password,
                                recipients, subject, body, attachments)
                    self.log(f"  ✅ 发送成功", "green")
                    ok_cnt += 1
                except Exception as e:
                    self.log(f"  ❌ 发送失败：{e}", "red")
                    err_cnt += 1

            # ─── 汇总 ───
            self.log(f"\n{'━' * 50}")
            self.log(f"  执行完成", "green")
            self.log(f"  成功：{ok_cnt} 封  │  失败：{err_cnt} 封")
            if coa_missing:
                self.log(f"  ⚠️  COA 缺失：{len(coa_missing)} 个", "red")
                for o in coa_missing:
                    self.log(f"     • {o}", "red")
            if pdf_missing:
                self.log(f"  ⚠️  PDF 无对应订单：{len(pdf_missing)} 个", "red")
                for o in pdf_missing:
                    self.log(f"     • {o}", "red")
            self.log(f"  输出目录：{output_dir}")
            self.status_label.config(
                text=f"  ✅ 完成 — 成功 {ok_cnt} 封，失败 {err_cnt} 封"
            )

        except Exception as e:
            self.log(f"\n❌ 执行出错：{e}", "red")
            import traceback
            self.log(traceback.format_exc())
            self.status_label.config(text="  ❌ 执行出错")

        finally:
            self.progress.stop()
            self.btn_run.config(state="normal")


# ══════════════════════════════════════════════════════════
#  启动
# ══════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailToolApp(root)
    root.mainloop()
