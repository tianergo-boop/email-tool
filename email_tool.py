import os
import sys
import zipfile
import smtplib
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import re
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# PDF & Excel
from pypdf import PdfReader, PdfWriter
from openpyxl import load_workbook

# ===================== 工具函数 =====================

def split_pdf_by_order(pdf_path, output_dir):
    """
    按订单号拆分PDF。
    每页提取文字，用正则找「订单号」后面的数字。
    同订单的多页合并为一个PDF。
    返回 {订单号(str): 输出pdf路径}
    """
    reader = PdfReader(pdf_path)
    order_pages = {}  # {order_no: [page_index, ...]}

    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        m = re.search(r'订单号\s*(\d+)', text)
        if m:
            order_no = m.group(1)
            order_pages.setdefault(order_no, []).append(i)

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
    """
    根据订单号，在 coa_dir 里找匹配的 COA 文件。
    匹配规则：文件名包含「装置-批次号」，不区分大小写。
    批次号可能用 + 连接多个，每个都要找。
    返回找到的文件路径列表。
    """
    found = []
    if not order_no or order_no not in order_batch_map:
        return found

    for device, batch_str in order_batch_map[order_no]:
        if not device or not batch_str:
            continue
        batches = [b.strip() for b in str(batch_str).split("+") if b.strip()]
        for batch in batches:
            # 在 coa_dir 里模糊匹配：文件名包含 device-batch
            target = f"{device}-{batch}"
            for fname in os.listdir(coa_dir):
                fpath = os.path.join(coa_dir, fname)
                if not os.path.isfile(fpath):
                    continue
                # 不区分大小写匹配
                if target.lower() in fname.lower():
                    found.append(fpath)
                    break  # 每个批次只取第一个匹配的文件

    # 去重（同一个文件可能被多个批次匹配）
    seen = set()
    unique = []
    for f in found:
        if f not in seen:
            seen.add(f)
            unique.append(f)
    return unique


def zip_files(file_paths, zip_path):
    """将多个文件打包为一个 ZIP"""
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fp in file_paths:
            zf.write(fp, os.path.basename(fp))


def refresh_pivot_tables(excel_path):
    """
    尝试刷新 Excel 中的数据透视表。
    Windows 下用 win32com（需要安装 pywin32，且本机有 Excel）。
    macOS/Linux 下无法自动刷新，返回 False 并提示用户手动刷新。
    """
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            wb.RefreshAll()
            wb.Save()
            wb.Close()
            excel.Quit()
            return True
        except Exception as e:
            print(f"自动刷新透视表失败: {e}")
            return False
    else:
        # macOS/Linux 无法用 COM 操控 Excel
        return False


def read_email_data(excel_path):
    """
    读取 Excel 两个 Sheet：
    - 「改单邮件」：按订单号去重，同一订单只发一封邮件
    - 「订单批次」：订单号、装置、批次号
    返回 (email_data列表, order_batch字典)
    """
    wb = load_workbook(excel_path, data_only=True)

    # --- 改单邮件 Sheet ---
    ws = wb["改单邮件"]
    # 用字典按订单号去重，同一订单只保留第一条
    order_email_map = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        order_no = row[0]
        if order_no is None:
            continue
        order_str = str(int(order_no)) if isinstance(order_no, (int, float)) else str(order_no).strip()
        if order_str in order_email_map:
            continue  # 同一订单已记录，跳过
        title = row[1] if row[1] else ""
        body  = row[3] if row[3] and str(row[3]).strip() != "." else ""
        order_email_map[order_str] = {
            "订单号": order_str,
            "邮件标题": str(title),
            "邮件正文": str(body) if body else "您好，请查收附件。"
        }

    email_data = list(order_email_map.values())

    # --- 订单批次 Sheet ---
    ws2 = wb["订单批次"]
    order_batch = {}

    for row in ws2.iter_rows(min_row=2, values_only=True):
        order_no = row[0]
        device   = row[1]
        batch    = row[2]
        if order_no is None or not device:
            continue
        order_str = str(int(order_no)) if isinstance(order_no, (int, float)) else str(order_no).strip()
        order_batch.setdefault(order_str, []).append((str(device), str(batch) if batch else ""))

    return email_data, order_batch


def send_email(smtp_server, smtp_port, username, password,
               recipients, subject, body, attachments=None):
    """发送一封邮件（带附件）"""
    msg = MIMEMultipart()
    msg["From"] = username
    msg["To"] = recipients
    # 标题用 Header 编码，支持中文
    msg["Subject"] = Header(subject, "utf-8")

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # 添加附件（处理中文文件名）
    if attachments:
        for fp in attachments:
            if not os.path.exists(fp):
                continue
            part = MIMEBase("application", "octet-stream")
            with open(fp, "rb") as f:
                part.set_payload(f.read())
            encoders.encode_base64(part)
            # RFC 5987 编码文件名，支持中文
            fname = os.path.basename(fp)
            part.add_header(
                "Content-Disposition",
                "attachment",
                filename=("utf-8", "", fname)
            )
            msg.attach(part)

    # 发送
    server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
    server.starttls()
    server.login(username, password)
    rcpt_list = [r.strip() for r in recipients.split(";") if r.strip()]
    server.sendmail(username, rcpt_list, msg.as_string())
    server.quit()


# ===================== GUI =====================

class EmailToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("改单邮件批量发送工具")
        self.root.geometry("750x800")
        self.root.resizable(False, False)

        # 所有界面变量
        self.excel_path  = tk.StringVar()
        self.coa_dir     = tk.StringVar()
        self.pdf_path    = tk.StringVar()
        self.smtp_server = tk.StringVar(value="smtp.qq.com")
        self.smtp_port   = tk.StringVar(value="587")
        self.email_user  = tk.StringVar()
        self.email_pass  = tk.StringVar()
        self.recipients  = tk.StringVar()

        self.build_ui()

    # -------- 搭建界面 --------
    def build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # ===== 文件选择区 =====
        frm = ttk.LabelFrame(self.root, text="文件选择")
        frm.pack(fill="x", padx=10, pady=(10, 5))

        ttk.Label(frm, text="Excel 文件：").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.excel_path, width=52).grid(row=0, column=1, **pad)
        ttk.Button(frm, text="浏览…", command=self.select_excel).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="COA 文件夹：").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.coa_dir, width=52).grid(row=1, column=1, **pad)
        ttk.Button(frm, text="浏览…", command=self.select_coa_dir).grid(row=1, column=2, **pad)

        ttk.Label(frm, text="大 PDF 文件：").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.pdf_path, width=52).grid(row=2, column=1, **pad)
        ttk.Button(frm, text="浏览…", command=self.select_pdf).grid(row=2, column=2, **pad)

        # ===== SMTP 配置区 =====
        frm = ttk.LabelFrame(self.root, text="SMTP 邮箱配置")
        frm.pack(fill="x", padx=10, pady=5)

        ttk.Label(frm, text="SMTP 服务器：").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.smtp_server, width=22).grid(row=0, column=1, **pad)
        ttk.Label(frm, text="端口：").grid(row=0, column=2, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.smtp_port, width=8).grid(row=0, column=3, **pad)

        ttk.Label(frm, text="邮箱账号：").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.email_user, width=22).grid(row=1, column=1, **pad)
        ttk.Label(frm, text="邮箱授权码：").grid(row=1, column=2, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.email_pass, width=18, show="*").grid(row=1, column=3, **pad)

        # ===== 收件人 =====
        frm = ttk.LabelFrame(self.root, text="收件人（多个用  ;  分隔）")
        frm.pack(fill="x", padx=10, pady=5)
        ttk.Entry(frm, textvariable=self.recipients, width=88).pack(padx=10, pady=8)

        # ===== 执行按钮 =====
        self.btn_run = ttk.Button(self.root, text="▶ 开始执行", command=self.on_start)
        self.btn_run.pack(pady=10)

        # ===== 进度条 =====
        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.pack(fill="x", padx=10)

        # ===== 日志区 =====
        frm = ttk.LabelFrame(self.root, text="执行日志")
        frm.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text = scrolledtext.ScrolledText(frm, height=22, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)

    # -------- 文件对话框 --------
    def select_excel(self):
        p = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if p:
            self.excel_path.set(p)

    def select_coa_dir(self):
        p = filedialog.askdirectory(title="选择 COA 文件夹")
        if p:
            self.coa_dir.set(p)

    def select_pdf(self):
        p = filedialog.askopenfilename(title="选择大 PDF 文件", filetypes=[("PDF 文件", "*.pdf")])
        if p:
            self.pdf_path.set(p)

    # -------- 日志输出 --------
    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    # -------- 点击"开始执行" --------
    def on_start(self):
        # 校验必填项
        missing = []
        if not self.excel_path.get(): missing.append("Excel 文件")
        if not self.coa_dir.get():   missing.append("COA 文件夹")
        if not self.pdf_path.get():   missing.append("大 PDF 文件")
        if not self.email_user.get(): missing.append("邮箱账号")
        if not self.email_pass.get(): missing.append("邮箱授权码")
        if not self.recipients.get(): missing.append("收件人")

        if missing:
            messagebox.showerror("缺少必填项", "请先填写：\n" + "\n".join(f"  • {m}" for m in missing))
            return

        # 在新线程执行，避免界面卡死
        threading.Thread(target=self.do_task, daemon=True).start()

    # -------- 主任务（后台线程） --------
    def do_task(self):
        self.btn_run.config(state="disabled")
        self.progress.start()

        try:
            excel_path  = self.excel_path.get()
            coa_dir    = self.coa_dir.get()
            pdf_path   = self.pdf_path.get()
            smtp_server = self.smtp_server.get()
            smtp_port   = int(self.smtp_port.get())
            username    = self.email_user.get()
            password    = self.email_pass.get()
            recipients  = self.recipients.get()

            # 输出目录：Excel同级的 output 文件夹
            work_dir = os.path.join(os.path.dirname(excel_path), "output")
            os.makedirs(work_dir, exist_ok=True)

            # Step 0：刷新透视表
            self.log("🔄 尝试刷新 Excel 数据透视表…")
            refreshed = refresh_pivot_tables(excel_path)
            if refreshed:
                self.log("   ✅ 透视表已刷新")
            else:
                self.log("   ⚠️  无法自动刷新透视表（仅 Windows+Excel 环境支持）")
                self.log("       请确保 Excel 文件中的透视表已手动刷新！")

            # Step 1：读取 Excel
            self.log("\n📖 读取 Excel 数据…")
            email_data, order_batch = read_email_data(excel_path)
            self.log(f"   改单邮件：{len(email_data)} 条")
            self.log(f"   订单批次：{len(order_batch)} 个订单")

            # Step 2：拆分 PDF
            self.log("\n✂️  拆分 PDF…")
            pdf_map = split_pdf_by_order(pdf_path, work_dir)
            self.log(f"   拆分完成，共 {len(pdf_map)} 个订单 PDF")

            # Step 3：逐订单打包 COA + 发邮件
            total  = len(email_data)
            ok_cnt = 0
            err_cnt = 0

            for i, item in enumerate(email_data, 1):
                order_no = item["订单号"]
                self.log(f"\n[{i}/{total}] 订单 {order_no}")

                # --- COA 打包 ---
                coa_files = find_coa_files(order_no, coa_dir, order_batch)
                zip_path  = os.path.join(work_dir, f"{order_no}.zip")
                coa_ok    = bool(coa_files)

                if coa_ok:
                    zip_files(coa_files, zip_path)
                    self.log(f"   ✅ COA 打包：{len(coa_files)} 个文件")
                else:
                    self.log(f"   ⚠️  未找到 COA 文件")

                # --- 邮件标题（COA 缺失时加前缀）---
                subject = item["邮件标题"]
                if not coa_ok:
                    subject = f"【COA待确认】{subject}"

                # --- 邮件正文 ---
                body = item["邮件正文"]

                # --- 附件列表 ---
                attachments = []
                pdf_file = pdf_map.get(order_no)
                if pdf_file and os.path.exists(pdf_file):
                    attachments.append(pdf_file)
                if coa_ok and os.path.exists(zip_path):
                    attachments.append(zip_path)

                # --- 发送 ---
                try:
                    send_email(smtp_server, smtp_port, username, password,
                               recipients, subject, body, attachments)
                    self.log(f"   ✅ 邮件发送成功")
                    ok_cnt += 1
                except Exception as e:
                    self.log(f"   ❌ 发送失败：{e}")
                    err_cnt += 1

            self.log(f"\n{'='*40}")
            self.log(f"🎉 全部完成！成功 {ok_cnt} 封，失败 {err_cnt} 封")
            self.log(f"输出文件在：{work_dir}")

        except Exception as e:
            self.log(f"\n❌ 执行出错：{e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.progress.stop()
            self.btn_run.config(state="normal")


# ===================== 主入口 =====================

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailToolApp(root)
    root.mainloop()
