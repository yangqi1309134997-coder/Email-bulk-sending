import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, colorchooser, simpledialog
import pandas as pd
import os
import time
import random
import threading
from itertools import cycle
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import pickle

# ========================= 配置常量 =========================
MAX_RETRIES = 3
DEFAULT_DELAY_MIN = 5
DEFAULT_DELAY_MAX = 15

SMTP_TEMPLATES = {
    "QQ邮箱": {"smtp_server": "smtp.qq.com", "port": 587, "use_tls": True},
    "163邮箱": {"smtp_server": "smtp.163.com", "port": 587, "use_tls": True},
    "Gmail": {"smtp_server": "smtp.gmail.com", "port": 587, "use_tls": True},
    "Outlook/Hotmail": {"smtp_server": "smtp-mail.outlook.com", "port": 587, "use_tls": True},
    "Yahoo": {"smtp_server": "smtp.mail.yahoo.com", "port": 587, "use_tls": True},
    "自定义SMTP": {"smtp_server": "", "port": 587, "use_tls": True},
}

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🚀 专业批量邮件发送工具 v3.0 - 完全稳定版")
        self.root.geometry("1300x900")
        self.root.state('zoomed')  # 默认最大化

        self.senders = []
        self.proxies = []
        self.recipients = []  # [{"email": "", "name": ""}]
        self.attachments = []
        self.send_report = []

        self.subject_template = ""
        self.body_html_template = ""

        self.body_images = []  # 保持插入图片的引用，防止消失

        self.build_ui()

    def build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # ==================== 发件人配置 ====================
        tab_sender = ttk.Frame(notebook)
        notebook.add(tab_sender, text="发件人配置")

        frame = tk.LabelFrame(tab_sender, text="添加发件人")
        frame.pack(fill="x", padx=10, pady=10)

        tk.Label(frame, text="邮箱类型：").grid(row=0, column=0, sticky="e", pady=5)
        self.sender_type = tk.StringVar(value="QQ邮箱")
        ttk.Combobox(frame, textvariable=self.sender_type, values=list(SMTP_TEMPLATES.keys()), state="readonly", width=20).grid(row=0, column=1, pady=5, padx=5)

        tk.Label(frame, text="发件邮箱：").grid(row=1, column=0, sticky="e", pady=5)
        self.sender_email = tk.StringVar()
        tk.Entry(frame, textvariable=self.sender_email, width=40).grid(row=1, column=1, pady=5, padx=5)

        tk.Label(frame, text="授权码/密码：").grid(row=2, column=0, sticky="e", pady=5)
        self.sender_pass = tk.StringVar()
        tk.Entry(frame, textvariable=self.sender_pass, show="*", width=40).grid(row=2, column=1, pady=5, padx=5)

        tk.Button(frame, text="添加发件人", bg="#007bff", fg="white", command=self.add_sender).grid(row=3, column=1, pady=10, sticky="e")

        self.sender_tree = ttk.Treeview(tab_sender, columns=("email",), show="headings", height=8)
        self.sender_tree.heading("email", text="发件人邮箱")
        self.sender_tree.column("email", width=600, anchor="w")
        self.sender_tree.pack(fill="both", expand=True, padx=10, pady=5)
        tk.Button(tab_sender, text="删除选中", command=self.remove_sender).pack(pady=5)

        # ==================== 收件人配置 ====================
        tab_recipient = ttk.Frame(notebook)
        notebook.add(tab_recipient, text="收件人配置")

        btn_frame = tk.Frame(tab_recipient)
        btn_frame.pack(fill="x", pady=10)
        tk.Button(btn_frame, text="从TXT/CSV/Excel导入（支持 邮箱 或 邮箱,姓名）", command=self.import_recipients, bg="#28a745", fg="white").pack(side="left", padx=10)
        tk.Button(btn_frame, text="清空列表", command=self.clear_recipients).pack(side="right", padx=10)

        self.recipient_tree = ttk.Treeview(tab_recipient, columns=("email", "name"), show="headings", height=18)
        self.recipient_tree.heading("email", text="收件人邮箱")
        self.recipient_tree.heading("name", text="姓名（可选）")
        self.recipient_tree.column("email", width=500)
        self.recipient_tree.column("name", width=200)
        self.recipient_tree.pack(fill="both", expand=True, padx=10)

        # ==================== 邮件内容 ====================
        tab_content = ttk.Frame(notebook)
        notebook.add(tab_content, text="邮件内容")

        tk.Label(tab_content, text="邮件主题（支持 {name} 变量）:", font=("Arial", 12, "bold")).pack(anchor="w", padx=20, pady=(20,5))
        self.subject_var = tk.StringVar(value="亲爱的{name}，您有一封重要邮件")
        tk.Entry(tab_content, textvariable=self.subject_var, font=("Arial", 12), width=100).pack(padx=20, pady=5, fill="x")

        tk.Label(tab_content, text="邮件正文（富文本编辑，支持 {name} 变量）:", font=("Arial", 12, "bold")).pack(anchor="w", padx=20, pady=(15,5))

        # 工具栏
        toolbar = tk.Frame(tab_content)
        toolbar.pack(fill="x", padx=20, pady=(0,8))

        tk.Button(toolbar, text="粗体", command=lambda: self.format_body("bold")).pack(side="left", padx=2)
        tk.Button(toolbar, text="斜体", command=lambda: self.format_body("italic")).pack(side="left", padx=2)
        tk.Button(toolbar, text="下划线", command=lambda: self.format_body("underline")).pack(side="left", padx=2)
        tk.Button(toolbar, text="颜色", command=self.set_body_color).pack(side="left", padx=2)

        tk.Label(toolbar, text="  字体大小:").pack(side="left", padx=(30,5))
        self.font_size_var = tk.IntVar(value=11)
        size_combo = ttk.Combobox(toolbar, textvariable=self.font_size_var, values=[8,10,11,12,14,16,18,20,24,28,32,36], width=6, state="readonly")
        size_combo.pack(side="left", padx=2)
        size_combo.bind("<<ComboboxSelected>>", lambda e: self.format_body("size"))

        tk.Button(toolbar, text="插入链接", command=self.insert_link_to_body).pack(side="left", padx=20)
        tk.Button(toolbar, text="插入图片", command=self.insert_image_to_body).pack(side="left", padx=2)

        # 正文编辑区
        self.body_text = scrolledtext.ScrolledText(tab_content, height=25, wrap="word", font=("Arial", 11), undo=True)
        self.body_text.pack(fill="both", expand=True, padx=20, pady=(0,15))

        # 默认内容（HTML格式）
        default_body = """<p>亲爱的<b>{name}</b>，</p>
<p>感谢您的关注与支持！</p>
<p>这是一封来自专业邮件工具的测试邮件。</p>
<p style="color:#0066cc;">祝您一切顺利！</p>
<p>—— 您的朋友</p>"""
        self.body_text.insert("end", default_body)

        # ==================== 附件与模板 ====================
        tab_attach = ttk.Frame(notebook)
        notebook.add(tab_attach, text="附件与模板")

        left_frame = tk.Frame(tab_attach)
        left_frame.pack(side="left", fill="both", expand=True, padx=10)

        tk.Label(left_frame, text="附件管理", font=("Arial", 12, "bold")).pack(anchor="w", pady=(0,10))
        tk.Button(left_frame, text="添加附件（支持多选）", command=self.add_attachment).pack(pady=5)
        self.attach_listbox = tk.Listbox(left_frame, height=15)
        self.attach_listbox.pack(fill="both", expand=True, pady=5)
        tk.Button(left_frame, text="删除选中附件", command=self.remove_attachment).pack(pady=5)

        right_frame = tk.Frame(tab_attach)
        right_frame.pack(side="right", fill="y", padx=30, pady=20)

        tk.Label(right_frame, text="模板管理", font=("Arial", 12, "bold")).pack(anchor="w", pady=(0,20))
        tk.Button(right_frame, text="保存当前为模板", command=self.save_template, bg="#ffc107", fg="black", height=2).pack(fill="x", pady=8)
        tk.Button(right_frame, text="加载模板", command=self.load_template, bg="#17a2b8", fg="white", height=2).pack(fill="x", pady=8)

        # ==================== 高级设置 ====================
        tab_advanced = ttk.Frame(notebook)
        notebook.add(tab_advanced, text="高级设置")

        tk.Label(tab_advanced, text="随机延迟范围（秒）:", font=("Arial", 11)).grid(row=0, column=0, sticky="w", pady=20, padx=20)
        self.delay_min = tk.IntVar(value=DEFAULT_DELAY_MIN)
        self.delay_max = tk.IntVar(value=DEFAULT_DELAY_MAX)
        tk.Entry(tab_advanced, textvariable=self.delay_min, width=8).grid(row=0, column=1, padx=5)
        tk.Label(tab_advanced, text=" ~ ").grid(row=0, column=2)
        tk.Entry(tab_advanced, textvariable=self.delay_max, width=8).grid(row=0, column=3, padx=5)

        tk.Label(tab_advanced, text="代理列表（每行一个 http://ip:port，可留空）:", font=("Arial", 11)).grid(row=1, column=0, sticky="nw", pady=20, padx=20)
        self.proxy_text = scrolledtext.ScrolledText(tab_advanced, height=12)
        self.proxy_text.grid(row=2, column=0, columnspan=4, sticky="nsew", padx=20, pady=5)
        tab_advanced.grid_rowconfigure(2, weight=1)
        tab_advanced.grid_columnconfigure(0, weight=1)

        # ==================== 发送日志 ====================
        tab_log = ttk.Frame(notebook)
        notebook.add(tab_log, text="发送日志与报告")

        self.log_text = scrolledtext.ScrolledText(tab_log, state="disabled", height=25)
        self.log_text.pack(fill="both", expand=True, padx=20, pady=10)

        self.progress = ttk.Progressbar(tab_log, mode='determinate')
        self.progress.pack(fill="x", padx=20, pady=10)

        btn_frame2 = tk.Frame(tab_log)
        btn_frame2.pack(pady=30)
        tk.Button(btn_frame2, text="🚀 开始发送", bg="#dc3545", fg="white", font=("Arial", 18, "bold"),
                  command=self.start_sending_thread, width=20, height=2).pack(side="left", padx=50)
        tk.Button(btn_frame2, text="📊 导出发送报告（CSV）", command=self.export_report, bg="#6c757d", fg="white", font=("Arial", 14)).pack(side="right", padx=50)

        tk.Label(tab_log, text="提示：使用 {name} 变量可实现个性化发送；支持插入图片、链接、富文本格式化", foreground="gray").pack(pady=10)

    # ====================== 富文本功能 ======================
    def format_body(self, style):
        try:
            if style == "bold":
                self.body_text.tag_add("bold", "sel.first", "sel.last")
                self.body_text.tag_config("bold", font=("Arial", 11, "bold"))
            elif style == "italic":
                self.body_text.tag_add("italic", "sel.first", "sel.last")
                self.body_text.tag_config("italic", font=("Arial", 11, "italic"))
            elif style == "underline":
                self.body_text.tag_add("underline", "sel.first", "sel.last")
                self.body_text.tag_config("underline", underline=True)
            elif style == "size":
                size = self.font_size_var.get()
                tag = f"size_{size}"
                self.body_text.tag_config(tag, font=("Arial", size))
                self.body_text.tag_add(tag, "sel.first", "sel.last")
        except tk.TclError:
            pass

    def set_body_color(self):
        color = colorchooser.askcolor(title="选择文字颜色")[1]
        if color:
            tag = f"color_{color.replace('#', '')}"
            self.body_text.tag_config(tag, foreground=color)
            try:
                self.body_text.tag_add(tag, "sel.first", "sel.last")
            except:
                pass

    def insert_link_to_body(self):
        url = simpledialog.askstring("插入链接", "请输入URL：")
        if url:
            self.body_text.insert("insert", url, "link")
            self.body_text.tag_config("link", foreground="blue", underline=True)
            self.body_text.tag_bind("link", "<Button-1>", lambda e: os.startfile(url))

    def insert_image_to_body(self):
        file = filedialog.askopenfilename(filetypes=[("图片文件", "*.png *.jpg *.jpeg *.gif *.bmp")])
        if file:
            try:
                img = tk.PhotoImage(file=file).subsample(3, 3)  # 缩小显示
                self.body_text.image_create("insert", image=img)
                self.body_text.insert("insert", "\n")
                self.body_images.append(img)  # 保持引用
            except Exception as e:
                messagebox.showerror("错误", f"无法加载图片：{e}")

    # ====================== 其他功能 ======================
    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def add_sender(self):
        email = self.sender_email.get().strip()
        pwd = self.sender_pass.get().strip()
        if not email or not pwd:
            messagebox.showwarning("警告", "请填写完整")
            return
        smtp_cfg = SMTP_TEMPLATES[self.sender_type.get()].copy()
        if self.sender_type.get() == "自定义SMTP":
            server = simpledialog.askstring("SMTP服务器", "请输入SMTP服务器地址：")
            port = simpledialog.askinteger("端口", "SMTP端口：", initialvalue=587)
            if not server or not port:
                return
            smtp_cfg["smtp_server"] = server
            smtp_cfg["port"] = port
        self.senders.append({"email": email, "password": pwd, "smtp": smtp_cfg})
        self.sender_tree.insert("", "end", values=(email,))
        self.sender_email.set("")
        self.sender_pass.set("")
        self.log(f"添加发件人：{email}")

    def remove_sender(self):
        sel = self.sender_tree.selection()
        if sel:
            idx = self.sender_tree.index(sel[0])
            del self.senders[idx]
            self.sender_tree.delete(sel[0])

    def import_recipients(self):
        file = filedialog.askopenfilename(filetypes=[("所有支持文件", "*.txt *.csv *.xlsx *.xls"), ("文本文件", "*.txt")])
        if not file:
            return
        try:
            new_recipients = []
            if file.endswith(".txt"):
                with open(file, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line or "@" not in line:
                            continue
                        parts = line.split(",", 1)
                        email = parts[0].strip()
                        name = parts[1].strip() if len(parts) > 1 else ""
                        new_recipients.append({"email": email, "name": name})
            else:
                if file.endswith(".csv"):
                    df = pd.read_csv(file, header=None)
                else:
                    df = pd.read_excel(file, header=None)
                for _, row in df.iterrows():
                    email = str(row.iloc[0]).strip()
                    name = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
                    if "@" in email:
                        new_recipients.append({"email": email, "name": name})

            self.recipients.extend(new_recipients)
            for r in new_recipients:
                self.recipient_tree.insert("", "end", values=(r["email"], r["name"] or "（无）"))
            self.log(f"成功导入 {len(new_recipients)} 个收件人")
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def clear_recipients(self):
        self.recipients.clear()
        for i in self.recipient_tree.get_children():
            self.recipient_tree.delete(i)
        self.log("收件人列表已清空")

    def add_attachment(self):
        files = filedialog.askopenfilenames()
        if files:
            added = 0
            for f in files:
                if f not in self.attachments:
                    self.attachments.append(f)
                    self.attach_listbox.insert("end", os.path.basename(f))
                    added += 1
            self.log(f"添加 {added} 个附件")

    def remove_attachment(self):
        sel = self.attach_listbox.curselection()
        if sel:
            idx = sel[0]
            del self.attachments[idx]
            self.attach_listbox.delete(idx)

    def save_template(self):
        name = simpledialog.askstring("保存模板", "请输入模板名称：")
        if not name:
            return
        data = {
            "subject": self.subject_var.get(),
            "body": self.body_text.get("1.0", "end"),
            "attachments": self.attachments[:]
        }
        with open(f"template_{name}.pkl", "wb") as f:
            pickle.dump(data, f)
        self.log(f"模板 '{name}' 已保存")

    def load_template(self):
        file = filedialog.askopenfilename(filetypes=[("模板文件", "*.pkl")])
        if not file:
            return
        try:
            with open(file, "rb") as f:
                data = pickle.load(f)
            self.subject_var.set(data.get("subject", ""))
            self.body_text.delete("1.0", "end")
            self.body_text.insert("1.0", data.get("body", ""))
            self.attachments = data.get("attachments", [])
            self.attach_listbox.delete(0, "end")
            for a in self.attachments:
                self.attach_listbox.insert("end", os.path.basename(a))
            self.log(f"模板已加载：{os.path.basename(file)}")
        except Exception as e:
            messagebox.showerror("加载失败", str(e))

    def personalize(self, text, name):
        return text.replace("{name}", name if name else "朋友")

    def add_attachments(self, msg):
        for path in self.attachments:
            try:
                with open(path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename=\"{os.path.basename(path)}\"")
                msg.attach(part)
            except Exception as e:
                self.log(f"附件失败 {os.path.basename(path)}: {e}")

    def send_with_smtp(self, sender, recipient):
        try:
            msg = MIMEMultipart()
            msg["From"] = sender["email"]
            msg["To"] = recipient["email"]
            msg["Subject"] = self.personalize(self.subject_template, recipient["name"])

            body = self.personalize(self.body_html_template, recipient["name"])
            msg.attach(MIMEText(body, "html", "utf-8"))

            self.add_attachments(msg)

            server = smtplib.SMTP(sender["smtp"]["smtp_server"], sender["smtp"]["port"], timeout=30)
            if sender["smtp"].get("use_tls", False):
                server.starttls()
            server.login(sender["email"], sender["password"])
            server.send_message(msg)
            server.quit()
            return True, ""
        except Exception as e:
            return False, str(e)

    def start_sending_thread(self):
        if not self.senders:
            messagebox.showwarning("警告", "请添加至少一个发件人")
            return
        if not self.recipients:
            messagebox.showwarning("警告", "请导入收件人")
            return

        self.subject_template = self.subject_var.get()
        self.body_html_template = self.body_text.get("1.0", "end").strip()
        if not self.subject_template or not self.body_html_template:
            messagebox.showwarning("警告", "请填写主题和正文")
            return

        proxy_lines = self.proxy_text.get("1.0", "end").strip().splitlines()
        self.proxies = [{"http": p.strip(), "https": p.strip()} for p in proxy_lines if p.strip()]

        self.send_report = []
        threading.Thread(target=self.send_batch, daemon=True).start()

    def send_batch(self):
        total = len(self.recipients)
        success = 0

        self.log(f"开始发送 {total} 封邮件，使用 {len(self.senders)} 个发件人，{len(self.attachments)} 个附件")
        sender_cycle = cycle(self.senders)
        proxy_cycle = cycle(self.proxies or [None])

        self.progress["maximum"] = total
        self.progress["value"] = 0

        for i, recipient in enumerate(self.recipients, 1):
            sender = next(sender_cycle)
            name = recipient["name"] or "无"

            self.log(f"[{i}/{total}] {sender['email']} → {recipient['email']} ({name})")

            sent = False
            error_msg = ""
            for attempt in range(MAX_RETRIES + 1):
                ok, err = self.send_with_smtp(sender, recipient)
                if ok:
                    self.log("✅ 发送成功")
                    success += 1
                    sent = True
                    break
                else:
                    error_msg = err
                    self.log(f"⚠️ 第{attempt+1}次失败: {err[:60]}...")
                if attempt < MAX_RETRIES:
                    time.sleep(2 ** attempt)

            self.send_report.append({
                "序号": i,
                "收件人邮箱": recipient["email"],
                "姓名": recipient["name"] or "",
                "状态": "成功" if sent else "失败",
                "失败原因": "" if sent else error_msg,
                "发送时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "发件人": sender["email"]
            })

            delay = random.uniform(self.delay_min.get(), self.delay_max.get())
            self.log(f"等待 {delay:.1f} 秒...")
            time.sleep(delay)

            self.progress["value"] = i
            self.root.update_idletasks()

        self.log(f"发送完成！成功 {success}/{total}")
        messagebox.showinfo("完成", f"批量发送完成！\n成功：{success}\n失败：{total-success}\n报告已记录，可点击“导出发送报告”保存")

    def export_report(self):
        if not self.send_report:
            messagebox.showinfo("提示", "暂无发送记录")
            return
        file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV文件", "*.csv")], title="保存发送报告")
        if file:
            pd.DataFrame(self.send_report).to_csv(file, index=False, encoding="utf-8-sig")
            self.log(f"发送报告已导出：{file}")
            messagebox.showinfo("成功", "报告导出成功！")


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()