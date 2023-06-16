import tkinter as tk
from tkinter import ttk
import smtplib
import socket
import requests
import datetime
import pyautogui
import threading
import os
import sys
import time
import ctypes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.header import Header

USER32 = ctypes.windll.user32
SW_SHOW = 5
SW_HIDE = 0

class EmailChecker:
    def __init__(self):
        self.sent_email = False
        self.stop_event = threading.Event()
        self.thread = None
        self.root = tk.Tk()
        self.root.title("Smilinghan-SPCWC")
        self.create_widgets()

    def create_widgets(self):
        mainframe = ttk.Frame(self.root, padding="20 10 20 10")
        mainframe.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))
        mainframe.columnconfigure(0, weight=1)
        mainframe.rowconfigure(0, weight=1)

        email_entry = ttk.Entry(mainframe, width=30)
        email_entry.grid(column=1, row=0, sticky=(tk.W, tk.E))

        email_label = ttk.Label(mainframe, text="Email:")
        email_label.grid(column=0, row=0, sticky=tk.E)

        submit_button = ttk.Button(mainframe, text="确定", command=lambda: self.submit_email(email_entry))
        submit_button.grid(column=2, row=0, sticky=tk.W)

        start_button = ttk.Button(mainframe, text="开始检测", command=self.start_thread)
        start_button.grid(column=0, row=2, sticky=tk.W)

        stop_button = ttk.Button(mainframe, text="停止检测", command=self.stop_thread)
        stop_button.grid(column=2, row=2, sticky=tk.E)

        output_text = tk.Text(mainframe, height=10, width=50)
        output_text.grid(column=0, row=3, columnspan=3, sticky=(tk.W, tk.E))

        for child in mainframe.winfo_children():
            child.grid_configure(padx=5, pady=5)

        email_entry.focus()
        self.root.bind('<Return>', lambda event: self.submit_email(email_entry))

        self.output_text = output_text

    def submit_email(self, email_entry):
        email = email_entry.get()
        self.output_text.insert(tk.END, "接收的电子邮件: " + email + "\n")

        with open("emails.txt", "w") as f:
            f.write(email + "\n")

        email_entry.delete(0, tk.END)
        email_entry.insert(0, "")

    def send_email(self, subject, content, receiver):
        sender = 'smilinghan@qq.com'
        password = 'password'

        with open("emails.txt", "r") as f:
            receivers = [f.read().strip()]

        message = MIMEMultipart()

        message['Subject'] = Header(subject, 'utf-8')
        message['From'] = Header(sender)
        message['To'] = Header(receiver)

        text = MIMEText(content, 'html', 'utf-8')
        message.attach(text)

        screenshot = pyautogui.screenshot()
        if screenshot is not None:
            screenshot.save('screenshot.png')
            with open('screenshot.png', 'rb') as f:
                img = MIMEImage(f.read())
                img.add_header('Content-ID', '<screenshot>')
                message.attach(img)

        try:
            smtpObj = smtplib.SMTP_SSL('smtp.qq.com', 465)
            smtpObj.login(sender, password)
            smtpObj.sendmail(sender, receivers, message.as_string())
            self.output_text.insert(tk.END, "邮件发送成功\n")
            self.sent_email = True
        except Exception as e:
            self.output_text.insert(tk.END, "邮件发送失败\n")
            self.output_text.insert(tk.END, str(e) + "\n")
            self.sent_email = False

    def check_system(self, receiver):
        try:
            while not self.stop_event.is_set():
                if USER32.GetForegroundWindow() != 0:
                    hostname = socket.gethostname()
                    ip_address = socket.gethostbyname(hostname)
                    response_ip = requests.get("https://www.90th.cn/api/ip")
                    if response_ip.status_code == 200:
                        public_ip = response_ip.json()["ip"]
                        address = response_ip.json()["address"]
                    else:
                        public_ip = "获取失败"
                        address = "获取失败"
                    login_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    subject = '系统状况报告'
                    content = f"""
                    <html>
                    <body>
                    <h3>系统状况报告</h3>
                    <p>主机名: {hostname}</p>
                    <p>外网IP: {public_ip}</p>
                    <p>内网IP: {ip_address}</p>
                    <p>归属地: {address}</p>
                    <p>发送时间: {login_time}</p>
                    <p>登录状态: 成功</p>
                    <p><img src="cid:screenshot"></p>
                    </body>
                    </html>
                    """
                    if not self.sent_email:
                        self.send_email(subject, content, receiver)
                    self.sent_email = True
                    time.sleep(5)
                else:
                    time.sleep(1)
                    self.sent_email = False
                    self.output_text.insert(tk.END, "电脑未唤醒\n")
        except Exception as e:
            self.output_text.insert(tk.END, "Error: 无法检测电脑唤醒状态\n")
            self.output_text.insert(tk.END, str(e) + "\n")
            self.sent_email = False

    def start_thread(self):
        if self.thread is None or not self.thread.is_alive():
            with open("emails.txt", "r") as f:
                receiver = f.read().strip()
            self.thread = threading.Thread(target=self.check_system, args=(receiver,), daemon=True)
            self.thread.start()
            self.output_text.insert(tk.END, "开始检测\n")

    def stop_thread(self):
        if self.thread is not None and self.thread.is_alive():
            self.stop_event.set()
            self.thread.join(timeout=3)
            self.output_text.insert(tk.END, "停止检测\n")

    def run(self):
        if getattr(sys, 'frozen', False):
            os.chdir(sys._MEIPASS)
        self.root.protocol("WM_DELETE_WINDOW", self.stop_check)
        self.root.mainloop()

    def stop_check(self):
        if not self.stop_event.is_set():
            self.stop_event.set()
        if self.thread is not None and self.thread.is_alive():
            self.thread.join(timeout=3)
        self.stop_thread()
        self.root.destroy()
        self.root.quit()
        for t in threading.enumerate():
            if t != threading.current_thread():
                t.join(timeout=3)

if __name__ == '__main__':
    checker = EmailChecker()
    checker.run()