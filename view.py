import tkinter as tk
from tkinter import ttk
import csv
import configparser

# config.init
config = configparser.ConfigParser()
config.read('config.ini')


class PayslipSenderView:
    def __init__(self):
        self.root = tk.Tk()
        self.root.iconbitmap('statics/logo.ico')
        self.root.title('薪 資 傳 送 者')
        width_value = self.root.winfo_screenwidth()
        height_value = self.root.winfo_screenheight()
        self.root.geometry('%dx%d+0+0' % (width_value/1.8, height_value/1.8))
        self.root.configure(bg='#ffffff')
        tab_control = ttk.Notebook(self.root)
        tab_control.pack(expand=1, fill="both")

    # tab1
        tab1 = ttk.Frame(tab_control)
        tab_control.add(tab1, text='郵寄薪資')

    # tab2
        tab2 = ttk.Frame(tab_control)
        tab_control.add(tab2, text='郵件設定')

    # frame
        fr1 = tk.Frame(tab1)
        fr1.grid(row=0, column=0)
        fr2 = tk.Frame(tab1)
        fr2.grid(row=1, column=0)
        fr3 = tk.Frame(tab2)
        fr3.grid(row=0, column=0)
        fr4 = tk.Frame(tab2)
        fr4.grid(row=0, column=1)

    # frame1
        lbl_payPeriod = tk.Label(fr1, text='發薪年月', font=("Arial", 14))
        lbl_payPeriod.grid(row=1, column=0)
        lbl_department = tk.Label(fr1, text='部門別', font=("Arial", 14))
        lbl_department.grid(row=1, column=1, padx=10)
        lbl_employee = tk.Label(fr1, text='員工', font=("Arial", 14))
        lbl_employee.grid(row=1, column=2, padx=10)

        self.payPeriod = tk.StringVar()
        self.cmb_payPeriod = ttk.Combobox(
            fr1,
            justify='center',
            width=10,
            textvariable=self.payPeriod
        )
        self.cmb_payPeriod.grid(row=2, column=0, padx=10)

        self.department = tk.StringVar()
        self.cmb_department = ttk.Combobox(
            fr1,
            justify='center',
            width=10,
            textvariable=self.department
        )
        self.cmb_department.grid(
            row=2,
            column=1,
            padx=10
        )

        self.employee = tk.StringVar()
        self.cmb_employee = ttk.Combobox(
            fr1,
            justify='center',
            width=10,
            textvariable=self.employee
        )
        self.cmb_employee.grid(
            row=2,
            column=2,
            padx=10
        )

        self.btn_send = tk.Button(
            fr1,
            padx=5,
            text='郵寄薪資',
            width=10,
            font=("Arial", 12),
        )
        self.btn_send.grid(row=2, column=3, padx=10)

    # frame2
        sb = tk.Scrollbar(fr2, width=25)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_logging = tk.Text(
            fr2,
            width=100,
            height=25,
            yscrollcommand=sb.set
        )
        self.txt_logging.pack(side=tk.LEFT, fill=tk.BOTH, padx=5, pady=5)
        sb.config(command=self.txt_logging.yview)
        self.txt_logging.config(state='disabled')

    # frame3
        lf1 = ttk.LabelFrame(
            fr3,
            text=' SMTP ',
            width=width_value/4.1,
            height=height_value/4
        )
        lf1.grid(row=0, column=0, padx=1, pady=1)
        lb_smtpserver = tk.Label(
            lf1,
            text='server',
            font=("Arial", 12)
        )
        lb_smtpserver.place(x=10, y=20, anchor="w")
        lb_port = tk.Label(
            lf1,
            text='port',
            font=("Arial", 12)
        )
        lb_port.place(x=10, y=50, anchor="w")
        lb_id = tk.Label(
            lf1,
            text='id',
            font=("Arial", 12)
        )
        lb_id.place(x=10, y=80, anchor="w")
        lb_pwd = tk.Label(
            lf1,
            text='pwd',
            font=("Arial", 12)
        )
        lb_pwd.place(x=10, y=110, anchor="w")
        self.value_smtpserver = tk.StringVar()
        self.entry_smtpserver = tk.Entry(
            lf1,
            font=("Arial", 12),
            width=25,
            state='normal',
            textvariable=self.value_smtpserver
        )
        self.entry_smtpserver.place(x=90, y=20, anchor="w")
        self.value_port = tk.StringVar()
        self.entry_port = tk.Entry(
            lf1,
            font=("Arial", 12),
            width=25,
            state='normal',
            textvariable=self.value_port
        )
        self.entry_port.place(x=90, y=50, anchor="w")
        self.value_id = tk.StringVar()
        self.entry_id = tk.Entry(
            lf1,
            font=("Arial", 12),
            width=25,
            state='normal',
            textvariable=self.value_id
        )
        self.entry_id.place(x=90, y=80, anchor="w")
        self.value_pwd = tk.StringVar()
        self.entry_pwd = tk.Entry(
            lf1,
            font=("Arial", 12),
            width=25,
            state='normal',
            textvariable=self.value_pwd,
            show='*'
        )
        self.entry_pwd.place(x=90, y=110, anchor="w")
        self.btn_set_smtp = tk.Button(
            lf1,
            text='修改',
            font=("Arial", 12),
        )
        self.btn_set_smtp.place(x=150, y=150, anchor="w")
        self.btn_save_smtp = tk.Button(
            lf1,
            text='儲存',
            font=("Arial", 12),
        )
        self.btn_save_smtp.place(x=210, y=150, anchor="w")
        lf2 = ttk.LabelFrame(
            fr3,
            text=' 郵件內容文字 ',
            width=width_value/4.1,
            height=height_value/4
        )
        lf2.grid(row=1, column=0, padx=1, pady=1)
        self.txt_email_content = tk.Text(
            lf2,
            font=("Arial", 14),
            width=28,
            height=6,
            background='#E4E4E4'
        )
        self.txt_email_content.place(x=10, y=70, anchor="w")
        self.btn_edit_email_content = tk.Button(
            lf2,
            text='修改',
            font=("Arial", 10),
        )
        self.btn_edit_email_content.place(x=150, y=155, anchor="w")
        self.btn_save_email_content = tk.Button(
            lf2,
            text='儲存',
            font=("Arial", 10),
        )
        self.btn_save_email_content.place(x=210, y=155, anchor="w")

    # frame4
        lf3 = ttk.LabelFrame(
            fr4,
            text=' 郵件清單 ',
            width=width_value/4.1,
            height=height_value/2
        )
        lf3.grid(row=0, column=0, padx=1, pady=1)
        sb2 = tk.Scrollbar(lf3, width=20)
        sb2.pack(side=tk.RIGHT, fill=tk.Y)
        self.tv = ttk.Treeview(lf3, height=17, columns=(3, 10), yscrollcommand=sb2.set)
        self.tv.pack()
        self.tv.heading('#0', text='員工編號', anchor='w')
        self.tv.column('#0', minwidth=0, width=100)
        self.tv.heading('#1', text='員工姓名', anchor='w')
        self.tv.column('#1', minwidth=0, width=100)
        self.tv.heading('#2', text='電子郵件', anchor='w')
        self.tv.column('#2', minwidth=0, width=180)
        self.tv.pack()
        sb2.config(command=self.tv.yview)
