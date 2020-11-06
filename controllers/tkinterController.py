import conn2MSSQL
import models
from view import PayslipSenderView
import csv
import tkinter as tk
from tkinter import messagebox
import configparser
import logging
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PyPDF2 import PdfFileWriter, PdfFileReader

# config.init
config = configparser.ConfigParser()
config.read('config.ini')

# logging
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
logging.basicConfig(
    level='DEBUG',
    filename='log\\payslip_sender.log',
    format=LOG_FORMAT
)
logger = logging.getLogger()


class Controller:
    def __init__(self):
        # create view instance
        self.view = PayslipSenderView()

        # create database connection
        self.connection = conn2MSSQL.Connection(
            config['database']['server_name'],
            config['database']['database'],
            config['database']['user_name'],
            config['database']['user_password']
        )
        self.conn = self.connection.conn
        logger.info(self.connection.connectionInfo)

        # create model instance
        self.model_employeeList = models.EmployeeListModel(self.conn)
        self.model_emailList = models.EmailListModel(self.conn)
        self.model_payslip = models.PayslipModel(self.conn)

        # set default value
        self.add_item_to_combobox()
        self.set_value_of_entry()
        self.insert_text_to_textBox()
        self.insert_data_into_treeview()

        # create event handler
        self.bind_function_to_combobox()
        self.bind_function_to_btn_send()
        self.bind_function_to_btn_smtp()
        self.bind_function_to_btn_email_content()

        # run Tkinter application
        self.view.root.mainloop()

# cascading combobox
    def add_item_to_combobox(self):
        list_payPeriods = self.model_employeeList.get_list_payPeriods()
        self.view.cmb_payPeriod['values'] = list_payPeriods
        self.view.cmb_payPeriod.set(list_payPeriods[0])

        list_departments = self.model_employeeList.get_list_departments(list_payPeriods[0])
        self.view.cmb_department['values'] = list_departments
        self.view.cmb_department.set('所有部門')

        list_employees = self.model_employeeList.get_list_employees(list_payPeriods[0], '所有部門')
        self.view.cmb_employee['values'] = list_employees
        self.view.cmb_employee.set('所有員工')

    def bind_function_to_combobox(self):
        self.view.cmb_payPeriod.bind(
            "<<ComboboxSelected>>",
            self.cmb_payPeriod_selected
        )
        self.view.cmb_department.bind(
            "<<ComboboxSelected>>",
            self.cmb_department_selected
        )
        self.view.cmb_employee.bind(
            "<<ComboboxSelected>>",
            self.cmb_employee_selected
        )

    def cmb_payPeriod_selected(self, event):
        payPeriod = self.view.payPeriod.get()
        self.list_department = self.model_employeeList.get_list_departments(payPeriod)
        self.view.cmb_department['values'] = self.list_department
        self.view.cmb_department.set('所有部門')
        department = self.view.department.get()
        self.list_employees = self.model_employeeList.get_list_employees(
            payPeriod,
            department
        )
        self.view.cmb_employee['values'] = self.list_employees
        self.view.employee.set('所有員工')

    def cmb_department_selected(self, event):
        payPeriod = self.view.payPeriod.get()
        department = self.view.department.get()
        self.view.cmb_employee['values'] = self.model_employeeList.get_list_employees(
            payPeriod,
            department
        )
        self.view.cmb_employee.set('所有員工')

    def cmb_employee_selected(self, event):
        payPeriod = self.view.payPeriod.get()
        department = self.view.department.get()
        self.view.cmb_employee['values'] = self.model_employeeList.get_list_employees(
            payPeriod,
            department
        )

# send button
    def bind_function_to_btn_send(self):
        self.view.btn_send.configure(command=self.btn_send_clicked)

    def btn_send_clicked(self):
        self.view.txt_logging.config(state='normal')
        self.view.txt_logging.insert(tk.INSERT, '\n------------------------------ 開始 ------------------------------')
        self.email_payslip()

    def email_payslip(self):
        # check smtp connection
        try:
            serverSMTP = smtplib.SMTP(config['smtp']['smtp_host'], config['smtp']['smtp_port'])
            serverSMTP.starttls()
            serverSMTP.login(config['smtp']['user_name'], config['smtp']['user_password'])
        except:
            connection_failed = 'SMTP Server 連線失敗。'
            print(connection_failed)
            logger.info(connection_failed)
            self.view.txt_logging.insert(tk.INSERT, ('\n' + connection_failed))
        else:
            connection_success = 'SMTP Server 連線成功!'
            print(connection_success)
            logger.info(connection_success)
            self.view.txt_logging.insert(tk.INSERT, '\n' + connection_success)

            # get list of employees
            payPeriod = self.view.payPeriod.get()
            department = self.view.department.get()
            employee = self.view.employee.get()
            list_employees = self.model_payslip.create_list_employees(
                payPeriod,
                department,
                employee
            )

            email_subject = payPeriod + '月份' + '薪資檔案'
            email_body = self.view.txt_email_content.get('1.0', 'end-1c')

            for employee in list_employees:
                self.create_payslip(payPeriod, employee)
                employee = employee.rstrip()
                email = self.model_payslip.get_email(employee)
                str_email = email.to_string(index=False, header=False)
                # Email header
                sender_email = config['smtp']['user_name']
                receiver_email = str_email
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = receiver_email
                msg['Subject'] = email_subject
                # Email content
                msg.attach(MIMEText(email_body, 'plain', 'utf-8'))
                fname = (
                    'payslip\\' +
                    employee +
                    '-' +
                    payPeriod +
                    '-' +
                    '薪資檔案.pdf'
                )
                attachment = open(fname, 'rb')

                part = MIMEBase('application', 'octet-stream')
                part.set_payload((attachment).read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=('gbk', '', fname)
                )
                msg.attach(part)
                text = msg.as_string()
                # Check whether the email exists
                if receiver_email.rstrip() == '':
                    msg = '員工' + employee + '無電子郵件資料, 請確實填寫'
                    messagebox.showwarning('小提醒', message=msg)
                    noData = '員工' + employee + '無電子郵件資料, 因此檔案傳送失敗'
                    print(noData)
                    logger.info(noData)
                    self.view.txt_logging.insert(tk.INSERT, '\n' + noData)
                else:
                    # Check whether email send success or not
                    try:
                        # send email
                        serverSMTP.sendmail(
                            sender_email,
                            receiver_email,
                            text
                        )
                    except:
                        delivery_failed = '電子郵件傳送失敗。'
                        print(delivery_failed)
                        logger.info(delivery_failed)
                        self.view.txt_logging.insert(tk.INSERT, delivery_failed)
                    else:
                        deliverey_success = '電子郵件傳送成功。'
                        print(deliverey_success)
                        logger.info(deliverey_success)
                        self.view.txt_logging.insert(tk.INSERT, deliverey_success)
        finally:
            serverSMTP.quit()
            self.view.txt_logging.insert(tk.INSERT, '\n------------------------------ 結束 ------------------------------')
            self.view.txt_logging.config(state='disable')

    def create_payslip(self, payPeriod, employee):
        profile = self.model_payslip.get_profile(employee)
        list_profile = profile.values.tolist()
        palti = self.model_payslip.get_palti(payPeriod, employee)
        list_palti = palti.values.tolist()
        paltj = self.model_payslip.get_paltj(payPeriod, employee)
        list_paltj = paltj.values.tolist()

        # 判斷津貼扣款項目正負值
        positive = 0
        negative = 0
        list1 = [['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1]]
        for i in range(len(list_paltj)):
            if list_paltj[i][0] == '':
                pass
            else:
                list1[i][0] = list_paltj[i][0]
                list1[i][1] = list_paltj[i][1]
                if list_paltj[i][2] < 0:
                    list1[i][2] = -1
                    negative = negative + (list1[i][1] * list_paltj[i][2])
                else:
                    positive = positive + list1[i][1]
        year = payPeriod[0:4]
        month = payPeriod[4:6]

        # add variables into the templates context
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template('payslip.html')
        template_vars = {
            'payDate': (year + '/' + str(int(month) + 1) + '/' + '01'),
            'payPeriod': (year + '/' + month + '/01-' + year + '/' + month + '/30'),
            'email': str(list_profile[0][5]),
            'employeeID': str(list_profile[0][0]),
            'name': str(list_profile[0][1]),
            'department': str(list_profile[0][2]),
            'dateOfBirth': str(list_profile[0][3]),
            'item1': str(int(list_palti[0][3])),
            'item2': str(int(list_palti[0][4])),
            'item3': str(int(list_palti[0][5])+int(list_palti[0][6])),
            'item4': str(int(list_palti[0][7]) * -1),
            'item5': str(int(list_palti[0][8]) * -1),
            'item6': str(int(list_palti[0][9]) * -1),
            'item7': str(int(list_palti[0][18]) * -1),
            'item8': str(list1[0][0]),
            'item9': self.determine_zero(list1[0][0], list1[0][1], list1[0][2]),
            'item10': str(list1[1][0]),
            'item11': self.determine_zero(list1[1][0], list1[1][1], list1[1][2]),
            'item12': str(list1[2][0]),
            'item13': self.determine_zero(list1[2][0], list1[2][1], list1[2][2]),
            'item14': str(list1[3][0]),
            'item15': self.determine_zero(list1[3][0], list1[3][1], list1[3][2]),
            'item16': str(list1[4][0]),
            'item17': self.determine_zero(list1[4][0], list1[4][1], list1[4][2]),
            'item18': str(list1[5][0]),
            'item19': self.determine_zero(list1[5][0], list1[5][1], list1[5][2]),
            'item20': str(list1[6][0]),
            'item21': self.determine_zero(list1[6][0], list1[6][1], list1[6][2]),
            'item22': str(list1[7][0]),
            'item23': self.determine_zero(list1[7][0], list1[7][1], list1[7][2]),
            'item24': str(list1[8][0]),
            'item25': self.determine_zero(list1[8][0], list1[8][1], list1[8][2]),
            'item26': str(list1[9][0]),
            'item27': self.determine_zero(list1[9][0], list1[9][1], list1[9][2]),
            'item28': str(list_palti[0][15]),
            'item29': str(list_palti[0][16]),
            'item30': str(int(list_palti[0][14])),
            'item31': str(int(positive) + int(list_palti[0][4]) + int(list_palti[0][5]) + int(list_palti[0][6])),
            'item32': str(int(negative) - int(list_palti[0][7]) - int(list_palti[0][8]) - int(list_palti[0][9]) - int(list_palti[0][18])),
            'item33': str(int(list_palti[0][12])),
            'item34': str(int(list_palti[0][10]) + int(list_palti[0][18])),
            'item35': str(int(list_palti[0][11]))
        }

        # Render the template into HTML
        html_out = template.render(template_vars)

        # create file name
        path = 'payslip\\'
        fname = path + str(employee).strip() + '-' + payPeriod + '-薪資檔案' + '.pdf'

        # rendering html to pdf
        HTML(string=html_out).write_pdf(fname, stylesheets=['style.css'])

        # encryption
        password = str(list_profile[0][6]).strip()
        self.encrypt(fname, fname, password)

    def determine_zero(self, value1, value2, value3):
        value = value2
        if value1 == '' and value2 == 0:
            value = ''
            return value
        else:
            return str(int(value * value3))

    def encrypt(self, input_pdf, output_pdf, password):
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(input_pdf)

        for page in range(pdf_reader.getNumPages()):
            pdf_writer.addPage(pdf_reader.getPage(page))
            pdf_writer.encrypt(user_pwd=password, owner_pwd=None, use_128bit=True)

        with open(output_pdf, 'wb') as fh:
            pdf_writer.write(fh)

# smtp server
    def set_value_of_entry(self):
        self.view.value_smtpserver.set(config['smtp']['smtp_host'])
        self.view.value_port.set(config['smtp']['smtp_port'])
        self.view.value_id.set(config['smtp']['user_name'])
        self.view.value_pwd.set(config['smtp']['user_password'])
        self.view.entry_smtpserver.config(state='disabled')
        self.view.entry_port.config(state='disabled')
        self.view.entry_id.config(state='disabled')
        self.view.entry_pwd.config(state='disabled')

    def bind_function_to_btn_smtp(self):
        self.view.btn_set_smtp.configure(command=self.set_smtp)
        self.view.btn_save_smtp.configure(command=self.save_smtp)

    def set_smtp(self):
        self.view.entry_smtpserver.config(state='normal')
        self.view.entry_port.config(state='normal')
        self.view.entry_id.config(state='normal')
        self.view.entry_pwd.config(state='normal')

    def save_smtp(self):
        config['smtp']['smtp_host'] = self.view.entry_smtpserver.get()
        config['smtp']['smtp_port'] = self.view.entry_port.get()
        config['smtp']['user_name'] = self.view.entry_id.get()
        config['smtp']['user_password'] = self.view.entry_pwd.get()
        self.view.entry_smtpserver.config(state='disabled')
        self.view.entry_port.config(state='disabled')
        self.view.entry_id.config(state='disabled')
        self.view.entry_pwd.config(state='disabled')
        with open('config.ini', 'w') as configfile:
            config.write(configfile)

# email list
    def insert_data_into_treeview(self):
        list_emailList = self.model_emailList.emailList
        for i in range(len(list_emailList)):
            self.view.tv.insert(
                '',
                0,
                text=list_emailList[i][0],
                values=(list_emailList[i][1], list_emailList[i][2])
            )

# email content
    def insert_text_to_textBox(self):
        with open('email_content.csv', 'rt')as f:
            data = csv.reader(f)
            alist = list(data)
            self.view.txt_email_content.insert(tk.END, alist[0][0])
        self.view.txt_email_content.config(state='disable')

    def bind_function_to_btn_email_content(self):
        self.view.btn_edit_email_content.configure(command=self.edit_email_content)
        self.view.btn_save_email_content.configure(command=self.save_email_content)

    def edit_email_content(self):
        self.view.txt_email_content.config(state='normal', bg='#FFFFFF')

    def save_email_content(self):
        self.view.txt_email_content.config(state='disabled', bg='#E4E4E4')
        with open('email_content.csv', 'w') as smtp:
            writer = csv.writer(smtp)
            writer.writerow([
                self.view.txt_email_content.get(1.0, tk.END)
            ])
