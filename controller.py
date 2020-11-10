import conn2MSSQL
import models
import csv
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
        self.default_payPeriod_list = self.model_employeeList.get_list_payPeriods()
        self.default_department_list = self.model_employeeList.get_list_departments(self.default_payPeriod_list[0])
        self.default_employee_list = self.model_employeeList.get_list_employees(self.default_payPeriod_list[0], '所有部門')

    def email_payslip(self, payPeriod, department, employee):
        log = '  開始執行'
        # check smtp connection
        try:
            serverSMTP = smtplib.SMTP(config['smtp']['smtp_host'], config['smtp']['smtp_port'])
            serverSMTP.starttls()
            serverSMTP.login(config['smtp']['user_name'], config['smtp']['user_password'])
        except:
            smtp_status = '\nSMTP Server 連線失敗。'
            print(smtp_status)
            logger.info(smtp_status)
            log += smtp_status
        else:
            smtp_status = '\nSMTP Server 連線成功!'
            print(smtp_status)
            logger.info(smtp_status)
            log += smtp_status
            # get list of employees
            list_employees = self.model_payslip.create_list_employees(
                payPeriod,
                department,
                employee
            )
            email_subject = payPeriod + '月份' + '薪資檔案'
            # email content
            email_body = config['email']['email_content']
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
                    emaildata = '員工' + employee + '無電子郵件資料, 因此檔案傳送失敗'
                    print(emaildata)
                    logger.info(emaildata)
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
                        delivery_status = '\n電子郵件傳送失敗。'
                        print(delivery_status)
                        logger.info(delivery_status)
                        log += delivery_status
                    else:
                        delivery_status = '\n電子郵件傳送成功。'
                        print(delivery_status)
                        logger.info(delivery_status)
                        log += delivery_status
        finally:
            serverSMTP.quit()
            log += '\n結束'
            return log

    def create_payslip(self, payPeriod, employee):
        profile = self.model_payslip.get_profile(employee)
        list_profile = profile.values.tolist()
        palti = self.model_payslip.get_palti(payPeriod, employee)
        list_palti = palti.values.tolist()
        paltj = self.model_payslip.get_paltj(payPeriod, employee)
        list_paltj = paltj.values.tolist()

        # 下面這段程式碼會先建立一個 10 * 3 的二維陣列 list1, 10 列中的每一列都會有 3 個元素, 每個元素分別代表[項目名稱, 項目金額, 正負值]
        # 接下來使用一個迴圈將 list_paltj 的值帶入 list1 中, 這個步驟是為了讓 list_paltj 的項目在小於 10 比的狀況下仍能正常顯示
        positive = 0
        negative = 0
        list1 = [['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1], ['', 0, 1]]
        for i in range(len(list_paltj)):
            if list_paltj[i][0] != '':
                list1[i][0] = list_paltj[i][0]
                list1[i][1] = list_paltj[i][1]
                if list_paltj[i][2] < 0:
                    list1[i][2] = -1
                    negative += (list1[i][1] * list_paltj[i][2])
                else:
                    positive += list1[i][1]

        year = payPeriod[0:4]
        month = payPeriod[4:6]

        # add variables into the templates context
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template('payslip.html')

        # 薪資條的加扣項:
        # item1 ~ item7 為固定的薪資項目, 其項目數量和金額的正負值是不變的
        # item8 ~ item27 為其他津貼扣款項目, 其項目的名稱、數量和金額的正負值都是變動的 (不同的員工, 項目的數量也會不同)
        # item28 ~ item35 為出勤明細及相關金額總和
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
