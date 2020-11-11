import conn2MSSQL
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
import pdf
import pandas as pd

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


class Payslip:
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

        # employee list
        self.employeeList = self.create_employeeList()
        self.default_payPeriod_list = self.get_list_payPeriods()
        self.default_department_list = self.get_list_departments(self.default_payPeriod_list[0])
        self.default_employee_list = self.get_list_employees(self.default_payPeriod_list[0], '所有部門')

# employeeList
    def create_employeeList(self):
        query_employeeList = (
            'SELECT DISTINCT TI002, TI004, TI001 ' +
            'FROM PALTI ' +
            'ORDER BY \'TI002\' DESC'
        )
        return pd.read_sql(query_employeeList, self.conn)

    def get_list_payPeriods(self):
        return self.employeeList[self.employeeList.columns[0]].unique().tolist()

    def get_list_departments(self, payPeriod):
        df_departments = self.employeeList.query(
            self.employeeList.columns[0] +
            ' == \'' +
            payPeriod +
            '\''
        )
        list_departments = df_departments[self.employeeList.columns[1]].unique().tolist()
        list_departments.sort()
        return list_departments

    def get_list_employees(self, payPeriod, department):
        if department == '所有部門':
            df_employees = self.employeeList.query(
                self.employeeList.columns[0] +
                ' == \'' +
                payPeriod +
                '\''
            )
        else:
            df_employees = self.employeeList.query(
                self.employeeList.columns[0] +
                ' == \'' +
                payPeriod +
                '\' and ' +
                self.employeeList.columns[1] +
                ' == \'' +
                department +
                '\''
            )
        list_employees = df_employees[self.employeeList.columns[2]].unique().tolist()
        list_employees.sort()
        return list_employees

# Payslip
    def get_profile(self, employee):
        query_profile = (
            'SELECT MV001, MV002, MV004, MV008, MV006, MV020, MV009 ' +
            'FROM CMSMV ' +
            'WHERE MV001 = \'' +
            employee +
            '\' ' +
            'ORDER BY \'MV001\' DESC'
            )
        return pd.read_sql(query_profile, self.conn)

    def get_palti(self, payPeriod, employee):
        query_palti = (
                'SELECT TI002 as \'發薪年月\', ' +
                'TI004 AS \'部門\', ' +
                'TI001 AS \'員工編號\', ' +
                'TI023 + TI024 AS \'底薪\', ' +
                'TI029 + TI030 AS \'全勤獎金\', ' +
                'TI025 + TI026 AS \'加班費免稅\', ' +
                'TI027 + TI028 AS \'加班費課稅\', ' +
                'TI031 + TI032 AS \'請假扣款\', ' +
                'TI033 AS \'健保費\', ' +
                'TI034 AS \'勞保費\', ' +
                'TI040 + TI041 + TI033 + TI034 AS \'應發金額\', ' +
                'TI040 + TI041 AS \'實發金額\', ' +
                'TI042 + TI043 AS \'課稅金額\', ' +
                'TI027 + TI028 + TI044 + TI045 AS \'津貼合計\', ' +
                'TI054 AS \'公司提繳\', ' +
                'TI015 + TI016 AS \'出勤天數\', ' +
                'TI007 + TI008 + TI009 + TI010 + TI011 + TI012 + TI062 + TI063 + TI064 + TI065 + TI066 + TI067 AS \'加班時數\', ' +
                'TI033 + TI034 AS \'扣款合計\', ' +
                'TI035 + TI036 AS \'所得稅\' ' +
                'FROM PALTI ' +
                'where TI001 = ' +
                '\'' +
                employee +
                '\' ' +
                'and ' +
                'TI002 = ' +
                payPeriod +
                ' ORDER BY \'員工編號\' ASC'
            )
        return pd.read_sql(query_palti, self.conn)

    def get_paltj(self, payPeriod, employee):
        query_paltj = (
            'SELECT DISTINCT TJ004 AS \'津貼扣款名稱\', ' +
            'TJ007 AS \'津貼扣款金額\', ' +
            'TJ005 AS \'加扣項\' ' +
            'FROM PALTJ ' +
            'WHERE TJ001 = ' +
            '\'' +
            employee +
            '\' ' +
            'AND TJ002 = ' +
            payPeriod
        )
        return pd.read_sql(query_paltj, self.conn)

    def get_email(self, employee):
        query_email = (
            'SELECT MV020 ' +
            'FROM CMSMV ' +
            'where MV001 = ' +
            '\'' +
            employee +
            '\' ' +
            'ORDER BY MV001 ASC'
        )
        return pd.read_sql(query_email, self.conn)

    def create_emailList(self, payPeriod, department, employee):
        if employee == '所有員工':
            if department == '所有部門': 
                query_employees = (
                    'SELECT DISTINCT TI001 ' +
                    'FROM PALTI ' +
                    'WHERE TI002 = \'' +
                    payPeriod +
                    '\' ' +
                    'ORDER BY \'TI001\' ASC'
                )
            else:
                query_employees = (
                    'SELECT DISTINCT TI001 ' +
                    'FROM PALTI ' +
                    'WHERE TI002 = \'' +
                    payPeriod +
                    '\' and TI004 = \'' +
                    department +
                    '\' ' +
                    'ORDER BY \'TI001\' ASC'
                )
        else:
            query_employees = (
                'SELECT DISTINCT TI001 ' +
                'FROM PALTI ' +
                'WHERE TI001 = \'' +
                employee +
                '\' ' +
                'ORDER BY \'TI001\' ASC'
            )
        df_employees = pd.read_sql(query_employees, self.conn)
        return df_employees['TI001'].tolist()

    def create_payslip(self, payPeriod, employee):
        profile = self.get_profile(employee)
        list_profile = profile.values.tolist()
        palti = self.get_palti(payPeriod, employee)
        list_palti = palti.values.tolist()
        paltj = self.get_paltj(payPeriod, employee)
        list_paltj = paltj.values.tolist()

        # 加扣項個別加總
        negative = paltj[paltj["加扣項"]<0]["津貼扣款金額"].sum()
        positive = paltj[paltj["加扣項"]>0]["津貼扣款金額"].sum()

        year = payPeriod[:4]
        month = payPeriod[4:]

        # add variables into the templates context
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template('payslip.html')

        # 薪資條的加扣項:
        # item1 ~ item7 為固定的薪資項目, 其項目數量和金額的正負值是不變的
        # item8 加扣款項目, 項目數量是變動的
        # item9 ~ item16 為出勤明細及相關金額總和
        year = int(year)
        month = int(month) + 1
        template_vars = {
            'payDate': f"{year}/{month}/01",
            'payPeriod': f"{year}/{month}/01-{year}/{month}/30",
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
            'item8': paltj.T.to_dict(),
            'item9': str(list_palti[0][15]),
            'item10': str(list_palti[0][16]),
            'item11': str(int(list_palti[0][14])),
            'item12': str(int(positive) + int(list_palti[0][4]) + int(list_palti[0][5]) + int(list_palti[0][6])),
            'item13': str(int(negative) - int(list_palti[0][7]) - int(list_palti[0][8]) - int(list_palti[0][9]) - int(list_palti[0][18])),
            'item14': str(int(list_palti[0][12])),
            'item15': str(int(list_palti[0][10]) + int(list_palti[0][18])),
            'item16': str(int(list_palti[0][11]))
        }

        # Render the template into HTML
        html_out = template.render(template_vars)

        # create file name
        fname = 'payslip\\' + str(employee).strip() + '-' + payPeriod + '-薪資檔案' + '.pdf'

        # rendering html to pdf
        HTML(string=html_out).write_pdf(fname, stylesheets=['style.css'])

        # encryption
        password = str(list_profile[0][6]).strip()
        pdf.encrypt(fname, fname, password)

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
            list_employees = self.create_emailList(
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
                email = self.get_email(employee)
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
