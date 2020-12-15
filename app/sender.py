import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import configparser
import logging
import erp

# config.init
config = configparser.ConfigParser()
config.read('../config.ini')

# logging
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
logging.basicConfig(
    level='DEBUG',
    filename='log\\payslip_sender.log',
    format=LOG_FORMAT
)
logger = logging.getLogger()

# emailList 資料範例：
#      | TI002(支薪期間) | TI004 (部門編號) | TI001(工號)| MV020(電子郵件)
#   ---┼--------------- ┼----- -----------┼-----------┼---------------------   
#     0| 202007         | 15010           | 106005    | XXX@XXXXXX.com
def send_email(emailList):
        log = ''
        # check smtp connection
        try:
            log += '  開始執行'
            print(log)
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

            payPeriod = emailList.iat[0,0]

            email_subject = payPeriod[:4] + '年' + payPeriod[4:] + '月份' + '薪資檔案'
            email_body = config['email']['email_content']

            payslip = erp.Payslip()

            for employee in emailList[emailList.columns[2]]:
                payslip.create_payslip(emailList.iat[0,0], employee)

                email = emailList.loc[emailList.TI001 == employee,'MV020'].values[0]

                # Email header
                sender_email = config['smtp']['user_name']
                receiver_email = email
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = receiver_email
                msg['Subject'] = email_subject

                # Email content
                msg.attach(MIMEText(email_body, 'plain', 'utf-8'))
                fname = (
                    'payslip\\' +
                    employee.strip() +
                    '-' +
                    emailList.iat[0,0].strip() +
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
                    emaildata = '員工' + employee + '無電子郵件資料, 因此檔案傳送失敗'
                    print(emaildata)
                    logger.info(emaildata)
                else:
                    # Check whether email send success or not
                    try:
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


if __name__ == '__main__':
    pass
