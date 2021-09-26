import json
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

with open("cfg.json", "r") as f:
    cfg = json.load(f)


def __form_msg(recipient, text):

    msg = MIMEMultipart()
    msg['From'] = 'stepanPyParser'
    msg['To'] = recipient
    msg['Date'] = formatdate(localtime=True)
    msg.attach(MIMEText(text))

    fp = open("exchange.xlsx", 'rb')
    part = MIMEBase('application', 'vnd.ms-excel')
    part.set_payload(fp.read())
    fp.close()
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="exchange.xlsx"')
    msg.attach(part)
    return msg


def __configureSmtp():
    smtp = smtplib.SMTP('smtp.gmail.com', '587')
    smtp.starttls()
    smtp.login(cfg['login'], cfg['pwd'])
    return smtp


def send_excel(recipient, text):
    smtpObj = __configureSmtp()
    msg = __form_msg(recipient, text)
    smtpObj.sendmail(msg['From'], msg['To'], msg.as_bytes())
    smtpObj.quit()
