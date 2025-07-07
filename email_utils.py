import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header, make_header
from typing import List
import os
from dotenv import load_dotenv
import time  # 添加这行导入
import logging  # 添加日志支持

load_dotenv(override=True)  # 添加 override=True 参数

def get_env_or_raise(key: str) -> str:
    value = os.getenv(key)
    if value is None:
        raise ValueError(f"环境变量 {key} 未设置，请检查 .env 文件！")
    return value

SMTP_SERVER = get_env_or_raise('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))
USERNAME_MAIL = get_env_or_raise('USERNAME_MAIL')
PASSWORD_MAIL = get_env_or_raise('PASSWORD_MAIL')
MAIL_SEND_TOO = os.getenv('MAIL_SEND_TOO')
MAIL_CCC = os.getenv('MAIL_CCC')


def send_email_with_attachments(subject: str, body: str, attachments: List[str]) -> None:
    """
    发送带附件的邮件，参数从环境变量读取
    :param subject: 邮件主题
    :param body: 邮件正文（支持HTML）
    :param attachments: 附件文件路径列表
    """
    msg = MIMEMultipart()
    msg['From'] = USERNAME_MAIL
    msg['To'] = MAIL_SEND_TOO
    msg['Cc'] = MAIL_CCC
    msg['Subject'] = subject

    # 邮件正文
    msg.attach(MIMEText(body, 'html'))

    # 附件
    for attachment in attachments:
        with open(attachment, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            filename = make_header([(attachment.split('/')[-1], 'utf-8')]).encode()
            part.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(part)
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(USERNAME_MAIL, PASSWORD_MAIL)
        server.sendmail(msg['From'], msg['To'].split(',') + msg['Cc'].split(','), msg.as_string())
        server.quit()
        print("邮件发送成功")
    except Exception as e:
        print(f"邮件发送失败: {e}") 