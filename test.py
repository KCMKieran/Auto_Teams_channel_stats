# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.base import MIMEBase
# from email import encoders
# from email.header import Header, make_header
# from typing import List
# import os
# from dotenv import load_dotenv
# import time  # 添加这行导入
# import logging  # 添加日志支持

# load_dotenv(override=True)  # 添加 override=True 参数

# def get_env_or_raise(key: str) -> str:
#     value = os.getenv(key)
#     if value is None:
#         raise ValueError(f"环境变量 {key} 未设置，请检查 .env 文件！")
#     return value

# SMTP_SERVER = get_env_or_raise('SMTP_SERVER')
# SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))
# USERNAME_MAIL = get_env_or_raise('USERNAME_MAIL')
# PASSWORD_MAIL = get_env_or_raise('PASSWORD_MAIL')
# MAIL_SEND_TOO = os.getenv('MAIL_SEND_TOO')
# MAIL_CCC = os.getenv('MAIL_CCC')

# print(SMTP_SERVER)
# print(SMTP_PORT)
# print(USERNAME_MAIL)
# print(PASSWORD_MAIL)
# print(MAIL_SEND_TOO)
# print(MAIL_CCC)

import datetime
import pytz

# 检查当前系统时区
print(f"系统当前时间: {datetime.datetime.now()}")
print(f"UTC当前时间: {datetime.datetime.utcnow()}")

# 设置香港时区
hk_tz = pytz.timezone('Asia/Hong_Kong')
now_hk = datetime.datetime.now(hk_tz)
print(f"香港当前时间: {now_hk}")

# 计算上周的时间范围（香港时间）
last_monday_hk = now_hk - datetime.timedelta(days=now_hk.weekday() + 7)
last_sunday_hk = last_monday_hk + datetime.timedelta(days=6)

# 设置准确的开始和结束时间
start_time_hk = last_monday_hk.replace(hour=0, minute=0, second=0, microsecond=0)
end_time_hk = last_sunday_hk.replace(hour=23, minute=59, second=59, microsecond=999999)

print(f"香港时间范围: {start_time_hk} 到 {end_time_hk}")

# 转换为UTC时间（用于API查询）
start_time_utc = start_time_hk.astimezone(pytz.UTC).replace(tzinfo=None)
end_time_utc = end_time_hk.astimezone(pytz.UTC).replace(tzinfo=None)

print(f"UTC时间范围: {start_time_utc} 到 {end_time_utc}")