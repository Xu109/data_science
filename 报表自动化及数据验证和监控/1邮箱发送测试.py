import pandas as pd
import numpy as np
import pymysql
import sqlalchemy
import smtplib  #邮件模块
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import xlrd
from xlutils.copy import copy
import xlsxwriter
import time

'''
POP3/SMTP服务
yrnkhduydjiygjie
'''
'''
IMAP/SMTP服务
dowomthvktfljddj
'''


'''
1.Python邮件测试
'''
host = 'smtp.qq.com' #  服务器地址
port = 465 # 端口
user = '1121134481@qq.com' # 发件人账号
password = 'yrnkhduydjiygjie' # 发件账号密码（授权码）
sender = '1121134481@qq.com' # 发件人账号
receivers = ['1121134481@qq.com']  # 收件人账号,此处设置为本人 
subject = 'Python邮件测试'  # 邮件标题
# 三个参数：第一个为文本内容，第二个 plain 设置文本格式，第三个 utf-8 设置编码
#message = MIMEText('Python 邮件发送测试', 'plain', 'utf-8')
try:
    message = MIMEText('Python 邮件发送测试', 'plain', 'gbk')
    message['Subject'] = Header(subject, 'gbk')
    message['From'] = 'Toby<1121134481@qq.com>'
    message['To'] = ';'.join(receivers)
    
    smtp_obj = smtplib.SMTP_SSL(host = 'smtp.qq.com') # 开启发信服务，加密传输
    smtp_obj.connect(host, port)
    smtp_obj.login(user, password) # 登录邮箱
    smtp_obj.sendmail(sender, receivers, message.as_string()) #发送邮件
    print ("邮件发送成功")
except smtplib.SMTPException:
    print ("邮件发送失败")

'''
2.Python连接数据库
'''
risk1 =pymysql.connect(host="localhost",user="root",
                      password="root",database="learn",
                      charset="utf8")
query1="""select * from learn.customer_detail"""
data1=pd.read_sql(query1,risk1)

risk2 = sqlalchemy.create_engine('mysql+pymysql://root:root@localhost:3306/?charset=utf8')
query2="""select * from learn.daily_report
           where 日期>'2018-03-31' and 日期<='2018-04-30'"""
data2=pd.read_sql(query2,risk2)

