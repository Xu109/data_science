# -*- coding: utf-8 -*-
"""
Created on Sat May 23 16:30:56 2020

@author: Administrator
"""

 
import pandas as pd
import pymysql
#import xlsxwriter
import smtplib  #邮件模块
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import time

while True:
    #配置时间
    ehour=15 #定时小时
    emin=37 #定时分钟
    esec=10 #定时秒
    current_time = time.localtime(time.time())  #当前时间date
    cur_time = time.strftime('%H:%M', time.localtime(time.time()))  #当前时间str
    
    # 邮件正文           
    smtpserver = 'smtp.qq.com'
    smtpport = 465
    username = '1121134481@qq.com'
    password = '*************'
    sender = 'Toby<1121134481@qq.com>'
    receiver = '1121134481@qq.com' 
    subject = '数据库表更新情况'
    
    message = MIMEMultipart()
    message['From'] = sender #发送
    message['To'] = receiver #收件
    message['Subject'] = Header(subject, 'utf-8')
                     
    if ((current_time.tm_hour == ehour) and (current_time.tm_min == emin) and (current_time.tm_sec == esec)):
       print ("开始")
       #执行
       try:
           risk_test = pymysql.connect(host="localhost",user="root",
                       password="root",database="learn",# revise to your own database
                       charset="utf8")
           query="""
                    SELECT
 table_name,
 update_time
FROM
 information_schema.`tables`
WHERE
 table_schema = 'learn'
AND table_name IN (
 'customer_info',
 'credit_loan',
 'repaying_plan_detail')
order by update_time;"""
           data=pd.read_sql(query,risk_test)
           html = """\
<html>
  <head></head>
  <body>
    {0}
  </body>
</html>
""".format(data.to_html())

           part1 = MIMEText(html, 'html')
           message.attach(part1)
    
           smtp = smtplib.SMTP_SSL(host = 'smtp.qq.com')
           smtp.connect(smtpserver, smtpport) #连接服务器
           smtp.login(username, password) #登录
           smtp.sendmail(sender, receiver, message.as_string())  #发送
           smtp.quit()
           print("邮件发送成功")
           
       except:
           print("邮件发送失败")
       print(cur_time)
    time.sleep(1)
