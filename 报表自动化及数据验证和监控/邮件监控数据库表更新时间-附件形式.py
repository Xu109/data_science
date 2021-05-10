# -*- coding: utf-8 -*-
"""
Created on Fri Apr 24 09:57:45 2020

@author: Administrator
"""

import pandas as pd
import pymysql
import smtplib  #邮件模块
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import time
#%%
while True:
    #配置时间
    ehour=15 #定时小时
    emin=19 #定时分钟
    esec=5 #定时秒
    current_time = time.localtime(time.time())  #当前时间date
    cur_time = time.strftime('%H:%M', time.localtime(time.time()))  #当前时间str
    
    smtpserver = 'smtp.qq.com'
    smtpport = 465
    username = '1121134481@qq.com'
    password = '**********'
    sender = 'Toby<1121134481@qq.com>'
    receiver = '1121134481@qq.com' 
    subject = '数据库表更新情况'
    
    message = MIMEMultipart() 
    message['From'] = sender #发送
    message['To'] = receiver #收件
    message['Subject'] = Header(subject, 'utf-8')
    message.attach(MIMEText("""Dear All:\n  \
    附件是截止今天的数据库表更新情况，请查阅！\n  \
    """, 'plain', 'utf-8'))# 邮件正文
                             
                            
    if ((current_time.tm_hour == ehour) and (current_time.tm_min == emin) and (current_time.tm_sec == esec)):
       print ("开始")
       #执行
       try:
           risk_test = pymysql.connect(host="localhost",user="root",
                       password="root",database="learn",
                       charset="utf8")
           query="""SELECT
	                       table_name,
	                       update_time
                     FROM
	                      information_schema.`tables`
                    WHERE
	                      table_schema = 'learn'
                     AND table_name IN (
	                                     'customer_info',
	                                     'credit_loan',
	                                     'repaying_plan_detail'
                                       )
                     order by update_time;"""
           data=pd.read_sql(query,risk_test)
           workbook = pd.ExcelWriter('update_time.xlsx') #新建一个excel文本
           data.to_excel(workbook,'update_time',index=False)
           workbook.save()
           workbook.close()  # 关闭报表 
           print("数据更新成功")
           
           # 构造附件
           att1 = MIMEText(open('update_time.xlsx','rb').read(), 'base64', 'utf-8')
           att1["Content-Type"] = 'application/octet-stream'
           att1["Content-Disposition"] = "attachment;filename=update_time.xls"
           message.attach(att1) 
           
           smtp = smtplib.SMTP_SSL(host = 'smtp.qq.com')
           smtp.connect(smtpserver, smtpport) #连接服务器
           smtp.login(username, password) #登录
           smtp.sendmail(sender, receiver, message.as_string())  #发送
           smtp.quit()
           print("邮件发送成功")
       except:
           print("数据更新失败")
       print(cur_time)
    time.sleep(1)
