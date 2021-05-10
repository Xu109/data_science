# -*- coding: utf-8 -*-
"""
Created on Tue Aug 27 17:11:49 2019

@author: Administrator
"""
import smtplib  #邮件模块
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import time

while True:
    
#配置时间
    ehour=14 #定时小时
    emin=18  #定时分钟
    esec=10  #定时秒
    current_time = time.localtime(time.time())  #当前时间date
    cur_time = time.strftime('%H:%M', time.localtime(time.time()))  #当前时间str

    smtpserver = 'smtp.qq.com'
    smtpport = 465
    username = '1121134481@qq.com'
    password = '**********'
    sender = 'Toby<1121134481@qq.com>'
    receiver = '1121134481@qq.com' 
    subject = '贷后日报'
    
    message = MIMEMultipart()
    message['From'] = sender #发送
    message['To'] = receiver #收件
    message['Subject'] = Header(subject, 'utf-8')
    message.attach(MIMEText("""Dear All:\n  \
    附件是截止今天的贷后日报，请查阅！\n  \
有任何疑问请随时与我联系，谢谢！ \
    """, 'plain', 'utf-8'))# 邮件正文
    # 构造附件
    att1 = MIMEText(open('vintage_report.xlsx','rb').read(), 'base64', 'utf-8')
    att1["Content-Type"] = 'application/octet-stream'
    att1["Content-Disposition"] = "attachment;filename=vintage_report.xlsx"
    message.attach(att1)

    #操作
    if ((current_time.tm_hour == ehour) and (current_time.tm_min == emin) and (current_time.tm_sec == esec)):
        print ("开始")
        #执行
        try:
            smtp = smtplib.SMTP_SSL(host = 'smtp.qq.com')
            smtp.connect(smtpserver, smtpport) #连接服务器
            smtp.login(username, password) #登录
            smtp.sendmail(sender, receiver, message.as_string())  #发送
            smtp.quit()
            print("发送成功")
        except:
            print("发送失败")
        print(cur_time)
    time.sleep(1)
