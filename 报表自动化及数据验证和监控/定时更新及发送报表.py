# -*- coding: utf-8 -*-
"""
Created on Wed Apr 22 18:19:31 2020

@author: Administrator
"""
 
import pandas as pd
import pymysql
import xlsxwriter
import smtplib  #邮件模块
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import time
#%%
while True:
    #配置时间
    ehour=14 #定时小时
    emin=29 #定时分钟
    esec=5 #定时秒
    current_time = time.localtime(time.time())  #当前时间date
    cur_time = time.strftime('%H:%M', time.localtime(time.time()))  #当前时间str
    
    smtpserver = 'smtp.qq.com'
    smtpport = 465
    username = '1121134481@qq.com'
    password = '**********'
    sender = 'Toby<1121134481@qq.com>'
    receiver = '1121134481@qq.com' 
    subject = 'vintage报表'
    
    message = MIMEMultipart()
    message['From'] = sender #发送
    message['To'] = receiver #收件
    message['Subject'] = Header(subject, 'utf-8')
    message.attach(MIMEText("""Dear All:\n  \
    附件是截止今天的vintage报表，请查阅！\n  \
有任何疑问请随时与我联系，谢谢！ \
    """, 'plain', 'utf-8'))# 邮件正文           
                           
    if ((current_time.tm_hour == ehour) and (current_time.tm_min == emin) and (current_time.tm_sec == esec)):
       print ("开始")
       #执行
       try:
           risk_test = pymysql.connect(host="localhost",user="root",
                       password="root",database="learn",
                       charset="utf8")
           query="""
                    select c.放款月,round(sum(放款金额/10000),2) 放款金额,
                           case when sum(放款金额)>0 then round(sum(mob1)/sum(放款金额),4) else null end as mob_1,
                           case when sum(放款金额)>0 then round(sum(mob2)/sum(放款金额),4) else null end as mob_2,
                           case when sum(放款金额)>0 then round(sum(mob3)/sum(放款金额),4) else null end as mob_3,
                           case when sum(放款金额)>0 then round(sum(mob4)/sum(放款金额),4) else null end as mob_4,
                           case when sum(放款金额)>0 then round(sum(mob5)/sum(放款金额),4) else null end as mob_5,
                           case when sum(放款金额)>0 then round(sum(mob6)/sum(放款金额),4) else null end as mob_6
                     from(
                           select 分期数,放款月,
                                  sum(case when mob=1 and 当前最大逾期天数>0 then 剩余本金 else 0 end) as mob1,
                                  sum(case when mob=2 and 当前最大逾期天数>0 then 剩余本金 else 0 end) as mob2,
                                  sum(case when mob=3 and 当前最大逾期天数>0 then 剩余本金 else 0 end) as mob3,
                                  sum(case when mob=4 and 当前最大逾期天数>0 then 剩余本金 else 0 end) as mob4,
                                  sum(case when mob=5 and 当前最大逾期天数>0 then 剩余本金 else 0 end) as mob5,
                                  sum(case when mob=6 and 当前最大逾期天数>0 then 剩余本金 else 0 end) as mob6
                           from ( 
                                  select m.分期数,m.剩余本金,m.放款月,m.观测月,m.当前最大逾期天数,
                                        case when substr(m.观测月,1,4)=substr(m.放款月,1,4) then substr(m.观测月,6,2)-substr(m.放款月,6,2)
                                             when substr(m.观测月,1,4)=substr(m.放款月,1,4)+1 then 12+substr(m.观测月,6,2)-substr(m.放款月,6,2)
                                        else 0 end as mob 
                                  from learn.repayment_sum_month m) a 
                          group by 分期数,放款月) b
                   join (select 合同期限,substr(放款日期,1,7) 放款月,
                                sum(合同金额) 放款金额,count(1) 放款量
                           from learn.customer_detail 
                          where 放款日期>='2017-11-01' 
                            and 放款日期<='2018-04-30'
                            and 合同期限=6
                          group by 合同期限,substr(放款日期,1,7)) c 
                     on b.分期数=c.合同期限 and b.放款月=c.放款月
                  group by c.放款月"""
           data=pd.read_sql(query,risk_test)

           workbook = xlsxwriter.Workbook('vintage_report.xlsx') #新建一个excel文本
           worksheet = workbook.add_worksheet("vintage_report")
           chart = workbook.add_chart({'type': 'line'})    #创建一个图表对象
       
           list_1=range(len(data)) 
           title = [u'放款月',u'放款金额',u'mob_1',u'mob_2',u'mob_3',u'mob_4',u'mob_5',u'mob_6']
           format=workbook.add_format()          #定义format格式对象
           format.set_border(1)        #定义format对象单元格边框加粗
           format_title=workbook.add_format()            #定义format_title格式对象
           format_title.set_border(1)         #定义format_title对象单元格边框加粗
           format_title.set_bg_color('#blue')           #定义format_title对象单元格背景颜色
           format_title.set_align('center')           #定义format_title对象单元格居中对齐
           format_title.set_bold()        #定义format_title对象单元格内容加粗
           format_title.set_font_color('white') 
           worksheet.write_row('A1',title,format_title)  
  

           for i in list_1:
               for j in range(1):
                   worksheet.write(i+1,j+0,data['放款月'][i],format)#写入EXCEL表格
                   worksheet.write(i+1,j+1,data['放款金额'][i],format)
                   worksheet.write(i+1,j+2,data['mob_1'][i],format)
                   worksheet.write(i+1,j+3,data['mob_2'][i],format)
                   worksheet.write(i+1,j+4,data['mob_3'][i],format)
                   worksheet.write(i+1,j+5,data['mob_4'][i],format)
                   worksheet.write(i+1,j+6,data['mob_5'][i],format)
                   worksheet.write(i+1,j+7,data['mob_6'][i],format)
               i += 1
           #定义图表数据系列函数
           def chart_series(cur_row):
               chart.add_series({
                   'categories': '=vintage_report!$C$1:$H$1',     
                   'values': '=vintage_report!$C$'+cur_row+':$H$'+cur_row,          
                   'name': '=vintage_report!$A$'+cur_row,            
               })
    
           for row in range(2, 8):     #数据域以第2~7行进行图表数据系列函数调用
               chart_series(str(row))
    
           chart.set_size({'width': 520, 'height': 300})            #设置图表大小
           chart.set_title ({'name': u'vintage报表'})          #设置图表(上方)大标题
           worksheet.insert_chart('A9', chart)          #在A8单元格插入图表
           workbook.close()  # 关闭报表
           print("报表更新成功")
          
           # 构造附件
           att1 = MIMEText(open('vintage_report.xlsx','rb').read(), 'base64', 'utf-8')
           att1["Content-Type"] = 'application/octet-stream' # 二进制编码
           att1["Content-Disposition"] = "attachment;filename=vintage_report.xlsx"
           message.attach(att1) 
    
           smtp = smtplib.SMTP_SSL(host = 'smtp.qq.com')
           smtp.connect(smtpserver, smtpport) #连接服务器
           smtp.login(username, password) #登录
           smtp.sendmail(sender, receiver, message.as_string())  #发送
           smtp.quit()
           print("邮件发送成功")
           
       except:
           print("报表更新失败")
       print(cur_time)
    time.sleep(1)
