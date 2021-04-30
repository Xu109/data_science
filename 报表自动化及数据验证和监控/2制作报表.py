import pandas as pd
import pymysql
import xlrd
# 第一种情况：新建一张报表
# 1、连接数据库
risk = pymysql.connect(host="localhost",user="root",
                      password="root",database="learn",
                      charset="utf8")

# 2、查询数据
query="""select * from learn.daily_report where 日期<='2018-03-31'"""
data=pd.read_sql(query,risk)


# 3、新建xls及sheet,把data写在这个sheet里
writer = pd.ExcelWriter('daily_report.xls') ###################此处文件存储位置需自定义########################
data.to_excel(writer,'daily_report',index=False)

# 4、保存报表
writer.save()

#%% 第二种情况：更新已有报表中的数据
import pandas as pd
import pymysql
from xlutils.copy import copy
# 1、复制原有的报表文件，formatting_info=True表示保留原文件格式
oldWb = xlrd.open_workbook('daily_report.xls',formatting_info=True);###################补充自定义的文件存放位置########################
newWb = copy(oldWb)
newWs = newWb.get_sheet('daily_report')

# 2、测出data_1长度、宽度，以range列出赋值给list_1、list2
list_1=range(len(data))
list_2=range(len(data.columns))

# 3、按照一定的格式和位置循环写入EXCEL表格
data['总进件']=data['总进件'].astype('float64')
for i in list_1:
    for j in range(1):
        newWs.write(i+1,j+1,data['总进件'][i])#写入EXCEL表格
        newWs.write(i+1,j+2,data['审批量'][i])
        newWs.write(i+1,j+3,data['准入拒绝量'][i])
        newWs.write(i+1,j+4,data['通过量'][i])
        newWs.write(i+1,j+5,data['拒绝量'][i])
        newWs.write(i+1,j+6,data['通过率'][i])
        newWs.write(i+1,j+7,data['批核金额'][i])
        newWs.write(i+1,j+8,data['批核日件均'][i])     
        newWs.write(i+1,j+9,data['放款量'][i])
        newWs.write(i+1,j+10,data['放款金额'][i])
        newWs.write(i+1,j+11,data['放款日件均'][i]) 
    i += 1
print ("write new values ok")

# 4、保存报表
newWb.save('daily_report.xls') ###################此处文件存储位置需自定义########################


