# 使用Oenyxl编写，没有使用pandas
### 第一部分汇总各国bulk广告数据到汇总表
#用于实际使用，从零开始导入10周最新的。创建新的文件，存到指定文件名。：测试OK。

# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np
import sqlite3


newdate=input('输入最新日期y-m-d',) #输入最新一周的日期
maxtime=datetime.datetime.strptime(newdate,'%Y-%m-%d')
print(newdate)
print(maxtime)

Allbulkpath='D:\\运营\\2生成过程表\\'  

conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db') 
today = datetime.date.today()
weeks_ago_30 = today - datetime.timedelta(weeks=27)

# 构建 SQL 查询并获取数据
query = f'SELECT * FROM "Bulkfiles" WHERE 日期 >= \'{weeks_ago_30}\''
Allbulk = pd.read_sql_query(query, conn)
Allbulk["日期"]=pd.to_datetime(Allbulk["日期"])

Allbulk["周数"]=(maxtime-Allbulk["日期"]).dt.days//7+1


#！！！！筛选汇总表的数据---后续可以按日期>某个日期来筛选
AllbulkD5=Allbulk[(Allbulk['Keyword or Product Targeting']. notna())&(Allbulk['周数']<27)]#定义了26周汇总
print(AllbulkD5)




AllbulkCampaignKeyword=AllbulkD5.groupby(["Country","Campaign","Keyword or Product Targeting","Match Type","Ad Group"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum") 
#AllbulkCampaignKeyword所有周数小于AllbulkD5定义的历史数据汇总

print(AllbulkCampaignKeyword.columns)

maxrow=len(AllbulkCampaignKeyword)
#maxrow1Week=len(AllbulkCampaignKeyword1Week)
AllbulkCampaignKeyword["zhuanhualv"]=AllbulkCampaignKeyword['Orders']/AllbulkCampaignKeyword['Clicks']
print(maxrow)
print(AllbulkCampaignKeyword)

AllbulkCampaignKeyword.to_excel(Allbulkpath+'周bulk广告数据汇总表huizong.xlsx') #生成分析用的报表



#从国家-bulkopration遍历
bulkoperationfilespath = 'D:\\运营\\1数据源\\bulkoperationfiles\\' #待处理的bulk文件--下载下来作为上传的基础的
bulkoperationfilesnewpath='D:\\运营\\4行动表\\bulkoperationfilesNEW\\'#修改后的bulk文件目录
countrylist1=["NEW-JP","NEW-IT","GV-MX","NEW-MX","NEW-FR","NEW-DE","NEW-ES"]
bulkfilecountrylist=[]#国家空列表
for bulkoperationfile in os.listdir(bulkoperationfilespath):
    clicktempt0=20
    clickatempt=10
    clicktemp=50
    clickenough1=12
    clickenough2=16#暂停条件为新站扶持国家,订单数<点击/16



    
    #提取国家名
    bulkfileCountry=os.path.basename(bulkoperationfile).split('_')[0]#遍历国家名
  
    print(bulkfileCountry,type(bulkfileCountry))
  
    
    bulkfilecountrylist.append(bulkfileCountry) #国家名追加到成国家清单
   
    print(bulkfilecountrylist) 

    dfbulkfile=pd.read_excel(bulkoperationfilespath+str(bulkoperationfile),engine="openpyxl",sheet_name=1)#用PD读取待操作的某国bulk文件。
    #dfbulkfile.replace(",",".",inplace=True)#替换逗号为标点，这个操作可能无效。
    dfbulkfile["更新内容"]=""
                        
    dfbulkfile["StatusNEW"]=""

    
  
######定义各种判断条件：用筛选的方式
    
    #####可以自合df筛选条件的字符串


    #bulkcondiion_new01='AllbulkCampaignKeyword["Orders"]<(AllbulkCampaignKeyword["Clicks"]/clickenough1).apply(np.floor))'
    bulcountryconditionstr='AllbulkCampaignKeyword[(AllbulkCampaignKeyword["Country"=="bulkfileCountry")'                            

    bulkcountrycondionclk0='(AllbulkCampaignKeyword["Clicks"]>(clickatempt-1)) &(AllbulkCampaignKeyword["zhuanhualv"]<0.02)]'
    bulkcountrycondionclk1str='(AllbulkCampaignKeyword["Clicks"]>(clickatempt-1)) &(AllbulkCampaignKeyword["Clicks"]<(clickenough+1))]'
    bulkcountrycondionclk21str='AllbulkCampaignKeyword["Clicks"]>clickenough'
    bulkcountrycondionzhl01str= '(AllbulkCampaignKeyword["zhuanhualv"]==0]'
    bulkcountrycondionzhl02str= '(AllbulkCampaignKeyword["zhuanhualv"]>0.2]'
    bulkcountrycondionzhl032str='(AllbulkCampaignKeyword["zhuanhualv"]>0.5]'
    

    
    #bulkconditiondic={}#

####################################################################################################3
    
##差广告：点击次数>16且订单数<点击次数/16 或者点击次数》12 且订单数<点击次数/12
    if bulkfileCountry in countrylist1:
        

        
        bocondition_new00df=AllbulkCampaignKeyword[(AllbulkCampaignKeyword["Country"]==str(bulkfileCountry)) & (AllbulkCampaignKeyword["Clicks"]>clicktemp) &(AllbulkCampaignKeyword["zhuanhualv"]<=0.03)]
    #符合第一个条件是有点击且订单数>点击数/16.#取整替代方法尝试(AllbulkCampaignKeyword["Clicks"]/clickenough2).round(0) #.astype("int")
    else:
        bocondition_new00df=AllbulkCampaignKeyword[(AllbulkCampaignKeyword["Country"]==str(bulkfileCountry)) & (AllbulkCampaignKeyword["Clicks"]>clicktemp) &(AllbulkCampaignKeyword["zhuanhualv"]<=0.04)]
                                                                                                             
   
    print("bocondition_new00df",bocondition_new00df)#打印符合条件1的列表
  
    
####从第一行开始遍历这个符合条件1的筛选表
                                                                                                    
    for boi in range(0,len(bocondition_new00df)):
                                                                                                    
        boi_keyword=bocondition_new00df.iloc[[boi],[2]].values[0][0]#获取第n行的关键词字段的值，遍历
        print(boi_keyword,boi)
         
        boi_campaign=bocondition_new00df.iloc[[boi],[1]].values[0][0]#获取第n行的关键词Campaign的值，遍历
        boi_matchtype=bocondition_new00df.iloc[[boi],[3]].values[0][0]
        boi_ADgroup=bocondition_new00df.iloc[[boi],[4]].values[0][0]                                                                                             


#将对bulkfile中符合筛选表1的camaign和词进行Status批量进行修改成暂停pause

        
        
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Status']=="enabled")&(dfbulkfile['Ad Group']==boi_ADgroup)&(dfbulkfile['Match Type']==boi_matchtype),"更新内容"]="enabled to paused"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Status']=="paused")&(dfbulkfile['Ad Group']==boi_ADgroup)&(dfbulkfile['Match Type']==boi_matchtype),"更新内容"]=" Keeping paused"
                                 
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Status']=="enabled")&(dfbulkfile['Ad Group']==boi_ADgroup)&(dfbulkfile['Match Type']==boi_matchtype),"StatusNEW"]="paused"
        
        
           
        print("01:将bulkfile中这个campaign-词的Status进行修改")




        
#########a将修改过的文件存入bulkoperationfilesnewpath
    dfbulkfile_notnull=dfbulkfile[(dfbulkfile['更新内容'].notna()) & (dfbulkfile['更新内容'] !="")]#只保留#只保留
    Today= datetime.datetime.today().strftime('%Y-%m-%d')
    
    dfbulkfile_notnull.to_excel(bulkoperationfilesnewpath+"Spl-Re_"+Today+"_"+str(bulkoperationfile),index=False)#生成简化版
    
    dfbulkfile.to_excel(bulkoperationfilesnewpath+"Re_"+Today+"_"+str(bulkoperationfile),index=False)

