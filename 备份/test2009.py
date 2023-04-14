
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np
import shutil
import numpy as np

#在这之前要生成汇总表，并且把每个国家的bulk表备份到D:\\运营\\bulkoperationfiles\\
newdate=input('输入最新日期y-m-d',) #输入最新一周的日期
maxtime=datetime.datetime.strptime(newdate,'%Y-%m-%d')
print(newdate)
print(maxtime)

Allbulkpath='D:\\运营\\2生成过程表\\'  
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表121.xlsx')



Allbulk["周数"]=(Allbulk["日期"]-maxtime).dt.days//7+1


#！！！！筛选汇总表的数据---后续可以按日期>某个日期来筛选
AllbulkD5=Allbulk[(Allbulk['Keyword or Product Targeting'].notna())&(Allbulk['周数']<26)]#定义了26周汇总
print(AllbulkD5)

# AllbulkCampaignKeyword1Week=Allbulk[(Allbulk['Keyword or Product Targeting'].notna())&(Allbulk['周数']==1)].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')

#AbulkCampaignKeyword1week["zhuanhualv"]=AllbulkCampaignKeyword1week['Orders']/AllbulkCampaignKeyword1week['Clicks']


AllbulkCampaignKeyword=AllbulkD5.groupby(["Country","Campaign","Keyword or Product Targeting","Match Type","Ad Group"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum") 
#AllbulkCampaignKeyword所有周数小于AllbulkD5定义的历史数据汇总

print(AllbulkCampaignKeyword.columns)

maxrow=len(AllbulkCampaignKeyword)
#maxrow1Week=len(AllbulkCampaignKeyword1Week)
AllbulkCampaignKeyword["zhuanhualv"]=AllbulkCampaignKeyword['Orders']/AllbulkCampaignKeyword['Clicks']
print(maxrow)
print(AllbulkCampaignKeyword)
