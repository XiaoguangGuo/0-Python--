
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil

#####生成各种Campaign summary



#定义bulk数据汇总表所在路径
Allbulkpath='D:\\运营\\'
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')
AllbulkD5=Allbulk[Allbulk['Keyword or Product Targeting'].notna()]

#生成统计用汇总表
#生成统计用汇总表
#生成统计用汇总表

AllbulkCampaign=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
#AllbulkCampaign1week=Allbulk[(Allbulk['Record Type']=="Campaign")&(Allbulk['周数']==1)].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')


AllbulkCampaignWEEK=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignKeywordWEEK=Allbulk.groupby(["Country","Campaign","Keyword or Product Targeting","Ad Group","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")

writer=pd.ExcelWriter(Allbulkpath+'周bulk数据Summary.xlsx')
AllbulkCampaign.to_excel(writer,"Campaign汇总")
AllbulkCampaignWEEK.to_excel(writer,"CampaignWEEK汇总")
AllbulkSKUWEEK.to_excel(writer,"SKU-WEEK")
AllbulkCampaignSKUWEEK.to_excel(writer,"SKU-Campaign-WEEK")
AllbulkCampaignKeywordWEEK.to_excel(writer,"Keyword-Campaign-WEEK")
writer.save()
#生成统计用汇总表完成


AllbulkWeek4=Allbulk[Allbulk['Keyword or Product Targeting'].notna() &(Allbulk['周数']<5)]
print(AllbulkWeek4)
input("?")
AllbulkCampaignKeyword=AllbulkD5.groupby(["Country","Campaign","Keyword or Product Targeting"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")
AllbulkCampaignKeywordweek4=AllbulkWeek4.groupby(["Country","Campaign","Keyword or Product Targeting"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")

 

AllbulkCampaignKeyword["zhuanhualv"]=AllbulkCampaignKeyword['Orders']/AllbulkCampaignKeyword['Clicks']
AllbulkCampaignKeyword["广告状态"]=" "
AllbulkCampaignKeywordweek4["zhuanhualv"]=AllbulkCampaignKeywordweek4['Orders']/AllbulkCampaignKeywordweek4['Clicks']
AllbulkCampaignKeywordweek4["广告状态"]=" "
clickatempt=10
clickenough=16
 


# condiion1 "点击大于9小于17,转化率=0，降价0.5" 
AllbulkCampaignKeyword.loc[(AllbulkCampaignKeyword["Clicks"]>(clickatempt-1)) &(AllbulkCampaignKeyword["Clicks"]<(clickenough+1))&(AllbulkCampaignKeyword["zhuanhualv"]==0),"广告状态"]="点击大于9小于17,转化率=0，降价0.5"
AllbulkCampaignKeywordweek4.loc[(AllbulkCampaignKeywordweek4["Clicks"]>(clickatempt-1)) &(AllbulkCampaignKeywordweek4["Clicks"]<(clickenough+1))&(AllbulkCampaignKeywordweek4["zhuanhualv"]==0),"广告状态"]="点击大于9小于17,转化率=0，降价0.5"
#  condiion2  "点击>16，转化率为0,paused"

AllbulkCampaignKeyword.loc[(AllbulkCampaignKeyword["Clicks"]>clickenough)&(AllbulkCampaignKeyword["zhuanhualv"]==0),"广告状态"]="点击>16，转化率为0,paused"
AllbulkCampaignKeywordweek4.loc[(AllbulkCampaignKeywordweek4["Clicks"]>clickenough)&(AllbulkCampaignKeywordweek4["zhuanhualv"]==0),"广告状态"]="点击>16，转化率为0,paused"
# condiion3
AllbulkCampaignKeyword.loc[(AllbulkCampaignKeyword["Clicks"]>clickenough)&(AllbulkCampaignKeyword["zhuanhualv"]>0.0625)&(AllbulkCampaignKeyword["zhuanhualv"]<0.125),"广告状态"]="点击>16，转化率低于0.125, 降低价*0.75"
AllbulkCampaignKeywordweek4.loc[(AllbulkCampaignKeywordweek4["Clicks"]>clickenough)&(AllbulkCampaignKeywordweek4["zhuanhualv"]>0.0625)&(AllbulkCampaignKeywordweek4["zhuanhualv"]<0.125),"广告状态"]="点击>16，转化率低于0.125, 降低价*0.75"
#condition4
AllbulkCampaignKeyword.loc[(AllbulkCampaignKeyword["Clicks"]>clickenough)&(AllbulkCampaignKeyword["zhuanhualv"]<0.0625)&(AllbulkCampaignKeyword["zhuanhualv"]>0),"广告状态"]="点击>16，转化率低于0.0625, 降低价*0.5"
AllbulkCampaignKeywordweek4.loc[(AllbulkCampaignKeywordweek4["Clicks"]>clickenough)&(AllbulkCampaignKeywordweek4["zhuanhualv"]<0.0625)&(AllbulkCampaignKeywordweek4["zhuanhualv"]>0),"广告状态"]="点击>16，转化率低于0.0625, 降低价*0.5"
#condition5
AllbulkCampaignKeyword.loc[AllbulkCampaignKeyword["zhuanhualv"]>0.2,"广告状态"]="广告状态"
AllbulkCampaignKeywordweek4.loc[AllbulkCampaignKeywordweek4["zhuanhualv"]>0.2,"广告状态"]="转化率>0.2，开启"

AllbulkCampaignKeyword.to_excel(Allbulkpath+'周bulk广告数据汇总表huizong2.xlsx')
AllbulkCampaignKeywordweek4.to_excel(Allbulkpath+'周bulk广告数据汇总表huizongweek4.xlsx')
