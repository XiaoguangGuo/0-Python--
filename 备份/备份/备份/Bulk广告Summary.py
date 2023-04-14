#使用Oenyxl编写，没有使用pandas
### 第一部分汇总各国bulk广告数据到汇总表
#用于实际使用，从零开始导入10周最新的。创建新的文件，存到指定文件名。：测试OK。

# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil



#####以下为处理bulk操作报表的程序#####以下为处理bulk操作报表的程序

#####以下为处理bulk操作报表的程序#####以下为处理bulk操作报表的程序

#####以下为处理bulk操作报表的程序#####以下为处理bulk操作报表的程序

 
print("以下为处理bulk操作报表的程序")

##先用临时文件测试后copy来
#此程序测试完后应该考入bulkoperation程序中

print("以下为处理bulk操作报表的程序")



#定义bulk数据汇总表所在路path='D:\\运营\\'
Allbulkpath='D:\\运营\\'

Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')
AllbulkBrand=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表-品牌.xlsx')
ListingPd=pd.read_excel(r'D:/2019plan/Listing.xlsx')

 

#####生成各种Campaign summary

 

AllbulkCampaign=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
#AllbulkCampaign1week=Allbulk[(Allbulk['Record Type']=="Campaign")&(Allbulk['周数']==1)].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkFenlei=pd.merge(Allbulk,ListingPd,how="left",left_on=["Country","SKU"],right_on=["COUNTRY","seller-sku"])

AllbulkCampaignWEEK=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkFenleiWeek=AllbulkFenlei.groupby(["Country","大类","小类","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignSKUWEEK_fenlei=pd.merge(AllbulkCampaignSKUWEEK,ListingPd,how="left",left_on=["Country","SKU"],right_on=["COUNTRY","seller-sku"])

AllbulkCampaignKeywordWEEK=Allbulk.groupby(["Country","Campaign","Keyword or Product Targeting","Ad Group","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")
AllbulkCampaignKeywordTOTAL=Allbulk.groupby(["Country","Campaign","Keyword or Product Targeting","Ad Group" ],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")

AllbulkCOUNTRYSKUCAMPAIGN=Allbulk.groupby(["Country","Campaign","SKU" ],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")


#品牌

AllbulkBrandCampaignWEEK=AllbulkBrand[AllbulkBrand['Record Type']=="Campaign"].groupby(["Country","Campaign","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkBrandCampaignKeywordWEEK=AllbulkBrand.groupby(["Country","Campaign","Keyword","Ad Group","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")




#写入Excel表格
writer=pd.ExcelWriter(Allbulkpath+'周bulk数据Summary.xlsx')
AllbulkCampaign.to_excel(writer,"Campaign汇总")
AllbulkCampaignWEEK.to_excel(writer,"CampaignWEEK汇总")
AllbulkSKUWEEK.to_excel(writer,"SKU-WEEK")
AllbulkCampaignSKUWEEK.to_excel(writer,"SKU-Campaign-WEEK")
AllbulkFenleiWeek.to_excel(writer,"AllbulkFenleiWeek")
AllbulkCampaignKeywordWEEK.to_excel(writer,"Keyword-Campaign-WEEK")
AllbulkCampaignKeywordTOTAL.to_excel(writer,"CAMPAIGN-KEYEORD-TOTAL")
AllbulkCOUNTRYSKUCAMPAIGN.to_excel(writer,"COUNTRY-CAMPAIGN-SKU")
AllbulkCampaignSKUWEEK_fenlei.to_excel(writer,"COUNTRY-CAMPAIGN-SKU_fenlei")

AllbulkBrandCampaignWEEK.to_excel(writer,"Brand-Campaign-week")
AllbulkBrandCampaignKeywordWEEK.to_excel(writer,"Brand-Campaign-keyword-week")




writer.save()
