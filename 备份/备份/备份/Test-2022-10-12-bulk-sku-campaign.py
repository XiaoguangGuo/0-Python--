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


#####生成各种Campaign summary

print("以下为处理bulk操作报表的程序")



#定义bulk数据汇总表所在路径
Allbulkpath='D:\\运营\\'
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')


AllbulkCampaignSKU1012=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign"],as_index=False)[['Orders']].agg('max')


writer=pd.ExcelWriter(Allbulkpath+'周bulk数据Summary1012.xlsx')

AllbulkCampaignSKU1012.to_excel(writer,"AllbulkCampaignSKU1012")

writer.save()

