
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np

Country_Adgroup_Maxid={"NEW-JP":100,"GV-MX":20,"NEW-MX":20,"GV-US":0.99,"GV-CA":0.99,"NEW-US":0.99,"NEW-CA":0.99,"NEW-UK":0.99,"NEW-IT":0.99, "NEW-ES":0.99,"NEW-DE":0.99,"NEW-DE":0.99,"HM-US":0.99}
Country_Keyword_Maxbid={"NEW-JP":80,"GV-MX":10,"NEW-MX":10,"GV-US":0.99,"GV-CA":0.88,"NEW-CA":0.88,"NEW-UK":0.88,"NEW-IT":0.88, "NEW-ES":0.88,"NEW-DE":0.88,"NEW-DE":0.88,"HM-US":0.88}
Country_DailyBudget={"NEW-JP":300,"GV-MX":300,"NEW-MX":300,"GV-US":3,"GV-CA":3,"NEW-CA":3,"NEW-UK":3,"NEW-IT":3, "NEW-ES":3,"NEW-DE":3,"NEW-DE":3,"HM-US":3}
bulkdatafilepath = 'D:\\运营\\1数据源\周bulk广告数据\\'

AllCountryActions=pd.read_excel(r'D:\\运营\\3数据分析结果\\国家汇总.xlsx',sheet_name="ProductActions")

AllCountryActions_CountryList=AllCountryActions["COUNTRY"].drop_duplicates().to_list()
AllCountryActions.dropna(subset=["SKU"],inplace=True)
AllCountryActions["SKU"].astype(str)

for AllCountryActions_Country in AllCountryActions_CountryList:
    print("现在处理"+AllCountryActions_Country)
    n=0                                             
    for bulkdatafile in os.listdir(bulkdatafilepath): #找bulkfile对应的国家文件

        Bulkfile_Country=bulkdatafile.split('_')[0]

        if Bulkfile_Country==AllCountryActions_Country:#if A1
            bulkfile_draft_1Country=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status", "Bidding strategy","Placement Type","Increase bids by placement","变更记录"]) 

            bulkfile=pd.read_excel(bulkdatafilepath+str(bulkdatafile),engine="openpyxl",sheet_name=1)
            bulkfile["变更记录"]=""
            bulkfile["之前状态"]=" "
            print("找到了要处理的文件")
            
#############################################################确保关闭##############################################################
            AllCountryActions_Country_SKU_Close_List=AllCountryActions.loc[(AllCountryActions["COUNTRY"]==AllCountryActions_Country)&(AllCountryActions["行动方案"].str.contains("关闭广告")),"SKU"].drop_duplicates().to_list()
            print(AllCountryActions_Country_SKU_Close_List)
            AllCountryActions_Country_SKU_Close_List_nocomma_list=[]
            for AllCountryActions_Country_SKU_Close_List_nocomma in AllCountryActions_Country_SKU_Close_List:
                print(AllCountryActions_Country_SKU_Close_List_nocomma)
                AllCountryActions_Country_SKU_Close_List_nocomma=AllCountryActions_Country_SKU_Close_List_nocomma
                if ',' in  AllCountryActions_Country_SKU_Close_List_nocomma:
                    print("包含,",AllCountryActions_Country_SKU_Close_List_nocomma)
                    chaifenlist=AllCountryActions_Country_SKU_Close_List_nocomma.split(",")
                    print(chaifenlist)
                    AllCountryActions_Country_SKU_Close_List_nocomma_list+=chaifenlist
                    print(AllCountryActions_Country_SKU_Close_List_nocomma_list)
                
                else:
                    print(AllCountryActions_Country_SKU_Close_List_nocomma)
                    chaifen=AllCountryActions_Country_SKU_Close_List_nocomma
                    AllCountryActions_Country_SKU_Close_List_nocomma_list+=[chaifen]
                        
                    
                    print(chaifen)
                    
                    
            AllCountryActions_Country_SKU_Close_List=AllCountryActions_Country_SKU_Close_List_nocomma_list
            print(AllCountryActions_Country_SKU_Close_List)
                
