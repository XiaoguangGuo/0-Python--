




# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np



zhuanhualv_bad={"GV-US":0.05,"GV-CA":0.05,"NEW-UK":0.05,"NEW-JP":0.03,"NEW-CA":0.05,"NEW-IT":0.05,"NEW-DE":0.05,"NEW-ES":0.05,"NEW-FR":0.05,"NEW-US":0.05,"HM-US":0.05,"GV-MX":0.03,"NEW-MX":0.03,"HM-US":0.05}}
bulkdatafilepath = 'D:\\运营\\1数据源\周bulk广告数据\\'

#底线策略
SearchTermAll=pd.read_excel(r'D:\\运营\\2生成过程表\\Sponsored Products Search term report.xlsx')

SearchTermAll["Clicks"].fillna(0,inplace=True)


SearchTermAll["Customer Search Term"].astype(str)
SearchTermAll_Sum_Targeting=SearchTermAll.groupby(["Country","Campaign Name", "Ad Group Name","Targeting","Customer Search Term"],as_index=False)[["Impressions","Clicks","Spend","7 Day Total Sales ","7 Day Total Orders (#)"]].agg("sum")

SearchTermAll_Sum=SearchTermAll.groupby(["Country","Campaign Name", "Ad Group Name","Customer Search Term"],as_index=False)[["Impressions","Clicks","Spend","7 Day Total Sales ","7 Day Total Orders (#)"]].agg("sum")


SearchTermAll_Sum.loc[SearchTermAll_Sum['Clicks']>0,"转化率"]=SearchTermAll_Sum["7 Day Total Orders (#)"]/SearchTermAll_Sum['Clicks']




SearchTermAll_Bad=SearchTermAll_Sum[(SearchTermAll_Sum["转化率"]<0.033)&(SearchTermAll_Sum["Clicks"]>30)]


for countryname in  SearchTermAll["Country"].drop_dulicates().to_list()#遍历国家



    SearchTermAll_Bad_Country=SearchTermAll_Sum[(SearchTermAll["Country"]==countryname)&(SearchTermAll_Sum["转化率"]<zhuanhualv_bad[countryname])&(SearchTermAll_Sum["Clicks"]>35)]
    
   
    

    print(SearchTermAll_Bad)

    for searchtermbad_oi in range(len(SearchTermAll_Bad_Country)):
        print(searchtermbad)

        
        
        
        seartchtermbad_Campaign=SearchTermAll_Bad_Country.iloc[[searchtermbad_oi],[1]].values[0][0]                                                      
        seartchtermbad_Campaign_Ad_Group=SearchTermAll_Bad_Country.iloc[[searchtermbad_oi],[2]].values[0][0]
        seartchtermbad=SearchTermAll_Bad_Country.iloc[[searchtermbad_oi],[3]].values[0][0]
        
       SearchTermAll_Sum_Targeting_List=SearchTermAll_Sum_Targeting.loc[(SearchTermAll_Sum_Targeting["Campaign"]==seartchtermbad_Campaign)&(SearchTermAll_Sum_Targeting["Ad Group"]==seartchtermbad_Campaign_Ad_Group)&(SearchTermAll_Sum_Targeting["Keyword or Product Targeting"]==searchtermbad)&,"Targeting"].drop_duplicates().to_list()
        
        
        print("处理",sertchtermbad_Campaign,seartchtermbad_Campaign_Ad_Group)
        
        if bulkfile[(bulkfile["Campaign"]==seartchtermbad_Campaign)&(bulkfile['Record Type']=="Campaign","Campaign Targeting Type"].to_list()[0]=="Manual":

            bulkfile.loc[(bulkfile["Campaign"]==seartchtermbad_Campaign)&(bulkfile["Keyword or Product Targeting"]==seartchtermbad_Campaign)&(bulkfile["Match Type"]=="exact"),"更改记录"]="暂停投放"
            bulkfile.loc[(bulkfile["Campaign"]==seartchtermbad_Campaign)&(bulkfile["Keyword or Product Targeting"]==seartchtermbad_Campaign)&(bulkfile["Match Type"]=="exact"),"Status"]="paused"
            
            #bulkfile.loc[(bulkfile["Campaign"]==seartchtermbad_Campaign)&(bulkfile["Keyword or Product Targeting"]==seartchtermbad_Campaign)&(bulkfile["Match Type"]=="phrase"),"更新记录"]="增加否定词"
            bulkfile=bulkfile.append({"Record Type":"Keyword","Campaign":seartchtermbad_Campaign,"Ad Group":seartchtermbad_Campaign_Ad_Group,"Match Type":"negative exact","Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":searchtermbad,"更改记录":"增加否定词"},ignore_index = True)


        #增加否定词
        "
            #重复bulkfile=bulkfile.append({"Record Type":"Keyword","Campaign":seartchtermbad_Campaign,"Ad Group":seartchtermbad_Campaign_Ad_Group,"Match Type":"negative exact","Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":searchtermbad,"更改记录":"增加否定词"},ignore_index = True)

         elif bulkfile[(bulkfile["Campaign"]==seartchtermbad_Campaign)&(bulkfile['Record Type']=="Campaign","Campaign Targeting Type"].to_list()[0]=="Auto":
                       
            bulkfile=bulkfile.append({"Record Type":"Keyword","Campaign":seartchtermbad_Campaign,"Ad Group":seartchtermbad_Campaign_Ad_Group,"Match Type":"negative exact","Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":searchtermbad,"更改记录":"增加否定词"},ignore_index = True)
        
