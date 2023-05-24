
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np



SearchTermAll=pd.read_excel(r'D:\\运营\\运行结果数据\\Sponsored Products Search term report.xlsx')

SearchTermAll["Clicks"].fillna(0,inplace=True)


SearchTermAll_Sum=SearchTermAll.groupby(["Country","Campaign Name", "Ad Group Name","Customer Search Term"],as_index=False)[["Impressions","Clicks","Spend","7 Day Total Sales ","7 Day Total Orders (#)"]].agg("sum")


SearchTermAll_Sum.loc[SearchTermAll_Sum['Clicks']>0,"转化率"]=SearchTermAll_Sum["7 Day Total Orders (#)"]/SearchTermAll_Sum['Clicks']


SearchTermAll_Good=SearchTermAll_Sum[SearchTermAll_Sum["转化率"]>0.2]




print(SearchTermAll_Sum)

print(SearchTermAll_Good)


#遍历SeachTermGood


Allbulkpath='D:\\运营\\'  
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')


Allbulk_Week1=Allbulk[Allbulk["周数"]==1]


SearchTermAll_Good_Country_list=SearchTermAll_Good["Country"].dropduplicates().to_list()
for countryname in SearchTermAll_Good_Country_list：
    SearchTermAll_Good_Country=SearchTermAll_Good[[SearchTermAll_Good["Country"]==str(countryname)]



    for Stoi in range(0,len(SearchTermAll_Good_Country)):

    
    Stoi_keyword=SearchTermAll_Good_Country.iloc[[Stoi],[4]].values[0][0]
    Stoi_Campaign=SearchTermAll_Good_Country.iloc[[Stoi],[2]].values[0][0]#获取第n行的关键词字段的值，遍历
    Stoi_AdGroup=SearchTermAll_Good_Country.iloc[[Stoi],[2]].values[0][0]
    avearageprice=SearchTermAll_Good_Country.iloc[[Stoi],[6]]/SearchTermAll_Good_Country.iloc[[Stoi],[5]
        
#下面判断bulkfile中 Exact是否 含有Stoi_keyword
#如果含有，是否是开着的？
                                                                                              
#如果没有则看是否有同一个产品的Camaign
  如果有，   是否是开着的？                                                                                         
#生成一行exact：maxbid=avearageprice；打开campaign
#如果没有Campaign,则                                                                                              
                                                                                              

 判断phrase中是否有
                                                  
#如果含有，是否是开着的？
#如果没有则生成一行phrase


生成的行加入Country Bulk 文件中                                                  
                                                  
循环第二个国家。

                                                  

                                                  
                                                  
                                                
