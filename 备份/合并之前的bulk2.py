# -*- coding: utf-8 -*-
import os

import pandas as pd
import openpyxl

import shutil
import datetime

bulkfileoldPath = r'D:\\周Bulk汇总表历史\\'

bulkfileold_List = os.listdir(bulkfileoldPath)
bulkfileold1=bulkfileold_List[0]
bulkfileold1file=pd.read_excel(bulkfileoldPath+bulkfileold1,usecols=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Impressions","Clicks","Spend","Orders","Total Units","Sales","ACoS","Bidding strategy","Placement Type","Increase bids by placement","Country","日期"],sheet_name="Sheet")
bulkfileold=pd.DataFrame(columns=bulkfileold1file.columns.to_list())
#bulkfileold_22=bulkfileold=pd.DataFrame(columns=bulkfileold1file.columns.to_list())
for bulkfileold_i in bulkfileold_List:
    bulkfileold_i=pd.read_excel(bulkfileoldPath+bulkfileold_i,sheet_name="Sheet")
    
    #if bulkfileold_i.columns.to_list()==bulkfileold.columns.to_list():
   
        #bulkfileold_i=bulkfileold_i.loc[(bulkfileold_i["日期"]<pd.Timestamp('2022-10-01'))&(bulkfileold_i["日期"]>=pd.Timestamp('2022-08-01'))]
        #bulkfileold_i.drop('周数',axis=1,inplace=True)
    bulkfileold=pd.concat([bulkfileold,bulkfileold_i],ignore_index=True)
    bulkfileold=bulkfileold.drop_duplicates()
        #bulkfileold_i_2=bulkfileold_i.loc[bulkfileold_i["日期"]<pd.Timestamp('2022-08-01')]
        #bulkfileold_i_2.drop('周数',axis=1,inplace=True)
        #bulkfileold_22=pd.concat([bulkfileold_22,bulkfileold_i_2],ignore_index=True)
        #bulkfileold_22=bulkfileold_22.drop_duplicates()                                  
                                          
                                        
    #else:
        #continue


#检验是否含有中文字符
#def is_contains_chinese(strs):
    #for _char in strs:
        #if '\u4e00' <= _char <= '\u9fa5':
            #return True
bulkfileold.to_excel(bulkfileoldPath+"bulkfileold"+".xlsx",index=False)
#bulkfileold_22.to_excel(bulkfileoldPath+"bulkfileold_22"+".xlsx",index=False)






    
