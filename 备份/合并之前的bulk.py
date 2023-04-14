# -*- coding: utf-8 -*-
import os

import pandas as pd
import openpyxl

import shutil
import datetime

bulkfileoldPath = r'D:\\周Bulk汇总表历史\\'

bulkfileold_List = os.listdir(bulkfileoldPath)
bulkfileold1=bulkfileold_List[0]
bulkfileold1file=pd.read_excel(bulkfileoldPath+bulkfileold1)
bulkfileold=pd.DataFrame(columns=bulkfileold1file.columns.to_list())
for bulkfileold_i in bulkfileold_List:
    bulkfileold_i=pd.read_excel(bulkfileoldPath+bulkfileold_i)
    if bulkfileold_i["Campaign Targeting Type"].str.contains("手动").any() or bulkfileold_i["Campaign Targeting Type"].str.contains("自动").any(): 
        print("中文版")
        continue
    else:
        bulkfileold_i=bulkfileold_i.loc[bulkfileold_i["日期"]<pd.Timestamp('2022-10-01')]
    
        bulkfileold=pd.concat([bulkfileold,bulkfileold_i],ignore_index=True)
        bulkfileold.drop_duplicates(inplace=True)

#检验是否含有中文字符
#def is_contains_chinese(strs):
    #for _char in strs:
        #if '\u4e00' <= _char <= '\u9fa5':
            #return True
    #return Falsebulkfileold.to_excel(bulkfileoldPath+"bulkfileold"+".xlsx",index=False)






    
