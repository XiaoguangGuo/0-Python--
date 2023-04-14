# -*- coding: utf-8 -*-
import pandas as pd
import os
import shutil
import openpyxl

src_dir_path_sales=r'D:\运营\广告BulkOperation'
#设置查找文件路径


dfadbulkskusum=pd.DataFrame(columns=["SKU","Spend","Country","Date"])
#取得原始目标文件的三列

#建立一个空的过渡文件

for file in os.listdir(src_dir_path_sales): 
    print(file)

    dfadbulkfile1= pd.read_excel(r'D:\\运营\\广告BulkOperation\\'+ str(file),usecols = ['SKU','Spend'],sheet_name="Sponsored Products Campaigns").assign(Country=os.path.basename(file).split('_')[0],Date=os.path.basename(file).split('-')[3])
    dfadbulkfile1=dfadbulkfile1[dfadbulkfile1['SKU'].notnull()]
    print(dfadbulkfile1)
    dfadbulkfile2=dfadbulkfile1.groupby('SKU','Country','date')['spend'].sum() 
    print(dfadbulkfile2)
    dfadbulkskusum=dfadbulkskusum.append (dfadbulkfile2,ignore_index=True)
    
print(dfadbulkskusum)
dfadbulkskusum.to_excel(r'D:/运营/Sku-Spend-Country-Date.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)                                                                                                                                              

                                                                                                                                                    
                                                                                                                                                    
