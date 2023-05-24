# -*- coding: utf-8 -*-
import pandas as pd
import os
import shutil
import openpyxl

src_dir_path_sales=r'D:\运营\广告BulkOperation'
#设置查找文件路径

campaignskucountry=pd.read_excel(r'D:\运营\Campaign-sku-Country.xlsx')
#先读出旧的目标文件

dfadbulk=pd.DataFrame(columns=["Campaign","SKU","Country"])

print(dfadbulk.columns)
#建立一个空的过渡文件

for file in os.listdir(src_dir_path_sales): 
    print(file)

    dfadbulkfile= pd.read_excel(r'D:\\运营\\广告BulkOperation\\'+ str(file),usecols = ['Campaign','SKU'],sheet_name="Sponsored Products Campaigns").assign(Country=os.path.basename(file).split('_')[0])
    
                                                                                                                                              
    #运行ok：dataframe = pd.read_excel(r'D:/运营/广告BulkOperation/US_bulk-a2ylsh1y5o0352-20210219-20210226-1614333746716.xlsx',usecols = ['Campaign','SKU'],sheet_name=1).assign(Country=os.path.basename(r'D:/运营/广告BulkOperation/US_bulk-a2ylsh1y5o0352-20210219-20210226-1614333746716.xlsx').split('_')[0])
    #sheet_name写成="sheet名"或者=1都可以;assign(Country=os.path.basename('US_bulk-a2ylsh1y5o0352-20210219-20210226-1614333746716.xlsx').split('_')[0]) basename直接写文件名即可。
    # dataframe = pd.read_excel(r'D:/运营/Sponsored Products Search term report.xlsx',
                          #usecols = ['Compaign Name','Customer Search Term'],Sheetname="Sponsored Product Search Term R"))
    # 有的文章说要加Sheetname="Sponsored Product Search Term R" ，但实际不用

                                                                                                                                                
    print(dfadbulkfile)

    #print(dfadbulk[dfadbulk['SKU'].notnull()])
    #运行ok可以用drop方法：dfadbulk=dfadbulk.dropna(axis=0,how='any')
    #运行ok可以用筛选方法：dfadbulk=dfadbulk[dfadbulk['SKU'].notnull()]
    dfadbulkfile=dfadbulkfile[dfadbulkfile['SKU'].notnull()]
    #去掉SKU列为空的行
    print(dfadbulkfile)
    
    #运行ok:dfadbulk.drop_duplicates(inplace=True)
    #运行ok:dfadbulk=dfadbulk.drop_duplicates(subset=None,keep='first',inplace=False)
    
    dfadbulk=dfadbulkfile.append (dfadbulk,ignore_index=True)
    dfadbulk.drop_duplicates(inplace=True)
    print(dfadbulk)
    #把文件加到过渡文件中
   

    shutil.move(r'D:\\运营\\广告BulkOperation\\'+ str(file), 'D:/运营/HistoricalData/广告BulkOperation')
    #将搜到的文件移动到历史文件夹
     
campaignskucountry= campaignskucountry.append (dfadbulk,ignore_index=True)
campaignskucountry.drop_duplicates(inplace=True)
# 将过渡文件追加到目标文件中；去掉重复项   
campaignskucountry.to_excel(r'D:/运营/Campaign-sku-Country.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
#导出到Excel

Campaignskucountryuse=campaignskucountry.drop_duplicates(subset=['Campaign','Country'], keep='last',inplace=False)
print(Campaignskucountryuse)
 
Campaignskucountryuse.to_excel(r'D:/运营/Campaign-sku-Country-use.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False) 



                                                                                                                                               
                                                                                                                                               


