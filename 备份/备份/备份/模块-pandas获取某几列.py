import pandas as pd
import os
dfadbulk= pd.read_excel(r'D:/运营/广告BulkOperation/US_bulk-a2ylsh1y5o0352-20210219-20210226-1614333746716.xlsx',usecols = ['Campaign','SKU'],sheet_name="Sponsored Products Campaigns").assign(Country=os.path.basename('US_bulk-a2ylsh1y5o0352-20210219-20210226-1614333746716.xlsx').split('_')[0])
#运行ok：dataframe = pd.read_excel(r'D:/运营/广告BulkOperation/US_bulk-a2ylsh1y5o0352-20210219-20210226-1614333746716.xlsx',usecols = ['Campaign','SKU'],sheet_name=1).assign(Country=os.path.basename(r'D:/运营/广告BulkOperation/US_bulk-a2ylsh1y5o0352-20210219-20210226-1614333746716.xlsx').split('_')[0])
#sheet_name写成="sheet名"或者=1都可以;assign(Country=os.path.basename('US_bulk-a2ylsh1y5o0352-20210219-20210226-1614333746716.xlsx').split('_')[0]) basename直接写文件名即可。
# dataframe = pd.read_excel(r'D:/运营/Sponsored Products Search term report.xlsx',
                          #usecols = ['Compaign Name','Customer Search Term'],Sheetname="Sponsored Product Search Term R"))
# 有的文章说要加Sheetname="Sponsored Product Search Term R" ，但实际不用。
print(dfadbulk)
# dfadbulk=dfadbulk['Campaign']=="Pr1-oven mitts1118"
print("CamainPr1-oven mitts1118",dfadbulk['Campaign']=="Pr1-oven mitts1118")
print(dfadbulk.loc[dfadbulk['SKU'] =="NaN", ['Campaign', 'SKU','Country']])
print(dfadbulk[dfadbulk['SKU']!=""])

#print(dfadbulk[dfadbulk['SKU'].notnull()])
#运行ok可以用drop方法：dfadbulk=dfadbulk.dropna(axis=0,how='any')
#运行ok可以用筛选方法：dfadbulk=dfadbulk[dfadbulk['SKU'].notnull()]
dfadbulk=dfadbulk[dfadbulk['SKU'].notnull()]
print(dfadbulk)
#运行ok:dfadbulk.drop_duplicates(inplace=True)
#运行ok:dfadbulk=dfadbulk.drop_duplicates(subset=None,keep='first',inplace=False)
dfadbulk.drop_duplicates(inplace=True)
print(dfadbulk)

#也可以用dropNA
#dfadbulk=dfadbulk.groupby('SKU')['SKU']!="NaN" 


