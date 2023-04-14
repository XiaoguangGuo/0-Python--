import pandas as pd
import os
import sqlite3


SKUlist=[]
country = 'GV-US'
countrydic={'GV-US':"US",'NEW-US':"US","NEW-CA":"CA","GV-CA":"CA"}
countryinTop1M=countrydic[country]
df_sku_Asin
file_path = r'D:\运营\2生成过程表\周bulk数据Summary.xlsx'
sheet_name = 'SKU-Campaign-Spend'
columns = ['Country', 'Campaign', 'SKU', 'Spend', 'Orders', 'Clicks', 'zhuanhualv', 'SKU-Campaign-zhuanhualv-ranking', 'Campaign-SKU_Spend_ranking', 'SKU_Campaign_Spend_ranking']
bulkdf = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns)
grouped = bulkdf.groupby(['Country', 'SKU'])[['Spend', 'Orders', 'Clicks']].sum()

for sku in SKUlist

   


    File_FanchaAsinTerms=r'D:\运营\1数据源\反查关键词\\GV-US_反查关键词-10.xlsx'
    File_selfAsinTerms=r'D:\运营\1数据源\反查关键词\\GV-US_反查关键词-B088H1W1W1-进阶版 (1).xlsx'

# 读取 Excel 文件并指定需要读取的列
 


# 按照 Country 和 SKU 进行分组，并对 'Spend'、'Orders'、'Clicks' 进行求和





# 输出指定国家和 SKU 的 Spend、Orders、Clicks 总和


result = grouped.loc[(country, sku)]
print(result)

file_pathSummary = r'D:\运营\2生成过程表\周bulk数据Summary.xlsx'
sheet_name="SKU-Campaign-Spend"
seartchtermSummary= r'D:\运营\2生成过程表\Search_Term_Summary.xlsx'
sheet_name_keywords='SeachTermWeekSum_Weeks'
columnsSummary = ['Country', 'Campaign', 'SKU', 'Spend', 'Orders', 'Clicks', 'zhuanhualv', 'SKU-Campaign-zhuanhualv-ranking', 'Campaign-SKU_Spend_ranking', 'SKU_Campaign_Spend_ranking']
seartchtermSummary_columns=['COUNTRY','Campaign Name',	'Ad Group Name','Targeting','Match Type','Customer Search Term','Impressions','Clicks','Spend',	'7 Day Total Sales','Clicks1','Orders1']

dfSummary = pd.read_excel(file_pathSummary, sheet_name=sheet_name, usecols=columnsSummary)
dfSummary=dfSummary[dfSummary['Country']==country]
seartchtermSummary_df=pd.read_excel(seartchtermSummary, sheet_name=sheet_name_keywords, usecols=seartchtermSummary_columns)

# 筛选 Campaign-SKU_Spend_ranking 为 1 的行
CampaigntoSKU = dfSummary.loc[dfSummary['Campaign-SKU_Spend_ranking'] == 1]
CampaigntoSKUBAoliu=CampaigntoSKU[['Country', 'Campaign', 'SKU']]


seartchtermSummary_df=pd.merge(seartchtermSummary_df,CampaigntoSKUBAoliu,left_on=['COUNTRY','Campaign Name'],right_on=['Country', 'Campaign'],how='left')

seartchtermSummary_df_grouped=seartchtermSummary_df.groupby(["COUNTRY","SKU",'Customer Search Term'],as_index=False)[['Impressions','Clicks','Spend',	'7 Day Total Sales','Clicks1','Orders1']].agg("sum")
seartchtermSummary_df_grouped_country=seartchtermSummary_df_grouped[(seartchtermSummary_df_grouped["COUNTRY"]==country)&(seartchtermSummary_df_grouped["SKU"]==sku)]

import pandas as pd
import os

# 指定需要查找的目录和文件名格式
directory = r'D:\运营\1数据源\周Bulk广告数据'
file_name_format = '{}_*.xlsx'

# 指定国家名并根据其查找文件
 
file_name = file_name_format.format(country)
file_path = None
for file in os.listdir(directory):
    if file.startswith(country):
        file_path = os.path.join(directory, file)
        break

if file_path:
    print(file_path)
    # 如果找到了文件，使用 pandas 打开它，并添加 SKU3 列
    df_bulk = pd.read_excel(file_path,sheet_name="Sponsored Products Campaigns")
    print(df_bulk.columns)
    df_bulk['SKU3'] = ''

    # 从 CampaigntoSKU 中遍历每一行数据，查找相同 Campaign 和 Ad Group 的行，并将其对应的 SKU 值赋给 SKU3 列
    for i, row in CampaigntoSKU.iterrows():
        campaign = row['Campaign']
        
        sku = row['SKU']
        df_bulk.loc[(df_bulk['Campaign'] == campaign), 'SKU3'] = sku
        
                         
    print(df_bulk.head())  # 查看前五行数据
    df_bulk.to_excel(r'D:\运营\关键词考察\\' + file)
    df_bulk_sku_words=df_bulk.loc[(df_bulk["SKU3"]==sku)&(df_bulk["Campaign Status"]=="enabled")&(df_bulk["Ad Group Status"]=="enabled")&(df_bulk["Status"]=="enabled")&(df_bulk["Keyword or Product Targeting"].notnull()),["SKU3","Keyword or Product Targeting","Status"]]
else:
    print('没有找到文件：{}'.format(file_name))

    


selfAsinTerms=pd.read_excel(File_selfAsinTerms)
selfAsinTerms_formerge=selfAsinTerms[["关键词","搜索热度"]]
FanchaAsinTerms=pd.read_excel(File_FanchaAsinTerms)
FanchaAsinTerms=pd.merge(FanchaAsinTerms,selfAsinTerms_formerge,on='关键词',how="left")

conn= sqlite3.connect('D:/运营/sqlite/Top1M.db')
Top1M_df = pd.read_sql_query('SELECT * FROM "Top1M" WHERE Country= countryinTop1M AND 日期 =(SELECT MAX(日期) FROM "Top1M" WHERE Country=countryinTop1M)', conn)

Top1M_df=Top1M_df[["Search Frequency Rank","Search Term","Top Clicked Product #1: ASIN","Top Clicked Product #1: Product Title","Top Clicked Product #1: Click Share","Top Clicked Product #1: Conversion Share"]]


output=pd.merge(FanchaAsinTerms,df_bulk_sku_words,left_on='关键词', right_on="Keyword or Product Targeting",how="left")
output=pd.merge(FanchaAsinTerms,seartchtermSummary_df_grouped_country,left_on='关键词', right_on="Customer Search Term",how="left") 
output=pd.merge(output,Top1M_df,left_on='关键词', right_on="Search Term",how="left")


output.to_excel(r'D:\运营\关键词考察\\' +"New"+ "GV-CA_反查关键词-10.xlsx")
