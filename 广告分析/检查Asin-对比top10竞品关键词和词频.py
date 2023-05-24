import pandas as pd
import os
import sqlite3
import shutil
import datetime
from collections import Counter


#1.注定国家 2，指定SKU 3。 修改第94行的Asin 4。设置竞品Asin反查词文件和自己Asin反查词文件路径

#后续修改计划：设置国家和SKU就直接运行。


Country = 'GV-US'  # 指定国:
countrydic={'GV-US':"US",'NEW-US':"US","NEW-CA":"CA","GV-CA":"CA"}
countryinTop1M=countrydic[Country]


CountrySummary=pd.read_excel(r'D:\运营\3数据分析结果\国家汇总.xlsx',sheet_name="ProductActions")

print(CountrySummary)
sku = '20210526-Gateron-Yellow65'  # 指定 SKU

CountrySummary['SKU']=CountrySummary['SKU'].dropna().astype(str)

if not CountrySummary.loc[CountrySummary['SKU'].str.contains(sku)].empty:
    row = CountrySummary.loc[CountrySummary['SKU'].str.contains(sku)].iloc[0]
else:
    print("没有找到满足条件的行")



# Get the value in the Asin column of that row
asin = row['Asin']



#竞品反查词文件
Top10File="GV-US_B08SBRRJBM-反查关键词-10.xlsx"
File_FanchaAsinTerms=r'D:\运营\1数据源\\\反查关键词\\'+Top10File
#自己的反查词文件
SelfAsinFile="GV-US_反查关键词-B095VSNS93-进阶版.xlsx"
File_selfAsinTerms=r'D:\运营\1数据源\反查关键词\\'+SelfAsinFile

#设置historyData路径
History_path=r'D:\\运营\\HistoricalData\\反查关键词\\'

# 读取 Excel 文件并指定需要读取的列
file_path = r'D:\运营\2生成过程表\周bulk数据Summary.xlsx'
sheet_name = 'SKU-Campaign-Spend'
columns = ['Country', 'Campaign', 'SKU', 'Spend', 'Orders', 'Clicks', 'zhuanhualv', 'SKU-Campaign-zhuanhualv-ranking', 'Campaign-SKU_Spend_ranking', 'SKU_Campaign_Spend_ranking']
bulkdf = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns)

# 按照 Country 和 SKU 进行分组，并对 'Spend'、'Orders'、'Clicks' 进行求和
grouped = bulkdf.groupby(['Country', 'SKU'])[['Spend', 'Orders', 'Clicks']].sum()

print(grouped)


file_pathSummary = r'D:\运营\2生成过程表\周bulk数据Summary.xlsx'
sheet_name="SKU-Campaign-Spend"
seartchtermSummary= r'D:\运营\2生成过程表\Search_Term_Summary.xlsx'
sheet_name_keywords='SeachTermWeekSum_Weeks'
columnsSummary = ['Country', 'Campaign', 'SKU', 'Spend', 'Orders', 'Clicks', 'zhuanhualv', 'SKU-Campaign-zhuanhualv-ranking', 'Campaign-SKU_Spend_ranking', 'SKU_Campaign_Spend_ranking']
seartchtermSummary_columns=['Country','Campaign Name',	'Ad Group Name','Targeting','Match Type','Customer Search Term','Impressions','Clicks','Spend','7 Day Total Orders (#)','Clicks1','Orders1']

dfSummary = pd.read_excel(file_pathSummary, sheet_name=sheet_name, usecols=columnsSummary)
dfSummary=dfSummary[dfSummary['Country']==Country]
seartchtermSummary_df=pd.read_excel(seartchtermSummary, sheet_name=sheet_name_keywords, usecols=seartchtermSummary_columns)

# 筛选 Campaign-SKU_Spend_ranking 为 1 的行
CampaigntoSKU = dfSummary.loc[dfSummary['Campaign-SKU_Spend_ranking'] == 1]
CampaigntoSKUBAoliu=CampaigntoSKU[['Country', 'Campaign', 'SKU']]


seartchtermSummary_df=pd.merge(seartchtermSummary_df,CampaigntoSKUBAoliu,left_on=['Country','Campaign Name'],right_on=['Country', 'Campaign'],how='left')

seartchtermSummary_df_grouped=seartchtermSummary_df.groupby(["Country","SKU",'Customer Search Term'],as_index=False)[['Impressions','Clicks','Spend','7 Day Total Orders (#)','Clicks1','Orders1']].agg("sum")
seartchtermSummary_df_grouped_country=seartchtermSummary_df_grouped[(seartchtermSummary_df_grouped["Country"]==Country)&(seartchtermSummary_df_grouped["SKU"]==sku)]

import pandas as pd
import os

# 指定需要查找的目录和文件名格式
directory = r'D:\运营\1数据源\周Bulk广告数据'
file_name_format = '{}_*.xlsx'

# 指定国家名并根据其查找文件
 
file_name = file_name_format.format(Country)
file_path = None
for file in os.listdir(directory):
    print(file)
    
    if file.startswith(Country):
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
        
        sku232 = row['SKU']
        df_bulk.loc[(df_bulk['Campaign'] == campaign), 'SKU3'] = sku232
        
                         
    print(df_bulk.head())  # 查看前五行数据
 
    df_bulk_sku_words=df_bulk.loc[(df_bulk["SKU3"]==sku)&(df_bulk["Keyword or Product Targeting"].notnull()),["SKU3","Keyword or Product Targeting","Campaign","Ad Group","Match Type","Campaign Status","Ad Group Status","Status"]]
    print(df_bulk_sku_words)
else:
    print('没有找到文件：{}'.format(file_name))

    


selfAsinTerms=pd.read_excel(File_selfAsinTerms)
selfAsinTerms_formerge=selfAsinTerms[["关键词","搜索热度","首页流量占比(%)",asin+"自然搜索绝对位置"]]
FanchaAsinTerms=pd.read_excel(File_FanchaAsinTerms) 
FanchaAsinTerms=pd.merge(FanchaAsinTerms,selfAsinTerms_formerge,on='关键词',how="left")

conn= sqlite3.connect('D:/运营/sqlite/Top1M.db')
Top1M_df = pd.read_sql_query('SELECT * FROM "Top1M" WHERE Country="US" AND 日期 =(SELECT MAX(日期) FROM "Top1M" WHERE Country= "US")', conn)

Top1M_df=Top1M_df[["Search Frequency Rank","Search Term","Top Clicked Product #1: ASIN","Top Clicked Product #1: Product Title","Top Clicked Product #1: Click Share","Top Clicked Product #1: Conversion Share"]]

 
output=pd.merge(FanchaAsinTerms,df_bulk_sku_words,left_on='关键词', right_on="Keyword or Product Targeting",how="outer")
output=pd.merge(output,seartchtermSummary_df_grouped_country,left_on='关键词', right_on="Customer Search Term",how="outer") 
output=pd.merge(output,Top1M_df,left_on='关键词', right_on="Search Term",how="left")



output_file = f"D:/运营/关键词考察/{Country}关键词考察表.xlsx"

# 检查文件是否存在，如果不存在则创建一个新文件
if not os.path.exists(output_file):
    empty_df = pd.DataFrame()
    empty_df.to_excel(output_file, index=False)

# 创建一个 ExcelWriter 对象，mode 设置为 'a' 以附加新的工作表
writer = pd.ExcelWriter(output_file, mode='a', engine='openpyxl')

# 在现有 Excel 文件中添加一个新的工作表


today = datetime.datetime.today()
Today = today.strftime("%Y%m%d")

output_name=Today+"_"+sku            
output.to_excel(writer, sheet_name=output_name,index=False)

outputcount=output.loc[(output["Search Frequency Rank"].notnull())&output["Search Frequency Rank"].notna(),"关键词"]
print(outputcount)
outputcount=outputcount.astype(str)
O_list=outputcount.drop_duplicates().to_list()

# 初始化一个空的 Counter 对象
word_counter = Counter()

# 遍历 "关键词" 列并更新计数器
for keywords in O_list:
    words = keywords.split()  # 使用空格分隔关键词
    word_counter.update(words)

# 按降序排列单词频率并打印结果
sorted_word_count = dict(sorted(word_counter.items(), key=lambda x: x[1], reverse=True))
print(sorted_word_count)
sorted_word_df=pd.DataFrame.from_dict(sorted_word_count,orient="index",columns=["Frequency"])
                      
sorted_word_name=Today+"简频"+sku
sorted_word_df.to_excel(writer,sheet_name=sorted_word_name)


outputcountnew=output.loc[(output["周搜索量"].notnull())&output["周搜索量"].notna(),["关键词","周搜索量"]]

outputcountnew=outputcountnew.drop_duplicates()

# 拆分关键词列为单词，并将它们与相应的搜索量一起放入一个新的 DataFrame 中
word_search_data = []

for index, row in outputcountnew.iterrows():
    words = row['关键词'].split()
    search_volume = row['周搜索量']
    
    for word in words:
        word_search_data.append([word, search_volume])

words_df = pd.DataFrame(word_search_data, columns=['单词', '周搜索量'])

# 使用 groupby 函数对单词进行分组，然后使用 sum 函数计算每个分组的搜索量总和
word_count = words_df.groupby('单词')['周搜索量'].sum().reset_index()

# 按周搜索量降序排序
word_count = word_count.sort_values(by='周搜索量', ascending=False)
word_count_name=Today+"权频"+sku
print(word_count)
word_count.to_excel(writer,sheet_name=word_count_name,index=False)


shutil.move(File_FanchaAsinTerms, History_path+Top10File)
shutil.move(File_selfAsinTerms, History_path+SelfAsinFile)

writer.close()
