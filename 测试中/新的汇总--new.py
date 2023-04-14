

import sqlite3
import pandas as pd
from datetime import datetime, timedelta
from datetime import datetime, date
import numpy as np


#获取数据库
conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')

today = date.today()
last_saturday = today - timedelta(days=(today.weekday() + 2) % 7)
print(last_saturday)

# 计算27周前的日期
weeks_ago_27 = last_saturday - timedelta(weeks=100)
print(weeks_ago_27)

# 修改查询，添加日期条件
df = pd.read_sql_query(f'SELECT * FROM "Bulkfiles" WHERE 日期 >= "{weeks_ago_27}"', conn)


df = df[df['日期'].notna()]

df=df.drop_duplicates()



df['Spend'] = df['Spend'].astype(float)
df["Max Bid"] = df["Max Bid"].astype(float)
df['Sales'] = df['Sales'].astype(float)

df['日期'] = pd.to_datetime(df['日期'])


def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday

def update_week_numbers(df):
    last_saturday = find_last_saturday()
    df['日期'] = pd.to_datetime(df['日期'])
    df['周数'] = ((last_saturday - df['日期']).dt.days // 7) + 1
    return df

updated_df = update_week_numbers(df)

updated_df= updated_df[updated_df["周数"]<100]

#df['周数'] = ((df['日期'] - latest_date) / np.timedelta64(1, 'W')).astype(int) + 1  #上次的写法
updated_df = updated_df.drop(["Campaign Status", "Ad Group Status", "Status"], axis=1)

pivot_df = updated_df.groupby(["Country", "Campaign", "Ad Group", "Keyword or Product Targeting",
                       "Match Type"]).agg({
                           "Impressions": 'sum',
                           'Clicks': 'sum',
                           'Spend': 'sum',
                           'Orders': 'sum',
                           "Total Units": 'sum',
                           'Sales': 'sum'
                       }).reset_index()

pivot_df['转化率'] = pivot_df['Orders'] / pivot_df['Clicks']
pivot_df['点击率'] = pivot_df['Clicks'] / pivot_df['Impressions']

pivot_df['标签'] = '无'
pivot_df.loc[(pivot_df['Clicks'] >= 10) & (pivot_df['转化率'] >=0.2),"标签"] = '好targeting'
pivot_df.loc[(pivot_df['Clicks'] >= 10) & ((pivot_df['转化率'] >= 0.1)&(pivot_df['转化率']< 0.2)), "标签"]   = '可用Targeting' 
pivot_df.loc[(pivot_df['Clicks'] >=20) & ((pivot_df['转化率'] >= 0.05)&(pivot_df['转化率']< 0.1)), "标签"]   = '差Targeting-挑选'             
pivot_df.loc[(pivot_df['Clicks'] >=20) & (pivot_df['转化率'] < 0.05), '标签'] = '差Targeting-淘汰'



 

Allbulkpath='D:\\运营\\2生成过程表\\'

pivot_df.to_excel(Allbulkpath+'周bulk数据testSummary0.xlsx',index=False)
# 按 "Campaign"，"Ad Group" 和 "SKU" 对 "Spend" 进行汇总
spend_summary = updated_df.groupby(["Country","Campaign", "Ad Group", "SKU"]).agg({"Spend": "sum"}).reset_index()

# 为每个 "Campaign" 和 "Ad Group" 找到具有最大 "Spend" 的 SKU
spend_summary = spend_summary.loc[spend_summary.groupby(["Country","Campaign", "Ad Group"])["Spend"].idxmax()]

# 将结果重命名为 "主要SKU"
spend_summary = spend_summary.rename(columns={"SKU": "主要SKU"})

# 将结果合并到原始数据集，创建一个新列 "主要SKU"
pivot_df = pivot_df.merge(spend_summary[["Country","Campaign", "Ad Group", "主要SKU"]], on=["Country","Campaign", "Ad Group"], how="left")

print(pivot_df)
Allbulkpath='D:\\运营\\2生成过程表\\'

pivot_df.to_excel(Allbulkpath+'周bulk数据testSummary1.xlsx',index=False)


Allbulk=updated_df
AllbulkSKU_Campaign=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend','Orders','Clicks']].agg('sum').reset_index()                          
AllbulkSKU_Campaign["Campaign-SKU_Spend_ranking"]=AllbulkSKU_Campaign.groupby(["Country","Campaign","Ad Group"],as_index=False)[['Spend']].rank(ascending=0,method='max')
print(AllbulkSKU_Campaign)
AllbulkSKU_Campaign_rank1=AllbulkSKU_Campaign.loc[AllbulkSKU_Campaign["Campaign-SKU_Spend_ranking"]==1,["Country","Campaign","Ad Group","SKU","Campaign-SKU_Spend_ranking"]]
print(AllbulkSKU_Campaign_rank1)                          
pivot_df=pivot_df.merge(AllbulkSKU_Campaign_rank1, on=["Country","Campaign", "Ad Group"], how="left")



 
pivot_df.to_excel(Allbulkpath+'周bulk数据testSummary2.xlsx',index=False)



# 筛选符合条件的行
#pivot_df_filtered = pivot_df[(pivot_df['Clicks'] > 10) & (pivot_df['Campaign Status'] == 'enabled') & 
                 #(pivot_df['Ad Group Status'] == 'enabled') & (pivot_df['Status'] == 'enabled')]
pivot_df_filtered = pivot_df[pivot_df['Clicks'] > 9]
# 计算每个国家每个主要SKU的转化率，并按照转化率排序

#pivot_df_filtered['Conversion Rate'] = df_filtered['Orders'] / df_filtered['Clicks']
pivot_df_grouped = pivot_df_filtered.groupby(['Country', '主要SKU'], as_index=False).apply(lambda x: x.sort_values('转化率', ascending=False)).reset_index(drop=True)
print(pivot_df_grouped)
 
# 取出每个国家每个主要SKU转化率排名前5的行，生成一个新的 dataframe
pivot_df_top= pivot_df_grouped.groupby(['Country', '主要SKU']).head(10).reset_index(drop=True)
#pivot_df_top5 = pivot_df_grouped.groupby(['Country', '主要SKU']), '主要SKU']).head(5)
 
pivot_df_top.to_excel(Allbulkpath+'周bulk数据testSummary3.xlsx',index=False)

import os
#################################################################################################################################33

# 读取选词表格
keywords_df = pivot_df_top[["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type","转化率","主要SKU"]]


# 获取选词表格中的国家列表
countries = keywords_df['Country'].unique()

# 指定Bulk文件所在的目录

bulk_dir = r'D:\运营\1数据源\周Bulk广告数据'

# 遍历国家列表
for country in countries:
    # 查找以国家名开头的Bulk文件
    bulk_file = None
    for file in os.listdir(bulk_dir):
        if file.startswith(country + "_"):
            bulk_file = os.path.join(bulk_dir, file)
            break

    if bulk_file is None:
        print(f"No bulk file found for country {country}")
        continue

    # 读取当前国家的Bulk文件
    amazon_bulk_df = pd.read_excel(bulk_file,sheet_name="Sponsored Products Campaigns")

    # 筛选当前国家的选词表格数据
    country_keywords_df = keywords_df[keywords_df['Country'] == country].drop(columns=['Country'])
    
    
    # 为简化处理，我们将Bulk文件和选词表格的列名进行统一
    #country_keywords_df.columns = ["Campaign", "Ad Group", "Match Type", "Keyword or Product Targeting"]
    amazon_bulk_df["变更记录"] = ""


# 获取手动广告系列列表
    manual_campaigns = amazon_bulk_df.loc[amazon_bulk_df["Campaign Targeting Type"] == "Manual", ["Campaign", "Campaign Targeting Type"]].drop_duplicates()
    manual_campaigns.rename(columns={"Campaign Targeting Type": "CampaignTargetingType_New"}, inplace=True)

    amazon_bulk_df_merge_keywords = amazon_bulk_df.merge(country_keywords_df, on=["Campaign", "Ad Group", "Match Type", "Keyword or Product Targeting"], how="left")
    amazon_bulk_df_merge_keywords = amazon_bulk_df_merge_keywords.merge(manual_campaigns, on=["Campaign"], how="left")

    


          
    mask1 = amazon_bulk_df_merge_keywords.loc[
    (amazon_bulk_df_merge_keywords["CampaignTargetingType_New"] == "Manual") & (amazon_bulk_df_merge_keywords["主要SKU"].isnull()) &
    (amazon_bulk_df_merge_keywords["Record Type"] == "Keyword") & ~((amazon_bulk_df_merge_keywords["Match Type"].str.contains("negtive") | amazon_bulk_df_merge_keywords["Match Type"].str.contains("Negtive")))]
 
    print(mask1)
    
    mask1["Status"] = "paused"
    mask1["变更记录"] = "暂停所有其他非选词"
    mask1 = mask1.drop("主要SKU", axis=1)
    pivot_df_df_SKU=pivot_df[["Campaign", "Ad Group","主要SKU"]].drop_duplicates()
    mask1=mask1.merge(pivot_df_df_SKU,on=["Campaign", "Ad Group"],how="left")
    mask1.to_excel(r'D:\\运营\\'+country+"mask1.xlsx")

    # 从Bulk文件中筛选出需要保留的行
    filtered_bulk_df = amazon_bulk_df.merge(country_keywords_df, on=["Campaign", "Ad Group", "Match Type", "Keyword or Product Targeting"], how="inner")

# 将“Status”列设置为“enabled”
    filtered_bulk_df["Status"] = "enabled"
    filtered_bulk_df["变更记录"] = "启用选词表格中的广告"

    
    # 将选词表格中有但Bulk文件中没有的内容添加到filtered_bulk_df
    missing_rows = country_keywords_df.loc[~country_keywords_df.apply(tuple, 1).isin(filtered_bulk_df.apply(tuple, 1))].copy()
# 根据Campaign列和Ad Group列匹配到Keyword or Product Targeting和Match Type列
    missing_rows = missing_rows.merge(amazon_bulk_df[['Campaign', 'Ad Group', 'Campaign Status', 'Ad Group Status']], on=['Campaign', 'Ad Group'], how='left')

# 把这些行的Campaign Status列、Ad Group Status列和Status列的状态改为enabled
    missing_rows['Campaign Status'] = 'enabled'
    missing_rows['Ad Group Status'] = 'enabled'
    missing_rows['Status'] = 'enabled'
    missing_rows['变更记录'] = "添加选词表格中有但Bulk文件中没有的内容并启用"
# 将missing_rows添加到filtered_bulk_df
    merged_df = pd.concat([filtered_bulk_df, missing_rows], ignore_index=True, sort=False)

    merged_df=pd.concat([merged_df,mask1], ignore_index=True, sort=False)

    # 将结果保存到新的Excel文件
    output_file = os.path.join("updated_bulk_files"+"20230413", f"{country}_updated_bulk_file.xlsx")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    
    merged_df.to_excel(output_file, index=False)


    
conn.close()
