import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import numpy as np

conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')

df = pd.read_sql_query('SELECT * FROM "Bulkfiles"', conn)
df = df[df['日期'].notna()]

df['Spend'] = df['Spend'].astype(float)
df["Max Bid"] = df["Max Bid"].astype(float)
df['Sales'] = df['Sales'].astype(float)

df['日期'] = pd.to_datetime(df['日期'])

latest_date = df['日期'].max()

df['周数'] = ((df['日期'] - latest_date) / np.timedelta64(1, 'W')).astype(int) + 1

pivot_df = df.groupby(["Country", "Campaign", "Ad Group", "Keyword or Product Targeting",
                       "Match Type", "Campaign Status", "Ad Group Status", "Status"]).agg({
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
pivot_df.loc[((pivot_df['Clicks'] > 20) & (pivot_df['转化率'] > 0.2)) | ((pivot_df['Clicks'] >= 8) & (pivot_df['Clicks'] < 20) & (pivot_df['转化率'] > 0.25)), '标签'] = '好targeting'
pivot_df.loc[(pivot_df['Clicks'] > 20) & (pivot_df['转化率'] < 0.05), '标签'] = '差Targeting'





# 获取 "Record Type" 列值为 "Campaign" 的 [Campaign] 列值对应的 "Campaign Status" 列的值
campaign_statuses = df.loc[df['Record Type'] == 'Campaign', ['Campaign', 'Campaign Status']]

# 去重，以确保每个 Campaign 只有一个对应的 Campaign Status
campaign_statuses = campaign_statuses.drop_duplicates(subset='Campaign')

# 将结果合并到 [Campaign] 列具有相同 "Campaign" 值的所有行
pivot_df = pivot_df.merge(campaign_statuses, on='Campaign', suffixes=('', '_merged'))


# 按 "Campaign"，"Ad Group" 和 "SKU" 对 "Spend" 进行汇总
spend_summary = df.groupby(["Campaign", "Ad Group", "SKU"]).agg({"Spend": "sum"}).reset_index()

# 为每个 "Campaign" 和 "Ad Group" 找到具有最大 "Spend" 的 SKU
spend_summary = spend_summary.loc[spend_summary.groupby(["Campaign", "Ad Group"])["Spend"].idxmax()]

# 将结果重命名为 "主要SKU"
spend_summary = spend_summary.rename(columns={"SKU": "主要SKU"})

# 将结果合并到原始数据集，创建一个新列 "主要SKU"
pivot_df = pivot_df.merge(spend_summary[["Campaign", "Ad Group", "主要SKU"]], on=["Campaign", "Ad Group"], how="left")



pivot_df.to_excel('D:\\运营\\output_summary.xlsx', index=False)

conn.close()
