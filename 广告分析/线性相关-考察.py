import sqlite3
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeRegressor
from sklearn.metrics import mean_squared_error
from sklearn.preprocessing import OneHotEncoder
import datetime  
#??Asin 周数 price   weekorders

def find_last_saturday():
    today = datetime.datetime.now()
    last_saturday = today - datetime.timedelta(days=today.weekday() + 2)
    return last_saturday

def update_week_numbers(df):
    last_saturday = find_last_saturday()
    df['日期'] = pd.to_datetime(df['日期'])
    df['周数'] = ((last_saturday - df['日期']).dt.days // 7) + 1
    return df

# 准备数据

country="GV-US"
sku="20200512-Deskorganizer-6CWhite"
Dataweek=100

#  获取数据
Sales_df=pd.read_excel(r'D:\2019plan\\周销售数据.xlsx')
Sales_df = update_week_numbers(Sales_df)
Sales_df=Sales_df.loc[Sales_df["周数"]<=Dataweek,["SKU","Units Ordered","Total Order Items","Ordered Product Sales","周数","日期"]]
Sales_df["Total Order Items"].fillna(0,inplace=True)
print(Sales_df)

conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')


query = f'SELECT * FROM "Bulkfiles" WHERE Country = "{country}"'

AdData = pd.read_sql_query(query, conn)

#赋予周数

AdData = AdData[AdData['日期'].notna()]





AdData['Spend'] =AdData['Spend'].astype(float)
AdData["Max Bid"] = AdData["Max Bid"].astype(float)
AdData["Orders"] = AdData["Orders"].astype(float)
AdData["Total Units"] = AdData["Total Units"].astype(float)
AdData['Sales'] = AdData['Sales'].astype(float)

AdData['日期'] = pd.to_datetime(AdData['日期'])



AdData = update_week_numbers(AdData)

AdData=AdData[AdData["周数"]<=Dataweek]

#  销售数据请理:得到SKU周数和销量
Sales_df=Sales_df.dropna()
# 选出SKU的数据
mask = Sales_df['SKU'].astype(str).str.contains(sku, case=False)
Sales_df_sku = Sales_df[mask]




Sales_df_taget=Sales_df_sku[["SKU","Total Order Items","周数"]]

Sales_df_taget["Total Order Items"].astype(float)
Sales_df_taget["周数"].astype(int)
#  广告数据请理 得到SKU 周数 关键词Targeting 广告状态 竞价 Match Type 展示量 点击量 广告订单量 参数是否可调

# 把camaign和SKU对应。
file_pathSummary = r'D:\运营\2生成过程表\周bulk数据Summary.xlsx'
sheet_name="SKU-Campaign-Spend"

columnsSummary = ['Country', 'Campaign', 'SKU', 'Spend', 'Orders', 'Clicks', 'zhuanhualv', 'SKU-Campaign-zhuanhualv-ranking', 'Campaign-SKU_Spend_ranking', 'SKU_Campaign_Spend_ranking']


dfSummary = pd.read_excel(file_pathSummary, sheet_name=sheet_name, usecols=columnsSummary)


CampaigntoSKU = dfSummary.loc[dfSummary['Campaign-SKU_Spend_ranking'] == 1]
CampaigntoSKUBAoliu=CampaigntoSKU.loc[CampaigntoSKU['Country']==country,['Country', 'Campaign', 'SKU']]

CampaigntoSKUBAoliu.rename(columns={"SKU": "SKU3"}, inplace=True)

# 把camaign和SKU对应完毕

AdData=pd.merge(AdData,CampaigntoSKUBAoliu,on=['Country', 'Campaign'],how='left')


AdData = AdData.loc[(AdData["SKU3"] == sku) & AdData["Keyword or Product Targeting"].notnull(), ["SKU3","Campaign","Max Bid", "Keyword or Product Targeting",
                                                                                                                                         "Match Type", "Campaign Status", "Ad Group Status", "Status",
                                                                                                                                         "Impressions", "Clicks", "Spend", "Orders", "Total Units","周数"]]

AdData["Max Bid"].fillna(0,inplace=True)
AdData["Ad Group Status"].fillna(1,inplace=True)
AdData["Campaign Status"].replace({"paused":0,"enabled":1},inplace=True) 
AdData["Ad Group Status"].replace({"paused":0,"enabled":1},inplace=True) 
AdData["Status"].replace({"paused":0,"enabled":1},inplace=True) 


Alldata=pd.merge(AdData,Sales_df_taget, left_on=["SKU3","周数"],right_on=["SKU","周数"],how="left")
Alldata.to_excel(r'D:\\运营\\'+sku+'ceshi.xlsx')
has_missing_values = Alldata.isna().any().any()

print("是否有缺失值？", has_missing_values)
          

grouped_data_campaign= Alldata.groupby(['Campaign', '周数']).agg({'Spend': 'sum', 'Total Order Items': 'sum'}).reset_index()
campaigns = grouped_data_campaign['Campaign'].unique()
correlations_campaign = {}

for campaign in campaigns:
    print(campaign)
    campaign_data = grouped_data_campaign[grouped_data_campaign['Campaign'] == campaign]
    print(campaign_data)
    correlation_campaign99 = campaign_data['Spend'].corr(campaign_data['Total Order Items'])
    correlations_campaign[campaign] = correlation_campaign99
     
correlations_df_campaign = pd.Series(correlations_campaign, name='Correlation').reset_index()
print(correlations_df_campaign)
correlations_df_campaign.columns = ['Campaign', 'Correlation']


grouped_data_keyword = Alldata.groupby(['Keyword or Product Targeting', '周数']).agg({'Clicks': 'sum', 'Total Order Items': 'sum'}).reset_index()


keywords = grouped_data_keyword['Keyword or Product Targeting'].unique()
correlations_keyword = {}

for keyword in keywords:
    keyword_data = grouped_data_keyword[grouped_data_keyword['Keyword or Product Targeting'] == keyword]
    
    # 检查数据量，确保每个分组至少有两个样本
    if len(keyword_data) >= 2:
        correlation_keyword = keyword_data['Clicks'].corr(keyword_data['Total Order Items'])
        correlations_keyword[keyword] = correlation_keyword
    else:
        correlations_keyword[keyword] = None  # 如果数据量过小，则将相关性设为 None

correlations_df_keyword = pd.Series(correlations_keyword, name='Correlation').reset_index()
correlations_df_keyword.columns = ['Keyword or Product Targeting', 'Correlation']


writer=pd.ExcelWriter(r'D:\\运营\\相关与机器学习.xlsx')


correlations_df_campaign.to_excel(writer,"correlations_df_campaign")


correlations_df_keyword.to_excel(writer,"correlations_df_keyword")

writer.close()
