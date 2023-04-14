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
Dataweek=26

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
 
Alldata.to_excel(r'D:\\运营\\ceshi.xlsx')
has_missing_values = Alldata.isna().any().any()

print("是否有缺失值？", has_missing_values)
          






# 独热编码
#encoder = OneHotEncoder()
#sku3_encoded = encoder.fit_transform(Alldata[["SKU3"]]).toarray()
#KeywordorProductTargeting_encoded=encoder.fit_transform(Alldata[["Keyword or Product Targeting"]]).toarray()
#MatchType_encoded=encoder.fit_transform(Alldata[["Match Type"]]).toarray()

# 整合独热编码后的数据
#data_encoded = pd.concat([pd.DataFrame(sku3_encoded), pd.DataFrame(KeywordorProductTargeting_encoded),pd.DataFrame(MatchType_encoded),Alldata[["Max Bid", "Campaign Status",
                                                                                                                                               #"Ad Group Status","Status","Impressions","Clicks",
                                                                                                                                               #"Spend","Orders","Total Units","周数",



data_encoded =  Alldata[["Max Bid", "Campaign Status","Ad Group Status","Status","Impressions","Clicks","Spend","Orders","Total Units","周数","Total Order Items"]]                                                                                                                          

data_encoded.loc[data_encoded["Clicks"]==0, "转化率"]=0
data_encoded.loc[data_encoded["Clicks"]>0, "转化率"]=data_encoded["Orders"]/data_encoded["Clicks"]

 
data_encoded=data_encoded.drop(["周数","Max Bid","Total Units","Spend","Impressions", "Campaign Status","Ad Group Status","Orders"],axis=1)




                                                                                                                                                
# 划分数据集
X = data_encoded.drop("Total Order Items", axis=1)
y = data_encoded["Total Order Items"]
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=42)
X_train.columns = X_train.columns.astype(str)
print(X_train.columns)
X_test.columns = X_test.columns.astype(str)

# 创建并训练模型
dt = DecisionTreeRegressor(max_depth=4,random_state=42)
dt.fit(X_train, y_train)

# 预测
y_pred = dt.predict(X_test)

# 评估模型性能
mse = mean_squared_error(y_test, y_pred)
print("Mean Squared Error:", mse)

import matplotlib.pyplot as plt
from sklearn.tree import plot_tree

plt.figure(figsize=(30, 15))

feature_names = data_encoded.columns[:-1].to_list()  # 获取除了最后一列（目标值）之外的所有列名


plot_tree(dt, filled=True, feature_names=feature_names, rounded=True)

plt.show()



from sklearn.tree import export_text

rules = export_text(dt, feature_names=feature_names )
print(rules)


# 设置不同参数评估模型性能

#from sklearn.tree import DecisionTreeRegressor

#dt = DecisionTreeRegressor(
    #max_depth=5,
    #min_samples_split=10,
    #min_samples_leaf=5,
    #max_features='auto',
    #criterion='mse',
    #splitter='best'
#)
#dt.fit(X_train, y_train)




