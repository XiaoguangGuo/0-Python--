import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import requests

def fetch_exchange_rates(app_id, currencies):
    api_url = 'https://openexchangerates.org/api/latest.json?app_id={}&symbols={}'
    rates = {}

    try:
        response = requests.get(api_url.format(app_id, ','.join(currencies)), headers={'Authorization': f'Token {app_id}'})
        response.raise_for_status()
        data = response.json()
        for currency in currencies:
            rate = data['rates'][currency]
            rates[currency] = rate
        print('连接API成功！以下是货币对美元的汇率：')
        print(rates)
    except requests.exceptions.HTTPError as error:
        print(f'连接API出错：{error}')
        
    return rates



if __name__ == '__main__':
    app_id = '438a43ad7170441aa0c7a00caebf086f'
    currencies = ['USD', 'CAD', 'EUR', 'GBP', 'JPY', 'MXN', 'SEK']
    exchange_rates = fetch_exchange_rates(app_id, currencies)


exchangerate_dic={"GV-US":"USD","GV-CA":"CAD","NEW-UK":"GBP","NEW-JP":"JPY","NEW-CA":"CAD","NEW-IT":"EUR","NEW-DE":"EUR","NEW-ES":"EUR","NEW-FR":"EUR","NEW-US":"USD","HM-US":"USD","GV-MX":"MXN","NEW-MX":"MXN"}

conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')

df = pd.read_sql_query('SELECT * FROM "Bulkfiles"', conn)
df = df[df['日期'].notna()]
df["Country"].replace("GX-MX","GV-MX",inplace=True)




df['Spend'] = df['Spend'].astype(float)
def convert_to_usd(row):
    country = row['Country']
    spend = row['Spend']
    currency = exchangerate_dic[country]
    exchange_rate = exchange_rates[currency]
    return spend / exchange_rate

df['Spend_USD'] = df.apply(convert_to_usd, axis=1)

print(df['Spend_USD'].head(5))


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

 

#df['周数'] = ((df['日期'] - latest_date) / np.timedelta64(1, 'W')).astype(int) + 1  #上次的写法

pivot_df = updated_df.groupby(["Country", "Campaign", "Ad Group", "Keyword or Product Targeting",
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

Allbulkpath='D:\\运营\\2生成过程表\\'

writer=pd.ExcelWriter(Allbulkpath+'周bulk数据Summary.xlsx')
 
pivot_df.to_excel(writer,"output_summary",index=False) 

conn.close()


############################################继续做summary####################################
print("以下生成Summary的程序")
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil

#在这之前要生成汇总表，并且把每个国家的bulk表备份到D:\\运营\\bulkoperationfiles\\




#定义bulk数据汇总表所在路径

Allbulk=updated_df[updated_df["周数"]<27]

 


AllbulkCountriesEffects=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","周数"],as_index=False)[['Impressions','Clicks','Spend_USD','Orders','Total Units','Sales']].agg('sum')

AllbulkCountriesEffectsCountries=AllbulkCountriesEffects["Country"].drop_duplicates()
for i in range(1,20):
    AllbulkCountriesEffects_i=AllbulkCountriesEffects[AllbulkCountriesEffects["周数"]==i]
    AllbulkCountriesEffectsCountries= pd.merge(AllbulkCountriesEffectsCountries,AllbulkCountriesEffects_i,on=["Country"] ,how="left")



AllbulkCampaign=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign","Campaign Targeting Type"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaign["zhuanhualv"]=AllbulkCampaign['Orders']/AllbulkCampaign['Clicks']
AllbulkCampaign_List=AllbulkCampaign["Campaign"].drop_duplicates().to_list()
AllbulkCampaign_Country_List=AllbulkCampaign["Country"].drop_duplicates().to_list()




#AllbulkCampaign1week=Allbulk[(Allbulk['Record Type']=="Campaign")&(Allbulk['周数']==1)].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')


AllbulkCampaignWEEK=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')

AllbulkSKUCampaignWEEK_1=Allbulk[(Allbulk['Record Type']=="Ad")&(Allbulk['周数']==1)].groupby(["Country","Campaign","SKU","Ad Group"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkSKUCampaignWEEK_1["zhuanhualv"]=AllbulkSKUCampaignWEEK_1["Orders"]/AllbulkSKUCampaignWEEK_1["Clicks"]

AllbulkCampaignSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign","周数","Campaign Status"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')



AllbulkCampaignSKUTotal=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign","Ad Group"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignSKUWEEK["Campaign Targeting Type"]=""
AllbulkCampaignSKUTotal["Campaign Targeting Type"]=""

###########################################################################################################

for AllbulkCampaign_Country_oi99 in AllbulkCampaign_Country_List:
    AllbulkCampaign_Country99=AllbulkCampaign.loc[AllbulkCampaign["Country"]==AllbulkCampaign_Country_oi99]
    AllbulkCampaign_Country99_CampaignList= AllbulkCampaign_Country99.loc[AllbulkCampaign_Country99["Country"]==AllbulkCampaign_Country_oi99,"Campaign"].drop_duplicates().to_list()
    
    for campaign_oi99 in AllbulkCampaign_Country99_CampaignList:
        campaigntype99=AllbulkCampaign_Country99.loc[(AllbulkCampaign_Country99["Country"]==AllbulkCampaign_Country_oi99)&(AllbulkCampaign_Country99["Campaign"]==campaign_oi99),"Campaign Targeting Type"].values[0] 
        AllbulkCampaignSKUTotal.loc[(AllbulkCampaignSKUTotal["Campaign"]==campaign_oi99)&(AllbulkCampaignSKUTotal["Country"]==AllbulkCampaign_Country_oi99),"Campaign Targeting Type"]=campaigntype99
############################################################################################################################campaigntype99


###########################################################################################################
 


AllbulkCampaignSKUTotal["zhuanhualv"]=AllbulkCampaignSKUTotal["Orders"]/AllbulkCampaignSKUTotal["Clicks"]

AllbulkCampaignSKUTotal["zhuanhualv_rank1"]=AllbulkCampaignSKUTotal.groupby(["Country","SKU","Campaign Targeting Type"],as_index=False)[["zhuanhualv"]].rank(ascending=0,method='max')
AllbulkCampaignSKUTotal["zhuanhualv_rank2"]=AllbulkCampaignSKUTotal.groupby(["Country","SKU","Campaign Targeting Type"],as_index=False)[["zhuanhualv"]].rank(ascending=0,method='dense')                                                                                       

AllbulkCampaignSKUTotalzhuanhualvMax=AllbulkCampaignSKUTotal.groupby(["Country","Campaign","Ad Group","SKU","Campaign Targeting Type"],as_index=False)[["zhuanhualv"]].agg('max')

AllbulkCampaignSKUWEEK["Campaign Targeting Type"]=""

########################################################################################################################################33                                               
for AllbulkCampaign_Country_oi in AllbulkCampaign_Country_List:
    
    AllbulkCampaign_Country88=AllbulkCampaign.loc[AllbulkCampaign["Country"]==AllbulkCampaign_Country_oi]
    AllbulkCampaign_Country_CampaignList88= AllbulkCampaign_Country88.loc[AllbulkCampaign_Country88["Country"]==AllbulkCampaign_Country_oi99,"Campaign"].drop_duplicates().to_list()

    
    for campaign_oi in AllbulkCampaign_Country_CampaignList88:
        campaigntype88=AllbulkCampaign_Country88.loc[(AllbulkCampaign_Country88["Country"]==AllbulkCampaign_Country_oi)&(AllbulkCampaign_Country88["Campaign"]==campaign_oi),"Campaign Targeting Type"].values[0]
        AllbulkCampaignSKUWEEK.loc[(AllbulkCampaignSKUWEEK["Campaign"]==campaign_oi)&(AllbulkCampaignSKUWEEK["Country"]==AllbulkCampaign_Country_oi),"Campaign Targeting Type"]=campaigntype88





                                               
AllbulkCampaignSKUWEEK["zhuanhualv"]=AllbulkCampaignSKUWEEK["Orders"]/AllbulkCampaignSKUWEEK["Clicks"]

AllbulkSKUMax=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].max()


                                                                              
AllbulkCampaignKeywordWEEK=Allbulk.groupby(["Country","Campaign","Keyword or Product Targeting","Ad Group","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")




#以下为建立Camaign和SKU联系的程序
ALLbulkCampaignSKU=Allbulk[['Country','Campaign','SKU']]

ALLbulkCampaignSKU=ALLbulkCampaignSKU.drop_duplicates()
ALLbulkCampaignSKU=ALLbulkCampaignSKU.dropna(axis=0,how='any')
print(ALLbulkCampaignSKU)

CamaignSKUAgg=ALLbulkCampaignSKU.groupby(["Country","Campaign"],as_index=False).agg({'SKU':[",".join]})#追加的新的汇总comaignSKU
New_columns=['Country',"Campaign",'SKU']


AllbulkSKUMax_CampaignWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","SKU","周数","Campaign Status"],as_index=False)[['Spend']].agg('max')
AllbulkSKU_Campaign=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","SKU"],as_index=False)[['Spend','Orders','Clicks']].agg('sum').reset_index()
AllbulkSKU_Campaign["zhuanhualv"]=AllbulkSKU_Campaign["Orders"]/AllbulkSKU_Campaign["Clicks"]
AllbulkSKU_Campaign["SKU-Campaign-zhuanhualv-ranking"]=AllbulkSKU_Campaign.groupby(["Country","SKU"],as_index=False)[['zhuanhualv']].rank(ascending=0,method='max')

AllbulkSKU_Campaign["Campaign-SKU_Spend_ranking"]=AllbulkSKU_Campaign.groupby(["Country","Campaign"],as_index=False)[['Spend']].rank(ascending=0,method='max')
AllbulkSKU_Campaign["SKU_Campaign_Spend_ranking"]=AllbulkSKU_Campaign.groupby(["Country","SKU"],as_index=False)[['Spend']].rank(ascending=0,method='max')

AllbulkSKU_Campaignrank1=AllbulkSKU_Campaign[AllbulkSKU_Campaign["Campaign-SKU_Spend_ranking"]==1]
AllbulkCampaignSKUWEEK["Spend_Order"]=AllbulkCampaignSKUWEEK.groupby(["Country","Campaign","SKU","周数","Campaign Status","Campaign Targeting Type"],as_index=False)[['Spend']].rank(ascending=0,method='max')
AllbulkCampaignSKUWEEK["Click_Order"]=AllbulkCampaignSKUWEEK.groupby(["Country","Campaign","SKU","周数","Campaign Status","Campaign Targeting Type"],as_index=False)[['Clicks']].rank(ascending=0,method='max')
#SKUCampaign_zhouzhuqanlv_Max

                                                                              
CamaignSKUAgg.columns=New_columns




AllbulkSKUCampaignWEEK_1.to_excel(writer,"SKUCampaignWEEK_1")
AllbulkCampaign.to_excel(writer,"Campaign汇总")
AllbulkCampaignWEEK.to_excel(writer,"CampaignWEEK汇总")
AllbulkSKUWEEK.to_excel(writer,"SKU-WEEK")
AllbulkCampaignSKUWEEK.to_excel(writer,"SKU-Campaign-WEEK",index=False)
AllbulkCampaignKeywordWEEK.to_excel(writer,"Keyword-Campaign-WEEK")
AllbulkSKU_Campaign.to_excel(writer,"SKU-Campaign-Spend")
AllbulkSKU_Campaignrank1.to_excel(writer,"SKUMax-Campaign")
AllbulkSKUMax_CampaignWEEK.to_excel(writer,"SKUMax-Campaign-WEEK") 
AllbulkCampaignSKUTotal.to_excel(writer,"AllSKUCampaign")
CamaignSKUAgg.to_excel(writer,"CamaignSKUAgg")#追加的新的汇总comaignSKU
AllbulkCampaignSKUTotalzhuanhualvMax.to_excel(writer,"CampaignSKUTotalzhuanhualvMax",index=False)
AllbulkCountriesEffectsCountries.to_excel(writer,"CountriesEffects2",index=False)

                                  
writer.close()
