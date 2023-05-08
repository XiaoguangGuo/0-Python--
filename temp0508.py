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

AllbulkCampaignSKUWEEK_allinfo=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign","Ad Group","周数","Campaign Status","Ad Group Status","Status"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
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
AllbulkCampaignKeywordWEEK_allinfo=Allbulk[Allbulk["周数"]<16].groupby(["Country","Campaign","Keyword or Product Targeting","Ad Group","Match Type","Campaign Status","Ad Group Status","Status","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")



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
AllbulkCampaignSKUWEEK_allinfo.to_excel(writer,"AllbulkCampaignSKUWEEK_allinfo",index=False)

AllbulkCampaignKeywordWEEK.to_excel(writer,"Keyword-Campaign-WEEK")
AllbulkCampaignKeywordWEEK_allinfo.to_excel(writer,"Keyword-Campaign-WEEK-allinfo")
AllbulkSKU_Campaign.to_excel(writer,"SKU-Campaign-Spend")
AllbulkSKU_Campaignrank1.to_excel(writer,"SKUMax-Campaign")
AllbulkSKUMax_CampaignWEEK.to_excel(writer,"SKUMax-Campaign-WEEK") 
AllbulkCampaignSKUTotal.to_excel(writer,"AllSKUCampaign")
CamaignSKUAgg.to_excel(writer,"CamaignSKUAgg")#追加的新的汇总comaignSKU
AllbulkCampaignSKUTotalzhuanhualvMax.to_excel(writer,"CampaignSKUTotalzhuanhualvMax",index=False)

writer.close()


###############BIAOTOU汇总################################################################3333333


CampaignSKU_Summary=pd.read_excel(r'D:/运营/2生成过程表/周bulk数据Summary.xlsx',sheet_name="SKU-Campaign-WEEK")

CampaignSKU_SummarySum=CampaignSKU_Summary.groupby(["Country","SKU","Campaign"],as_index=False)[["Impressions","Clicks","Spend","Orders","Total Units","Sales"]].agg('sum')



CampaignSKU_Summary["皮质层标签"]=" "

CampaignSKU_Summary["Zhouzhuanlv"]=CampaignSKU_Summary["Orders"]/CampaignSKU_Summary["Clicks"]

CampaignSKU_SummarySum["Zhouzhuanlv"]=CampaignSKU_SummarySum["Orders"]/CampaignSKU_SummarySum["Clicks"]


CampaignSKU_Summary.loc[(CampaignSKU_Summary["Clicks"]>0) &(CampaignSKU_Summary["Zhouzhuanlv"]>0.15),"皮质层标签"] = CampaignSKU_Summary["皮质层标签"].astype(str)+"好广告"


CampaignSKU_Summary.loc[(CampaignSKU_Summary["Clicks"]>0) &(CampaignSKU_Summary["Zhouzhuanlv"]<0.05),"皮质层标签"] = CampaignSKU_Summary["皮质层标签"].astype(str)+"差广告"

#CampaignSKU_Summary10=CampaignSKU_Summary.loc[(CampaignSKU_Summary["周数"]<5)&(CampaignSKU_Summary["Country"]=="GV-US")]

CampaignSKU_Summary_biaotou=CampaignSKU_Summary[["Country","SKU","Campaign"]].drop_duplicates()
CampaignSKU_Summary_biaotou=pd.merge(CampaignSKU_Summary_biaotou,CampaignSKU_SummarySum,on=["Country","SKU","Campaign"] ,how="left")
print(CampaignSKU_Summary_biaotou)

for i in range(1,20):
    #CampaignSKU_Summary_i=CampaignSKU_Summary["Clicks","Orders"].loc[(CampaignSKU_Summary["周数"]==i)]
    CampaignSKU_Summary_i=CampaignSKU_Summary.loc[(CampaignSKU_Summary["周数"]==i)]
    
    #CampaignSKU_Summary_i=CampaignSKU_Summary_i["Country","SKU","Campaign","Clicks","Orders"]
    #更改列名

    CampaignSKU_Summary_i.rename(columns = {'Clicks':'Clicks'+str(i), 'Orders':'Orders'+str(i),'Spend':'Spend'+str(i),'Impressions':'Impressions'+str(i)}, inplace = True)

    CampaignSKU_Summary_biaotou=pd.merge(CampaignSKU_Summary_biaotou,CampaignSKU_Summary_i,on=["Country","SKU","Campaign"] ,how="left")
    




#CampaignSKU_Summary_pivot10=CampaignSKU_Summary10.pivot_table(values=["Clicks","Orders"], index=['Country','SKU','Campaign'],columns="周数", aggfunc = 'sum', fill_value=None, margins=False, dropna=False,margins_name='All').reset_index() # 是否启用总计行/列# 值

print(CampaignSKU_Summary_biaotou)


CampaignSKU_Summary_biaotou.to_excel(r'D:\\运营\\2生成过程表\\CampaignSKU_Summary_biaotou.xlsx',sheet_name="sheet1",startrow=0,header=True,index=True)













