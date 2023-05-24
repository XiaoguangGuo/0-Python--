


# -*- coding:utf-8 –*-

import pandas as pd
import os

import datetime 

import numpy as np

Campaign_SKU=pd.read_excel(r'D:\\运营\\2生成过程表\\周Bulk数据Summary.xlsx',sheet_name="SKUMax-Campaign")

SearchTermAll=pd.read_excel(r'D:\\运营\\2生成过程表\\Sponsored Products Search term report.xlsx')

SearchTermAll["Clicks"].fillna(0,inplace=True)


SearchTermAll["Customer Search Term"].astype(str)

All_Campaign_SearchTerm=SearchTermAll.groupby(["Country","Campaign Name", "Customer Search Term"],as_index=False)[["Impressions","Clicks","Spend","7 Day Total Sales ","7 Day Total Orders (#)"]].agg("sum")

#AllbulkCampaignKeywordWEEK=Allbulk.groupby(["Country","Campaign","Keyword or Product Targeting","Ad Group","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")



All_Campaign_SearchTerm.loc[All_Campaign_SearchTerm['Clicks']>0,"转化率"]=All_Campaign_SearchTerm["7 Day Total Orders (#)"]/All_Campaign_SearchTerm['Clicks']


All_Campaign_SearchTerm.loc[(All_Campaign_SearchTerm["转化率"]>=0.25)&(All_Campaign_SearchTerm["Clicks"]>=5),"转化率好坏"]="好词"


All_Campaign_SearchTerm.loc[(All_Campaign_SearchTerm["转化率"]<0.05)&(All_Campaign_SearchTerm["Clicks"]>=15),"转化率好坏"]="差词"


All_Campaign_SearchTerm=pd.merge(All_Campaign_SearchTerm,Campaign_SKU,how="left",left_on=["Country","Campaign Name"],right_on=["Country","Campaign"])

GoodWord=All_Campaign_SearchTerm[All_Campaign_SearchTerm["转化率好坏"]=="好词"].groupby(["Country","SKU"],as_index=False)["转化率好坏"].agg("count")

SeachTermWeekSum=SearchTermAll.groupby(["Country","Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term","周数"],as_index=False)[["Impressions","Clicks","Spend","7 Day Total Sales ","7 Day Total Orders (#)"]].agg("sum")
SeachTermWeekSumTotal=SearchTermAll.groupby(["Country","Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term",],as_index=False)[["Impressions","Clicks","Spend","7 Day Total Sales ","7 Day Total Orders (#)"]].agg("sum")

SeachTermWeekSumTotal["zhuanhualv"]=SeachTermWeekSumTotal["7 Day Total Orders (#)"]/SeachTermWeekSumTotal["Clicks"]

SeachTermWeekSum["zhuanhualv"]=SeachTermWeekSum["7 Day Total Orders (#)"]/SeachTermWeekSum["Clicks"]

SeachTermWeekSum_Biaotou=SeachTermWeekSum[["Country","Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term"]].drop_duplicates()
SeachTermWeekSum_Biaotou=pd.merge(SeachTermWeekSum_Biaotou,SeachTermWeekSumTotal,on=["Country","Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term"] ,how="left")
    
max_week=SearchTermAll["周数"].max()

for i in range(1,max_week):
    #CampaignSKU_Summary_i=CampaignSKU_Summary["Clicks","Orders"].loc[(CampaignSKU_Summary["周数"]==i)]
    SeachTermWeekSum_i=SeachTermWeekSum.loc[(SeachTermWeekSum["周数"]==i)]

     

    SeachTermWeekSum_i=SeachTermWeekSum_i[["Country","Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term","Clicks","7 Day Total Orders (#)","zhuanhualv"]]
    SeachTermWeekSum_i.rename(columns = {"Clicks":'Clicks'+str(i),"7 Day Total Orders (#)":'Orders'+str(i),  "zhuanhualv":'zhuanhualv'+str(i)}, inplace = True)
     
    #合并

    SeachTermWeekSum_Biaotou=pd.merge(SeachTermWeekSum_Biaotou,SeachTermWeekSum_i,on=["Country","Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term"] ,how="left")
    



writer=pd.ExcelWriter(r'D:\\运营\\2生成过程表\\Search_Term_Summary.xlsx')

All_Campaign_SearchTerm.to_excel(writer,"All_Campaign_SearchTerm",index=False)
GoodWord.to_excel(writer,"GoodWord",index=False)
SeachTermWeekSum_Biaotou.to_excel(writer,"SeachTermWeekSum_Weeks")
writer.close()
