

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime


CountrDic={"加拿大":"NEW-CA","美国":"NEW-US","英国":"NEW-UK","意大利":"NEW-IT","德国":"NEW-DE","法国":"NEW-FR","西班牙":"NEW-ES","日本":"NEW-JP","墨西哥":"NEW-MX"}

exchangerate_20221217={"GV-US":1,"GV-CA":1.3701,"NEW-UK":0.8223,"NEW-JP":136.6790,"NEW-CA":1.3701,"NEW-IT":0.9457,"NEW-DE":0.9457,"NEW-ES":0.9457,"NEW-FR":0.9457,"NEW-US":1,"HM-US":1,"GV-MX":19.774,"NEW-MX":19.774}

plan=pd.read_excel(r'D:\运营\1数据源\plan.xlsx',sheet_name="Sheet1")
plan["COUNTRY"].replace("CA","GV-CA",inplace=True)
plan["COUNTRY"].replace("US","GV-US",inplace=True)
plan["COUNTRY"].replace("MX","GV-MX",inplace=True)



SailingstarPlan=pd.read_excel(r'D:\运营\2生成过程表\All_Product_Analyzefile_Weeks排序.xlsx',sheet_name="sheet1")


SailingstarPlan.rename(columns = {'站点':'COUNTRY','MSKU':'SKU','销量1':'1','销量2':'2','销量3':'3','销量4':'4','销量5':'5','销量6':'6','销量7':'7','销量8':'8','销量9':'9','销量10':'10'},inplace=True)

SailingstarPlan.rename(columns = {'广告花费1':'广告1','广告花费2':'广告2','广告花费3':'广告3','广告花费4':'广告4','广告花费5':'广告5','广告花费6':'广告6','广告花费7':'广告7','广告花费8':'广告8','广告花费9':'广告9','广告花费10':'广告10'},inplace=True)
SailingstarPlan.rename(columns = {'FBA可售':'Fufillable'},inplace=True)

SailingstarPlan=SailingstarPlan.loc[~SailingstarPlan["COUNTRY"].isnull()]

for countryname99 in SailingstarPlan["COUNTRY"].drop_duplicates().to_list():
    SailingstarPlan.loc[SailingstarPlan["COUNTRY"]==countryname99,'COUNTRY']=CountrDic[countryname99]

SailingstarPlan=SailingstarPlan.loc[~SailingstarPlan["COUNTRY"].isnull()]

plan=pd.concat([plan,SailingstarPlan],ignore_index=True)

CampaignWeek1=pd.read_excel(r'D:\运营\2生成过程表\周Bulk数据Summary.xlsx',sheet_name="SKU-Campaign-WEEK")
CampaignWeek1=CampaignWeek1[CampaignWeek1["周数"]==1]
print(CampaignWeek1.columns)
CampaignWeek1CampaignTotalCount=CampaignWeek1.groupby(["Country","SKU","Campaign Status"],as_index=False)["Campaign"].agg("count").reindex()#groupby(["Country","Campaign","周数"],as_index=False)

print(CampaignWeek1CampaignTotalCount.columns)
